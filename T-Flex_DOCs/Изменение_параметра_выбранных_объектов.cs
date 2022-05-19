using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using TFlex.DOCs.Model.Macros;
using TFlex.DOCs.Model.Macros.ObjectModel;
using TFlex.DOCs.Model.Desktop;
using TFlex.DOCs.Model.References;
using TFlex.DOCs.Model.References.Files;
using TFlex.DOCs.Model.Classes;
using TFlex.DOCs.Model;
using TFlex.DOCs.Model.Search;
using TFlex.DOCs.Model.Structure;

public class MacroSpe : MacroProvider
{
    public MacroSpe(MacroContext context)
        : base(context)
    {
    }

    public override void Run()
    {
        if (Вопрос("Хотите запустить в режиме отладки?"))
        {
            System.Diagnostics.Debugger.Launch();
            System.Diagnostics.Debugger.Break();
        }
    }

    public void ЗапуститьИзмененияПараметраВыбранныхОбъектов()
    {
        ДиалогВвода диалог = СоздатьДиалогВвода("Выбор параметров");

        диалог.ДобавитьСтроковое("Имя параметра","",false, true);
        диалог.ДобавитьСтроковое("Новое значение параметра");
        if(!диалог.Показать())
            return;

        EditObjectParameters(Context.GetSelectedObjects().ToList(), диалог["Имя параметра"], диалог["Новое значение параметра"]);
    }

    public void ЗаменитьНаЭлектроннуюСтруктуру()
    {

        EditObjectParameters(Context.GetSelectedObjects().ToList(), "Тип Номенклатуры", 6);
    }

    private void EditObjectParameters(List<ReferenceObject> editObjectsColletion, string parameterName, object stringValue)
    {   
        if(!editObjectsColletion.Any())
            Ошибка("Нет объектов");

        var firstObj = editObjectsColletion.FirstOrDefault();

        var parameterInfoCollection = firstObj.Reference.ParameterGroup.OneToOneParameters;
        
        var parameterInfo = parameterInfoCollection.FindByName(parameterName);
        if(parameterInfo == null)
            Ошибка("В справочнике не найден параметр с наименованием: " + parameterName);

        if (parameterInfo.Type.IsNumber)
        {
            if (stringValue is string)
            {
                int value;
                if(!Int32.TryParse((string)stringValue, out value))
                    Ошибка("Новое значение параметра невозможно записать в параметр");
            }
            else if(!(stringValue is int))
            { 
                Ошибка("Новое значение параметра невозможно записать в параметр");
            }
        }
    
        ДиалогОжидания.Показать("Подождите", true);

        List<ReferenceObject> saveObjects = new List<ReferenceObject>();
 
        int current = 0;
        int total = editObjectsColletion.Count;

        foreach (var referenceObject in editObjectsColletion)
        {
            current++;
            if (!ДиалогОжидания.СледующийШаг(string.Format("Подождите идет изменение объектов: {0} из {1}", current, total)))
                return;

            if (!referenceObject.CanEdit)
                continue;

            referenceObject.BeginChanges();

            referenceObject[parameterInfo].Value = stringValue;

            saveObjects.Add(referenceObject);
        }

        ДиалогОжидания.СледующийШаг(
            string.Format("Подождите идет сохранение объектов на сервер.{0}Это может занять несколько минут",
                Environment.NewLine));

        Reference.EndChanges(saveObjects);
    }
}
