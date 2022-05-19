using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using TFlex.DOCs.Model.Macros;
using TFlex.DOCs.Model.Macros.ObjectModel;
using TFlex.DOCs.Model.Desktop;
using TFlex.DOCs.Model.References;
using TFlex.DOCs.Model.References.Files;
using TFlex.DOCs.Model.Classes;
using TFlex.DOCs.Model;

public class MacroSpe : MacroProvider
{
    private StringBuilder _errorStringBuilder = new StringBuilder();
    public MacroSpe(MacroContext context)
        : base(context)
    {
    }

    public override void Run()
    {
        ДиалогВвода диалог = СоздатьДиалогВвода("Выбор параметра");
        диалог.ДобавитьСтроковое("Параметр", "Обозначение");

        if (!диалог.Показать())
            return;

        string наименованиеПараметра = диалог["Параметр"];

        var parameterInfo = Context.Reference.ParameterGroup.OneToOneParameters.FindByName(наименованиеПараметра);
        if (parameterInfo == null)
            Ошибка("У выбранного справочника нет параметра: " + наименованиеПараметра);

        HashSet<ReferenceObject> saveList = new HashSet<ReferenceObject>();

        foreach (ReferenceObject currentReferenceObject in Context.GetSelectedObjects())
        {
            var parameterValue = currentReferenceObject[parameterInfo].GetString();

            //Если не начинается с 8А или не содержит . то идем дальше
            if (!parameterValue.StartsWith("8А") || parameterValue.Contains("."))
            {
                AddErrorString(String.Format("Значение {0} некорректное", parameterValue));
                continue;
            }

            // Изменяем значение
            parameterValue = parameterValue.Insert(3, ".").Insert(7, ".");
            try
            {
                currentReferenceObject.BeginChanges(false);
                currentReferenceObject[parameterInfo].Value = parameterValue;
                saveList.Add(currentReferenceObject);
            }
            catch
            {
                AddErrorString(String.Format("Ошибка при изменении объекта {0}", parameterValue));
            }
        }

        //Пакетное сохранение объекта
        if (saveList.Count > 0)
            Reference.EndChanges(saveList);

        ShowAvailableErrorsMessage();
    }

    private void ShowAvailableErrorsMessage()
    {
        if (_errorStringBuilder.Length > 0)
            Сообщение("Предупреждение", _errorStringBuilder.ToString());
    }

    private void AddErrorString(string errorString)
    {
        _errorStringBuilder.Append(errorString + Environment.NewLine);
    }

}
