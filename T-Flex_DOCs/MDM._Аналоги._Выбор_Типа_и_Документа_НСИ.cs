using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using TFlex.DOCs.Model.Macros;
using TFlex.DOCs.Model.Macros.ObjectModel;

public class MacroMDMSelectTypeNSI : MacroProvider
{
    public MacroMDMSelectTypeNSI(MacroContext context)
        : base(context)
    {
        if (Context.Connection.ClientView.HostName == "MOSINS")
            if (Вопрос("Хотите запустить в режиме отладки?"))
            {

                Debugger.Launch();
                Debugger.Break();
            }
    }

    public override void Run()
    {
        Объекты выбранныеОбъекты = ВыполнитьМакрос("ea0d8d3c-395b-48d5-9d42-dab98371522c", "ПолучитьВыбраныеОбъекты", "item13");
        if (выбранныеОбъекты.Count == 0)
            return;

        var группы = выбранныеОбъекты.GroupBy(ob => ob["Тип НСИ"]);
        if (группы.Count() > 1)
            Ошибка("Невозможно выполнить операцию для объектов с разными Типами НСИ. Выберите группу объектов с одинаковым Типом НСИ");

        Объект выбранныйОбъект = выбранныеОбъекты.FirstOrDefault();

        Guid гуидТипаНСИ = выбранныйОбъект["Тип НСИ"];
        if (гуидТипаНСИ == Guid.Empty)
            гуидТипаНСИ = ПолучитьТипНСИ(выбранныйОбъект.Справочник.Guid);

        if (гуидТипаНСИ == Guid.Empty)
            return;

        string фильтрДокументовНСИ = String.Format("[Параметры документа НСИ]->[Тип НСИ]->[Guid] = '{0}'", гуидТипаНСИ);
        //Этого не может быть,но на всякий случай проверка
        var ПоискДокументовНСИ = НайтиОбъекты("Документы НСИ", фильтрДокументовНСИ);
        if (ПоискДокументовНСИ.Count == 0)
            Ошибка("Не найден документ НСИ для выбранного типа НСИ");

        Guid гуидДокументаНСИ = ПолучитьДокументНСИ(фильтрДокументовНСИ);
        if (гуидДокументаНСИ == Guid.Empty)
            return;

        if (!Вопрос("Параметры: Документ НСИ, Тип НСИ будут перезаписаны." + Environment.NewLine + "Продолжить выполнение?"))
            return;

        StringBuilder errors = new StringBuilder();
        foreach (var объект in выбранныеОбъекты)
        {
            try
            {
                объект.Изменить();
                объект["Тип НСИ"] = гуидТипаНСИ;
                объект["Документ НСИ"] = гуидДокументаНСИ;
                объект.Сохранить();
            }
            catch (Exception e)
            {
                errors.Append(String.Format("Ошибка при обработке объекта: {0}{1}{2}{1}",
                    объект, Environment.NewLine, e.Message));
            }
        }

        if (errors.Length > 0)
            Сообщение("Предупреждение", errors.ToString());
    }

    private Guid ПолучитьДокументНСИ(string фильтрДокументаНСИ)
    {
        var диалог = СоздатьДиалогВыбораОбъектов("Документы НСИ");
        диалог.Фильтр = фильтрДокументаНСИ;
        диалог.ПоказатьПанельКнопок = false;
        диалог.МножественныйВыбор = false;
        if (диалог.Показать())
            return диалог.ФокусированныйОбъект["Guid"];
        return Guid.Empty;
    }

    /// <summary>
    /// Находит тип НСИ для указанно справочника показывает диалог с выбором типа НСИ
    /// </summary>
    /// <param name="referenceGuid"></param>
    /// <returns>Возвращает гуид типа НСИ+</returns>
    public Guid ПолучитьТипНСИ(Guid referenceGuid)
    {
        string filterString = String.Format("[Использование в пользовательских системах]->[Справочник-аналог] = '{0}'", referenceGuid);

        var найденыеТипы = НайтиОбъекты("00bf7ef0-6080-4edd-a548-95b44df465c4", filterString);
        if (найденыеТипы.Count == 0)
            Сообщение("Предупреждение", String.Format("Нет связанных типов НСИ, выбор будет осуществляться из всех доступных"));

        var диалог = СоздатьДиалогВыбораОбъектов("00bf7ef0-6080-4edd-a548-95b44df465c4");//Типы НСИ
        //Если нет подходящих типов, то отображаем все
        if (найденыеТипы.Count != 0)
            диалог.Фильтр = filterString;
        диалог.ПоказатьПанельКнопок = false;
        диалог.МножественныйВыбор = false;
        if (диалог.Показать())
            return диалог.ФокусированныйОбъект["Guid"];

        return Guid.Empty;
    }
}
