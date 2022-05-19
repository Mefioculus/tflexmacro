using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using TFlex.DOCs.Model.Macros;
using TFlex.DOCs.Model.Macros.ObjectModel;
using TFlex.DOCs.Model.References;
using TFlex.DOCs.Model.References.Nomenclature;

public class Macro8 : MacroProvider
{
    public Macro8(MacroContext context)
        : base(context)
    {
    }

    public void НазначитьНомер()
    {
        int номер = (int)ГлобальныйПараметр["Номер выгрузки в FoxPro"] + 1;
        ТекущийОбъект["Номер выгрузки"] = номер;
        ГлобальныйПараметр["Номер выгрузки в FoxPro"] = номер;
    }

    public void Экспорт()
    {
        int номер = ТекущийОбъект["Номер выгрузки"];

        ОбменДанными.Экспортировать(
            "Выгрузка в базу данных спецификации",
            "SPEC_OUT",
            String.Format("[Номер журнала выгрузки] = '{0}'", номер),
            показыватьДиалог: false);

        Сообщение("Обмен данными", "Выгрузка завершена");
    }

    public void ВыгрузитьНоменклатуру()
    {
        var объекты = ВыбранныеОбъекты;

        if (объекты.Count == 0)
        {
            Сообщение("Обмен данными", "Нет выделенных объектов");
            return;
        }

        var всеОбъекты = ПолучитьВсеДочерние(объекты);
        var журнал = СоздатьОбъект("Журнал выгрузок в FoxPro", "Запись журнала выгрузок");
        журнал.Сохранить();
        int номерЖурнала = журнал["Номер выгрузки"];

        var дополнительныеДанные = new Dictionary<string, dynamic>()
        {
            { "Number", номерЖурнала },
        };

        ОбменДанными.ИмпортироватьОбъекты("Выгрузка структуры в Fox|Pro", всеОбъекты,
            показыватьДиалог: false,
            дополнительныеДанные: дополнительныеДанные);

        ПоказатьДиалогСвойств(журнал);
    }

    private List<object> ПолучитьВсеДочерние(Объекты объекты)
    {
        var objects = объекты.To<ReferenceObject>();
        var result = objects.OfType<object>().ToList();

        foreach (var ro in objects)
            result.AddRange(ro.Children.RecursiveLoadHierarchyLinks());

        return result;
    }

}
