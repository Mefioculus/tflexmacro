using System.Collections.Generic;
using System.Linq;
using TFlex.DOCs.Model.Macros;
using TFlex.DOCs.Model.Macros.ObjectModel;
using TFlex.DOCs.Model.References.Nomenclature;

public class Macro : MacroProvider
{
    private List<string> _списокОбозначений;

    public Macro(MacroContext context)
        : base(context)
    {
    }

    public void ИспользоватьОбозначение()
    {
        if (ТекущееПодключение == null || ТекущееПодключение.РодительскийОбъект == null)
            return;

        _списокОбозначений = new List<string>();
        ПолучитьСписокОбозначений();
        if (!_списокОбозначений.Any())
            return;

        ДиалогВвода диалог = СоздатьДиалогВвода("Обозначения");
        диалог.ДобавитьВыборИзСписка("Список обозначений", _списокОбозначений.First(), true, _списокОбозначений.ToArray());
        if (!диалог.Показать())
            return;

        string выбранноеОбозначение = диалог["Список обозначений"];
        if (string.IsNullOrEmpty(выбранноеОбозначение))
            return;

        var объектНоменклатуры = Context.ReferenceObject as NomenclatureObject;
        if (объектНоменклатуры != null && объектНоменклатуры.Class.LinkedClass != null)
        {
            string suffix = объектНоменклатуры.Class.LinkedClass.Attributes.GetValue<string>("Suffix");
            if (!string.IsNullOrEmpty(suffix))
                выбранноеОбозначение += suffix;
        }

        ТекущийОбъект["Обозначение"] = выбранноеОбозначение;
    }

    private void ПолучитьСписокОбозначений()
    {
        ЗаполнитьОбозначение(ТекущееПодключение.РодительскийОбъект);
        РекурсивноПолучитьСписокОбозначений(ТекущееПодключение.РодительскийОбъект);
    }

    private void РекурсивноПолучитьСписокОбозначений(Объект объект)
    {
        foreach (var родитель in объект.РодительскиеОбъекты)
        {
            ЗаполнитьОбозначение(родитель);
            РекурсивноПолучитьСписокОбозначений(родитель);
        }
    }

    private void ЗаполнитьОбозначение(Объект объект)
    {
        string обозначение = объект["Обозначение"];
        if (!string.IsNullOrEmpty(обозначение))
        {
            if (!_списокОбозначений.Contains(обозначение))
                _списокОбозначений.Add(обозначение);
        }
    }
}

