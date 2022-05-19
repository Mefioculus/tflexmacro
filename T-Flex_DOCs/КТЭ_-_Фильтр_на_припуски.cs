using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using TFlex.DOCs.Model;
using TFlex.DOCs.Model.Configuration.Variables;
using TFlex.DOCs.Model.Macros;
using TFlex.DOCs.Model.Macros.ObjectModel;
using TFlex.DOCs.Model.References;
using TFlex.DOCs.Model.Search;
using TFlex.Model.Technology.Macros;
using TFlex.Model.Technology.Macros.ObjectModel;

public class Macro : MacroProvider
{
    public Macro(MacroContext context)
        : base(context)
    {
    }

    public override void Run()
    {
    
    }    
    
    public double ПолучитьЗначениеПрипуска(string xmlФильтр, string гуидСправочника, Параметр[] параметрыПерехода, int класс, int квалитет, string путькПараметру)
    {
        double значениеПрипуска = 0;
        if (!String.IsNullOrEmpty(xmlФильтр))
        {
            Filter фильтр = Filter.Deserialize(xmlФильтр, Context.Connection);
            // Устанавливаем значения переменных фильтра
            foreach (Variable переменная in фильтр.Variables)
            {
                switch (переменная.Name)
                {
                    case "Класс":
                        переменная.Value = класс;
                        break;

                    case "Квалитет":
                        переменная.Value = квалитет;
                        break;

                    default:
                        // Для остальных переменных
                        Параметр искомыйПараметр = параметрыПерехода.FirstOrDefault(параметр => параметр.Наименование == переменная.Name);
                        переменная.Value = искомыйПараметр != null ? искомыйПараметр.Номинал : 0;
                        break;
                }
            }


            ReferenceInfo info = Context.Connection.ReferenceCatalog.Find(гуидСправочника);
            Reference reference = info.CreateReference();

            // ищем припуск по фильтру
            List<ReferenceObject> списокПрипусков = reference.Find(фильтр);
            if (списокПрипусков.Count > 0)
            {
                string[] списокГуидов = путькПараметру.Split('.');
                string гуидИскомогоПараметра = списокГуидов.Length > 0 ? списокГуидов[списокГуидов.Length - 1].Trim('[', ']') : String.Empty;
                if (String.IsNullOrEmpty(гуидИскомогоПараметра)) // искомый параметр не указан, берем параметр справочника, отображаемый по умолчанию
                {
                    значениеПрипуска = списокПрипусков.Min(объектПрипуска => (double)объектПрипуска.ParameterValues[info.Description.DefaultVisibleParameter.Guid]);
                }
                else
                {
                    значениеПрипуска = списокПрипусков.Min(объектПрипуска => (double)объектПрипуска.ParameterValues[new Guid(гуидИскомогоПараметра)]);
                }
            }
        }
        return значениеПрипуска;
    }
}
