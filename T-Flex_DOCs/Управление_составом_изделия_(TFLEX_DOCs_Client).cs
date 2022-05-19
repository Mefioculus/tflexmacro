using System.Collections.Generic;
using System.Linq;
using TFlex.DOCs.Model.Macros;
using TFlex.DOCs.Model.Macros.ObjectModel;
using TFlex.DOCs.Model.References;
using TFlex.DOCs.Model.References.Nomenclature;
using TFlex.DOCs.Client.ViewModels;
using TFlex.DOCs.Client.ViewModels.References;
using TFlex.DOCs.Client.ViewModels.ObjectProperties;

public class Macro : MacroProvider
{
    private List<string> _списокОбозначений;

    public Macro(MacroContext context)
        : base(context)
    {
    }

    public void ИспользоватьОбозначение()
    {
    	var uiMacroContext = Context as UIMacroContext;
    	
    	var parent = uiMacroContext.HierarchyLink?.ParentObject ;
    	
    	if (parent is null)
  		{
    		var parentViewModel = uiMacroContext.OwnerViewModel;
    		
    		while (parentViewModel != null)
    		{
    			if (parentViewModel is IPanel panel && panel.Object is ReferenceObjectViewModel roViewModel) 
    			{ 
    				parent = roViewModel.ReferenceObject;
    				break;
    			}
    			
    			parentViewModel = parentViewModel.Owner as LayoutViewModel;
    		}    		
        }    

    	if (parent is null)
    		return;
    	
        _списокОбозначений = new List<string>();
        ПолучитьСписокОбозначений(parent);
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

    private void ПолучитьСписокОбозначений(ReferenceObject parent)
    {
        ЗаполнитьОбозначение(parent);
        РекурсивноПолучитьСписокОбозначений(parent);
    }
    

    private void РекурсивноПолучитьСписокОбозначений(ReferenceObject объект)
    {
        foreach (var родитель in объект.Parents)
        {
            ЗаполнитьОбозначение(родитель);
            РекурсивноПолучитьСписокОбозначений(родитель);
        }
    }

    private void ЗаполнитьОбозначение(ReferenceObject объект)
    {
        string обозначение = объект.GetObjectValue("Обозначение")?.Value?.ToString();
        if (!string.IsNullOrEmpty(обозначение))
        {
            if (!_списокОбозначений.Contains(обозначение))
                _списокОбозначений.Add(обозначение);
        }
    }
}

