using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using TFlex.DOCs.Model.Macros;
using TFlex.DOCs.Model.Macros.ObjectModel;
using TFlex.DOCs.Model.References;
using TFlex.DOCs.Model.References.Users;
using TFlex.DOCs.References.TechnologicalRouteElements;
using TFlex.DOCs.UI.Objects.Managers;

public class Macro : MacroProvider
{
    public Macro(MacroContext context)
        : base(context)
    {
    }

    public override void Run()
    {
    }
    
    public void ПолучениеДанныхИзАналога(ReferenceObject selectedObject)
    {
    	if(selectedObject == null)
    		return;
    	
    	Объект аналог = Объект.CreateInstance(selectedObject, Context);
    	foreach(var цехопереход in аналог.СвязанныеОбъекты["Цехопереходы"])
    	{
    		var переход = ТекущийОбъект.СоздатьОбъектСписка("Цехопереходы", "Цехопереход");
    		переход["Длительность"] = цехопереход["Длительность"];
    		переход["Заключительное время"] = цехопереход["Заключительное время"];
    		переход["Подготовительное время"] = цехопереход["Подготовительное время"];
    		переход["Наименование"] = цехопереход["Наименование"];
    		переход.СвязанныйОбъект["Подразделение"] = цехопереход.СвязанныйОбъект["Подразделение"];
    		переход.Сохранить();
    	}
    	
    	ТекущийОбъект["Маршрут"] = аналог["Маршрут"];
    	ТекущийОбъект["Разделитель"] = аналог["Разделитель"];

    	ТекущийОбъект["Квота"] = аналог["Квота"];
    	ТекущийОбъект["Объем"] = аналог["Объем"];
    }

    public void ОбновлениеДиалога()
    {
    	if(ИзмененныйПараметр == "Маршрут")
    	{
    		var uiContext = Context as UIMacroContext;
    		if(uiContext != null)
    		{	
    			var dialog = uiContext.Dialogs.FirstOrDefault();
    			if(dialog != null)
    				dialog.RefreshCollectionControls();
    		}
        }
    }
    
    public void ЗавершениеИзмененияСвязиВЦехопереходе()
    {
       if (Context.ChangedLink.LinkGroup.Guid == TechnologicalRouteElementsReferenceObject.RelationKeys.ProductionUnit)
       {
           var productionUnit = ((ObjectLinkChangedEventArgsBase)Context.ModelChangedArgs)?.AddedObject as ProductionUnit;
           var technologicalRouteElement = (TechnologicalRouteElementsReferenceObject)CurrentObject;
           
           technologicalRouteElement.Name.Value =
           	productionUnit != null ?
           	productionUnit.Code.Value :
           	String.Empty;
       }
    }
}
