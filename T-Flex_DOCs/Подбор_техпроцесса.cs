using System;
using System.Linq;
using TFlex.DOCs.Model.Macros;
using TFlex.DOCs.Model.Macros.ObjectModel;
using TFlex.DOCs.Model.References;

public class Macro : MacroProvider
{
    public Macro(MacroContext context)
        : base(context)
    {
    }
    
    public override void Run()
    {
    	var uiContext = Context as TFlex.DOCs.Client.ViewModels.UIMacroContext;
        if (uiContext == null)
        	return;
        
        var dse = uiContext.Reference.LinkInfo.Master as ReferenceObject;
    	
    	ДиалогВыбораОбъектовИзСправочников диалог = СоздатьДиалогВыбораОбъектовИзСправочников();
    	
    	if (диалог.Показать())
        {
    		if (!диалог.ВыбранныеОбъекты.Any())
    			return;
    			
    		var ТП = диалог.ВыбранныеОбъекты.First();
    		Объект ДСЕ = Объект.CreateInstance(dse, Context);
    		ТП.Изменить();
    		ТП.СвязанныйОбъект["Изготавливаемая ДСЕ"] = ДСЕ;
    		
    		ТП.Сохранить();
        }
    }
}
