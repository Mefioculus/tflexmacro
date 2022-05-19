using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using TFlex.DOCs.Model.Macros;
using TFlex.DOCs.Model.Macros.ObjectModel;

public class Macro : MacroProvider
{
    public Macro(MacroContext context)
        : base(context)
    {
    }

    public override void Run()
    {
    
    }
    
    public void СоздатьВходимость()
    {
    	//System.Diagnostics.Debugger.Launch();
    	//System.Diagnostics.Debugger.Break();
    	
    	string шифрБезТочек = ТекущийОбъект["Обозначение"].ToString().Replace(".", "");
    	var аналоги = НайтиОбъекты("KAT_VHOD", string.Format("[IZD_ST] = '{0}'", шифрБезТочек));
    	var документ = ТекущийОбъект.СвязанныйОбъект["Связанный объект"];
    	if (документ == null)
    		return;
    	
    	var технологическиеДокументы = документ.СвязанныеОбъекты["~[Связанные документы]"];
    	foreach (var технологическийДокумент in технологическиеДокументы)
    	{
    		технологическийДокумент.Изменить();
    		var входящиеИзделия = технологическийДокумент.СвязанныеОбъекты["Входящие изделия"];
    		foreach (var изделие in входящиеИзделия)
    		{
    			bool flag = false;
    			foreach (var аналог in аналоги)
    			{
    				if (изделие["Наименование"].ToString().Contains(аналог["SHIFR"].ToString().Replace(".", "")))
    				{
    					flag = true;
    					break;
    				}
    			}
    			if (!flag)
    				continue;
    			
    			изделие.Изменить();
        		изделие["Есть ГПП"] = true;
        		
        		изделие.Сохранить();
    		}
    		
    		технологическийДокумент.Сохранить();
    	}
    }
}
