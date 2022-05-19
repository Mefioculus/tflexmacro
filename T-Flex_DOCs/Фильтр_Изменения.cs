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
//#if DEBUG
//        System.Diagnostics.Debugger.Launch();
//        System.Diagnostics.Debugger.Break();
//#endif
    }

    public override void Run()
    {
    
    }
    
    public List<int> ВыдатьСписокИзменений(Объект эси)
    {
        List<int> списокИдентификаторов = new List<int>{};
       
        string фильтр = "[Объект ЭСИ]->[Объект] = '" + эси["Объект"] + "'";
        Объекты вводящиеИзменения = НайтиОбъекты("Изменения", фильтр);
        if(вводящиеИзменения == null)
        	return null;
        foreach(Объект вводящееИзменение in вводящиеИзменения)
        {
        	 списокИдентификаторов.Add(вводящееИзменение["ID"]);
	         Объекты исходныеРевизии = вводящееИзменение.СвязанныеОбъекты["Исходные ревизии"];
	         if(исходныеРевизии != null)
	         {
	         	foreach(Объект исходнаяРевизия in исходныеРевизии)
    	         	ВыдатьСписокИзмененийРевизии(исходнаяРевизия.СвязанныйОбъект["Ревизия"], ref списокИдентификаторов);
	         }

        }
        	
        return списокИдентификаторов;
    }
    
    private void ВыдатьСписокИзмененийРевизии(Объект эси, ref List<int> списокИдентификаторов)
    {
        string фильтр = "[Объект ЭСИ]->[Объект] = '" + эси["Объект"] + "'";
        Объекты вводящиеИзменения = НайтиОбъекты("Изменения", фильтр);
        if(вводящиеИзменения == null)
            return;
        foreach(Объект вводящееИзменение in вводящиеИзменения)
        {
             списокИдентификаторов.Add(вводящееИзменение["ID"]);
             Объекты исходныеРевизии = вводящееИзменение.СвязанныеОбъекты["Исходные ревизии"];
             if(исходныеРевизии != null)
             {
                 foreach(Объект исходнаяРевизия in исходныеРевизии)
                     ВыдатьСписокИзмененийРевизии(исходнаяРевизия.СвязанныйОбъект["Ревизия"], ref списокИдентификаторов);
             }
        }
    }
    
    
}
