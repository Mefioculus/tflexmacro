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
    Объект ном = ТекущийОбъект.Владелец;
        if (ном == null)
            return;
    	ТекущийОбъект["Наименование"] = ном["Наименование"] + "_" + "(ЗАК)";
    	ТекущийОбъект["Обозначение"] = ном["Обозначение"];
    	Объекты файлы = ном.СвязанныеОбъекты["[Связанный объект].[Документы]->[Файлы]"];
    	Объекты поставщики = ном.СвязанныеОбъекты["Поставщики"];
    	 
        if (файлы.Count > 0)
        {
            Объект файл = файлы.FirstOrDefault();
            ТекущийОбъект.СвязанныйОбъект["9bc695f9-fa6f-4bfd-bc6f-9229b75938da"] = файл;
         }
    	
    	if (поставщики.Count > 0)
    		{
    		foreach (var поставщик in поставщики)
    			ТекущийОбъект.Подключить("Поставщики", поставщик);
    		}
    }
}
