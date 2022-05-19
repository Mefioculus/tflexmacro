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
	private static string _имяРодителя;
		
    public Macro(MacroContext context)
        : base(context)
    {
    }

    public override void Run()
    {    
    }
    
    public ICollection<Объект> ПолучитьИзделияЗаготовки()
    {
    	_имяРодителя = ТекущийОбъект["Наименование"];
    	return ТекущийОбъект.ДочерниеОбъекты.Where(t => (bool)t["[Подключения].[Изделие-заготовка]"]).ToArray();
    }
    
    public string ПолучитьИмяИзделеяЗаготовки()
    {
    	string имяЗаготовки = ТекущийОбъект["Наименование"];
        return string.Format("{0} (заготовка для {1})", имяЗаготовки, _имяРодителя);
    }
}
