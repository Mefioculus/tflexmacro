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
    
    public ButtonValidator Validate()
    {
    	return new ButtonValidator()
    		{
    		Visible = ТекущийОбъект != null,
    		Enable = ТекущийОбъект != null
    		};
    }
    
    public void ЗаполнитьНаименование()
    {
    	Параметр["Наименование"] = ТекущийОбъект.СвязанныйОбъект["Тип статуса"].Параметр["Наименование"];
    	
    	var startDate = (DateTime)Параметр["Дата начала"];
    	if(startDate != DateTime.MinValue)
    		Параметр["Наименование"] += $" с {startDate:d}";
    	
    	var endDate = (DateTime)Параметр["Дата окончания"];
    	if(endDate != DateTime.MinValue)
    		Параметр["Наименование"] += $" по {endDate:d}";
    }
    
    public void Закрыть()
    {
    	Параметр["Дата окончания"] = DateTime.Now;
    }
}
