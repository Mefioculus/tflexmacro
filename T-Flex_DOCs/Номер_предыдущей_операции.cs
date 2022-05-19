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
    	int num = (int)ТекущийОбъект["№"];
    	//MessageBox.Show(num.ToString());
    	Объект операция = null;
    	if (num > 1)
    	{
        	операция = ТекущийОбъект.РодительскийОбъект.ДочерниеОбъекты.FirstOrDefault(t => (int)t["№"] == num - 1);
        	ТекущийОбъект["Номер предыдущей операции"] = операция["Номер"];
        }
    }
}
