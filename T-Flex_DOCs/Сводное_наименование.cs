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
    
    public void osn()
    	{
    	
Объект объект = ТекущийОбъект;
/*
if (объект["Обозначение"].ToString()=="*")
объект["Сводное наименование"] = объект["Наименование"].ToString() +" "+ объект["Обозначение"].ToString();
else
*/
	объект["Сводное наименование"] = объект["Наименование"].ToString() +" "+ объект["Обозначение"].ToString();
 

        }
    
}
