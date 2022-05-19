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
    
    public int GetNormativesID()
    {
		if(TFlex.DOCs.Model.ReferenceCatalog.FindReference("Нормативы ТОиР") == null)
		{
		    return ТекущийОбъект["ID"];
		}
		else
		{
			return -1;
		}
    }

    public int GetReglamentsID()
    {
		if(TFlex.DOCs.Model.ReferenceCatalog.FindReference("Регламенты ТОиР") == null)
		{
		    return ТекущийОбъект["ID"];
		}
		else
		{
			return -1;
		}
    }
}

