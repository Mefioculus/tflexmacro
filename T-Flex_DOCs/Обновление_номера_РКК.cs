using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using TFlex.DOCs.Model.Macros;
using TFlex.DOCs.Model.Macros.ObjectModel;
using TFlex.DOCs.Model.References.RecordControlCards;

public class Macro : MacroProvider
{
    public Macro(MacroContext context)
        : base(context)
    {
    }

    public override void Run()
    {
        var rcc = Context.ReferenceObject as RCCReferenceObject;
        if(rcc == null)
        	return;
        rcc.RegenerateRegNumber();
    }
}
