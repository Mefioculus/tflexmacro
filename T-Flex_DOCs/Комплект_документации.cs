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
    
    public void FillLinkTP()
    {
    	var комплект = СоздатьОбъект("Комплекты документов", "Технологический комплект");
    	комплект.СвязанныйОбъект["Элемент технологии"] = ТекущийОбъект;
    	
    	ПоказатьДиалогСвойств(комплект);
    }
}
