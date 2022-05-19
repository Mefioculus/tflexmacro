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
    	// Номенклатура и изделия из Списка номенклатуры
    	ОбменДанными.ИмпортироватьОбъекты("51a17213-ff56-4cef-b35a-b4d1fe72d9e9", ВыбранныеОбъекты);
    	
    	// Изменения в конструкторском составе KAT_IZV
    	//ОбменДанными.ИмпортироватьОбъекты("3fa16424-c980-413e-968e-d1bd9c0d71b5", ВыбранныеОбъекты);
    	
    	// Изменения в конструкторском составе (печатных плат) KAT_IZVP
    	//ОбменДанными.ИмпортироватьОбъекты("6dd7d16b-5ea3-43aa-8a12-01b3a29af8ec", ВыбранныеОбъекты);
    	
    	ВыполнитьМакрос("2c2d4f88-428c-4819-8d91-ead0b2a80494", "СоздатьКонструкторскиеИзменения");
    }
}
