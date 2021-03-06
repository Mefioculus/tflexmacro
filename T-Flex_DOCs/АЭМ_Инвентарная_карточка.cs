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
    
    public void ОткрытьАрхивнуюКарточку15()
    {
    	var документ = ТекущийОбъект.СвязанныйОбъект["Связанный объект"];
    	var карточка = документ.СвязанныйОбъект["d708e1b4-2a1a-499c-aaaf-be5828e6377e"];
    	if (карточка == null)
    	{
    		карточка = СоздатьОбъект("Инвентарная книга", "Карточка учёта конструкторских документов");
    		карточка.СвязанныйОбъект["Документ"] = документ;
    		карточка["Наименование документа"] = документ["Наименование"];
    		карточка["Обозначение документа"] = документ["Обозначение"];
    	}
    	
    	ПоказатьДиалогСвойств(карточка);
    }
    
public void ОткрытьАрхивнуюКарточку()
    {
        var документ = ТекущийОбъект.СвязанныйОбъект["Связанный объект"];
        var карточка = НайтиОбъект("Инвентарная книга", "Guid логического объекта", документ["Guid логического объекта"]);
        if (карточка == null)
        {
            карточка = СоздатьОбъект("Инвентарная книга", "Карточка учёта конструкторских документов");
            карточка["Guid логического объекта"] = документ["Guid логического объекта"];
            карточка["Наименование документа"] = документ["Наименование"];
            карточка["Обозначение документа"] = документ["Обозначение"];
        }
        
        ПоказатьДиалогСвойств(карточка);
    }

}
