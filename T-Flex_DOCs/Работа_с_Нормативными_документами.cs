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
    
    public void СоздатьИнвентарнуюКарточку()
    {
    	var документ = ТекущийОбъект;
    	var карточка = документ.СвязанныйОбъект["Запись инвентарной книги"];
    	if (карточка != null)
    		return;
    	
    	карточка = СоздатьОбъект("Инвентарная книга", "Карточка учета нормативных документов");
    	карточка["Наименование документа"] = документ["Наименование"];
    	карточка["Обозначение документа"] = документ["Обозначение"];
    	
    	карточка.Сохранить();
    	
    	документ.СвязанныйОбъект["Запись инвентарной книги"] = карточка;
    	
    	ПоказатьДиалогСвойств(карточка);
    }
}
