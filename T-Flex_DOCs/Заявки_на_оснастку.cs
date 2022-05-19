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
    
    public void CreateCatalogSTO()
    {
    	Объект заявка = ТекущийОбъект;
    	Объект оснащение = НайтиОбъект("Каталог оснащения", string.Format("[Наименование] = '{0}' И [Обозначение] = '{1}'", заявка["Наименование"], заявка["Обозначение"]));
    	if (оснащение == null)
    		оснащение = СоздатьОбъект("Каталог оснащения", "Приспособления");
    	else
    		оснащение.Изменить();
    	
    	оснащение["Наименование"] = заявка["Наименование"];
    	оснащение["Обозначение"] = заявка["Обозначение"];
    	оснащение.СвязанныйОбъект["Объект номенклатуры"] = заявка.СвязанныйОбъект["Номенклатура оснастки"];
    	Объект изготовитель = заявка.СвязанныйОбъект["Изготовитель"];
    	if (изготовитель != null)
    		оснащение["Производитель"] = изготовитель["Наименование"];
    	
    	оснащение.Сохранить();
    	
    	if (оснащение != null)
    		Сообщение("", "Данные добавлены в каталог.");
    }
}
