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
    
    //ВыполнитьМакрос("Фильтрация", "Извещения", ТекущийОбъект);
    
    /* Учёт изменений в инвентарной книге */
    public List<int> Извещения (Объект инвентарная_книга)
    {
    	//Список, в который будут записываться ID извещений, относящихся к данной инвентарной книге
    	List<int> idList = new List<int>();
    	idList.Add(0);
    	
    	//Не учитывать фильтр в других справочниках
    	if (инвентарная_книга == null)
    		return idList;
    	
    	//Получение связанного объекта номенклатуры
    	Объект номенклатура = инвентарная_книга.СвязанныйОбъект["[Документ]->[Объект номенклатуры]"];
    	if (номенклатура == null)
    		return idList;
    	string nomenclatureID = номенклатура.Параметр["ID"];
    	
    	//Поиск всех связанных извещений (можно добавить ограничение на то, что извещения проведены)
    	Объекты извещения = НайтиОбъекты("Извещения об изменениях", "[Объекты в извещении]->[ID] = '" + nomenclatureID + "'");
    	
    	//Добавление найденных (+ дополнительные к ним) извещений в список
    	int changeID;
    	foreach (var извещение in извещения)
    	{
    		changeID = извещение.Параметр["ID"];
    		idList.Add(changeID);
    		Объекты дополнительные = извещение.ДочерниеОбъекты;
    		foreach (var дополнительный in дополнительные)
    		{
    			changeID = дополнительный.Параметр["ID"];
    			idList.Add(changeID);
    		}
    	}
    	//Возвращение результата
    	return idList;
    }
}
