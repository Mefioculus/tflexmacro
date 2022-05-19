/*Ссылки
TFlex.Model.Technology.dll
TFlex.Reporting.Technology.dll*/

using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using TFlex.DOCs.Model.Macros;
using TFlex.DOCs.Model.Macros.ObjectModel;

using TFlex.DOCs.Model.References;
using TFlex.DOCs.Model.References.Nomenclature;
using TFlex.Model.Technology.Macros.ObjectModel;
using TFlex.DOCs.Model.Parameters;
using GetReferenceObjects;


public class Macro : MacroProvider
{
    public Macro(MacroContext context)
        : base(context)
    {
    }
 
    
    public override void Run()
    {
    	ОчиститьЗамечания();
    	Dictionary<string, double> Технологические_переделы = Сформировать_словарь();
    	
    	//Функция фильтрации собираемых объектов
    	Func<ReferenceObject, bool> filter = (referenceObject) => 
        {
        	bool isPackage = referenceObject.GetObjectValue("Является покупным изделием").ToBoolean();
        	return (bool)(referenceObject.Class.IsInherit("Материальный объект") && !isPackage);
        };
    	
    	//Получение состава изделия 
		List<ReferenceObjectWithData> productComposition = ObjectsFinder.FindObjects(Context.ReferenceObject, filter);
		
    	//Обработка объектов в технологии
    	Technology (productComposition, ref Технологические_переделы);
    	
    	//Подготовка данных для вывода на диаграмму
    	double total = 0;
    	string[,] Вывод_на_диаграмму = Преобразовать_словарь(Технологические_переделы, ref total);
    	
    	ВывестиЗамечания();
    	ОчиститьЗамечания();
    	
    	//Заголовок
    	string изделие = Параметр["Обозначение"] + " " + Параметр["Наименование"];
    	string заголовокДиаграммы = "Экспресс-анализ продолжительности изготовления\r\nизделия '" + изделие
    		+ "' по технологическим переделам.";
    	
    	//Создание диаграммы
    	ВыполнитьМакрос("Графика", "ПоказатьДиаграмму", заголовокДиаграммы, Вывод_на_диаграмму, total);
    }
    
    /* Больше не используется. Заменено на библиотеку
    private void Recursion (ReferenceObject referenceObject, List<ReferenceObject> objectList, List<int> countList, int nkol)
    {
    	//Является ли изделие покупным
    	bool isPackage = referenceObject.GetObjectValue("Является покупным изделием").ToBoolean();
    	
    	//Рекурсию выполняем только для изготавливаемых изделий
    	if (referenceObject.Class.IsInherit("Материальный объект") && !isPackage)
    	{
    		//Проверяем, добавлен ли уже рассматриваемый объект в список
			int index = objectList.IndexOf(referenceObject);
			if (index != -1)	//если добавлен, то изменяем количество
			{
				countList[index] += nkol;
			}
			else				//если ещё не добавлен, то добавляем
			{
				objectList.Add(referenceObject);
				countList.Add(nkol);
			}
			
			//Вспомогательная переменная для расчёта количества
			int kol;
			double dkol;
			
			//ID рассматриваемого, т.е. родительского объекта
			//int parentId = referenceObject.SystemFields.Id;
			ReferenceObjectCollection childrenCollection = referenceObject.Children;
    		foreach (ReferenceObject children in childrenCollection)
    		{
    			//ID нового, т.е. дочернего объекта
    			//int childId = children.SystemFields.Id;
    			
    			//Получение параметра подключения (количество)
    			NomenclatureHierarchyLink hLink = children.GetParentLink(referenceObject) as NomenclatureHierarchyLink;
    			if (hLink != null)
    			{	
    				dkol = hLink.Amount;
    				kol = (int)dkol;
    				
    				//Изменение накопленного значения и углубление по дереву
	    			nkol = nkol*kol;
	    			Recursion (children, objectList, countList, nkol);
	    			nkol = nkol/kol;
    			}
    			else
    				Ошибка("Не удалось определить количество между " + referenceObject.ToString()
    				       + " и " + children.ToString() + ".\r\nРасчёт будет остановлен.");
    			
    		}
		}
    }
    */
	
    private Guid guidLinkTPNomenclature = new Guid("e1e8fa07-6598-444d-8f57-3cfd1a3f4360");
    
    //Получение списка технологических процессов изготовления изделия	
	private void Technology (List<ReferenceObjectWithData> productComposition, 
                             ref Dictionary<string, double> Технологические_переделы)
    {
    	//Вспомогательные переменные
    	ReferenceObject referenceObject;	//найденный объект в структуре
    	int count;							//количество таких объектов в изделии
    	List<ReferenceObject> technologyList = new List<ReferenceObject>();	//список технологических процессов, по которым изготавливается объект
    	ReferenceObject technologyObject;	//используемый в расчёте техпроцесс
    	int technologyId;					//ID этого техпроцесса
    	
    	
    	//Ищем технологию изготовления для каждого из найденных объектов структуры изделия
    	foreach (ReferenceObjectWithData element in productComposition)
    	{
    		//Получение информации об объекте и его количестве в составе изделия
    		count = element.Count;
    		referenceObject = element.TargetObject;
    		
    		//Получение списка техпроцессов
    		technologyList = referenceObject.GetObjects(guidLinkTPNomenclature);
    		if (technologyList.Count > 0)
    		{
    			//Поиск среди списка единичного либо структурированного техпроцесса (+ порождённые от них)
    			technologyObject = technologyList.FirstOrDefault(tp => 
    			     tp.Class.IsInherit("Единичный технологический процесс") || tp.Class.Name == "Технологический процесс 2010"
    			     || tp.Class.IsInherit("Структурированный техпроцесс") || tp.Class.IsInherit("Технологический процесс"));
    			if (technologyObject != null)
    			{
    				//Приведение типов
    				technologyId = technologyObject.SystemFields.Id;
    				Объект тп = НайтиОбъект("Технологические процессы", "[ID] = '" + technologyId.ToString() + "'");
    				ТехнологическийПроцесс техпроцесс = (ТехнологическийПроцесс)тп;
    				
    				//Обработка операций техпроцесса с учётом количества изготавливаемых деталей
    				Обработка_списка_операций(техпроцесс, count, ref Технологические_переделы);
    			}
    		}
    	}
    }
    
    
    //Используется АПИ Технологии, поскольку техпроцесс может быть как 2010, так и 2012
    private void Обработка_списка_операций(ТехнологическийПроцесс техпроцесс, int количество, 
                                           ref Dictionary<string, double> Технологические_переделы)
    {
    	//Вспомогательные переменные
    	string код;
    	double Тшт;
    	double Тпз;
    	double Тшк;
    	string tpinfo = СводноеНаименованиеТехпроцесса(техпроцесс);
    	
    	//Объём в запуске
    	int объёмВЗапуске = техпроцесс.Параметр["Объем в запуске"];
    	if (объёмВЗапуске < 1)
    	{
    		ДобавитьЗамечания("В техпроцессе '" + tpinfo + "' указан недопустимое значение объёма в запуске: "
    		                  + объёмВЗапуске.ToString() + ". В расчёте будет принят объём в запуске равным 1.");
    		объёмВЗапуске = 1;
    	}
    	
    	//Получение списка операций к указанному техпроцессу
    	Операция[] операцииТехпроцесса = техпроцесс.Операции;
    	foreach (Операция операция in операцииТехпроцесса)
    	{
    		код = операция.Код;
    		string operinfo = СводноеНаименованиеОперации(операция);
    		
    		if (код.Length == 4)
    		{
    			код = код.Substring(0, 2);
    			if (!Технологические_переделы.ContainsKey(код))
    			{
    				ДобавитьЗамечания("В операции '" + operinfo + "' техпроцесса '" + tpinfo
    			                + "' указан код, не соответствующий Общероссийскому классификатору операций. "
    			                + "Параметры операции перенесены в раздел \"Без указания\".");
    				код = "00";
    			}
    		}
    		else
    		{
    			ДобавитьЗамечания("В операции '" + operinfo + "' техпроцесса '" + tpinfo
    			                + "' не указан код. Параметры операции перенесены в раздел \"Без указания\".");
    			код = "00";
    		}
    		
    		//Расчёт времени, затрачиваемого на данный вид операции
    		Тшт = операция.Тшт;
    		Тпз = операция.Тпз;
    		Тшк = Тшт + Тпз/объёмВЗапуске;
    		//MessageBox.Show(operinfo + ": " + (Тшк*количество).ToString());
    		Технологические_переделы[код] += Тшк*количество;
    	}
    }
    
    
    //Вывод содержания и значения пар параметров словаря
    private void Показать_содержание_словаря(string title, Dictionary<string, double> Output)
    {
    	string myFormat = "\r\n{0} = {1}";
    	string result = title + "\r\nПараметр		Значение";
    	foreach (KeyValuePair<string, double> kvp in Output)
    	{
    		result += string.Format(myFormat, kvp.Key, kvp.Value);
    	}
    	MessageBox.Show(result);
    }
    
    
    private Dictionary<string, double> Сформировать_словарь ()
    {
	    Dictionary<string, double> Технологические_переделы = new Dictionary<string, double>()
	    {
	    	{"00", 0}, 
	    	{"01", 0}, 
	    	{"02", 0}, 
	    	{"03", 0}, 
	    	{"04", 0}, 
	    	{"06", 0}, 
	    	{"07", 0}, 
	    	{"08", 0}, 
	    	{"10", 0}, 
	    	{"21", 0}, 
	    	{"41", 0}, 
	    	{"42", 0}, 
	    	{"50", 0}, 
	    	{"51", 0}, 
	    	{"55", 0}, 
	    	{"60", 0}, 
	    	{"65", 0}, 
	    	{"71", 0}, 
	    	{"73", 0}, 
	    	{"74", 0}, 
	    	{"75", 0}, 
	    	{"80", 0}, 
	    	{"81", 0}, 
	    	{"85", 0}, 
	    	{"88", 0}, 
	    	{"90", 0}, 
	    	{"91", 0}
	    };
	    return Технологические_переделы;
    }
    
    private string[,] Преобразовать_словарь (Dictionary<string, double> Технологические_переделы, ref double total)
    {
    	IEnumerable<KeyValuePair<string, double>> Заполненные_виды_операций = Технологические_переделы.Where(kvp => (kvp.Value > 0));
    	int size = Заполненные_виды_операций.Count();
    	
    	string[,] Вывод_на_диаграмму = new string[size, 3];
    	if (size == 0)
    	{
    		ДобавитьЗамечания("ВНИМАНИЕ. Нет данных для формирования диаграммы.");
    		return Вывод_на_диаграмму;
    	}
    	
    	int i = 0;
    	string key;
    	int dop;
    	string признак;
    	
    	foreach (KeyValuePair<string, double> kvp in Заполненные_виды_операций)
    	{
    		key = kvp.Key;
    		Объект передел = НайтиОбъект("Технологические переделы (виды обработки)", "[Код]", key);
    		if (передел == null)
    		{
    			ДобавитьЗамечания("ВНИМАНИЕ. Технологический передел '" + key + "'не найден.");
    			continue;
    		}
    		
    		//Код вида операции
    		Вывод_на_диаграмму[i, 0] = key;
    		
    		//Количество времени, затрачиваемого на указанный вид
    		total += kvp.Value;
    		Вывод_на_диаграмму[i, 1] = kvp.Value.ToString();
    		
    		//Текст, отображаемый на диаграмме
    		dop = передел.Параметр["Дополнительные признаки"];
    		if (dop > 0)
    		{
    			ReferenceObject referenceObject = (ReferenceObject)передел;
    			Parameter parameter = referenceObject[new Guid("d7abf8df-6c69-4e98-8330-246a66b1a771")];
				признак = parameter.ParameterInfo.ValueList.GetName(parameter.Value);
				if (key == "90" || key == "91")
					признак = передел.Параметр["Технологический передел"] + "(" + признак.ToLower() + ")";
				Вывод_на_диаграмму[i, 2] = признак;
    		}
    		else
    			Вывод_на_диаграмму[i, 2] = передел.Параметр["Технологический передел"];
    		
    		i++;
    	}
    	return Вывод_на_диаграмму;
    }
    
    
    //Переменная, в которую записываются замечания
    private static string замечания;
    
    
    /****************************************************************************************************************************/
    //Функция получения замечаний, найденных при генерации карт
    /****************************************************************************************************************************/
    public void ВывестиЗамечания ()
    {
        if (замечания != "")
            MessageBox.Show(замечания, "Список замечаний", MessageBoxButtons.OK, MessageBoxIcon.Information);
    }
    //==========================================================================================================================//
    
    
    
    /****************************************************************************************************************************/
    //Процедура очистки замечаний, найденных при генерации карт
    /****************************************************************************************************************************/
    public void ОчиститьЗамечания ()
    {
        замечания = "";
    }
    //==========================================================================================================================//
    
    
    
    /****************************************************************************************************************************/
    //Процедура добавления замечаний (из других макросов)
    /****************************************************************************************************************************/
    public void ДобавитьЗамечания (string errors)
    {
        замечания += ">> " + errors + "\r\n";
    }
    //==========================================================================================================================//
    
    
    
    /****************************************************************************************************************************/
    //Функция формирования сводного наименования техпроцесса
    /****************************************************************************************************************************/
    public string СводноеНаименованиеТехпроцесса (ТехнологическийПроцесс рассматриваемыйТехпроцесс)
    {			
        string tpinfo = рассматриваемыйТехпроцесс.Обозначение + " "
                          + рассматриваемыйТехпроцесс.Наименование;
        return tpinfo;
    }
    //==========================================================================================================================//
    
    
    /****************************************************************************************************************************/
    //Функция формирования сводного наименования операции
    /****************************************************************************************************************************/
    public string СводноеНаименованиеОперации (Операция рассматриваемаяОперация)
    {			
        string operinfo = рассматриваемаяОперация.Номер
                         + рассматриваемаяОперация.Индекс + " "
                          + рассматриваемаяОперация.Наименование;
        return operinfo;
    }
    //==========================================================================================================================//
}	
