using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using TFlex.DOCs.Model.Macros;
using TFlex.DOCs.Model.Macros.ObjectModel;
using TFlex.Reporting.Technology.Macros;
using TFlex.Reporting.CAD.MacroGenerator.Macros;
using Newtonsoft.Json;

public class Macro : ReportMacroProvider
{
    public Macro(ReportGenerationMacroContext context)
        : base(context)
    {
    }

    public override void Run()
    {
        //Получение текущего объекта (Технологический процесс)
    	Объект ТП = ТекущийОбъект;
    	
    	//Заполнение графы "Обозначение ТП"
    	Переменная["$ОбозТП"] = ТП.Параметр["Обозначение"];
    	Переменная["$Составил"] = ТекущийПользователь.Параметр["Короткое имя"];
    	Переменная["$Заготовка"] = ТП.Параметр["Сводное наименование материала"];  	
    		
    	//Получение данных из буфера обмена, которые должны поступить из диалога ввода параметров, инициализированного в макросе "Формирование сопроводительной карты"
        string json = Clipboard.GetText();

        
        
        
        //Попытка десериализации объета и получение из него данных
        Data newDataReport;
    	try
    	{
    		newDataReport = JsonConvert.DeserializeObject<Data>(json);

            Переменная["$НомСопрКарт"]  = newDataReport.номерКарты;
            Переменная["$НомЗаказ"] = newDataReport.заказ;
            Переменная["$НомПроизЗаказ"] = newDataReport.заказ1С;
            Переменная["$ДатаЗапуск"] = newDataReport.датаЗап;
            Переменная["$СвидетЦЗЛ"] = newDataReport.свидетельство;
            Переменная["$ОбозначИзд"] = newDataReport.обозначение;
            Переменная["$НаимИзд"] = newDataReport.наименование;

    	}
    	catch
    	{}

    	
    	
        //Получение связанного ДСЕ и зполнение параметров, относящихся к ДСЕ    	
    	Объект ДСЕ = ТП.СвязанныйОбъект["Изготавливаемая ДСЕ"];
    	if (ДСЕ != null)
    		{
            Переменная["$Обознач"] = ДСЕ["Обозначение"];
            Переменная["$Наим"] = ДСЕ["Наименование"];
            Переменная["$Материал"] = ДСЕ.Параметр["Марка"];
            }
    	else
    		{
    		Переменная["$Обознач"] = "";
            Переменная["$Наим"] = "";
            Переменная["$Материал"] = "";
    		}
    	
    	
    	/*
        //Получение данных об обозначении и наименовании изделия, к которому относится
        //данная ДСЕ
        var tempobj = дсе;
        string наименование = дсе.Параметр["Наименование"];
        string обозначение = дсе.Параметр["Обозначение"];
        
        
        while (true)
        	{
            try
            	{
                if (tempobj.РодительскийОбъект["Тип"] == "Папка")
                	{
                     break;
                    }
                tempobj = tempobj.РодительскийОбъект;
                наименование = tempobj.Параметр["Наименование"];
                обозначение = tempobj.Параметр["Обозначение"];
                }
            catch
            	{
                break;
                }
            }
        
        */

            
        

        
        //Заполнение данных по операции
        var текст = Текст["Текст1"];
        var шаблонСтроки = текст["НомЦех"];
        string номерЦеха = "";
        foreach (var цехозаход in ТП.ДочерниеОбъекты)
        {
            foreach (var операция in цехозаход.ДочерниеОбъекты)
            {
                	var строкаОперация = текст.Таблица.ДобавитьСтроку(шаблонСтроки);
                	// Добавление проверки на ошибки, чтобы отчет генерировался даже в том случае, если не указано подразделение
                	try
                	{
                        строкаОперация["НомЦех"].Текст = операция.СвязанныйОбъект["Производственное подразделение"]["Номер"];
                    }
                	catch
                	{
                	   	строкаОперация["НомЦех"].Текст = "-";
                	}
                	
                    строкаОперация["НомОпер"].Текст = операция["Номер"];
                    строкаОперация["Шифр"].Текст = операция["Код операции"];
                    строкаОперация["НаименованиеОперации"].Текст = операция["Наименование"];
                    
                    // Получение списка исполнителей, для того, чтобы с первого из исполнителей получить данные о разряде работ
                    try
                    {
                    	Объекты списокИсполнителей = операция.СвязанныеОбъекты["Исполнители операции"];
                    	
                    	if (списокИсполнителей.Count > 0)
                    	{
                    		строкаОперация["РазРаб"].Текст = списокИсполнителей[0]["Разряд работ"];
                    	}
                    }
                    catch
                    {}
                    
                   
                    if (операция["Штучное время"] != 0.0)
                    {
                    	строкаОперация["Тшт"].Текст = операция["Штучное время"];
                    }
                    
                    // Проверка на подключенное подразделение
                    try
                    {
                        номерЦеха = операция.СвязанныйОбъект["Производственное подразделение"]["Номер"].ToString();
                    }
                    catch
                    {
                    	номерЦеха = "-";
                    }
            }
        
            //Определение переменной номера цеха. Эта переменная будет определяться из номера цеха последней операции технологического процесса
        
            Переменная["$НомЦех"] = номерЦеха;
        }

            

        
        //Все переменные, которые требуется заполнить в данный момент
        
       
        //Переменная["$КолДет"]
        //Переменная["$НомПарт"]
        //Переменная["$КолНаИзд"]
        //Переменная["$КолПоПлан"]
        	
    }

    public class Data
    {
    	public string обозначение { get; set; }
    	public string наименование { get; set; }
        public string заказ { get; set; }
        public string заказ1С { get; set; }
        public string датаЗап { get; set; }
        public string номерКарты { get; set; }
        public string свидетельство { get; set; }

    }


}






