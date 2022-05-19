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
    public void АРМ_АнализСимволовКолонки()
    {
    	//ВыполнитьМакрос(MDM. АРМ. Взять объект ячейку)
    	Объекты записи = ВыполнитьМакрос("ea0d8d3c-395b-48d5-9d42-dab98371522c", "ПолучитьВыбраныеОбъекты", "Записи");
    	string параметр = ВыполнитьМакрос("ea0d8d3c-395b-48d5-9d42-dab98371522c", "ПолучитьКолонку", "Записи");
    	АнализСимволов(записи, параметр);
    }
    
    public void АРМ_АнализСимволов()
    {
    	//ВыполнитьМакрос(MDM. АРМ. Взять объект ячейку)
        Объекты записи = ВыполнитьМакрос("ea0d8d3c-395b-48d5-9d42-dab98371522c", "ПолучитьВыбраныеОбъекты", "Записи");
    	string параметр = "";
        АнализСимволов( записи, параметр);
    }
    
    private void АнализСимволов(Объекты записи, string параметр)
    {
    	string имя = "";
    	string строка = "";
    	foreach (Объект запись in записи)
    		{
    		if (параметр == "")
    			имя = запись["Сводное наименование"] + " - " + запись["Комментарий"];
    		else 
    		    имя = запись[параметр];
    				
    		if (string.IsNullOrEmpty(имя))
                continue; 
    		
    		строка = СтрокаHTML(имя);
    		 
        	запись.Изменить();
        	запись["Анализ символов"] = "<!DOCTYPE html PUBLIC \"-//W3C//DTD XHTML 1.0 Transitional//EN\" \"http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd\"> " +
                "<html xmlns=\"http://www.w3.org/1999/xhtml\"> " +
                	"<head>" +
                		"<meta http-equiv=\"Content-Type\" content=\"text/html; charset=utf-8\" /><title> " +
                		"</title>" +
                		"<style type=\"text/css\"> " +
                			".cs2654AE3A{text-align:left;text-indent:0pt;margin:0pt 0pt 0pt 0pt}" +
                			".csFD04CD6A{color:#FF0000;background-color:transparent;font-family:Calibri;font-size:11pt;font-weight:normal;font-style:normal;}" +  // красный
                			".csF8E8676A{color:#0000FF;background-color:transparent;font-family:Calibri;font-size:11pt;font-weight:normal;font-style:normal;}" +  // синий
                			".csF96BE9E{color:#008000;background-color:transparent;font-family:Calibri;font-size:11pt;font-weight:normal;font-style:normal;}" +  // зеленый
                			".csC8F6D76{color:#000000;background-color:transparent;font-family:Calibri;font-size:11pt;font-weight:normal;font-style:normal;}" +  // черный
                		"</style>" +
                	"</head>" +
                	"<body>" +
                		строка +
                    "</body>" +
                "</html>";
        	
        	запись.Сохранить();
    		}
    	ОбновитьЭлементыУправления("Записи");
    	//ОбновитьОкноСправочника();
    }
    
    private string СтрокаHTML(string имя)
    {
        int типбыл = ТипСимвола(имя[0]);  // 0 - символ черный, 1 - цифра зеленый 2 - кирилица синий 3 - латиница красный     
        int типстал = типбыл;  
        string строка = "<p class=\"cs2654AE3A\">";
        if (типстал == 0)
        	строка = строка + "<span class=\"csC8F6D76\">" + имя[0];
        else if (типстал == 1)
            строка = строка + "<span class=\"csF96BE9E\">" + имя[0];
        else if (типстал == 2)
            строка = строка + "<span class=\"csF8E8676A\">" + имя[0];
        else
            строка = строка + "<span class=\"csFD04CD6A\">" + имя[0];
        for (int i = 1; i < имя.Length ; i++)
            {
        	if (имя[i] != ' ')
        		{
            	 типстал = ТипСимвола(имя[i]);
            	// Сообщение("строка", строка + "; был " + типбыл + "; стал " + типстал + "; символ " + имя[i] +"; номер " + i);
                    if (типбыл != типстал)
                    {
                    	строка = строка + "</span>";
                        
                        if (типстал == 0)
                        	строка = строка + "<span class=\"csC8F6D76\">";
                        else if (типстал == 1)
                            строка = строка + "<span class=\"csF96BE9E\">";
                        else if (типстал == 2)
                            строка = строка + "<span class=\"csF8E8676A\">";
                        else
                            строка = строка + "<span class=\"csFD04CD6A\">";
                       типбыл = типстал ;
                    }
                }
             строка = строка + имя[i]; 
            }        
    	return строка + "</span></p>";
    }
    
    private int ТипСимвола(char символ)
    {
    	if (((символ >= 'a') && (символ <= 'z')) || ((символ >= 'A') && (символ <= 'Z')))
           return 3;
    	else if (((символ >= 'а') && (символ <= 'я')) || ((символ >= 'А') && (символ <= 'Я')))
    		return 2;
    	else if ((символ >= '0') && (символ <= '9'))
    		return 1;
    	else 
    		return 0;
    }
    
}
