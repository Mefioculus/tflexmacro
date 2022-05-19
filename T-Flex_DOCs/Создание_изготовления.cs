using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using TFlex.DOCs.Model.Macros;
using TFlex.DOCs.Model.References;
using TFlex.DOCs.Model.Macros.ObjectModel;

public class Macro_AEM_TP : MacroProvider
{
    public Macro_AEM_TP(MacroContext context)
        : base(context)
    {
    }

    public override void Run()
    {
        Объект ном = ТекущийОбъект.Владелец;
        if (ном == null)
        {
        	ном = ТекущийОбъект.СвязанныйОбъект["Изготавливаемая ДСЕ"];
        	if (ном == null)
        		return;
        }
        
        ТекущийОбъект["Масса"] = ном["Масса"];
        ТекущийОбъект["Наименование"] = ном["Наименование"];
        ТекущийОбъект["Обозначение"] = ном["Обозначение"];
        
        Объекты файлы = ном.СвязанныеОбъекты["[Связанный объект].[Документы]->[Файлы]"];
        if (файлы.Count > 0)
        {
            Объект файл = файлы.FirstOrDefault();
            ТекущийОбъект.СвязанныйОбъект["9bc695f9-fa6f-4bfd-bc6f-9229b75938da"] = файл;
        }
    }
    
    public void ГенерацияМаршрута()
    {
    	var Текущий = ТекущийОбъект;
		string обозначениеТО = Текущий["Обозначение"].ToString();
    	if (обозначениеТО.StartsWith("33") || 
    	    обозначениеТО.StartsWith("34") || 
    	    обозначениеТО.StartsWith("42") || 
    	    обозначениеТО.StartsWith("60") || 
    	    обозначениеТО.StartsWith("61") || 
    	    обозначениеТО.StartsWith("62") || 
    	    обозначениеТО.StartsWith("63") || 
    	    обозначениеТО.StartsWith("66") || 
    	    обозначениеТО.StartsWith("68") || 
    	    обозначениеТО.StartsWith("46") || 
    	    обозначениеТО.StartsWith("75"))
    	{
    		Объект изготовление = СоздатьОбъект("Технологические процессы", "Изготовление");
    		изготовление["Наименование"] = Текущий["Наименование"];
    		изготовление["Масса"] = Текущий["Масса"];
    		изготовление["Обозначение"] = обозначениеТО;
    		изготовление.СвязанныйОбъект["Изготавливаемая ДСЕ"] = Текущий;
    		
    		var файлы = Текущий.СвязанныеОбъекты["[Связанный объект].[Документы]->[Файлы]"];
            if (файлы.Count > 0)
            {
                var файл = файлы.FirstOrDefault();
                изготовление.СвязанныйОбъект["9bc695f9-fa6f-4bfd-bc6f-9229b75938da"] = файл;
            }
    		
    		изготовление.Сохранить();
    		
    		var цехопереход = СоздатьОбъект("Технологические процессы", "Цехопереход", изготовление);
    		var подразделение = НайтиОбъект("Группы и пользователи", "[Номер] = '107'");
    		if (подразделение != null)
    		{
    			цехопереход.СвязанныйОбъект["Производственное подразделение"] = подразделение;
    			цехопереход["0ab2fc52-94b2-459b-a83c-15a5483c2a88"] = подразделение["Номер"];
    			цехопереход["1a2617ee-db99-4713-8115-e805ff3417ed"] = подразделение["Наименование"];
    		}
    		цехопереход.Сохранить();
    		
    		изготовление.Изменить();
    		CalculateProcessTimes((ReferenceObject)изготовление);
    		изготовление.Сохранить();
		}
    	
    	
    	var current = Context.ReferenceObject;
    	foreach (var dse in current.Children.RecursiveLoad())
    	{
    		var ДСЕ = Объект.CreateInstance(dse, Context);
    		string обозначение = ДСЕ["Обозначение"].ToString();
        	if (обозначение.StartsWith("33") || 
        	    обозначение.StartsWith("34") || 
        	    обозначение.StartsWith("42") || 
        	    обозначение.StartsWith("60") || 
        	    обозначение.StartsWith("61") || 
        	    обозначение.StartsWith("62") || 
        	    обозначение.StartsWith("63") || 
        	    обозначение.StartsWith("66") || 
        	    обозначение.StartsWith("68") || 
        	    обозначение.StartsWith("46") || 
        	    обозначение.StartsWith("75"))
        	{
        		Объект изготовление = СоздатьОбъект("Технологические процессы", "Изготовление");
        		изготовление["Наименование"] = ДСЕ["Наименование"];
        		изготовление["Масса"] = ДСЕ["Масса"];
        		изготовление["Обозначение"] = обозначение;
        		изготовление.СвязанныйОбъект["Изготавливаемая ДСЕ"] = ДСЕ;
        		
        		var файлы = ДСЕ.СвязанныеОбъекты["[Связанный объект].[Документы]->[Файлы]"];
                if (файлы.Count > 0)
                {
                    var файл = файлы.FirstOrDefault();
                    изготовление.СвязанныйОбъект["9bc695f9-fa6f-4bfd-bc6f-9229b75938da"] = файл;
                }
        		
        		изготовление.Сохранить();
        		
        		var цехопереход = СоздатьОбъект("Технологические процессы", "Цехопереход", изготовление);
        		var подразделение = НайтиОбъект("Группы и пользователи", "[Номер] = '107'");
        		if (подразделение != null)
        		{
        			цехопереход.СвязанныйОбъект["Производственное подразделение"] = подразделение;
        			цехопереход["0ab2fc52-94b2-459b-a83c-15a5483c2a88"] = подразделение["Номер"];
        			цехопереход["1a2617ee-db99-4713-8115-e805ff3417ed"] = подразделение["Наименование"];
        		}
        		цехопереход.Сохранить();
        		
        		изготовление.Изменить();
        		CalculateProcessTimes((ReferenceObject)изготовление);
        		изготовление.Сохранить();
        	}
    	}
    }
    
    private void CalculateProcessTimes(ReferenceObject process)
    {
        //var original = process;
        //System.Windows.Forms.MessageBox.Show("");
        // Вариант - пусто; тип - порожден от "Технологическая операция" 

        try
        {
            var routes = process.Children.RecursiveLoad().Where(o => o.Class.IsInherit(new Guid("459ae48b-165b-44fd-8b3e-890298f2c3d7"))).ToArray();       // Цехопереходы

            if (routes.Any())
            {
                List<string> Executors = new List<string>();
                foreach (var route in routes)
                {
                    if (route.GetObject(new Guid("30888ac1-d215-478f-aaf2-915be9aa9066")) != null)
                        Executors.Add(route.GetObject(new Guid("30888ac1-d215-478f-aaf2-915be9aa9066"))[new Guid("1ff481a8-2d7f-4f41-a441-76e83728e420")].ToString());
                    else
                        Executors.Add(" ");
                }
                process[new Guid("cf0eb573-a7e1-4025-b05d-699e2ce69a1a")].Value = string.Join("-", Executors);
            }
        }
        catch
        { }
    }
}
