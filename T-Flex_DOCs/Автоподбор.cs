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
    
    
    
        public void IOT()
            	{
       	            
        	System.Diagnostics.Debugger.Launch();
System.Diagnostics.Debugger.Break();
//System.Windows.Forms.MessageBox.Show("RAZRAB"+Переменная["$Razrab_FIO"]);
                   Объект текущий = ТекущийОбъект;
                   var kodOper= текущий["Код операции"];
                   var  cex  =  текущий.РодительскийОбъект["Код подразделения"];
                 
                      if (kodOper!="" && cex!="")
                          	{
                                var IOT = НайтиОбъекты("OperAem",String.Format("[F1] = '{0}' И [F6] = '{1}'", kodOper.ToString(), cex.ToString()));
                    
                                
                                if (IOT.Count != 0)
                                    {
                                      //  Сообщение("Текущий", "{0} {1} {2} ИОТ  {3}", текущий["Код операции"].ToString(), текущий.РодительскийОбъект["Код подразделения"].ToString(), IOT[0]["F5"], IOT.ToString().Trim());
                                        var IOTDOC = НайтиОбъекты("Документы", String.Format("[Тип] = 'Инструкция' И [Обозначение] содержит '{0}'", IOT[0]["F5"].ToString().Trim()));
                                       // Сообщение("Текущий", IOT.Count.ToString() + " " + IOT + " --  " + IOTDOC.Count.ToString());
                        
                                        текущий.Изменить();
                                        if (IOTDOC != null)
                                            {
                                                текущий.Подключить("Документы", IOTDOC[0]);
                                                текущий.Сохранить();
                                            }
                                       }
          	
       
                             }
                 }
        
       public void Prof()
       {
       	Объект текущий = ТекущийОбъект;
        var prof=(текущий.СвязанныйОбъект["Каталог оборудования"]).СвязанныеОбъекты["Профессии"][0];
        var oper= текущий.Владелец;
     //   Сообщение("Текущий",prof.ToString());
     //   Сообщение("Операция",oper.ToString());
         
            oper.Изменить();
            Объект исполнитель = oper.СоздатьОбъектСписка ("Исполнители операции", "Исполнители операции");
            исполнитель["Наименование"]=prof["Наименование профессии"];
            // исполнитель["Разряд работ"]=prof["Разряд"];
            исполнитель.Подключить("Профессия рабочего", prof);
            
            //исполнитель["Разряд рабочего"]=1;
            исполнитель.Сохранить();
            oper.Сохранить();
       }
        
}
