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
    
    public string result=""; 
    public Объект result2=null;
    public Macro(MacroContext context)
        : base(context)
    {
    }

    public override void Run()
    {
    
    }
    
    public void NumberZakaz()
        {
        
        //СвязьсДСЕ();
    //    System.Diagnostics.Debugger.Launch();
    //    System.Diagnostics.Debugger.Break();
        int nextnum =0;
        Объект деталь = ТекущийОбъект;
        DateTime date1 = DateTime.Today;
        int year = date1.Year;
        int num=0;
       
                 
                 
       Объекты оснастка = НайтиОбъекты("Заказы на оснастку","[ID] > '0' ");
      
       
       foreach (Объект объект in оснастка)
           {
              if (объект["ID"]>num)
                      {
                          num=объект["ID"];
                   }
        }
    
       Объект оснастк2 = НайтиОбъект("Заказы на оснастку","[ID] = '"+num.ToString()+"' ");
   
       
       if (!(оснастк2["Номер"].ToString().Substring(оснастк2["Номер"].ToString().IndexOf("-")+1,2).Equals(year.ToString().Substring(2, 2))))
            nextnum =1;
       else          
            nextnum=Int32.Parse(оснастк2["Номер"].ToString().Substring(0, оснастк2["Номер"].ToString().IndexOf("-")))+1;
       
       
       деталь.Изменить();
       if (nextnum<100)
            деталь["Номер"]="0"+nextnum.ToString()+"-"+year.ToString().Substring(2,2);
       else
            деталь["Номер"]=nextnum.ToString()+"-"+year.ToString().Substring(2,2);
       changeNumOsn(деталь);
       деталь.Сохранить();    
        
       }
    
    
    public void Copy()
        {
            DateTime date = DateTime.Today;        
            Объект текущий = ТекущийОбъект;
            Объект копия=текущий.Копия();
            копия["Дата"]=date;
            копия.Сохранить();
           
            
        }
    
    
           public Объект ParentObj(Объект объект)
            {
            
            if (объект.РодительскиеОбъекты.Count>0)
                          {
                        foreach (var подкл in объект.РодительскиеОбъекты)
                                {
                                                                                          
                                         ParentObj(подкл);
                                 }       
                        }
                        
    
    
        
    
            if (объект.РодительскиеОбъекты.Count==0)
                    result2 = объект;
            
            return result2;
            }
           
           
        public string Parent(Объект объект)
            {
            //string result="";      
            if (объект.РодительскиеОбъекты.Count>0)
                          {
                        foreach (var подкл in объект.РодительскиеОбъекты)
                                {
                                                                                          
                                         Parent(подкл);
                                 }       
                        }
                        
    
    
        
    
            if (объект.РодительскиеОбъекты.Count==0)
                    result = объект["Обозначение"].ToString();
            
            return result;
            }
     
     
     
     public void Test()
         {
         
               Объект оснастка = ТекущийОбъект;
              Объект объект= оснастка.СвязанныйОбъект["Связь с ЭСИ"];
        if ( объект!=null)
               {
                   Объект izd=ParentObj(объект);
                   оснастка.Изменить();
                   оснастка["Индекс"]=izd["Обозначение"].ToString();
                   оснастка["Изделие"]=izd["Наименование"].ToString();
                   оснастка.Сохранить();
               }
                
         }
     
     
         public void НумерацияВДокументыОГТ2()
    {
        // Для начала получаем текущий объект (как я понимаю, это должен быть объект, который в данный момент создается
        Объект запись = ТекущийОбъект;
        
        
         // Далее, получаем список всех объектов данного справочника
        Объекты существующиеЗаписи = НайтиОбъекты("500d4bcf-e02c-4b2e-8f09-29b64d4e7513", "Тип", "Документы ОГТ");   
        Сообщение("Информация",существующиеЗаписи.Max().ToString());
//         if (существующиеЗаписи != null)
//        {
//            
//            List<int> list_num = new List<int>();
//            // Мы получаем номер записи с самым большим значением
//            foreach(Объект существующаяЗапись in существующиеЗаписи)
//                      list_num.Add(существующаяЗапись["Номер"]);
//            Сообщение("Информация",list_num.Max().ToString());
//           // ТекущийОбъект.Параметр["Номер"] = list_num.Max() + 1;
//        }
     }  

         
         public void СвязьсДСЕ()
             {
             
                  Объект заказ = ТекущийОбъект;
                 // Message("Оснастка",заказ.ToString());
                  Объект связь= заказ.СвязанныйОбъект["Связь с ЭСИ"];
                 Объект оснастка = заказ.Владелец;    
              //   Message("Оснастка",оснастка.ToString());    
           
              //   Message("Связь",связь.ToString());    
                 
                 Объект владелецЦехпереход = оснастка.Владелец;
                 Объект изготовление=владелецЦехпереход.РодительскийОбъект;
                 Объект деталь=изготовление.Владелец;
         
                
                if (деталь == null)
                    {
                   деталь = изготовление.СвязанныйОбъект["Изготавливаемая ДСЕ"];
                   if (деталь == null)
                      return;
                    }
         
        
                ТекущийОбъект.СвязанныйОбъект["Связь с ЭСИ"]=деталь;
                ТекущийОбъект["Наименование ДСЕ"]=деталь["Наименование"];
                ТекущийОбъект["Деталь"]=деталь["Обозначение"];
                          
             }
         
         
              public void changeNumOsn(Объект currentObj)
                    {
       //           System.Diagnostics.Debugger.Launch();
    //    System.Diagnostics.Debugger.Break();
                  
                       // Объект currentObj= ТекущийОбъект;
                 //       currentObj.Изменить();
                        currentObj["NumId"]=changing(currentObj["Номер"].ToString());
                       // currentObj.Сохранить();    
                    }
              
              
             static public int changing(string number)
            {
                if (number != "" & number.Contains('-'))
                {
                    var parsenum = number.Split(new char[] { '-' })[1] + number.Split(new char[] { '-' })[0];
                    int result;
                    if (int.TryParse(parsenum, out result))
                        return result;
                    else
                        return 0;
                }
                else
                    return 0;

                return 0;
            }
         
                        
    
    
        
    

         
         


         
     
     
}
