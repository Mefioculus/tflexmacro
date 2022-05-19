using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using TFlex.DOCs.Model.Macros;
using TFlex.DOCs.Model.Macros.ObjectModel;
using System.Text;
using System.Text.RegularExpressions;
using TFlex.DOCs.Model.References;
using TFlex.DOCs.Model.Desktop;

public class Macro : MacroProvider
{
	
	public  List<Объект> DCE = new List<Объект>();
    public Macro(MacroContext context)
        : base(context)
    {
        System.Diagnostics.Debugger.Launch();
        System.Diagnostics.Debugger.Break();
    }

    public override void Run()
    {
    
    }
    
    
    
    
    
    
     public void Test()
    	{
     	
  // 	System.Diagnostics.Debugger.Launch();
//    System.Diagnostics.Debugger.Break();
    ОбновитьОкноСправочника();        //8А3119055
            //8А3.119.055
             var dataStart = DateTime.Now;
            ДиалогОжидания.Показать("Пожалуйста, подождите", true);
                        
            StringBuilder strBuilder = new StringBuilder();
            var текущий= ТекущийОбъект;
            текущий.Изменить();
            
            SetZag(текущий);
            SetIzmM(текущий);
            //izg.СвязанныеОбъекты["Материалы/Заготовки"];
            
        /*    
            
             Объект обор = опер.СоздатьОбъектСписка ("Оснащение", "Оборудование"); 
                                       // if (oborTF!=null)   
                               //        Сообщение("",trud["NAIM_ST"]) ;
                                        обор["Строка оснащения"]=trud["NAIM_ST"].ToString();
                                        обор.Сохранить();
                                        опер.Сохранить();
          
       */  
          //  текущий.
          //  текущий.Сохранить();
         //   текущий.Изменить();
       //     Сообщение("",текущий.ToString());
       //     Сообщение("",текущий.Тип.ToString());
       //     Сообщение("",текущий.ДочерниеОбъекты.Count.ToString());
          



                  // /*
                   Объекты цехозаходы= текущий.ДочерниеОбъекты;
                        //AddOper(цехозаходы[0]);
                        foreach (var cexp in цехозаходы)
                        	{
                               // strBuilder.AppendLine("Код №"+cexp["Код подразделения"]+" --- "+cexp["Номер цехоперехода"]);
                                AddOper(cexp,текущий);
                                 ДиалогОжидания.СледующийШаг("Идёт формирование технологического процесса");
                                //AddOperT(cexp);
                            //   GetOper(cexp);
                           //   Сообщение("","CHEXP");
                            }
                        
                                    
                    //               Сообщение("",strBuilder.ToString());
                   // */
         текущий.Сохранить();
         ДиалогОжидания.Скрыть();
         var dataStop = DateTime.Now;
         Сообщение("",String.Format("Технологический процесс сформирован {0}",(dataStop-dataStart).ToString()));
        }
    
     
    public void Test2()
        {
    	 // Объект obj= ТекущийОбъект;
    	  //8А8.603.299
    	  //8А8.942.348
    	  Объект obj = НайтиОбъект("Технологические процессы", "Обозначение = '8А8.942.348'");
    	  Сообщение("{1}",obj.ToString());
    	  Сообщение("{2}",obj.Тип.ToString());
    	  //8А8.942.348
    	  
    	  
    	  var списокматериалов = obj.СвязанныеОбъекты["Материалы/Заготовки"];
    	  
    	  Сообщение("",списокматериалов.Count().ToString() );
    	  Сообщение("",списокматериалов[0].ToString() );
    	  Сообщение("",String.Format("{0}  {1}  {2}",списокматериалов[0].ToString(),
    	                                    списокматериалов[0]["Обозначение"].ToString(),
    	                                    списокматериалов[0]["Размеры"].ToString()
    	                            )
    	           );
//    	                             ,
////                    	                             списокматериалов[0]["Обозначение"].ToString(),
 //                   	                             списокматериалов[0]["Размеры"].ToString()  ));
    	  //Обозначение
    	}
 

    public void GetCZ(Объект TP)
    	{
    	
 //  	System.Diagnostics.Debugger.Launch();
 //   System.Diagnostics.Debugger.Break();
            //8А3119055
            //8А3.119.055
        //    StringBuilder strBuilder = new StringBuilder();
        //    var текущий= ТекущийОбъект;
       //     Сообщение("",текущий.ToString());
       //     Сообщение("",текущий.Тип.ToString());
       //     Сообщение("",текущий.ДочерниеОбъекты.Count.ToString());
            Объекты цехозаходы= TP.ДочерниеОбъекты;
            //AddOper(цехозаходы[0]);
            foreach (var cexp in цехозаходы)
            	{
                   // strBuilder.AppendLine("Код №"+cexp["Код подразделения"]+" --- "+cexp["Номер цехоперехода"]);
                    AddOper(cexp,TP);
                  //  AddOperT(cexp);
                //   GetOper(cexp);
               //   Сообщение("","CHEXP");
                }
            
                        
        //               Сообщение("",strBuilder.ToString());
        }    
   
        
  public void test6()
  	{
  	    
  	      ДиалогОжидания.Показать("Пожалуйста, подождите", true);
  	      StringBuilder strBuilder2 = new StringBuilder();   
          Объект деталь = ТекущийОбъект;
          //  itemsParent=деталь;
          DCE.Add(деталь);            
          RecursESI(деталь);
            
           // Сообщение("",деталь.РодительскиеОбъекты.Count.ToString());
            Сообщение("",DCE.Count.ToString());
             var dataStart = DateTime.Now;
            foreach (var s in DCE) 
            	{ 
            	//if (s.Тип.ToString().Equals("Деталь") || s.Тип.ToString().Equals("Сборочная единица") || s.Тип.ToString().Equals("Прочее изделие") )
                    strBuilder2.AppendLine("ДСЕ "+ s["Обозначение"].ToString()+s.Тип.ToString());            	
                    test5(s);
                    ДиалогОжидания.СледующийШаг("Идёт формирование технологического процесса");
                }

    //       Сообщение("str2", strBuilder2.ToString());
          // Сообщение("str", strBuilder.ToString());

  	
  	  ДиалогОжидания.Скрыть();
  	  var dataStop = DateTime.Now;
      Сообщение("",String.Format("Технологический процесс сформирован {0}",(dataStop-dataStart).ToString()));
    }
  
  
   public void test7()
        {
    	      StringBuilder strBuilder2 = new StringBuilder();   
            Объект деталь = ТекущийОбъект;
          //  itemsParent=деталь;
             DCE.Add(деталь);            
            RecursESI(деталь);
            Сообщение("",DCE.Count.ToString());
          }
  
  
          public void RecursESI(Объект текущий)
		{
        	
        	
        		if (текущий.ДочерниеОбъекты.Count != 0)
        			{
                		foreach (var подкл in текущий.ДочерниеОбъекты)
                            {       
                                    RecursESI(подкл);  
                                    if (подкл.Тип.ToString().Equals("Деталь") || подкл.Тип.ToString().Equals("Сборочная единица") || подкл.Тип.ToString().Equals("Прочее изделие") )
                                    DCE.Add(подкл);
                               	   // strBuilder.AppendLine(подкл["Обозначение"].ToString() +" - "  +подкл.РодительскийОбъект.ToString());
                               	   //strBuilder.AppendLine(подкл["Обозначение"].ToString()+"-"+itemsParent["Обозначение"].ToString());
                                  // itemsParent=подкл;
                            }
        	       	}
        		
		}

    
      
  public void test5(Объект деталь)
        	{
  	
  	
  	
  	
           //var текущий= ТекущийОбъект;
          
           foreach(Объект tp in деталь.СвязанныеОбъекты["Технология"])
               	{
                    
                     GetCZ(tp);
                }
           
  //        Сообщение("",текущий.GetType().ToString());
          
  //        Сообщение("",текущий["Обозначение"].ToString());
  //        Сообщение("",текущий["Обозначение"].ToString());
  //        Сообщение("",текущий["Обозначение"].ToString().Replace(".",""));
         //  текущий["Текст перехода"]=tper;
           //текущий.Значение("Текст перехода")="TEST";
        //   Сообщение("Изменяем текст перехода",текущий["Текст перехода"]);
        //   текущий.Сохранить();
        //   Сообщение("После сохранения",текущий["Текст перехода"]);            
        	
           }
  
  

        
   /*     
        public void test4()
        	{
                     var текущий= ТекущийОбъект;
           Сообщение("Исходный текст перехода",текущий["Текст перехода"]);
           var tper= текущий["Текст перехода"];
           
           
           Сообщение("Текст перехода переменная",tper+" "+tper.GetType().ToString());
           tper=tper+"TEXT2";
          // Сообщение("Текст перехода переменная",);
          Сообщение("Текст перехода переменная2",tper+" "+tper.GetType().ToString());
           текущий["Текст перехода"]=tper;
           //текущий.Значение("Текст перехода")="TEST";
           Сообщение("Изменяем текст перехода",текущий["Текст перехода"]);
           текущий.Сохранить();
           Сообщение("После сохранения",текущий["Текст перехода"]);            
        	
            }
            */
    
/*        
        public void Test3()
    	{
    	
    //	System.Diagnostics.Debugger.Launch();
   //     System.Diagnostics.Debugger.Break();

         //   StringBuilder strBuilder = new StringBuilder();
         
          var Изг = ТекущийОбъект;
          Сообщение("ТП",Изг.Тип.ToString()); 
          
          Объекты заход = Изг.ДочерниеОбъекты;
          Сообщение("Заход",заход.Count.ToString()+" "+заход[0].Тип.ToString()+" "+заход[0].ToString()); 
          
          Объекты Опер= заход[0].ДочерниеОбъекты;
          Сообщение("Операция",Опер.Count.ToString()+" "+Опер[0].Тип.ToString()+" "+Опер[0].ToString());  
            
          Объекты перех= Опер[0].ДочерниеОбъекты;
          Сообщение("Переход",перех.Count.ToString()+" "+перех[0].Тип.ToString()+" "+перех[0].ToString());  

Изг.Изменить();
заход[0].Изменить();
Опер[0].Изменить();
перех[0].Изменить();
перех[0]["Текст перехода"]="22291020";
перех[0]["Текст перехода"]="22291020";
Сообщение("Переход",перех.Count.ToString()+" "+перех[0].Тип.ToString()+" "+перех[0].ToString()); 
перех[0].Сохранить();
Опер[0].Сохранить();
заход[0].Сохранить();
Изг.Сохранить();
          // Опер.Изменить();            
          
          // перех[0].Изменить();
         //  Сообщение("",перех[0]["Текст перехода"].ToString());
       //    перех[0]["Текст перехода"]="222";
       //    Сообщение("",перех[0]["Текст перехода"].ToString());
       //    перех[0].Сохранить();
       //    Опер.Сохранить();
        }
        */
    
    /*
     public void AddOperT(Объект cehzah)
        {
 cehzah.Изменить();
                        	//цехозаход.Подключить(Oper[0]);
                            Объект опер = СоздатьОбъект("Технологические процессы","Технологическая операция", cehzah); 
                            опер["Тип нумерации"]=1;
                        	опер["Номер"] ="001"; 
                            опер["Код операции"] ="001"; 
                            опер["Наименование"] ="Операция"; 
                                                  	                       	
                        	

                        	
                        	
                        	опер.Сохранить();
                        	//опер.Изменить();
                        	Объект перех = СоздатьОбъект("Технологические процессы","Технологический переход",опер);
                        	перех["Текст перехода"]="Зачистить заусенцы";
                        	перех.Сохранить();
                        	//перех.Изменить();
                        	
                        	
                        	//перех.Сохранить();
                        	//Объект оснастка = СоздатьОбъект("Технологические процессы","Оснащение", опер);
                        //	опер.Сохранить();
                            //Подключение подкл2 = СоздатьПодключение(цехозаход, Oper[0]);
                            cehzah.Сохранить();

        }
    */
    /*
        public void GetOper(Объект cehzah)
    {
        		StringBuilder strBuilder3 = new StringBuilder();
        Сообщение("Операций в цехозаходе", cehzah.ДочерниеОбъекты.Count.ToString());
        foreach (Объект oper in cehzah.ДочерниеОбъекты)
        {
            strBuilder3.AppendLine("Номер+КОД= " + oper["Номер"] + oper["Код операции"] );
          
            
        }
         Сообщение("",strBuilder3.ToString());
    }
      */  
    
     
     
//Проверка есть ли в текущем цехозаходе добавляемая операция
    public bool GetOper(Объект cehzah,Объект AddOper)
      	{

      	StringBuilder strBuilder3 = new StringBuilder();

      	bool result=false;
      	ОбновитьОбъектыСправочника(cehzah.ДочерниеОбъекты);
      	ОбновитьОкноСправочника();
      	var refobj = (ReferenceObject)cehzah;
     // 	 Сообщение("Мод",refobj.IsModified.ToString());
      
//	СохранитьВсё(применитьИзменения, "комментарий", показыватьДиалогПодтверждения);
      	foreach (Объект oper in cehzah.ДочерниеОбъекты)
      		{
                  //  ()oper.ApplyChanges
      		    if ((oper["Номер"]+oper["Код операции"]).Equals(AddOper["NUM_OP"] +AddOper["SHIFR_OP"]))
                    { 

      		    	result=true;
      		    	break;
      		    	}
          	}

      	      return result;
        }
    
    
    
     public   void Perehod(Объект trud, Объект oper)
     	{
            //Сообщение("",trud["OP_OP"]);
            
    /*        
             	Array array = trud["OP_OP"].ToString().Split(';');
                               	
                               	foreach (string s in array)
                               		{
                                        Объект оснащение2 = опер.СоздатьОбъектСписка ("Оснащение", "Оснащение"); 
                                       	оснащение2["Строка оснащения"]=s;
                                      // 	Сообщение("AddOSn",оснащение2["Строка оснащения"].ToString()+" "+s)    ;
                                        оснащение2.Сохранить();
                                        опер.Сохранить();
                               	    }
   */         
            
  
         //   string pattern = @"(\S*[а-я])[.]([А-Я][а-я]\S*)";
         //   string replace = "$1.\r\n$2";
         //   Regex regex = new Regex(pattern);
  
         //   string modpereh = Regex.Replace(trud["OP_OP"].ToString(), pattern, replace);
            
           Array arrpereh = trud["OP_OP"].ToString().Split(new string[] { "\r\n" }, StringSplitOptions.RemoveEmptyEntries);;
                               	
                               	foreach (string s in arrpereh)
                               		{
                                        Объект переход = СоздатьОбъект("Технологические процессы","Технологический переход",oper); 
                                       	переход["Текст перехода"]=s;
                                        переход["Пользовательский шаблон текста перехода"]=s;
                                        переход.Сохранить();
                               	    }
                

        }
    
    
      public   void AddOper(Объект cehzah, Объект деталь)
    {
      	//cehzah.Изменить();
 StringBuilder strBuilderoper = new StringBuilder();
 
 
      	// 8А3.119.055
      	//8А7.064.013
      	
    
      	
    	string shifr = деталь["Обозначение"].ToString().Replace(".","");


    	 var Trud = НайтиОбъекты("TRUD",
    	   String.Format("[SHIFR] = '{0}' ", shifr));                     

    	 
     	var list = new List<int>();
    	var zah = new List<int[]>();

        var listTrud = new List<Объект>();

            foreach (var trud in Trud)
    		{
                	int result;
                	if (Int32.TryParse(trud["NUM_OP"], out result))
               		list.Add(result);
            }	
//сортировка операций
             list.Sort();

 // переносим оперции в массив в отсортированом виде             
               foreach (int i in list)
               {

                    foreach (var trud in Trud)
        		      {   
                    	int result;
                    	if (Int32.TryParse(trud["NUM_OP"], out result))
                    		{
                    	           if (i==result)
                    	           	{
                    	           	   
                    	               listTrud.Add(trud);
                    	            }
                            }                    	   	
                        }
               
               }
               

                
                int lt1=1;
             //   string previzg ="";
                for (int lt=0; lt<listTrud.Count;lt++)
                	{
              
                    
                	
                	  if (lt > 0)
            {
                if (listTrud[lt - 1]["IZG"] != listTrud[lt]["IZG"])
                    lt1 = lt1 + 1;
            }
                	
                	listTrud[lt]["NPER"]=lt1;
                	listTrud[lt].Сохранить();
               	
                    }
               
 

//System.Diagnostics.Debugger.Break();


//Для переданного цехозахода добавляем операцию
foreach (var trud in listTrud)
	
    		{
                	

	                var Oper = НайтиОбъект("Типы технологических операций",
                    String.Format("[Код] = '{0}' И [Id] > 1365", trud["SHIFR_OP"]));
                	
/*
strBuilderoper.AppendLine("Getoper =="+
	                      GetOper(cehzah,trud)+ "Oper" +
	                      cehzah["№"].ToString()  +
	                      " и "+
	                      trud["NPER"].ToString()+
	                      " is " +
	                      ((int)cehzah["№"]==(int)trud["NPER"])+
	                      "Num oper "+
	                       cehzah["Код подразделения"].ToString()+
	                       " == "+
	                       trud["IZG"].ToString()+
	                       " is "+
	                       ((int)cehzah["Код подразделения"]==(int)trud["IZG"])

);
*/	
                //Номер цехоперехода
                	if  ((int)cehzah["№"]==(int)trud["NPER"] & !GetOper(cehzah,trud) & (int)cehzah["Код подразделения"]==(int)trud["IZG"])
                		{
                	
         //     System.Diagnostics.Debugger.Break();
          //    Сообщение("","ADD"+Oper["Название"]);
   
   
                            cehzah.Изменить();
                        	                          Объект опер = СоздатьОбъект("Технологические процессы","Технологическая операция", cehzah); 
                        	 var ЕИтштопер = НайтиОбъект("Единицы измерения","[Код ОКЕИ] = 355");    
                            
                            опер["Тип нумерации"]=1;
                        	опер["Номер"] =trud["NUM_OP"];
                        	опер["Штучное время"]=trud["NORM_T"]*60+trud["NORM_M"];
                              if (опер["Номер"]=="010")  
                                      {                              	
                        	       Message("",trud["NORM_M"].ToString());
                        	       Message("",опер["Штучное время"].ToString());
                        	       }
                        	//Штучное время
                        	опер.Подключить("ЕИ суммарного штучного времени", ЕИтштопер);
                        	//Сообщение("",опер.СвязанныйОбъект["ЕИ суммарного штучного времени"]["Guid"].ToString()); c63f08aa-ae67-4350-ba75-41b7d494d7ca


                        	
                           if (Oper!=null)  
                                    {                           	
                                        опер["Код операции"] =Oper["Код"]; 
                                        опер["Наименование"] =Oper["Название"]; 
                                    }
                            else 
                                {
                                    	опер["Код операции"] =trud["SHIFR_OP"];
                                    
                                      
                                      //0770 Операции нет в БД
                                }
                            
                        
                        	опер.Сохранить();
//                    
                            Perehod(trud, опер);
                   


//Добавляем исполнителя   

                if  (trud["PROF"]!=null)
                	{
                           	         var profTrud = НайтиОбъект("Профессии",String.Format("[Наименование профессии] = '{0}' И [ID] > '131'", trud["PROF"]));    
                      //    Сообщение("Исполнитель",profTrud.ToString())    ;                        
              	         Объект исполнитель = опер.СоздатьОбъектСписка ("Исполнители операции", "Исполнители операции"); 
                            if (profTrud!=null)
              	         	   {
              	         //  Сообщение("AddProf","AddProf")    ; 
              	         //  Сообщение("A",profTrud["Наименование профессии"].ToString());
              	         исполнитель["Наименование"]=profTrud["Наименование профессии"].ToString();
                         исполнитель["Разряд работ"]=trud["RAZR"].ToString();
                       //  Сообщение("1",trud["RAZR"].ToString());
                             }   
                        else 
                                {
                        	исполнитель["Наименование"]=trud["PROF"].ToString();
                        	исполнитель["Разряд рабочего"]=trud["RAZR"].ToString();
                      //  	Сообщение("2",trud["RAZR"].ToString());
                                }
                                        исполнитель.Сохранить();
                                        опер.Сохранить();
                  }    


                
           
//Добавляем оснастку      

               if (!trud["OSN_TARA"].ToString().Equals(""))
                       	{
               	                    //устанавлеваем разделитель ;
               	                string patternosn = @"(,\s)([А-Яа-я/\d][а-я\d]*\b)";
//               	                 	@"(,\s)([А-Я][а-я]*\b)";
                                 string replaceosn = ";$2";
                     		     string modOSN = Regex.Replace(trud["OSN_TARA"].ToString(), patternosn, replaceosn);
               	
                               	Array array = modOSN.ToString().Split(';');
                               	
                               	foreach (string s in array)
                               		{
                               		              
                               		
                                        Объект оснащение2 = опер.СоздатьОбъектСписка ("Оснащение", "Оснащение"); 
                                       	оснащение2["Строка оснащения"]=s.Trim();
                                      // 	Сообщение("AddOSn",оснащение2["Строка оснащения"].ToString()+" "+s)    ;
                                        оснащение2.Сохранить();
                                        опер.Сохранить();
                               	    }
                       	}
               
//Добавляем оборудование               
               if (!trud["NAIM_ST"].ToString().Equals(""))
                       	{
                              	
                  //  var oborTF = НайтиОбъект("Средства технологического оснащения",String.Format("[Наименование] = '{0}'", trud["NAIM_ST"]));                               
               	        
                               	
                                        Объект обор = опер.СоздатьОбъектСписка ("Оснащение", "Оборудование"); 
                                       // if (oborTF!=null)   
                               //        Сообщение("",trud["NAIM_ST"]) ;
                                        обор["Строка оснащения"]=trud["NAIM_ST"].ToString();
                                        обор.Сохранить();
                                        опер.Сохранить();


                       	}
               
               
               //Добавляем ИОТ
               //[Тип] = 'Инструкция' И [Обозначение] = 'ИОТ-311'
var OppAnalog = НайтиОбъект("KAT_OPER",String.Format("[SHIFROP] = '{0}'",опер["Код операции"]));
   //Сообщение("OppAnalog",OppAnalog.ToString()+" "+OppAnalog["NUM_IOT"].ToString().Replace("ИОТ  ","ИОТ-"));  
    if    (OppAnalog!=null)
    	{
            var ИОТ = НайтиОбъект("Документы",String.Format("[Тип] = 'Инструкция' И [Обозначение] = '{0}'",OppAnalog["NUM_IOT"].ToString().Replace("ИОТ ","ИОТ-")));
            //Сообщение(">>",ИОТ.ToString());
            if (ИОТ!=null)
            	{
            	опер.Подключить("Документы", ИОТ);
            	опер.Сохранить();
            //	Сообщение("",ИОТ.ToString());
            	}
            	
         }          
            //   Сообщение("",опер["Штучное время"].ToString());
                            cehzah.Сохранить();
                            
                        }
                	
            }	
// Сообщение("",strBuilder1.ToString());
//Сообщение("",strBuilderoper.ToString());
 //*/   
    	
    }
 


// Переименовывает имена папок так  чтобы  наименование папки соответствовало  типу операции

  public void ChangeFolderOpp()
    	{
    	
             var TipOper = new Dictionary<string, string>
            {
                {"Операциии общего назначения", "Общего назначения" },
                {"Технический контроль", "Технический контроль"},
                {"Перемещение", "Перемещение"},
                {"Литьё металов и сплавов", "Литьё металлов и сплавов"},
                {"Обработка давлением", "Обработка давлением"},
                {"Обработка резанием", "Обработка резанием"},
                {"Испытание", "Испытания"},
                {"Получение покрытий (не органических)", "Получение покрытий (металлических и неметаллических неорганических)"},
                {"Получение покрытий органических{лакокрасочных}", "Получение покрытий органических(лакокрасочных)"},
                {"Пайка", "Пайка"},
                {"Электромонтаж", "Электромонтаж"},
                {"Сборка", "Сборка"},
                {"Сварка", "Сварка"},
                {"Термообработка", "Термическая обработка"},
                {"Консервация и упакование","Консервация и упаковывание"}
            };
             
             var OpersTip = НайтиОбъекты("Типы технологических операций","[ID] > '1361' И [Тип] = 'Группа'");
    	     
              foreach (Объект ot in  OpersTip)
                      	{
                             // if (ot["Название"].ToString().Equals("Операциии общего назначения"))
                             
                           if  (TipOper.ContainsKey(ot["Название"].ToString()))
                                       	{
                                       	    ot.Изменить();
                                       	    ot["Название"]=TipOper[ot["Название"].ToString()];
                                       	    ot.Сохранить();
                                        }
                        }
             
             
             
        }
  
  
// Создаем отсутсвующие операции из аналогов 
      public void CreateOper()
    	{
  // 	System.Diagnostics.Debugger.Launch();
  // System.Diagnostics.Debugger.Break();
    	 var OpersTip = НайтиОбъекты("Типы технологических операций","[ID] > '1361' И [Тип] = 'Группа'");
    	StringBuilder strBuilderoper = new StringBuilder();
            var OpersAnalog = НайтиОбъекты("KAT_OPER","[ID] > '0'"); 
             //var OpersAnalog = НайтиОбъекты("KAT_OPER","[SHIFROP] = '0424'"); 
         
                foreach (Объект OperAnalog in OpersAnalog)
                	
                	{
                	   var OpersAemDocs = НайтиОбъект("Типы технологических операций",String.Format("[ID] > '1363' И [Код] = '{0}'", OperAnalog["SHIFROP"]));
                	   
                        	if (OpersAemDocs==null)
                    //   		System.Diagnostics.Debugger.Break();
                            		{
                            		 //  strBuilderoper.AppendLine(OperAnalog["SHIFROP"]+" "+OperAnalog["NAIMOP"]);
                                		
///*
                                        foreach (Объект ot in  OpersTip)
                                        		   	{
                                        //	strBuilderoper.AppendLine(ot["Код"].ToString());
       ///*                                    		   
                                            	//OperAnalog["SHIFROP"].ToString().Substring(0, 2);
                                        		   	  strBuilderoper.AppendLine(ot["Код"].ToString().IndexOf(OperAnalog["SHIFROP"].ToString().Substring(0, 2)).ToString()+" "+
                                        		   	                          ot["Код"].ToString()+" "+
                                            		   	                        OperAnalog["SHIFROP"].ToString().Substring(0, 2) +" "+
                                                                                ot["Название"].ToString()
                                            		   	                       );
                                                         if (ot["Код"].ToString().IndexOf(OperAnalog["SHIFROP"].ToString().Substring(0, 2))!=-1)
                                                         	{
                                                         	 Объект NewOper=null;
                                                                 if (ot["Название"].ToString().Equals("Прочие"))
                                                                        {
                                                                            NewOper = СоздатьОбъект("Типы технологических операций", "Общего назначения", ot);
                                                                        }
                                                                else
                                                                        {
                                                                            NewOper = СоздатьОбъект("Типы технологических операций", ot["Название"].ToString(), ot);
                                                                        }
                                                                NewOper["Название"]=OperAnalog["NAIMOP"];
                                                                NewOper["Код"]=OperAnalog["SHIFROP"];
                                                                NewOper.Сохранить();                                                         	    
                                                         	}
                                                         
                                                         
                                                              
                                                                strBuilderoper.AppendLine("Добавить"+OperAnalog["SHIFROP"]+" "+OperAnalog["NAIMOP"]);
   // */ 
                                                    }
        
// */
                                    }
                    }
            
            Сообщение("",strBuilderoper.ToString());


      //      Объект документ1 = СоздатьОбъект("Типы технологических операций","Деталь",папка);

        }
      
       public void SetZag(Объект izg)
        {
       	
       	
        
        
     //  	Message("",izg["Обозначение"].ToString().Replace(".", string.Empty));
       	var zag=НайтиОбъект("KAT_RAZM",String.Format("[SHIFR] = '{0}'", izg["Обозначение"].ToString().Replace(".", string.Empty)));
       // Message("",zag["OKP"].ToString());
       	if (zag !=null) 
                {
                       		Объект заготовка = izg.СоздатьОбъектСписка ("Материалы/Заготовки", "Материал-Заготовка");
                                                           заготовка["Обозначение"]=zag["OKP"].ToString();
                                                           заготовка["Наименование"]=zag["NAME"].ToString();
                                                           заготовка["Размеры"]=zag["RAZM"].ToString();
                                                           заготовка["Норма расхода"]=zag["NMAT"].ToString();
                                                           заготовка["Основной"]=true;
                                                           //zag["EDIZM"].ToString();
                                                           //zag["NMAT"].ToString();
                                                           //zag["MASS"].ToString();
                                                          // NMAT
                                                           	//MASS
                                                           //
                                                      //     Сообщение("AddOSn",оснащение2["Строка оснащения"].ToString()+" "+s)    ;
                                                      
                                                      
                                                      
                                                        //var edizDocs=заготовка.СвязанныйОбъект["ЕИ количества"];
                                                       
                                                        
                                                        // НайтиОбъект("KAT_RAZM"
                                                        //[KAT_EDIZ]->[KOD] = '127'
                                                        
                                                        var edizDocs=НайтиОбъект("Единицы измерения",String.Format("[KAT_EDIZ]->[KOD] = '{0}'", zag["EDIZM"].ToString()));
                                                        if (edizDocs != null) 
                                                                заготовка.СвязанныйОбъект["ЕИ количества"]=edizDocs;
                                                        //заготовка.СвязанныйОбъект["KAT_EDIZ"]=eedizDocs.СвязанныйОбъект["KAT_EDIZ"];
                                                        var material=НайтиОбъект("Материалы",String.Format("[Код / обозначение] = '{0}'", zag["OKP"].ToString()));
                                                                if (material != null)  
                                                                        заготовка.СвязанныйОбъект["Материал"]=material;
                                                        
                                                        
                                                        var edizFox= edizDocs.СвязанныйОбъект["KAT_EDIZ"];
                                                       // var material=заготовка.СвязанныеОбъекты["Материал"];
                                                   //     Message("",edizDocs.СвязанныйОбъект["KAT_EDIZ"]["KOD"].ToString());
                                                //        Message("",edizDocs.ToString());
                                                 //       Message("",material.ToString());
                                                        
                                                        
                                                        заготовка.Сохранить();
                                        }
                                      
       	}
       
       
       
       public void SetIzmM(Объект izg)
       	    {
       	Объект izv=null;
       	var izmmfox=НайтиОбъект("KAT_IZVM",String.Format("[SHIFR] = '{0}'", izg["Обозначение"].ToString().Replace(".","") ));
        if    (izmmfox!=null)     
        	{
        	Объект izmm=СоздатьОбъект("Изменения","Технологическое изменение");
        	
   
            izmm["Внедрить"]=izmmfox["VNEDR"];
        	izmm["№ изменения"]=izmmfox["N_IZV"];
     
        	if (izmmfox["SH_IZM"]!="")
        		{
        	       izv=СоздатьОбъект("Извещения об изменениях","Извещение об изменении");
        	       izv["Обозначение"]=izmmfox["SH_IZM"];
        	       izv.Сохранить();
        	       DesktopObject izvd = (DesktopObject)izv;
                   if (izvd.CanCheckIn)
                   Desktop.CheckIn(izvd, "+", false);

        		}
        	
        	izv.Сохранить();

            izmm.Сохранить();
            
        
            izg.Изменить();
            izg.Подключить("Изменения",izmm);
            izg.Сохранить();

          
          }


//       	izg
       }
       
      
     /* 
      public void GetZag(Объект izg)
        {
        Razm_out zagrazm=new Razm_out();
        
        
        
        var списокматериалов = izg.СвязанныеОбъекты["Материалы/Заготовки"];
         
        //Сообщение("{}",String.Format("{0}  {1} {2} {3} ",zag[0].ToString(),zag[0]["Обозначение"].ToString(),  zag[0]["Размеры"].ToString()  ));
        
        
         Сообщение("",String.Format("{0}  {1}  {2}",списокматериалов[0].ToString(),
                                            списокматериалов[0]["Обозначение"].ToString(),
                                            списокматериалов[0]["Размеры"].ToString() )
                                    );
        
        
           foreach (var zagitem in списокматериалов)
               {
                  Сообщение("",String.Format("{0} {1}",zagitem.Тип,zagitem["Основной"]));
               if (zagitem.Тип=="Материал-Заготовка" && zagitem["Основной"])
                       {
                   var edizDocs=zagitem.СвязанныйОбъект["ЕИ количества"];
                   var edizFox= edizDocs.СвязанныйОбъект["KAT_EDIZ"]["KOD"].ToString();
                   var material=zagitem.СвязанныеОбъекты["Материал"];
                   
                   Сообщение("ediz",edizFox);
                   Сообщение("mat",material.Count.ToString());
                   
                           zagrazm.EDIZM=edizFox;
                        //zagrazm.GOST
                        //zagrazm.MARKA
                        //zagrazm.MAS
                        zagrazm.NAME=izg["Наименование"].ToString();
                        //zagrazm.NMAT
                        //zagrazm.NOTH
                        //zagrazm.NPOT
                        //zagrazm.NVES
                        zagrazm.OKP=zagitem["Обозначение"].ToString();
                        //zagrazm.POKR_H
                        //zagrazm.POKR_S
                        zagrazm.RAZM=zagitem["Размеры"].ToString();
                     
                        zagrazm.SHIFR=izg["Обозначение"].ToString();
                        //zagrazm.VID
                        Сообщение("{}",String.Format("{0}  {1} {2}",zagitem.ToString(),zagitem["Обозначение"].ToString(),
                                                                       zagitem["Размеры"].ToString()  ));
                       }
               }
           
           Save_kat_razm_out(zagrazm);
           Export_kat_razm_out();
        }
        
        */
      

      
}
