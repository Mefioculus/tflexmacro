using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using TFlex.DOCs.Model.Macros;
using TFlex.DOCs.Model.Macros.ObjectModel;
using System.Text.RegularExpressions;
using TFlex.DOCs.Model.References.Files;
//--------
using TFlex.DOCs.Model.Search;
using TFlex.DOCs.Model;
using TFlex.DOCs.Model.References;
using TFlex.DOCs.Model.Classes;
using TFlex.DOCs.Model.Structure;
using System.Text;


public class Macro : MacroProvider
{
    public Macro(MacroContext context)
        : base(context)
    {
    }

    
    public static class Guids
        {

            public static class Parameters
            {
          
                public static readonly Guid Number = new Guid("9e1a8c39-ec48-489d-b6d3-eec2a34f7dbb");
                public static readonly Guid NumberId = new Guid("b8fbcd2b-8910-4684-ae5f-61ba1b051995");
            }

        }
    
    
    public override void Run()
    {
    
    }
    
           
                	
                         
  public  void Test()
                        {
                	   string value1;
                	   string value2;
                          Объект документ = ТекущийОбъект;
                          string scan=документ["Скан документа"].ToString();

                          //Message("",документ["ID"].ToString());
                        if (scan.ToLower()!=@"\\mainserver\ogtmk\" && scan!="")
                             TryPath(scan,документ["Номер"].ToString());

//\\mainserver\ogtMK\MK\8А7\8А7779\8A7779090
             //\\Mainserver\ogtmk\MK\8A7\8А7732\8А7732272                      
                /*             
                        if (scan.ToLower().StartsWith(@"\\mainserver\ogtmk\mk\8"))
                    {
                    //Console.WriteLine("8a index "+ scan.ToLower().IndexOf(@"ogtmk\mk\8а"));
                    //Console.WriteLine("text " + scan.Substring(scan.ToLower().IndexOf(@"ogtmk\mk\8а") +9));
                    //value1 = scan.Substring(scan.ToLower().IndexOf(@"ogtmk\mk\8а") + 9);
                    //Console.WriteLine(@"\\mainserver\ogtMK\MK\" + value1.Replace("8A","8А"));
                    
                    value1 = scan.Substring(scan.ToLower().IndexOf(@"ogtmk\mk\8") + 9);
                    документ.Изменить();                                                
                    документ["Скан документа"]=(@"\\mainserver\ogtMK\MK\" + value1.Replace("8A","8А"));
                    документ.Сохранить();
                }
           / */
                        
                         
/*

                	 string pattern = @"УЯИС[\\]УЯИС[.]\d*[\\]УЯИС[.]\d*[\\]";
                    Match m = Regex.Match(scan, pattern);

 
                    if (m.Success)
                    {
                         value2 = scan.Substring(scan.IndexOf(@"\УЯИС\УЯИС.") + 6, 8);
                            документ.Изменить();
                            документ["Скан документа"] =(scan.Replace(value2 + "\\" + value2, value2)) ;
                            документ.Сохранить();
                            //value1
                            Message(" ",(scan.Replace(value2 + "\\" + value2, value2)));
                    }                       
                          	
  */
                        	
                          //    документ.Изменить();
                           //    документ["Номер"] = документ["ID"];
                           //    документ.Сохранить();

                            /*
                        
                            if (scan.ToLower()!=@"\\mainserver\ogtmk\" && scan!="")
                            	 {
                                      
     
                            	
                            	
                                    if (scan.ToLower().StartsWith(@"\\mainserver\ogtmk\tи"))
                                  	     {
                                            value1 = scan.Substring(0, scan.IndexOf(@"TИ\")+2).ToLower();
                                            value2 = scan.Substring(scan.IndexOf(@"TИ\") + 3,scan.Length-(scan.IndexOf(@"TИ\") + 3));
                                  
                                            //Сообщение("Файл найден",value1.Replace(@"\\mainserver\ogtmk\tи", @"\\mainserver\ogtMK\ТИ\")+value2);
                                            документ.Изменить();
                                  	        документ["Скан документа"] = value1.Replace(@"\\mainserver\ogtmk\tи", @"\\mainserver\ogtMK\ТИ\")+value2;
                                  	        документ.Сохранить();
                                  	      }
                                    
                                    
                                    if (scan.ToLower().StartsWith(@"\\mainserver\ogtmk\ти\8a."))
                                            {
                                                value1 = scan.Substring(scan.ToLower().IndexOf(@"ogtmk\ти\8a.") + 9);
                                                документ.Изменить();                                                
                                                документ["Скан документа"]=(@"\\mainserver\ogtMK\ТИ\" + value1.Replace("8A","8А"));
                                                документ.Сохранить();
                                             }
                                    
                                    
                                    
                              
                                    if ((scan.ToLower().StartsWith(@"\\mainserver\ogtmk\mk\уяис\уяис.")))
                                       {
                                            value1 = scan.Substring(scan.ToLower().IndexOf(@"уяис\уяис.") + 5);
                                            документ.Изменить();
                                  	        документ["Скан документа"] = @"\\mainserver\ogtmk\MK\УЯИС\"+value1.Substring(0, 8)+"\\"+ value1;
                                  	        документ.Сохранить();
                                            //Message("TEST",(@"\\mainserver\ogtmk\MK\УЯИС\"+value1.Substring(0, 8)+"\\"+ value1));
                                        }
                                    
                                    
                                    
                                      }   
                                    
                                     */	


                                                    /*                                 
                                    if ((scan.ToLower().StartsWith(@"\\mainserver\ogtmk\mk\уяис\уяис.") && scan.Length == 43))
                                           

                                            {
                                                string pattern = @"УЯИС[.]\d{3}.{8}$";

               
                
                                            Match m = Regex.Match(scan, pattern);
               
                                            value1=(scan.Substring(0, scan.IndexOf(@"УЯИС\") + 5) +
                                            m.Value.Substring(0, 8) + "\\" +
                                            m.Value);                               
                                    	    документ.Изменить();
                                  	        документ["Скан документа"] = value1;
                                  	        документ.Сохранить();

                                            //Сообщение("Файл найден",value1);
                                                
                                            }
                                 
                                        */

                                    
                                    
                                    	/*
                                    {
                                            string pattern = @"УЯИС\d{3}[.]\d{3}[.]\d{3}";
                                            Regex regex = new Regex(pattern);
                                            Match m = Regex.Match(scan, pattern);
                                            value1= (
                                            scan.Substring(0, scan.IndexOf(@"УЯИС\") + 5) +
                                            (m.Value.Substring(0, 4) + '.' + m.Value.Substring(4, 3)) + "\\" +
                                            (m.Value.Substring(0, 4) + '.' + m.Value.Substring(4, 3)) +
                                             m.Value.Substring(7, m.Length - 7)
                                                    );
                                            Сообщение("Файл найден",value1);
                                    
                                    
                                         }
                                         */    
                                        
                                        
                                        

                                   	 

 }
    	
  public  void Test2()
  	{
 Объект файл = ТекущийОбъект;
    Message("",файл["Наименование"].ToString());
    
      }
                                  
                public  void Izm()
                        {
                	        Объект файл = ТекущийОбъект;
                	        Объекты файлы2 =НайтиОбъекты("Изменения ОГТ","[Файл]->[Наименование] = '"+файл["Наименование"].ToString()+"' И [Тип документа] = 'МК'");
                	        string filename=файл["Наименование"].ToString();
                	        filename = filename.Substring(0, filename.IndexOf(".pdf"));
                             if (filename.StartsWith("8A"))
                                                          filename = filename.Replace("8A", "8А");
                               
                               
                	        string version="";
                	      //  Message("",filename);
                	        if (файлы2.Count==0)
                	         
                	         	{
                    	         //	Сообщение("Связь","Нет объектов");
                	           
                    	         	
                            
                                      
                            
                            
                                     if (filename.IndexOf('(')>0)
                                     	{
                                     	//Message("",filename.Split('.')[0].IndexOf(')').ToString());
                                       // Message("",filename.Split('.')[0].IndexOf('(').ToString());
                                        version=filename.Substring(filename.IndexOf('(')+1,  
                                                 filename.IndexOf(')') - filename.IndexOf('(') -1 );
                                         filename=filename.Substring(0,filename.IndexOf('('));
                                         
                                        // version=filename.Substring(filename.Split('.')[0].IndexOf('(') + 1,
                                       //          filename.Split('.')[0].IndexOf(')') - (filename.Split('.')[0].IndexOf('(') + 1));
                                           
                                                   
                                         }
                                     
                                       //  Сообщение("Файл1",filename);
                            
                            
                            //"Удаляем точки "+filename.Replace(".", string.Empty )
                                    //[Обозначение детали, узла] содержит '8А7774083' И [Тип документа] = 'МК' И [Шифр извещения] содержит текст
                                    Объект документОГТ = НайтиОбъект("Документы ОГТ","[Обозначение детали, узла] = '"+filename.Replace(".", string.Empty )+"' И [Тип документа] = 'МК'");
                                  //  Сообщение("Файл1",документОГТ.ToString());
                            
                            
                            
                                    Объект измененияОГТ = СоздатьОбъект("Изменения ОГТ","Изменения ОГТ");
        	                        измененияОГТ["Номер"]=документОГТ["Номер"];
        	                        измененияОГТ["Тип документа"]="МК";
  
        	                        if (version != "")
                        	                    измененияОГТ["Номер изменения"]=Convert.ToInt32(version); 

                                 
        	                        
          	                        
          	                        //измененияОГТ.СвязанныйОбъект["Файл"]=файл;
          	                        
          	                        измененияОГТ.Подключить("Файл",файл);
          	                        //Подключить("Документ ОГТ", документОГТ)
                                    измененияОГТ.Сохранить();
          	                        
                                  
                            
                                 }
  
                            	

                            
      
                    	
                         }            
                                  
            
                                         
               public void СвязьДокументОГТ_ИзмененияОГТ()
                    {
               	            Объект документОГТ = ТекущийОбъект;
               	      //      Message("dd1",документОГТ["Номер"]);
               	            
               	           Объекты измОГТ = НайтиОбъекты("Изменения ОГТ","[Номер] = '"+документОГТ["Номер"]+"'");
                  //         Message("dd1",измОГТ.Count.ToString());
                           
                            документОГТ.Изменить();
                            //документОГТ.СвязанныйОбъект["Изменения"]=измененияОГТ;
                            foreach (Объект объект in измОГТ)
                            	документОГТ.Подключить("Изменения", объект);
                            	//СвязанныеОбъекты["Изменения"]=измОГТ;
                            документОГТ.Сохранить();
                     }               	
                            
               
                            public void СвязьИзмененияОГТ_ДокументОГТ()
                    {
               	            Объект ИзмененияОГТ = ТекущийОбъект;
               	            
               	         //   Message("dd1",ИзмененияОГТ["Номер"]);
               	            
                   	        Объект документОГТ = НайтиОбъект("Документы ОГТ","[Номер] = '"+ИзмененияОГТ["Номер"]+"'");
                           
                            ИзмененияОГТ.Изменить();
                            //документОГТ.СвязанныйОбъект["Изменения"]=измененияОГТ;
                            ИзмененияОГТ.Подключить("Документ ОГТ", документОГТ);
                            //ИзмененияОГТ.СвязанныйОбъект["Документ ОГТ"]=документОГТ;
                           
                            ИзмененияОГТ.Сохранить();
                     }  
                                	                   	
                    	
             
                public void Создать_Изменения()
                	{
                //	Объект user = ТекущийПользователь;
                //	Message("",user.ToString());
                	// Исходная папка в файловой системе - из нее будут копироваться файлы
                	// 41af4b79-6737-485a-a707-15ced9e18a96 GuID папка User7
                  
                      string FolderPath="";                    
                        const string FolderMK = "41af4b79-6737-485a-a707-15ced9e18a96"; // папка User7 
                      //const string FolderMK = "342b20d6-3578-4eb7-98e8-31f6c7d74392";  // папка Старые версии МК  
                      	//"342b20d6-3578-4eb7-98e8-31f6c7d74392";

                        Объект ДокументОГТ = ТекущийОбъект;
                     //   Message("",ДокументОГТ.ToString());
                    //    Message("",ДокументОГТ["Номер"].ToString());
                    //    Message("",ДокументОГТ["Шифр извещения"].ToString());
                    //    Message("",ДокументОГТ["Тип документа"].ToString());

                       
                                    Объект измененияОГТ = СоздатьОбъект("Изменения ОГТ","Изменения ОГТ");
        	                        

                                    измененияОГТ["Номер"]=ДокументОГТ["Номер"];
        	                        измененияОГТ["Тип документа"]=ДокументОГТ["Тип документа"].ToString();
                                    измененияОГТ["Шифр"]=ДокументОГТ["Шифр извещения"].ToString();
                                    измененияОГТ.Сохранить();
                                    
                                    ДокументОГТ.Изменить();
                                    ДокументОГТ.Подключить("Изменения", измененияОГТ);
                                    ДокументОГТ.Сохранить();
               
                    }

 
    
                public  void Linkcopy()
                        {
                	
                	
                            Объект документ = ТекущийОбъект;
                            string scan=документ["Скан документа"].ToString();
                      
                            if (scan.ToLower()!=@"\\mainserver\ogtmk\" && scan!="")
                            	{

                                  
                                   
                                       string value1=scan.Substring(2, (scan.Length-2)).Replace(scan.Substring(2, 16), "Архив ОГТ");
                                       string filename=value1.Split('\\').Last();
                                       value1=value1+"\\"+filename+".pdf";
                                       Сообщение("val",value1);
                                       Сообщение("val2",filename);
                                  
                                       //Архив ОГТ\MK\8А6\8А6123\8А6123605\8А6123605.pdf
            
                                  
            
                                  //        Сообщение("путь",value1);
                                           
                                        Объект файлы = НайтиОбъект("Файлы","[Относительный путь] содержит '"+filename+".pdf'");
                                        
                
                                           Сообщение("Файл найден",файлы.ToString());
                                 
                                 	документ.Изменить();
                                   	документ.СвязанныйОбъект["Скан документа1"] = файлы;
                                   	документ.Сохранить();
                                 }       	                   	
                    	
                         }
                
                 public  void StrLower()
                          {

                          Объект документ = ТекущийОбъект;
                          документ.Изменить();
                          документ["Скан документа"]=документ["Скан документа"].ToString().ToLower();
                          документ.Сохранить();
                        
                           }   
                 
                 
                 //@"\\DESKTOP-0G59NHM\"
   public void TryPath(string path,string id)
               	{
                        string dirName = @path;
                        string writePath = @"C:\trypath\path.txt";
                        string text = "";
                        if (!Directory.Exists(@"C:\trypath"))
                            Directory.CreateDirectory(@"C:\trypath");


                        if (!Directory.Exists(dirName))
                             text ="отсутствует "+id+" "+dirName;
                        else
                           text =dirName;


                     try
                        {
                            using (StreamWriter sw = new StreamWriter(writePath, true, System.Text.Encoding.Default))
                        {
                            sw.WriteLine(text);
                        }

                        }
                    catch (Exception e)
                        {
                            Console.WriteLine(e.Message);
                        }
    
    
                }
                
                
     public void Replace()
        {
     	
     	Объект документОГТ=ТекущийОбъект;
     	
             if (документОГТ["Обозначение детали, узла"].ToString().StartsWith("8A"))
                {
             	документОГТ.Изменить();
                документОГТ["Обозначение детали, узла"] = документОГТ["Обозначение детали, узла"].ToString().Replace("8A", "8А");
                документОГТ.Сохранить();
                }
        }
     
     
      public void TrimString()
        {
     	
     	Объект текущий=ТекущийОбъект;
     	
             текущий.Изменить();
             текущий["Обозначение"]=текущий["Обозначение"].ToString().Trim();
             текущий.Сохранить();
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

            
          public void changeOsnNumAll()
          {
          	Message("","");
      string referenceGuid = "8d727772-d7e5-4058-b7e1-046c510e7f76";//заказы на оснастку
     
      ReferenceInfo info = Context.Connection.ReferenceCatalog.Find(referenceGuid);
      
      Reference reference = info.CreateReference();
      ReferenceObjectCollection listObj = reference.Objects;
      
      foreach (var cobj in listObj)
                    {
                        var changingObject = cobj;

                        if (!changingObject.CanEdit)
                            return;

                        // Перевод объекта в состояние редактирования
                        changingObject.BeginChanges();

                        try
                        {

                            var number = changingObject.GetObjectValue("Номер").Value.ToString();

                            cobj[Guids.Parameters.NumberId].Value = changing(number);
                            changingObject.EndChanges();

                        }

                        catch
                        {
                            Console.WriteLine("Ошибка изменения");
                        }
                    }
          	
          }
      
         public void changeNumOsn()
    	{
           var currentObj= ТекущийОбъект;
            currentObj.Изменить();
            currentObj["NumberId"]=changing(currentObj["Номер"].ToString());
            currentObj.Сохранить();	
        }
         
         
             public  void LinkcopyOboz()
                        {
                    
                    
                            Объект документ = ТекущийОбъект;
                            string Oboz=документ["Обозначение детали, узла"].ToString();
                      
                          
                            // if (scan.ToLower()!=@"\\mainserver\ogtmk\" && scan!="")
                           //   if (scan.ToLower()!=@"\\mainserver\ogtmk\" && scan!="")
                                

                                  
                                   
                                 //      string value1=scan.Substring(2, (scan.Length-2)).Replace(scan.Substring(2, 16), "Архив ОГТ");
                               //        string filename=value1.Split('\\').Last();
                                     //  value1=value1+"\\"+filename+".pdf";
                                     //  Сообщение("val",value1);
                                    //   filename=filename.Replace(".","");
                                    //   Сообщение("val2",filename);
                                  
                                       //Архив ОГТ\MK\8А6\8А6123\8А6123605\8А6123605.pdf
            
                                  
            
                                  //        Сообщение("путь",value1);
                                        try
                                        {                                  
                                        Объект файлы = НайтиОбъект("Файлы","[Относительный путь] содержит '"+Oboz+".pdf'");
                                        документ.Изменить();
                                           документ.СвязанныйОбъект["Скан документа1"] = файлы;
                                           документ.Сохранить();
                                        }
                                        catch(Exception e)
                                        {
                                            Сообщение("Файл найден",String.Format("{e.ToString} {файлы.ToString()}"));
                                        }
                                     
                                                                   
                        
                         }
      
           	                   	
    
}
