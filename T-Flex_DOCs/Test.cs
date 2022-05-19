using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using TFlex.DOCs.Model.Macros;
using TFlex.DOCs.Model.Macros.ObjectModel;
using TFlex.DOCs.Model.References.Files;
using TFlex.DOCs.Model.References;
using System.Text;

public class Macro : MacroProvider
{
	
	
	public  List<Объект> DCE = new List<Объект>();
	        private static class Guids
        {
            public static class References
            {
                public static readonly Guid Изменения = new Guid("c9a4bb1b-cacb-4f2d-b61a-265f1bfc7fb9");
                public static readonly Guid ИнвентарнаяКнига = new Guid("f1d28c5d-6b08-4061-b81f-cde490088e1f");
            }

            public static class Classes
            {
                public static readonly Guid ТехнологическийПроцесс = new Guid("3e93d599-c214-48c8-854f-efe4b475c4d8");
                public static readonly Guid ТехнологическаяОперация = new Guid("f53c9d73-18bb-4c59-a260-61fea65f6ed9");
                public static readonly Guid КарточкаУчетаТехнологическихДокументов = new Guid("d2297aab-a159-45bf-8601-7a7f1f27a38c");
                public static readonly Guid ИзвещениеОбИзменении = new Guid("52ccb35c-67c5-4b82-af4f-e8ceac4e8d02");
                public static readonly Guid ТехнологическоеИзменение = new Guid("a8c61089-dc59-43c8-9296-f2629fec2c4e");
                public static readonly Guid ТехнологическийКомплект = new Guid("dc1cf2a0-6c01-400d-9a42-9642b7496404");
            }
            
           }
	
	
    public Macro(MacroContext context)
        : base(context)
    {
    }

    public override void Run()
    {
    
    }
    
    public  void Msg1()
    {
    Сообщение("Событие 1","Событие 1");
    }
    
    
    public  void Msg2()
    {
    Сообщение("Событие 2","Событие 2");
    }
    
    public  void Msg3()
    {
    Сообщение("Событие 3","Событие 3");
    }
    
    public  void Msg4()
    {
    Сообщение("Событие 4","Событие 4");
    }
    
    public  void Msg5()
    {
    Сообщение("Событие 5","Событие 5");
    }
    
    public  void Msg6()
    {
    Сообщение("Событие 6","Событие 6");
    }
    
    public  void Msg7()
    {
    Сообщение("Событие 7","Событие 7");
    }
    
    
    public  void Msg8()
    {
    Сообщение("Событие 8","Событие 8");
    }
    
    
        public  void Operall()
    {
        	
        	 var текущий= ТекущийОбъект;
       //     Сообщение("",текущий.ToString());
       //     Сообщение("",текущий.Тип.ToString());
       //     Сообщение("",текущий.ДочерниеОбъекты.Count.ToString());
            Объекты цехозаходы= текущий.ДочерниеОбъекты;
            
    Сообщение("Событие 8","Событие 8");
    }
    
    
    
    public void Link()
    	{
            Объект файл= ТекущийОбъект;
            Сообщение("",файл.ToString());
            //Сообщение("",файл.СвязанныйОбъект["Чертеж детали"].ToString());
            
         FileObject file = (ReferenceObject)файл as FileObject;
    	file.GetHeadRevision();
    	string filePath = file.LocalPath;
        Сообщение("",filePath);
        }
    
    public void Save_Sing()
    	{
    	 Объект user=ТекущийОбъект;
 
        if (user["Изображение подписи"].ToString()!="")
        	{
 
 	          filewrite(user);
 	        }
        }
    
     public void filewrite(Объект user)

        {
    	
    	   byte[] array=user["Изображение подписи"];
            // создаем каталог для файла
            string path = @"C:\SomeDir2";
            DirectoryInfo dirInfo = new DirectoryInfo(path);
            if (!dirInfo.Exists)
            {
                dirInfo.Create();
            }
        //    Console.WriteLine("Введите строку для записи в файл:");
           // string text = "EE";

            // запись в файл
           // Сообщение("","C:\\SomeDir2\\"+user["login"]);
            using (FileStream fstream = new FileStream("C:\\SomeDir2\\"+user["Логин"].ToString()+".png", FileMode.OpenOrCreate))
            {
                // преобразуем строку в байты
              //  byte[] array = System.Text.Encoding.Default.GetBytes(text);
                // запись массива байтов в файл
                fstream.Write(array, 0, array.Length);
            //    Console.WriteLine("Текст записан в файл");
            }
           }
     
     public void FIORUKOV()
     	{
          Объект user=НайтиОбъект("Группы и пользователи","[Сокращённое название] = 'ОТиЗ'");          
           Сообщение("",user.СвязанныйОбъект["Руководитель"].ToString());
           Сообщение("",user.СвязанныйОбъект["Руководитель"]["Короткое имя"].ToString());
           Сообщение("",НайтиОбъект("Группы и пользователи","[Сокращённое название] = 'ОТиЗ'").СвязанныйОбъект["Руководитель"]["Короткое имя"].ToString());
     }
     
     
     public void tip()
     	{
     	ReferenceObject currentObj = Context.ReferenceObject;
           Объект текущий = ТекущийОбъект;
           Сообщение("1",текущий.ToString());
           Сообщение("2",текущий.Тип.ToString());
           Сообщение("3",currentObj.ToString());
           Сообщение("4",currentObj.GetType().ToString());
           Сообщение("4",currentObj.Class.ToString());
           if (!currentObj.Class.IsInherit(Guids.Classes.ТехнологическийПроцесс))
                Error(string.Format("Текущий объект '{0}' не является технологическим процессом.", currentObj));
          // if (!currentTP.Class.IsInherit(Guids.Classes.ТехнологическийПроцесс))
     }
     
     
     public void Mess(Объекты объекты)
     	{
     	var s="01010";
        //    Message("","");
        //    System.Windows.Forms.MessageBox.Show("TEXT");
            StringBuilder strbild= new StringBuilder();
            strbild.AppendLine("1");
            strbild.AppendLine("2");
            strbild.AppendLine("3");
            strbild.AppendLine("4");
            
             foreach(Объект объект in объекты)
          	{
           	            strbild.AppendLine(String.Format("Объект {0}",объект.ToString()));
           	            strbild.AppendLine(String.Format("Тип объекта {0}",объект.Тип.ToString()));
          }
            
      //  Ошибка(String.Format("-------{0}",объекты.Count.ToString()));
        Ошибка(String.Format("-------{0}",strbild.ToString()    ));
            
           
         }
     
          
      public void Подписи()
      	{
      	
      	 Объект текущий = ТекущийОбъект;
      	string namesing="Начальник цеха";
      	//Подписи sings = текущий.Подписи["Разраб."];
      	foreach(Подпись подп in текущий.Подписи)
      		if 	(подп.ТипПодписи=="Разраб.")
      		Сообщение("4",подп.Пользователь["Короткое имя"].ToString());
          }
     
     
  /*   
     
      public void ПринятьНаХранениеТП(ReferenceObject currentObj)
        {
		
		   if (currentObj == null)
                currentObj = Context.ReferenceObject;
            ReferenceObject currentTP = currentObj;
            if (!currentTP.Class.IsInherit(Guids.Classes.ТехнологическийПроцесс))
                Error(string.Format("Текущий объект '{0}' не является технологическим процессом.", currentTP));

            var rootDocuments = currentTP.GetObjects(Guids.Links.ТПДокументация);
            if (!rootDocuments.Any())
                Error(string.Format("Технологический процесс '{0}' не содержит комплект документов.", currentTP));

            List<ReferenceObject> allTechnologyDocuments = new List<ReferenceObject>();
            allTechnologyDocuments.AddRange(rootDocuments);
            foreach (var rootDocument in rootDocuments)
            {
                var childDocuments = rootDocument.Children.ToList();
                if (childDocuments.Any())
                    allTechnologyDocuments.AddRange(childDocuments);
            }

            // комплект документов
            var setOfDocuments = rootDocuments.FirstOrDefault(document => document.Class.IsInherit(Guids.Classes.ТехнологическийКомплект));
            // проверка на заполнение параметров
            string checkResult = CheckTD(setOfDocuments);

            if (!Question(string.Format("Принять на хранение?{0}{0}{1}", Environment.NewLine, checkResult)))
                return;

            //Поиск учетной карточки для технологического комплекта документов
            ReferenceObject inventoryCard = GetInventoryCard(setOfDocuments);
            // заполнить данные инвентарной карточки
            FillInventoryCard(inventoryCard);
            if (!ShowPropertyDialog(RefObj.CreateInstance(inventoryCard, Context)))
                return;

            // все объекты, у которых меняем стадию на хранение
            List<ReferenceObject> storageObjects = new List<ReferenceObject>();

            // обрабатываем файлы комплекта документов
            if (!DoActionWithSetOfDocument(setOfDocuments, ref storageObjects))
                return;

            var operations = currentTP.Children.ToList();
            operations.ForEach(operation => storageObjects.AddRange(operation.Children.ToList()));

            storageObjects.Add(currentTP);
            storageObjects.AddRange(operations);
            storageObjects.AddRange(allTechnologyDocuments);

            try
            {
                //Перевод на стадию "Хранение" ТП и Комплект документов
                ChangeStage(storageObjects, "Хранение");
            }
            catch
            {
                //%%TODO записать в лог?
                //LogWriter
            }
            currentTP.Reload();
        }
*/
       public  void testZag()
    {
        	
        	 var ТП= ТекущийОбъект;
           Сообщение("",ТП.ToString());
           Сообщение("",ТП.Тип.ToString());
           var материалы = ТП.СвязанныеОбъекты["8f505469-c1b4-4caf-a7c1-3bc7ea8c2bbe"];
           Сообщение("",материалы[0].ToString());
       //     Сообщение("",текущий.Тип.ToString());
       //     Сообщение("",текущий.ДочерниеОбъекты.Count.ToString());
       //     Объекты цехозаходы= текущий.ДочерниеОбъекты;
            
  //  Сообщение("Событие 8","Событие 8");
    }
       
       
       
      public  void testKK()
    {
        	
           var ТП= ТекущийОбъект;
     //      Сообщение("",ТП.ToString());
     //      Сообщение("",ТП.Тип.ToString());
           Объекты цехозаходы= ТП.ДочерниеОбъекты;
           
       //    Сообщение("Цехозаходы",цехозаходы[0].Тип.ToString());
   //        var операции= цехозаходы[0].ДочерниеОбъекты;
     //      Сообщение("Опер",операции[0].ToString());
           
         //  var комплектующие = операции[0].СвязанныеОбъекты["90242708-0a8a-4897-939e-505d09e559bf"];
        //   Сообщение("комплектующие",комплектующие[0].ToString());
           
            StringBuilder strbild= new StringBuilder();
            
          //  strbild.AppendLine("1");
            
///*        

            foreach (var item in цехозаходы)
            	{
                  var операции= item.ДочерниеОбъекты;
                   foreach (var опер in операции)
                   	    {
                   	        var комплектующие = опер.СвязанныеОбъекты["90242708-0a8a-4897-939e-505d09e559bf"];
                   	        
                            if (комплектующие!=null) 
                                {
                            	   foreach (var компл in комплектующие)
                            	       strbild.AppendLine(опер.ToString()+" "+компл.ToString());
                            	}
                   	    }
                }
                
 //*/
            
             Сообщение("Опер",String.Format("{0}",strbild.ToString()));
            
           //var материалы = ТП.СвязанныеОбъекты["8f505469-c1b4-4caf-a7c1-3bc7ea8c2bbe"];
           //Сообщение("",материалы[0].ToString());
       //     Сообщение("",текущий.Тип.ToString());
       //     Сообщение("",текущий.ДочерниеОбъекты.Count.ToString());
       //     Объекты цехозаходы= текущий.ДочерниеОбъекты;
            
  //  Сообщение("Событие 8","Событие 8");
    }
      
      
      
      
     public void Struct()
          {
          String shifr="УЯИС.794344.073"; //"УЯИС.711351.083";
         // [Обозначение] = 'УЯИС.521473.002'
           Сообщение("","[Обозначение] = '"+shifr+"'");
        var деталь=НайтиОбъект("Электронная структура изделий","[Обозначение] = '"+shifr+"'"); 
          Сообщение("",деталь.ToString());
            //Объект деталь = ТекущийОбъект;
         //   var objs = деталь.ChildObjects;
         //var objs = num.ChildObjects;
            DCE.Add(деталь);  
            //  Рекурсивно получаем все дочерние объекты        
            RecursStruct(деталь);
      
          
          
           foreach (var ТекущийОбъект in DCE)
                              {
               
               	               //if (ТекущийОбъект.Тип.ПорожденОт("Стандартное изделие") || ТекущийОбъект.Тип.ПорожденОт("Прочее изделие")  || ТекущийОбъект.Тип.ПорожденОт("Электронный компонент") || (ТекущийОбъект.Тип.ПорожденОт("Другое"))
               if (ТекущийОбъект.Тип.ПорожденОт("Другое"))
                   {
               ТекущийОбъект.Изменить();
               ТекущийОбъект["Объект"]=ТекущийОбъект["Наименование"];
               ТекущийОбъект.Сохранить();
               }
               
     //Разработка
     //Хранение
                          //      item.ИзменитьСтадию("Разработка", true);
                            
                              }
                  
        

        }
        
        
 
           
        
        
         
         
               public void RecursStruct(Объект текущий)
        {
            
            
                if (текущий.ДочерниеОбъекты.Count != 0)
                    {
                        foreach (var подкл in текущий.ДочерниеОбъекты)
                            {       
                                    RecursStruct(подкл);  
                                    DCE.Add(подкл);
                                 
                            }
                       }
         }
               
               
               
          public void create_link_Material_KLASM()
    {
         // [Код / обозначение]	
          	
        Объект Объект = ТекущийОбъект;
        Объекты Объекты=ВыбранныеОбъекты;
      //  Сообщение("",Объект["OKP"].ToString());
        Объект material = НайтиОбъект("Материалы", "[Код / обозначение] = '"+Объект["OKP"].ToString()+"'");
       // Сообщение("",material["Код / обозначение"].ToString());
        Объект.Изменить();
        //Сообщение("",material.СвязанныеОбъекты["Аналоги материала"].Count.ToString());
        //Сообщение("",Объект.СвязанныйОбъект["Материал"].ToString());
        Объект.СвязанныйОбъект["Материал"]=material;
        //material.СвязанныеОбъекты["Аналоги материала"]=Объекты;
        Объект.Сохранить();
        
    }
          
          
          
          
          
                    public void Material_KLASM()
    {
         // [Код / обозначение]    
              
        Объект Объект = ТекущийОбъект;
        
      //  Сообщение("",Объект.СвязанныеОбъекты["Аналоги материала"].Count.ToString());
      //  Сообщение("",Объект.СвязанныеОбъекты["Аналоги материала"][0].ToString());
     //   Сообщение("",Объект.СвязанныеОбъекты["Аналоги материала"][0]["NAME"].ToString());
     //   Сообщение("",Объект.СвязанныеОбъекты["Аналоги материала"][0]["GOST"].ToString());
        
        string name=Объект.СвязанныеОбъекты["Аналоги материала"][0]["NAME"].ToString().Trim();
        string gost=Объект.СвязанныеОбъекты["Аналоги материала"][0]["GOST"].ToString().Trim();
        Объект.Изменить();
        Объект["Сводное наименование"]=String.Format("{0} {1}",name,gost);
        Объект.Сохранить();
        
       // Объекты Объекты=ВыбранныеОбъекты;
      //  Сообщение("",Объект["OKP"].ToString());
     //   Объект material = НайтиОбъект("Материалы", "[Код / обозначение] = '"+Объект["OKP"].ToString()+"'");
       // Сообщение("",material["Код / обозначение"].ToString());
     //   Объект.Изменить();
    //    Сообщение("",material.СвязанныеОбъекты["Аналоги материала"].Count.ToString());
        //Сообщение("",Объект.СвязанныйОбъект["Материал"].ToString());
      //  Объект.СвязанныйОбъект["Материал"]=material;
        //material.СвязанныеОбъекты["Аналоги материала"]=Объекты;
      //  Объект.Сохранить();
        
    }
    
                    
                    
                    
                          
                    public void Material_KLASM_Nomenclature()
    {
         // [Код / обозначение]    
              
        Объект Объект = ТекущийОбъект;
        string okp= ТекущийОбъект["Обозначение"].ToString();
        Объект material = НайтиОбъект("Материалы", "[Код / обозначение] = '"+okp+"'");
        Объект.Изменить();
        Объект["Наименование"]=String.Format("{0}",material["Сводное наименование"]);
        Объект["Объект"]=String.Format("{0}",material["Сводное наименование"]);
        Объект.Сохранить();
        Сообщение("",Объект["Наименование"].ToString());
      //  Сообщение("",Объект.СвязанныеОбъекты["Аналоги материала"].Count.ToString());
      //  Сообщение("",Объект.СвязанныеОбъекты["Аналоги материала"][0].ToString());
     //   Сообщение("",Объект.СвязанныеОбъекты["Аналоги материала"][0]["NAME"].ToString());
     //   Сообщение("",Объект.СвязанныеОбъекты["Аналоги материала"][0]["GOST"].ToString());
        Сообщение("",material.ToString());
        
        
        //string name=Объект.СвязанныйОбъект["Аналог материала"]["NAME"].ToString().Trim();
        //string gost=Объект.СвязанныйОбъект["Аналог материала"]["GOST"].ToString().Trim();
        //Объект.Изменить();
        //Объект["Наименование"]=String.Format("{0} {1}",name,gost);
        //Объект.Сохранить();
        
       // Объекты Объекты=ВыбранныеОбъекты;
      //  Сообщение("",Объект["OKP"].ToString());
     //   Объект material = НайтиОбъект("Материалы", "[Код / обозначение] = '"+Объект["OKP"].ToString()+"'");
       // Сообщение("",material["Код / обозначение"].ToString());
     //   Объект.Изменить();
    //    Сообщение("",material.СвязанныеОбъекты["Аналоги материала"].Count.ToString());
        //Сообщение("",Объект.СвязанныйОбъект["Материал"].ToString());
      //  Объект.СвязанныйОбъект["Материал"]=material;
        //material.СвязанныеОбъекты["Аналоги материала"]=Объекты;
      //  Объект.Сохранить();
        
    }
                    
                    
                    
                    
       public void Material_KLASM_LINK_EDIZ()

      {
       	Объект klasmobj = ТекущийОбъект;
       	Объект ediz = НайтиОбъект("KAT_EDIZ", "[KOD] = '"+klasmobj["EDIZM"].ToString()+"'");
      //  Сообщение("",String.Format("{0} {1} ",ediz["KOD"].ToString(),klasmobj["EDIZM"].ToString() ));
        
        //Объект klasmobj = ТекущийОбъект
        //klasmobj.СвязанныйОбъект["КАТ_EDIZ"];
        //var link =klasmobj.СвязанныйОбъект["КАТ_EDIZ"].ToString();
        // Сообщение("",link);
        
        klasmobj.Изменить();
        	klasmobj.СвязанныйОбъект["КАТ_EDIZ"]=ediz;
        klasmobj.Сохранить();
        
       }
       
  public void previewFile()
    {  	
           var файл = НайтиОбъект("Файлы", "Относительный путь", "Изображения\\TopSystems.png"); // Находим файл
        if (файл != null)
        ОткрытьОкноПросмотра(файл); // Открываем окно просмотра с указанным файлом
    }
       

}
