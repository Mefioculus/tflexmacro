using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using TFlex.DOCs.Model.Mail;
using TFlex.DOCs.Model.Macros;
using TFlex.DOCs.Model.Macros.ObjectModel;
using TFlex.DOCs.Model.References.Users;
using System.Text;
using TFlex.DOCs.Model.References;
using TFlex.DOCs.Model.References.Documents;
using TFlex.DOCs.Model;
using TFlex.DOCs.Model.Processes.Events.Contexts.Data;




using TFlex.DOCs.Common;

using TFlex.DOCs.Model.Processes.Events.Contexts;

using TFlex.DOCs.Model.References.ActiveActions;

using TFlex.DOCs.Model.References.Nomenclature;

public class Macro : MacroProvider
{
    public Macro(MacroContext context)
        : base(context)
    {
    }
  
    public override void Run()
    {
    
    }
    
    public void ОтправитьЗаданиеИсполнителю ()
    {
    	
        Объект Заказ_Проектирование = ТекущийОбъект;
        Сообщение ("Номер заказа на проектирование", Заказ_Проектирование.СвязанныйОбъект["Заказы на оснастку"].Параметр["Номер"]);
        string заголовок = string.Format("Открыт заказ №{0} на проектирование оснастки", Заказ_Проектирование.СвязанныйОбъект["Заказы на оснастку"].Параметр["Номер"]);
        string текстПисьма = string.Format("Для работы с новым заданием необходимо: \n 1.Нажать на кнопку «Принять» для того, чтобы взять задание в работу. \n 2.Далее, когда оснастка спроектирована и осуществлена выгрузка её состава изделия в T-FLEX DOCs (или состав изделия был создан вручную в справочнике «Электронная структура изделия»), следует открыть карточку справочника «Заказ на проектирование оснастки». Для этого следует в поиске ввести номер заказа и нажать два раза левой кнопкой мыши для открытия документа. \n  3.Далее в открывшемся документе следует привязать объект, соответствующий номеру головного чертежа, нажав на ссылку у параметра «Шифр и наименование оснастки». \n 4.Выбрав нужный чертежный номер, следует сохранить изменения в документе. После этого закрыть задание, нажав на кнопку «Завершить»" );        
        if (Заказ_Проектирование.Параметр["245fe15f-3fec-475b-89eb-583a30582aab"] == true)
        {
        	//Сообщение ("Информация","Макрос запускается");
        	
        	//Сообщение("Информация", Заказ_Проектирование.Параметр["3cef804f-a4da-48f8-a8cd-fc82bc112926"]);
        
        	if (Заказ_Проектирование.Параметр["6537d551-cbd1-4c5f-b09f-9f40e82fddc6"] == "0")
        	{
        		//Сообщение ("Информация1","Статус Новый");
        		 // Создаем новое задание
        MailTask task = new MailTask(Context.Connection)
        {
            Subject = заголовок,
            Body = текстПисьма
        };

        // Добавляем исполнителя задания
           Сообщение ("Пользователь", Заказ_Проектирование.СвязанныйОбъект["Группы и пользователи"].Параметр["Короткое имя"]);
         task.Executors.Add(new MailTaskExecutor((User)Заказ_Проектирование.СвязанныйОбъект["Группы и пользователи"]));
         
       // Сообщение ("Информация6", Заказ_Проектирование.СвязанныйОбъект["Группы и пользователи"]);
        // Добавляем исполнителя задания
       // task.Executors.Add(new MailTaskExecutor(_administrator));

        // Дата начала выполнения задания
        task.StartDate = DateTime.Now.AddDays(1);
        // Срок выполнения задания
        task.EndDate = DateTime.Now.AddDays(2);
        // Контрольный срок задания 
        task.CheckDate = DateTime.Now.AddDays(3);

        // task.Priority = MailTaskPriority.Hight; //Важность задания
    /*  ReferenceObject attachment = (ReferenceObject)Заказ_Проектирование;
        if (attachment != null)
         task.Attachments.Add(new ObjectAttachment(attachment));
        else
        	Сообщение("Ошибка", "Объект для вложения не был найден"); */
        	                
        // Прикрепляем к заданию файл
        // string filePath = @"C:\testFile.grb";
       
     //  ReferenceObject refObj = new DocumentReference(Context.Connection).Objects[0];
     //  ReferenceObject refObj = (ReferenceObject)Заказ_Проектирование;
        
        ReferenceInfo referenceInfo = Context.Connection.ReferenceCatalog.Find("85482ef8-6ed6-42a2-adb7-36db7c9a06ff");
        // Создание объекта справочника для работы с данными
        Reference reference = referenceInfo.CreateReference();
        
        ReferenceObject refObj = reference.Find((int)ТекущийОбъект.Параметр["Id"]);

    /*   if (refObj != null)
            task.Attachments.Add(new ObjectAttachment(refObj));
        else Сообщение("", string.Format("Объект не был найден {0}", (int)ТекущийОбъект.Параметр["Id"]));
        Заказ_Проектирование.Параметр["Guide_Задания"] = task.Guid;
        Сообщение("", Заказ_Проектирование.Параметр["Guide_Задания"].ToString()); */

        // if (File.Exists(filePath))
        //    task.Attachments.Add(new FileAttachment(filePath));
        
        // Отправляем задание
        task.Send();
        
        
      /*  List<MailTask> tasks = Context.Connection.Mail.DOCsAccount.GetTasks().Where(t => t.Status == MailTaskStatus.InProgress).ToList();
        Сообщение("Входящие задания", string.Join(Environment.NewLine, tasks.Select(GetMailItemString))); */

    	   	}
        	else
        	{
        		//Сообщение ("Информация2", "У документа не статус Новый");
        	}
        	Заказ_Проектирование.Параметр["245fe15f-3fec-475b-89eb-583a30582aab"] = false;
        	return;
        }
    	
    }
  /*  private string GetMailItemString(MailItem item)
    {
        StringBuilder sb = new StringBuilder();
        sb.AppendLine("Заголовок: " + item.Subject);
        sb.AppendLine("Дата отправки: " + item.SentDate);
        sb.AppendLine("От: " + item.From);
        sb.AppendLine("Текст: " + (item.Body.Length > 15 ? item.Body.Substring(0, 15) + "..." : item.Body));
              
        return sb.ToString();
    } */
  
    public void ПроверкаСвязи ()
    	{
    	Объект ном = ТекущийОбъект;
    //	Объект изготовление = ном.СвязанныйОбъект["aa724569-dc58-48d0-b030-2364d2f07b59"];
    	Объекты изготовление = ном.СвязанныеОбъекты["aa724569-dc58-48d0-b030-2364d2f07b59"];
    	Сообщение ("Информация", ном.ToString());
    	Сообщение ("Информация", ном.Тип.ToString());
    	//Сообщение ("Информация1", изготовление.ToString());
    	Сообщение ("Информация2", изготовление.Count.ToString());
    /*	foreach (var link in ном.ChildHierarchyLinks)
    	{
    		Сообщение("", link.ToString());
    	} */
    	//Сообщение ("Информация", ном.Параметр["Наименование"]);
    //	Сообщение ("Связь с записью справочника Изготовление", ном.СвязанныйОбъект["Изготовление оснастки"].Параметр["Наименование"].ToString());
    //	Сообщение ("Связь с записью справочника Изготовление", string.Format("На изделие {0} найдены связанные объекты {1}", ном.Параметр["Обозначение"], ном.СвязанныйОбъект["Изготовление оснастки"].Параметр["Исполнитель"]));
   
        if (изготовление != null)
    		{
    		
    		// Сообщение ("Связь с записью справочника Изготовление", string.Format("Исполнитель заказа {0}", изготовление.Параметр["Guid"]) /*ном.СвязанныйОбъект["Изготовление оснастки"].Параметр["Наименование"].ToString()*/);
    	
    		}
    	else
    		{
    		Сообщение ("Информация", string.Format("На изделие {0} не было найдено связанных объектов", ном.ToString()));
    		} 
    	
    	}
  /*  public void ФормированиеОтчёта ()
        	{
        	;
        	} */
    public void ИзменитьСтатусУСправочникаИзготовление ()
    	{
    	 var eventContext = Context as EventContext;
        if (eventContext == null)
            return;

        var data = eventContext.Data as StateContextData;
        // Текущее действие
        var activeAction = data.ActiveAction;
        // Данные текущего действия
        ActiveActionData activeActionData = activeAction.GetData<ActiveActionData>();
        // Объекты, подключенные к БП
        List<ReferenceObject> processObjects = activeActionData.GetReferenceObjects().ToList();
        
        
        
        
        foreach (ReferenceObject ном in processObjects)
        	{
        	   //Сообщение("Информация", ном.ToString());
        	//   ReferenceObject[] заказ = null;
        	   List<ReferenceObject> заказы = null;
        	   ном.TryGetObjects(new Guid("aa724569-dc58-48d0-b030-2364d2f07b59"), out заказы);
        	   
        	   if (заказы != null)
        		{
        		Сообщение ("Связь с записью справочника 'Изготовление'", "aa724569-dc58-48d0-b030-2364d2f07b59");
        		foreach (ReferenceObject заказ in заказы)
        			{
        		заказ.BeginChanges();
        		заказ[new Guid("6537d551-cbd1-4c5f-b09f-9f40e82fddc6")].Value = "2"; 
        		заказ[new Guid("e17fa3c4-3f43-4be4-9ec7-8d2aa24f67a7")].Value = DateTime.Now;
        	/*	List<MailTask> tasks = Context.Connection.Mail.DOCsAccount.GetTasks().Where(t => t.Guid == new Guid((string)заказ[new Guid("98462bf5-01d6-4245-9821-1a4a4adcd202")].Value)).ToList();
                //Сообщение("Входящие задания", string.Join(Environment.NewLine, tasks.Select(GetMailItemString)));
               if (tasks.Count == 1)
                	{
                	tasks[0].Complete(string.Format("Закрытие заявки"));
                	}
                else 
                	{
                	Сообщение ("Информация", "Поручения не найдены");
                	return;
                	}  */
                                           	
        		заказ.EndChanges();
        		}
        		
        		}
        	   else 
        	   	{
        	   	   Сообщение("Информация", string.Format("На изделие {0} не было найдено связанных объектов", ном.ToString()));
        	   	}
                     
        	}
    	}
  
}
