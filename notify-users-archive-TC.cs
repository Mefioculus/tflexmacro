using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using TFlex.DOCs.Model.Macros;
using TFlex.DOCs.Model.Macros.ObjectModel;


using System.Text;
using TFlex.DOCs.Model.Mail;
using TFlex.DOCs.Model.References;
using TFlex.DOCs.Model.References.Users;
using TFlex.DOCs.Model.References.Documents;

public class Macro : MacroProvider {
	 public Macro(MacroContext context)
        : base(context)
	 	{}
	
	
    public override void Run() {
        // Получаем текущий объект - маршрутную карту
        ReferenceObject currentObject = Context.ReferenceObject;

        // Инициализируем новую коллекцию для хранения пользователей
        List<User> пользователи = new List<User>();

        // Получаем группу рассылки по связи "Группы рассылки"
        ReferenceObject mailGroup = currentObject.GetObject(new Guid("df88c8e4-b032-488d-b8ed-90855d83f0cd"));
        if (mailGroup == null)
            return;

        // Получаем всех пользователей, прикрепленных к группе рассылок
        foreach (ReferenceObject userRecord in mailGroup.GetObjects(new Guid("e837ec33-6aa8-4e02-be5f-75a8cb54e566"))) {
            User currentUser = Context.Connection.References.Users.Find((Guid)userRecord[new Guid("1e0036f9-adb0-4ddd-9ccf-ba210d9d951e")].Value) as User;
            if (currentUser != null)
                пользователи.Add(currentUser);
        }

        // Выводим сообщение с списком пользователей, которым будет производиться отправка
        Message(
                "Адресаты рассылки",
                $"Оповещение об изменении будет отправлено следующим пользователям:\n{string.Join("\n", пользователи.Select(user => user.ToString()))}");

        // Формируем заголовок
        string заголовок = String.Format(
                "Изменения в \"Архиве ogtMK\" '{0} - {1} - {2}'",
                ((int)currentObject[new Guid("7131d5fd-4080-4df4-b0cb-ee094ad9603f")].Value).ToString(), // Номер
                (string)currentObject[new Guid("c11b5a98-c22c-42bc-8375-be30052ffba2")].Value, // Обозначение детали, узла
                (string)currentObject[new Guid("e6d133be-e21e-445c-8651-5f35d2068f74")].Value // Наименование ДСЕ
                );

        // Формируем текст письма
    	string текстПисьма = string.Format(
                "В электронный \"Архив ogtMK\" добавлен новый документ. Ссылка на сканированный документ: <html><body><a href=file:///{0}>file:///{0}</a></body></html>",
                ((string)currentObject[new Guid("5947d0ce-b096-4791-96a4-e3ac03f9c49c")].Value).Replace(" ", "%20"), // Скан документа
                (string)currentObject[new Guid("c11b5a98-c22c-42bc-8375-be30052ffba2")].Value // Обозначение детали, узла
                );

        /*
        string текстПисьма = String.Format(
            "в сетевой папке ogtMK произошли изменения в позиции '{0} - {1}\nСсылка на файл: {2}",
            мк.Параметр["Наименование ДСЕ"],
            мк.Параметр["Оборзачение детали, узла"],
            мк.Параметр["Скан документа"]
        );
        */

        //Сообщение("", мк.Параметр["Скан документа"].ToString().Replace(" ", "%20"));
        
       
        
        
        //\\mainserver\ogtmk\MK\УЯИС\УЯИС.757\УЯИС.757.564.022-03(1)
        ОтправитьСообщение(пользователи, заголовок, текстПисьма);

    }
    

    public void ОтправитьСообщение(List<User> пользователи, string заголовок, string текстПисьма) {

    	// Создаем новое сообщение
        MailMessage message = new MailMessage(Context.Connection.Mail.DOCsAccount) {
            Subject = заголовок,
            Body = текстПисьма
        };
        
        message.BodyType = MailBodyType.Html;
        

        // Добавляем адресатов сообщения
        foreach (User пользователь in пользователи) {
            message.To.Add(new MailUser(пользователь));
            message.To.Add(new EMailAddress(пользователь.Email)); 
        }
        
        /*
        // Прикрепляем к сообщению объект справочника
        ReferenceObject refObj = new DocumentReference(Context.Connection).Objects[0];
        if (refObj != null)
            message.Attachments.Add(new ObjectAttachment(refObj));
        */
        
        message.Send();
    }

    /*
    public string GetUrlObj(ReferenceObject obj) {
    	if (obj == null)
    		return string.Empty;
    	
    	string server = Context.Connection.ConnectionParameters.GetServerAddress();
    	string objGuid = obj.SystemFields.Guid.ToString();
    	string url = string.Format("docs://{0}/OpenReferenceWindow/?refId={1}&objId={2}", server, obj.Reference.Id, objGuid);
        
    	return url;
    }
    */
}    

