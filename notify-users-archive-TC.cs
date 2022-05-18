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

public class Macro : MacroProvider
{
	 public Macro(MacroContext context)
        : base(context)
	 	{}
	
	
    public override void Run()
    {
        
        Объект мк = ТекущийОбъект;
        Объекты пользователи = new Объекты();
        
        try
        {
            пользователи  = мк.СвязанныйОбъект["Группы рассылки"].СвязанныеОбъекты["Группы и пользователи"];
        }
        catch
        {
            Сообщение("Информация", "Не указана группа рассылки, письма не будут отправлены");            
        	return;
        }
        
     //  string urlObjDOCs = GetUrlObj(Context.ReferenceObject);
        
        string message = "";
        foreach(Объект пользователь in пользователи)
        {
           message = message + String.Format("{0}\n", пользователь.Параметр["Наименование"]);
           //message = message + мк.Параметр["Скан документа"];
        }
        Сообщение("", message);
        message = "";
        
        
        
        string заголовок = String.Format("Изменения в \"Архиве ogtMK\" '{0} - {1} - {2}'", мк.Параметр["Номер"].ToString(), мк.Параметр["Обозначение детали, узла"], мк.Параметр["Наименование ДСЕ"]);
   //    string текстПисьма = String.Format("в сетевой папке ogtMK произошли изменения в позиции '{0} - {1}\n Ссылка на файл: {2}", мк.Параметр["Наименование ДСЕ"], мк.Параметр["Обозначение детали, узла"], мк.Параметр["Скан документа"]);
       
    	string текстПисьма = string.Format("В электронный \"Архив ogtMK\" добавлен новый документ. Ссылка на сканированный документ: <html><body><a href=file:///{0}>file:///{0}</a></body></html>", мк.Параметр["Скан документа"].ToString().Replace(" ", "%20"), мк.Параметр["Обозначение детали, узла"]);
   //     Сообщение("", мк.Параметр["Скан документа"].ToString().Replace(" ", "%20"));
        
       
        
        
        //\\mainserver\ogtmk\MK\УЯИС\УЯИС.757\УЯИС.757.564.022-03(1)
        ОтправитьСообщение(пользователи, заголовок, текстПисьма);

    }
    

    public void ОтправитьСообщение(Объекты пользователи, string заголовок, string текстПисьма)
    {

    	// Создаем новое сообщение
        MailMessage message = new MailMessage(Context.Connection.Mail.DOCsAccount)
        {
            Subject = заголовок,
            Body = текстПисьма
        };
        
        message.BodyType = MailBodyType.Html;
        

        // Добавляем адресатов сообщения
        foreach (Объект пользователь in пользователи)
        {
            message.To.Add(new MailUser((User)пользователь));
            message.To.Add(new EMailAddress(пользователь.Параметр["Электронная почта"].ToString())); 
        }
        
     /*  // Прикрепляем к сообщению объект справочника
        ReferenceObject refObj = new DocumentReference(Context.Connection).Objects[0];
        if (refObj != null)
            message.Attachments.Add(new ObjectAttachment(refObj)); */


        
        message.Send();
    }

    /* public string GetUrlObj(ReferenceObject obj)
    {
    	if (obj == null)
    		return string.Empty;
    	
    	string server = Context.Connection.ConnectionParameters.GetServerAddress();
    	string objGuid = obj.SystemFields.Guid.ToString();
    	string url = string.Format("docs://{0}/OpenReferenceWindow/?refId={1}&objId={2}", server, obj.Reference.Id, objGuid);
        
    	return url;
    }	*/
}    

