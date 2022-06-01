using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Forms = System.Windows.Forms;
using TFlex.DOCs.Model.Desktop;
using TFlex.DOCs.Model.Macros;
using TFlex.DOCs.Model.Macros.ObjectModel;
using TFlex.DOCs.Model.References;
// using TFlex.DOCs.Model.Classes; // Для работы с классами
using TFlex.DOCs.Model.Stages; // Для работы со стадиями

using TFlex.DOCs.Model.FilePreview.CADExchange;
using TFlex.DOCs.Model.FilePreview.CADService;

using TFlex.DOCs.Model.References.Reporting;
using TFlex.DOCs.Common;

// Для отправки уведомлений
using TFlex.DOCs.Model.Mail;
using TFlex.DOCs.Model.References.Users;
using TFlex.DOCs.Model.References.Files;


//using TFlex.DOCs.UI.Utils.Helpers;


/*
Для работы макроса так же потрубеются дополнительные ссылки на внешние библиотеки
TFlex.DOCs.UI.Utils.dll
TFlex.DOCs.Common.dll
TFlex.Reporting.dll
*/


public class Macro : MacroProvider
{
    string directoryWithPDF;
    string[] listOfFilesOnProcessing;
    Reference reportReference;
    Reference archiveReference;
    Report reportMI;
    Report reportFMI;
    
    // Guids
    private static class Guids {
        public static class Reports {
            public static Guid Свидетельство = new Guid("a9e67d98-ce6c-4ad4-8c94-328e71089099");
        }

        public static class Types {
            public static Guid ТипPDF = new Guid("58e7a26a-cf5f-445b-b08a-885f1bcf7f12");
            public static Guid ПротоколМеталлографическойЛаборатории = new Guid("80debbb1-d94d-4926-a7d8-da6e22cd277f");
            public static Guid ПротоколФизикомеханическойЛаборатории = new Guid("9233e1f0-6584-49a5-9ffb-cedabbc579af");
            public static Guid ПротоколСпектральнойЛаборатории = new Guid("f674fa70-afe0-4be5-aa63-4729c5676dbc");
            public static Guid ПротоколХимическойЛаборатории = new Guid("0a98f565-a37d-4115-b7dc-547c051cd3e6");
            public static Guid ПротоколГальваническойЛаборатории = new Guid("08f5dc3f-7760-4098-8e20-20e12011ac27");
            public static Guid ПротоколМагнитнойЛаборатории = new Guid("3a2d970e-99b5-4cf7-8813-7c86609aaf6b");
            public static Guid ПротоколЭлектрическойЛаборатории = new Guid("c7ae898a-5e8a-4891-9dc2-9bd5710ae9fb");
            
        }

        public static class References {
            public static Guid АрхивЦЗЛ = new Guid("ca66f59f-e077-4ae4-91fe-645dd47c2300");
            public static Guid Отчеты = new Guid("d3396686-2cb9-44ff-994b-d446c0a42515");
        }

        public static class ListsOfObjects {
            public static Guid Образцы = new Guid("97d5d377-48ee-41d0-b6fc-ab43b7abdd6e");
            public static Guid Показатели = new Guid("7d2616a0-73c2-4f4a-933b-10c0bcc1f37a");
            public static Guid КонтролируемыеПараметры = new Guid("93b19910-3553-4f36-9dd7-f81e42e2739d");
            public static Guid ФактическиеПоказания = new Guid("a6f224ab-6044-472b-8e32-5780d2148e39");
            public static Guid ОбразцыМагнитнойЛаборатории = new Guid("784e5bb5-8247-4a47-bf78-7fc5033a2a6c");
            public static Guid КомпонентыЭлектролита = new Guid("fd3ca50d-79c1-4237-91f5-b65ecf706d27");
            public static Guid СписокОборудования = new Guid("93e97cfd-25b6-4b2a-98b0-21ae8ed24eb3");
            public static Guid СписокПользователей = new Guid("e837ec33-6aa8-4e02-be5f-75a8cb54e566");
        }

        public static class Props {
            public static Guid НомерПротокола = new Guid("d662eed7-c2a2-41fc-9b35-05e527349cc7");
            public static Guid ДатаПротокола = new Guid("3ba03308-123b-4840-bc78-bc0fdcb1de4f");
            public static Guid АвторПротокола = new Guid("799ce8dc-6936-4d0f-b837-acf5525fef40");
            public static Guid СводноеНаименованиеПротокола = new Guid("7b4b4de4-70b0-4a1c-83ce-20d2adf9f4f6");
            public static Guid Заказчик = new Guid("04deee06-ef48-4538-b82f-5e0e3b463687");
            public static Guid Оборудование = new Guid("058ab970-1352-4e10-82b4-e49d8847004f");
            // Параметры для передачи данных в отчет
            public static Guid ПараметрВидТаблицы = new Guid("5ee7e25b-e56a-4f62-abb0-f092fc0bdb27");

            // Параметры объектов списка пользователей
            public static Guid GuidПользователя = new Guid("1e0036f9-adb0-4ddd-9ccf-ba210d9d951e");

            // Параметр типа Оборудование списка объектов Оборудование
            public static Guid НаименованиеОборудованияВСписке = new Guid("93e75517-2c3b-4b36-ba32-c39847b861f7");

            // Параметры типа Образец
            // Для расчета сводного размера
            public static Guid ФормаОбразца = new Guid("03b50c57-5856-4761-807e-8bf07ef6c50f");
            public static Guid Длина = new Guid("6b5e98e3-81a5-4ec8-908a-ff73c8cad31d");
            public static Guid Ширина = new Guid("b391cbe3-0012-4f52-9c58-76e6e90fe30b");
            public static Guid Толщина = new Guid("5d3e8b14-bc35-408d-bcc2-53ec81d51d1e");
            public static Guid Диаметр = new Guid("50c7f6c1-e8da-496a-9327-f32bbbbee29f");
            public static Guid СводныйРазмер = new Guid("f91d8f9c-0d56-414f-b452-1032cc3cd0cd");
            // Для передачи в отчет
            public static Guid Испытание = new Guid("4cfe5d7a-93ad-4db2-a9d4-ab2fda219a36");
            public static Guid ПределПрочности = new Guid("d205d64d-ba0e-43e9-b7b3-52b250a5c8f8");
            public static Guid ПределТекучести = new Guid("20fb0b1a-be6d-4c0d-b78f-f8fc7cf0fe9c");
            public static Guid ОтносительноеУдлинение = new Guid("86bc650f-101e-47c2-9686-2ef8f5ccacee");
            public static Guid ПрочностьПриИзгибе = new Guid("885eb736-578a-427e-924e-044bafea03f6");

            // Параметры типа Показатель
            public static Guid Наименование = new Guid("759d291e-5a1c-4cb0-bc6b-6794cf1f3cb2");
            public static Guid ДанныеПоТУ = new Guid("6d25bb62-1975-4bae-82ae-20bfee8ecaf9");
            public static Guid ФактическиеДанные = new Guid("534a9e41-174b-46b9-9ba6-a12d336897a4");

            // Параметры типа Контролируемый параметр
            public static Guid КонтролируемыйПараметр = new Guid("a24c5b02-374f-40aa-be54-f5103f46ef77");
            public static Guid ПунктТребованийНТД = new Guid("971dbda8-a091-4e8f-ad82-4dcc37f3fcbb");
            public static Guid Норма = new Guid("39a0b103-2964-4e64-b8d1-754b2f477c5d");
            
            // Параметры типа Фактическое значение
            public static Guid НаименованиеИзмеряемогоОбъекта = new Guid("8e1c6e2e-d885-4b43-8dd7-6d5eba283783");
            public static Guid ФактическоеЗначение = new Guid("878df0b8-f018-41e6-a5a4-5f673dedf4ad");

            // Параметры протокола магнигной лаборатории
            public static Guid ТипИсследования = new Guid("1f61dd57-ec89-403b-a9d8-01d22ed71736");

            // Параметры образца магнитной лаборатории
            public static Guid РазмерМатериалаМагнитнаяЛаборатория = new Guid("fe9d24b0-e1b0-489c-b029-b7b0b76878cf");
            public static Guid СводноеНаименованиеОбразцаМагнитнойЛаборатории = new Guid("43b442a0-5901-4fe7-8ad1-6e072bd350d6");
            public static Guid ТипРазмераОбразцаМагнитнойЛаборатории = new Guid("59f6ea61-d864-4439-8f70-53e8937a2643");
            public static Guid КоэрцитивнаяСилаОбразца = new Guid("b0d15b96-45f4-43b8-95e8-633c7f8e4821");
            public static Guid МаксимальнаяПроницаемостьОбразца = new Guid("403b8795-5d09-4ac6-91c9-682313c8e3e5");
            public static Guid МагнитнаяИндукцияОбразца1 = new Guid("4793f354-8820-4027-a599-bd70fc7c8693");
            public static Guid МагнитнаяИндукцияОбразца2 = new Guid("568b10a1-9bf8-438c-902b-1077fa213f7b");
            public static Guid МагнитнаяИндукцияОбразца3 = new Guid("ca6b8770-4810-403b-b05e-b6571e1f3133");
            public static Guid МагнитнаяИндукцияОбразца4 = new Guid("214ffd17-4f90-4ce6-917a-7e88840b0527");
            public static Guid МагнитнаяИндукцияОбразца5 = new Guid("41b86f16-5dec-4272-85ee-0258239febb5");
            public static Guid МагнитнаяИндукцияОбразца6 = new Guid("9381da32-9691-403e-8fca-47c7b6bdc39b");
            public static Guid УдельныеПотери1 = new Guid("a061fa6e-be09-4529-bb78-8e1e26f1d4d4");
            public static Guid УдельныеПотери2 = new Guid("7b5b64d0-43a9-43c0-b4e8-55260a9e043c");
            
            // Параметры материала магнитной лаборатории
            public static Guid МаркаМатериалаМагнитнаяЛаборатория = new Guid("89eacccd-ad18-4e14-b383-a2846cbacaff");
            public static Guid МагнитнаяИндукция1 = new Guid("8e272ff4-207a-43ed-adf1-24320948a7d2");
            public static Guid МагнитнаяИндукция2 = new Guid("220f3c7a-ef3f-44c0-81aa-9290b27c78dd");
            public static Guid МагнитнаяИндукция3 = new Guid("7bd70dc0-37b8-44ea-ba06-6fdbe0c5db14");
            public static Guid МагнитнаяИндукция4 = new Guid("119d00a0-96ae-471d-b90a-e19e6f3d20df");
            public static Guid МагнитнаяИндукция5 = new Guid("040eb1c3-9610-4685-bddc-944b9a1b81eb");
            public static Guid МагнитнаяИндукция6 = new Guid("7dc4b5d1-b9c9-42e3-b255-2834a4ac1a7a");
            public static Guid МинМагнитнаяИндукция1 = new Guid("3dfc300e-dd47-4c15-9a30-7cd97d78ce48");
            public static Guid МаксМагнитнаяИндукция1 = new Guid("5752b855-61cd-482e-a114-0b2925da8f1d");
            public static Guid МинМагнитнаяИндукция2 = new Guid("e8b759cc-4aff-4b98-9638-87eeaf8166c1");
            public static Guid МаксМагнитнаяИндукция2 = new Guid("a070a605-48df-416c-80c5-8219fdad524a");
            public static Guid МинМагнитнаяИндукция3 = new Guid("d49c7a51-72a1-43f8-81c0-b6b953962234");
            public static Guid МаксМагнитнаяИндукция3 = new Guid("19f43123-7281-475a-a47a-f326b4e0e30b");
            public static Guid МинМагнитнаяИндукция4 = new Guid("b3d8cd52-e512-4371-924a-d825056fd358");
            public static Guid МаксМагнитнаяИндукция4 = new Guid("67a5dc74-5e1f-4b4b-adbe-3bf54710223c");
            public static Guid МинМагнитнаяИндукция5 = new Guid("ce3be873-fd6b-4243-9be2-dd2ccb9c5ab3");
            public static Guid МаксМагнитнаяИндукция5 = new Guid("6ad6c59b-1de9-4c6d-a477-4c1f58373d62");
            public static Guid МинМагнитнаяИндукция6 = new Guid("308ff472-62e6-4009-bbf6-ab1f83c1567b");
            public static Guid МаксМагнитнаяИндукция6 = new Guid("59cf59e3-435b-4581-8f2d-0dc66d0fbeda");
            public static Guid МинКоэрцитивнаяСила = new Guid("789fbfea-85e6-4930-8824-7c8a6c81a77e");
            public static Guid Напряженность1 = new Guid("094af1de-177e-4623-94f8-4140c46ef4a0");
            public static Guid Напряженность2 = new Guid("5fb1907e-bd8b-4b97-8a73-011aa06276dd");
            public static Guid Напряженность3 = new Guid("dcd9aeca-a88b-4772-a21e-e6c2dafd5bdf");
            public static Guid Напряженность4 = new Guid("5ff64532-65c5-4b6b-8bed-8a5f3bb94aeb");
            public static Guid Напряженность5 = new Guid("25a0764c-ea45-472e-a0b0-1b77ebb622ab");
            public static Guid Напряженность6 = new Guid("34b4f05e-bde3-4342-89d3-846b4e298cc3");
            public static Guid МаксКоэрцитивнаяСила = new Guid("f85b7135-21b6-4231-bad5-f83e1ee0e6b6");
            public static Guid МаксимальнаяПроницаемость = new Guid("d66b6261-22e4-4226-b8e8-cdd6294fecd2");
            public static Guid КоэрцитивнаяСила = new Guid("db5f2f57-bb99-48cb-9668-14a7e033c584");
            public static Guid ТипСтандартаМагнитнаяЛаборатория = new Guid("2c1890c0-e0d4-40dc-b680-94e49130efe9");
            public static Guid СтандартМагнитнаяЛаборатория = new Guid("7892a0b0-9d5a-4e4d-a0c4-ccbbda5d185b");
            public static Guid КоличествоЗамеров = new Guid("6afc0add-c244-43bd-9e48-847695a180ae");

            // Параметры для химической лаборатории

            // Параметры типа Электролит
            public static Guid НаименованиеЭлектролита = new Guid("1b4b1aac-1ac6-4b6e-aee6-4f337632dce9");
            // Параметры типа Компонент
            public static Guid НаименованиеКомпонента = new Guid("16da1dd1-71c6-43b9-9c52-58c220c8b4e7");
            public static Guid МинимальноеСодержаниеКомпонента = new Guid("7a9a6229-0c85-4acf-ae52-51030bbb6d1a");
            public static Guid МаксимальноеСодержаниеКомпонента = new Guid("dc52485d-6c5c-4d5b-bddf-21b4e76f4ec9");
            public static Guid СводноеСодержаниеКомпонента = new Guid("06e87a6e-902c-4a78-9212-9f021faaae36");
            // Параметры результатов замеров
            public static Guid СодержаниеКомпонента1 = new Guid("12dd0c84-1e58-4302-81ef-f1b8af872f43");
            public static Guid СодержаниеКомпонента2 = new Guid("4eebf0b0-2471-4ff3-9651-d5f10a5da28d");
            public static Guid СодержаниеКомпонента3 = new Guid("e41b599d-5296-4828-848c-9fb1e4d44515");
            public static Guid СодержаниеКомпонента4 = new Guid("fa735a63-9df3-493d-a060-d9a00cdf33cf");
            public static Guid СодержаниеКомпонента5 = new Guid("80d8b00f-9a4b-4fa0-91d9-3e1c147a4826");
            public static Guid СодержаниеКомпонента6 = new Guid("3730bbb3-f58e-4566-b162-b127d50b3da8");
            public static Guid СодержаниеКомпонента7 = new Guid("98a206e5-4cc4-4572-966d-95979fcc7812");
            public static Guid СодержаниеКомпонента8 = new Guid("c15a0f03-5738-4e48-9fbd-fc57736fe46c");

            // Параметры каталога оборудования
            public static Guid НаименованиеОборудования = new Guid("b9b13d1f-397b-4467-bd41-f4f606fe01a8");
            public static Guid ЗаводскойНомерОборудования = new Guid("f436ae6d-7840-4ae0-8593-63e69f55ba7a");
            public static Guid ОкончаниеПоверкиОборудования = new Guid("1fc6089b-073b-4183-a5d2-ea0414e90d28");
            public static Guid НомерСвидетельстваОборудования = new Guid("19bf4e5b-afbf-4f64-af0e-29cff1e4d141");
            public static Guid СводноеНаименованиеОборудования = new Guid("836e0691-cbd3-445b-95c3-f8db5b6c919b");
        }

        public static class Links {
            public static Guid СправочныеМатериалыМагнитнаяЛаборатория = new Guid("c543586f-17ce-4731-9690-bfddd0f10a4b");
            public static Guid ПодлинникПротокола = new Guid("5545694e-602c-4090-bb9f-1453aa54845b");
            public static Guid СправочныеМатериалыХимическаяЛаборатория = new Guid("56c4a7e2-3007-4c94-a9f0-ad1857064a4f");
            public static Guid ГруппыРассылки = new Guid("4685b9f7-1e97-48e7-a192-52ed786eadb3");
            public static Guid ГруппыПользователей = new Guid("861bbc2e-ab60-4bd9-bc09-404d6c90aa95");
        }

        public static class Stages {
            public static Guid Разработка = new Guid("527f5234-4c94-43d1-a38d-d3d7fd5d15af");
            public static Guid Согласование = new Guid("c89e6ee8-f060-4091-b576-00025be0846a");
            public static Guid Корректировка = new Guid("18df455a-0dc8-43a9-b256-c0fd6898df1b");
            public static Guid Хранение = new Guid("9826e84a-a0a7-404b-bf0e-d61b902e346a");
        }
    }

    // Properties
    Reference ReportReferenceInstance {
        get {
            if (reportReference == null) {
                reportReference = Context.Connection.ReferenceCatalog.Find(Guids.References.Отчеты).CreateReference();
            }
            return reportReference;
        }
    }

    Reference ArchiveReferenceInstance {
        get {
            if (archiveReference == null) {
                archiveReference = Context.Connection.ReferenceCatalog.Find(Guids.References.АрхивЦЗЛ).CreateReference();
            }
            return archiveReference;
        }
    }


    public Macro(MacroContext context)
        : base(context)
    {
    }

    public override void Run()
    {
    }
    
    #region Методы для запуска генерации отчета

    // Формирование отчета из БП по окончанию согласования
    public void СформироватьОтчетДляБП(Объекты объекты) {

        foreach (Объект вложение in объекты) {
            ReferenceObject attachment = (ReferenceObject)вложение;
            
            if (attachment == null) {
                Message("Ошибка", "Не получилось обработать вложение");
                continue;
            }

            // Получаем контекст формирования отчета
            ReportGenerationContext reportContext = new ReportGenerationContext(attachment, null);
            reportContext.OpenFile = false;

            Report report = ReportReferenceInstance.Find(Guids.Reports.Свидетельство) as Report;
            if (report == null)
                Error("Не удалось найти объект отчета в справочнике 'Отчеты'. Обратитесь к системному администратору");

            // Перед началом формирования отчета переводим протокол на стадию корректировка для того, чтобы сформированный отчет можно было прикрепить к
            // обрабатываемому объекту

            ChangeStage(attachment, Guids.Stages.Корректировка);

            report.Generate(reportContext);

            // Автоматическое применение изменений файла отчета
            Desktop.CheckIn(reportContext.ReportFileObject, string.Format("Автоматическое формирование отчета для свидетельства №{0} архива ЦЗЛ", attachment[Guids.Props.СводноеНаименованиеПротокола].Value.ToString()), false);
            
            // Отправка уведомлений
            SendNotification(attachment);

        }
    }

    public void TestSendNotification() {
        // Получаем объект для тестирования
        Reference czlReference = Context.Connection.ReferenceCatalog.Find(new Guid("f7f43d73-857c-41f9-b449-38ee72caa221")).CreateReference();
        ReferenceObject testObject = czlReference.Find(new Guid("af8603b9-9786-4a66-a73d-a107c7477ac3"));

        Message("Проверка содержимого аккаунта", Context.Connection.Mail.DOCsAccount.ToString());
        Message("Аккаунты, которые есть в MailService", string.Join("\n", Context.Connection.Mail.Accounts.Select(acc => acc.ToString())));

        SendNotification(testObject);
        Message("Информация", "Работа макроса завершена");
    }

    private void SendNotification(ReferenceObject protocol) {
        // Получаем пользователей, прикрепленных через группы рассылок к протоколу
        ReferenceObject mailGroup = protocol.GetObject(Guids.Links.ГруппыРассылки);
        if (mailGroup == null)
            return;

        List<User> users = new List<User>();

        foreach (ReferenceObject user in mailGroup.GetObjects(Guids.ListsOfObjects.СписокПользователей)) {
            User findedUser = Context.Connection.References.Users.Find((Guid)user[Guids.Props.GuidПользователя].Value) as User;
            if (findedUser != null)
                users.Add(findedUser);
        }

        SendMailTo(users, protocol);
    }

    private void SendMailTo(List<User> users, ReferenceObject protocol) {
        if ((users.Count == 0) || (users == null))
            return;
        // Получаем название протокола
        string protocolName = (string)protocol[Guids.Props.СводноеНаименованиеПротокола].Value;
        String clientName = (string)protocol[Guids.Props.Заказчик].Value;

        // Создаем новое сообщение
        // INFO: Для того, чтобы отправить сообщение с общего аккаунта, нужно будет под ним произвести вход в систему
        // Открыть новое ServerConnection с учетными данными того пользователя, который требуется
        MailMessage message = new MailMessage(Context.Connection.Mail.DOCsAccount) {
            Subject = $"Протокол {protocolName}",
            Body = $"Согласование протокола '{protocolName}' для заказчика '{clientName}' завершено"
        };

        // Добавляем адресатов
        foreach (User user in users) {
            MailUser mUser = new MailUser(user);
            message.To.Add(new EMailAddress(user.Email));
            message.To.Add(mUser);
        }

        FileObject fileOfProtocol = protocol.GetObject(Guids.Links.ПодлинникПротокола) as FileObject;

        if (fileOfProtocol != null) {
            fileOfProtocol.GetHeadRevision();
            string pathToAttachment = fileOfProtocol.LocalPath;

            // Прикрепляем к письму файл протокола
            if (File.Exists(pathToAttachment))
                message.Attachments.Add(new FileAttachment(pathToAttachment));
        }

        // Отправляем сообщение
        string errors = string.Empty;
        try {
            message.Send();
        }
        catch (Exception e) {
            errors += $"Возникла ошибка при отправке письма следующим адресатам:\n{message.To.ToString()}\n\nТекст ошибки:\n{e.Message}";
        }

        if (errors != string.Empty) {
            Message("Ошибка", errors);
        }
    }

    private void ChangeStage(ReferenceObject refObj, Guid guidOfStage) {
        // Метод для изменения стадии объекта
        
        Stage stage = Stage.Find(Context.Connection, guidOfStage);

        if (stage == null) {
            Error("При изменении стадии объекта возникла ошибка. Указанная стадия не найдена");
        }

        stage.Change(new List<ReferenceObject>() {refObj}); // Охренеть логика. Я думал записимость идет от объекта, а они идет от стадии
    }

    //Формирование отчета для предварительного просмотра
    public void СформироватьОтчетДляПредварительногоПросмотра(ReferenceObject protocol) {
        // Получаем текущий объект
        UpdateEquipment(protocol); // Вызов метода для того, чтобы в отчет попало актуальное состояние добавленного оборудования
        if (protocol == null) {
            Message("Ошибка", "Не получилось обратиться к текущему объекту для формирования отчета");
            return;
        }

        // Получаем контекст формирования отчета
        ReportGenerationContext reportContext = new ReportGenerationContext(protocol, null);
        reportContext.OpenFile = true;

        // Получаем объект отчета
        Report report = ReportReferenceInstance.Find(Guids.Reports.Свидетельство) as Report;
        if (report == null)
            Error("Не удалось найти объект отчета в справочнике 'Отчеты'. Обратитесь к системному администратору");
        report.Generate(reportContext);

        Desktop.CheckIn(reportContext.ReportFileObject, "Предварительный просмотр", false);
    }

    public void СформироватьОтчетДляПредварительногоПросмотра() {
        СформироватьОтчетДляПредварительногоПросмотра(Context.ReferenceObject);
    }

    public void СформироватьОтчетДляПредварительногоПросмотра(Объекты объекты) {
        foreach (Объект вложение in объекты)
            СформироватьОтчетДляПредварительногоПросмотра((ReferenceObject)вложение);
    }

    #endregion Методы для запуска генерации отчета

    #region Общие методы для всех протоколов

    //Генерация сводного наименования протокола

    // Данный метод используется по событию, при создании объекта в справочнике Архив ЦЗЛ
    public void GenerateSummaryName() {
        // Данный метод на основа порядкового номера протокола, его даты, а так же типа протокола будет
        // формировать сводное наименование протокола вида Мет-2-02/20

        // Данные, из которых будет формироваться сводное наименование
        int orderNumber;
        DateTime date;
        string typeOfRecord;

        // Получаем текущий объект
        ReferenceObject currentObject = Context.ReferenceObject;

        // Получаем у текущего объекта его порядковый номер, дату, а так же его тип
        orderNumber = (int)currentObject[Guids.Props.НомерПротокола].Value;
        date = (DateTime)currentObject[Guids.Props.ДатаПротокола].Value;
        typeOfRecord = GetStringOfType(currentObject);

        //currentObject.BeginChanges();
        currentObject[Guids.Props.СводноеНаименованиеПротокола].Value = string.Format("{0}-{1}-{2}/{3}", typeOfRecord, orderNumber, date.Month, date.Year);
        //currentObject.EndChanges();
        
    }

    public void UpdateEquipment(ReferenceObject currentProtocol) {
        //ReferenceObject currentProtocol = Context.ReferenceObject;

        string result = string.Join(
                ":\n",
                currentProtocol.GetObjects(Guids.ListsOfObjects.СписокОборудования)
                    .OrderBy(equip => equip.SystemFields.CreationDate)
                    .Select(equip => (string)equip[Guids.Props.НаименованиеОборудованияВСписке].Value)
                );

        if ((string)currentProtocol[Guids.Props.Оборудование].Value != result) {
            bool wasEditable = true;
            if (!currentProtocol.Changing) {
                currentProtocol.BeginChanges();
                wasEditable = false;
            }
            currentProtocol[Guids.Props.Оборудование].Value = result;
            if (!wasEditable) {
                currentProtocol.EndChanges();
            }
        }
    }

    private string GetStringOfType(ReferenceObject referenceObject) {
        // Получаем тип переданного объекта и возвращаем его сокращенное название
        if (referenceObject.Class.IsInherit(Guids.Types.ПротоколМеталлографическойЛаборатории))
            return "МЕТ";
        else if (referenceObject.Class.IsInherit(Guids.Types.ПротоколФизикомеханическойЛаборатории))
            return "ФИЗ";
        else if (referenceObject.Class.IsInherit(Guids.Types.ПротоколСпектральнойЛаборатории))
            return "СП";
        else if (referenceObject.Class.IsInherit(Guids.Types.ПротоколХимическойЛаборатории))
            return "ХИМ";
        else if (referenceObject.Class.IsInherit(Guids.Types.ПротоколГальваническойЛаборатории))
            return "ГАЛ";
        else if (referenceObject.Class.IsInherit(Guids.Types.ПротоколМагнитнойЛаборатории))
            return "МАГ";
        else if (referenceObject.Class.IsInherit(Guids.Types.ПротоколЭлектрическойЛаборатории))
            return "ЭЛ";
        else {
            Message("Ошибка", "Тип данного объекта не входит в список предусмотренных для формирования сводного наименования");
            return "ОШИБКА";
        }
    }

    // Генерация сводного наименования в каталоге оборудования ЦЗЛ
    public void ГенерацияСводногоОбозначенияОборудования() {
        ReferenceObject оборудование = Context.ReferenceObject;

        // Получаем параметры для заполнения сводного обозначения
        string наименование = (string)оборудование[Guids.Props.НаименованиеОборудования].Value;
        string обозначение = (string)оборудование[Guids.Props.ЗаводскойНомерОборудования].Value;
        string поверка = (string)оборудование[Guids.Props.НомерСвидетельстваОборудования].Value;
        string годенДо = ((DateTime)оборудование[Guids.Props.ОкончаниеПоверкиОборудования].Value).ToString("dd.MM.yyyy");

        // Заполняем сводное обозначение
        оборудование[Guids.Props.СводноеНаименованиеОборудования].Value = поверка != string.Empty ?
            $"{наименование} зав. №{обозначение} поверка №{поверка} до {годенДо}" :
            $"{наименование} зав. №{обозначение}";

        оборудование[Guids.Props.СводноеНаименованиеОборудования].Value =
            наименование +
            (обозначение != string.Empty ? $"зав. №{обозначение}" : string.Empty) +
            (поверка != string.Empty ? $"поверка №{поверка} до {годенДо}" : string.Empty);
    }



    #endregion Общие методы для всех протоколов

    #region Методы для формирования магнитного протокола

    //Формирование сводного размера образца
    public void ФормированиеСводногоРазмераОбразца () {
        ReferenceObject sample = Context.ReferenceObject;
        
        const int бесформенный = 0;
        const int плоский = 1;
        const int прямоугольный = 2;
        const int круглый = 3;
        // Получаем у текущего объекта параметры размеров и вида формы
        int formType = (int)sample[Guids.Props.ФормаОбразца].Value;
        double area = 0.0;
        double length = 0.0;
        double width = 0.0;
        double thickness = 0.0;
        double diameter = 0.0;

        switch (formType) {
            case плоский:
                // Получаем значение длины и ширины, остальные значения обнуляем
                length = (double)sample[Guids.Props.Длина].Value;
                width = (double)sample[Guids.Props.Ширина].Value;

                if ((length != 0.0) && (width != 0.0))
                    area = length * width;
                
                sample[Guids.Props.СводныйРазмер].Value = string.Format("{0}Х{1} (S={2})", length, width, area);
                // Обнуляем остальные значения
                sample[Guids.Props.Толщина].Value = 0.0;
                sample[Guids.Props.Диаметр].Value = 0.0;
                break;
            case прямоугольный:
                length = (double)sample[Guids.Props.Длина].Value;
                width = (double)sample[Guids.Props.Ширина].Value;
                thickness = (double)sample[Guids.Props.Толщина].Value;

                if ((length != 0.0) && (width != 0.0))
                    area = length * width;

                sample[Guids.Props.СводныйРазмер].Value = string.Format("{0}Х{1}Х{2} (S={3})", length, width, thickness, area);
                // Обнуляем остальные значения
                sample[Guids.Props.Диаметр].Value = 0.0;
                break;
            case круглый:
                diameter = (double)sample[Guids.Props.Диаметр].Value;
                
                if (diameter != 0.0) {
                    double radius = diameter / 2;
                    area = System.Math.PI * radius * radius;
                }
                sample[Guids.Props.СводныйРазмер].Value = string.Format("d{0} (S={1})", diameter, area);
                // Обнуляем остальные значения
                sample[Guids.Props.Длина].Value = 0.0;
                sample[Guids.Props.Ширина].Value = 0.0;
                sample[Guids.Props.Толщина].Value = 0.0;
                break;
            case бесформенный:
                sample[Guids.Props.СводныйРазмер].Value = string.Empty;
                sample[Guids.Props.Длина].Value = 0.0;
                sample[Guids.Props.Ширина].Value = 0.0;
                sample[Guids.Props.Толщина].Value = 0.0;
                sample[Guids.Props.Диаметр].Value = 0.0;
                break;
            default:
                Error("При попытке сформировать сводный размер возникла ошибка. Выбран неизвестный тип формы образца");
                break;
        }
    }

    public void ФормированиеСводногоРазмераКомпонентаХимическойЛаборатории() {
        //TODO: реализовать метод
        ReferenceObject component = Context.ReferenceObject;
        double minValue = (double)component[Guids.Props.МинимальноеСодержаниеКомпонента].Value;
        double maxValue = (double)component[Guids.Props.МаксимальноеСодержаниеКомпонента].Value;

        var rangeParam = component[Guids.Props.СводноеСодержаниеКомпонента];

        if (minValue == maxValue) {
            if (minValue != 0)
                rangeParam.Value = minValue.ToString();
            else
                rangeParam.Value = "-";
            return;
        }
        if (minValue == 0) {
            rangeParam.Value = $"до {maxValue.ToString()}";
            return;
        }
        if (maxValue == 0) {
            rangeParam.Value = $"от {minValue.ToString()}";
            return;
        }
        if (minValue > maxValue) {
            rangeParam.Value = "Ошибка";
            return;
        }

        rangeParam.Value = $"{minValue.ToString()}-{maxValue.ToString()}";
    }

    //Формирование сводного наименования образца магнитной лаборатории
    public string ФормированиеСводногоНаименованияОбразцаМагнитнойЛаборатории() {

        // Строка, которая будет содержать сводное наименование
        string summaryName = string.Empty;
        ReferenceObject sample = Context.ReferenceObject;
        if (sample == null)
            return "Не удалось получить образец";

        // Получаем протокол, к которому подключен данный образец
        ReferenceObject protocol = sample.MasterObject;
        if (protocol == null)
            return "Не удалось получить протокол";

        // Получаем материал, который привязан к протоколу
        ReferenceObject material = protocol.GetObject(Guids.Links.СправочныеМатериалыМагнитнаяЛаборатория);
        if (material == null)
            return "В протоколе не указан материал";

        string labelOfMaterial = material[Guids.Props.МаркаМатериалаМагнитнаяЛаборатория].Value.ToString();
        string sizeOfMaterial = sample[Guids.Props.РазмерМатериалаМагнитнаяЛаборатория].Value.ToString();

        // Определение типа размера для подстановки корректной переменной в сводное наименование
        int typeOfSize = (int)sample[Guids.Props.ТипРазмераОбразцаМагнитнойЛаборатории].Value;
        string typeOfSizeVariable = string.Empty;

        switch (typeOfSize) {
            // Случай, когда размер - толщина
            case 0:
                typeOfSizeVariable = "s";
                break;
            // Случай, когда размер - диаметр
            case 1:
                typeOfSizeVariable = "d";
                break;
            // Дефолтным значением будет значение толщины
            default:
                typeOfSizeVariable = "s";
                break;
        }

            summaryName = string.Format("{0}, {1} = {2} мм", labelOfMaterial, typeOfSizeVariable, sizeOfMaterial);
            
            // Присваиваем значение для параметра сводное наименование
            sample[Guids.Props.СводноеНаименованиеОбразцаМагнитнойЛаборатории].Value = summaryName;
            // Возвращаем значение сводного наименования
            return string.Format("Сводное наименование: {0}", summaryName);
    }

    //Вывод данных о допустимых значениях в диалоге образца магнитой лаборатории
    public string ОтображениеДопустимогоЗначенияОбразцаМагнитнойЛаборатории(string parameter) {
        ReferenceObject sample = Context.ReferenceObject;
        if (sample == null)
            return "Не удалось вернуть образец";

        ReferenceObject protocol  = sample.MasterObject;
        if (protocol == null)
            return "Не удалось получить протокол";

        ReferenceObject material = protocol.GetObject(Guids.Links.СправочныеМатериалыМагнитнаяЛаборатория);
        if (material == null)
            return "-";

        string result = string.Empty;

        switch (parameter) {
            case "1":
                result = material[Guids.Props.МагнитнаяИндукция1].Value.ToString();
                break;
            case "2":
                result = material[Guids.Props.МагнитнаяИндукция2].Value.ToString();
                break;
            case "3":
                result = material[Guids.Props.МагнитнаяИндукция3].Value.ToString();
                break;
            case "4":
                result = material[Guids.Props.МагнитнаяИндукция4].Value.ToString();
                break;
            case "5":
                result = material[Guids.Props.МагнитнаяИндукция5].Value.ToString();
                break;
            case "6":
                result = material[Guids.Props.МагнитнаяИндукция6].Value.ToString();
                break;
            // Убрал данный параметр, так как им не будут пользоваться
            //case "Проницаемость":
                //result = material[Guids.Props.МаксимальнаяПроницаемость].Value.ToString();
                //break;
            case "Коэрцитивная сила":
                result = material[Guids.Props.КоэрцитивнаяСила].Value.ToString();
                break;
            default:
                result = "Ошибка";
                break;
        }

        return result;
    }

    //ФормированиеСводногоДопустимогоРазмера
    public string ПолучитьДиапазонДопустимыхЗначенийПараметра(string nameOfParameter) {
        ReferenceObject material = Context.ReferenceObject;
        string result = string.Empty;
        switch (nameOfParameter) {
            case "Индукция 1":
                result = ПолучитьСтрокуДиапазона(material, Guids.Props.МинМагнитнаяИндукция1, Guids.Props.МаксМагнитнаяИндукция1);
                material[Guids.Props.МагнитнаяИндукция1].Value = result;
                break;
            case "Индукция 2":
                result = ПолучитьСтрокуДиапазона(material, Guids.Props.МинМагнитнаяИндукция2, Guids.Props.МаксМагнитнаяИндукция2);
                material[Guids.Props.МагнитнаяИндукция2].Value = result;
                break;
            case "Индукция 3":
                result = ПолучитьСтрокуДиапазона(material, Guids.Props.МинМагнитнаяИндукция3, Guids.Props.МаксМагнитнаяИндукция3);
                material[Guids.Props.МагнитнаяИндукция3].Value = result;
                break;
            case "Индукция 4":
                result = ПолучитьСтрокуДиапазона(material, Guids.Props.МинМагнитнаяИндукция4, Guids.Props.МаксМагнитнаяИндукция4);
                material[Guids.Props.МагнитнаяИндукция4].Value = result;
                break;
            case "Индукция 5":
                result = ПолучитьСтрокуДиапазона(material, Guids.Props.МинМагнитнаяИндукция5, Guids.Props.МаксМагнитнаяИндукция5);
                material[Guids.Props.МагнитнаяИндукция5].Value = result;
                break;
            case "Индукция 6":
                result = ПолучитьСтрокуДиапазона(material, Guids.Props.МинМагнитнаяИндукция6, Guids.Props.МаксМагнитнаяИндукция6);
                material[Guids.Props.МагнитнаяИндукция6].Value = result;
                break;
            case "Коэрцитивная сила":
                result = ПолучитьСтрокуДиапазона(material, Guids.Props.МинКоэрцитивнаяСила, Guids.Props.МаксКоэрцитивнаяСила);
                material[Guids.Props.КоэрцитивнаяСила].Value = result;
                break;
            default:
                result = "Ошибка";
                break;
        }
        return result;
    }

    private string ПолучитьСтрокуДиапазона(ReferenceObject material, Guid minValueOfParameter, Guid maxValueOfParameter) {
        // Получаем текущий объект для получения значения параметров
        string result = string.Empty;

        double minValue = (double)material[minValueOfParameter].Value;
        double maxValue = (double)material[maxValueOfParameter].Value;

        if ((minValue == 0.0) && (maxValue == 0.0)) {
            return result;
        }
        if ((minValue != 0.0) && (maxValue != 0.0)) {
            return string.Format("{0} - {1}", minValue, maxValue); 
        }

        result = minValue == 0.0 ?
            string.Format("< {0}", maxValue.ToString()) : string.Format("> {0}", minValue.ToString());

        return result;
    }

    //ОтображениеНапряженностиВОбразце()
    public string ОтображениеНапряженностиВОбразце(int number) {
        ReferenceObject sampleObject = Context.ReferenceObject;
        ReferenceObject material = sampleObject.MasterObject.GetObject(Guids.Links.СправочныеМатериалыМагнитнаяЛаборатория);

        string result = string.Empty;

        if (material == null)
            return result;

        // Получаем значение напряженности в зависимости от номера замера, для которого требуется ее получить
        switch (number) {
            case 1:
                result = material[Guids.Props.Напряженность1].Value.ToString();
                break;
            case 2:
                result = material[Guids.Props.Напряженность2].Value.ToString();
                break;
            case 3:
                result = material[Guids.Props.Напряженность3].Value.ToString();
                break;
            case 4:
                result = material[Guids.Props.Напряженность4].Value.ToString();
                break;
            case 5:
                result = material[Guids.Props.Напряженность5].Value.ToString();
                break;
            case 6:
                result = material[Guids.Props.Напряженность6].Value.ToString();
                break;
            default:
                result = "Ошибка";
                break;
        }

        return result;
    }

    #endregion Методы для формирования магнитного протокола

    #region Методы для формирования химического протокола

    public string ОтображениеКомпонентаХимическаяЛаборатория(int index) {
        ReferenceObject electrolite = Context.ReferenceObject.GetObject(Guids.Links.СправочныеМатериалыХимическаяЛаборатория);
        if (electrolite == null)
            return string.Empty;
        // Получаем компоненты электролита
        List<ReferenceObject> components = electrolite.GetObjects(Guids.ListsOfObjects.КомпонентыЭлектролита);
        if (index > components.Count)
            return string.Empty;
        ReferenceObject component = components[index - 1];
        string name = (string)component[Guids.Props.НаименованиеКомпонента].Value;
        string value = (string)component[Guids.Props.СводноеСодержаниеКомпонента].Value;
        return $"{name} (по ТИ: '{value}')";
    }

    public bool СкрытиеЭлементаХимическаяЛаборатория(int index) {
        ReferenceObject electrolite = Context.ReferenceObject.GetObject(Guids.Links.СправочныеМатериалыХимическаяЛаборатория);
        if (electrolite == null)
            return true;
        List<ReferenceObject> components = electrolite.GetObjects(Guids.ListsOfObjects.КомпонентыЭлектролита);
        if (index > components.Count)
            return true;
        return false;
    }

    #endregion Методы для формирования химического протокола

    #region Получение данных для отчета

    // Данный метод создан для использлования в "Выполнить макрос", который выполняется при формировании отчета
    // (вернее, при попытке получить параметр DataTable)
    // (для получения всех данных в вложенных в объект списках в виде единой строки)
    public string ПолучитьДанныеДляОтчета() {
        // Способ получения данных для отчета будет определяться в зависимости от типа отчета
        
        ReferenceObject protocol = Context.ReferenceObject;
        Guid guidOfType = protocol.Class.Guid;

        string result = string.Empty;

        if (protocol.Class.IsInherit(Guids.Types.ПротоколМеталлографическойЛаборатории)) {
            result = GetDataStringDefault(protocol);
        }
        else if (protocol.Class.IsInherit(Guids.Types.ПротоколФизикомеханическойЛаборатории)) {
            result = GetDataStringDefault(protocol);
        }
        else if (protocol.Class.IsInherit(Guids.Types.ПротоколСпектральнойЛаборатории)) {
        }
        else if (protocol.Class.IsInherit(Guids.Types.ПротоколХимическойЛаборатории)) {
            result = GetDataStringForChemProtocol(protocol);
        }
        else if (protocol.Class.IsInherit(Guids.Types.ПротоколГальваническойЛаборатории)) {
        }
        else if (protocol.Class.IsInherit(Guids.Types.ПротоколМагнитнойЛаборатории)) {
            result = GetDataStringForMagnetProtocol(protocol);
        }
        else if (protocol.Class.IsInherit(Guids.Types.ПротоколЭлектрическойЛаборатории)) {
        }
        else {
        }

        return result;
    }

    private string GetDataStringDefault(ReferenceObject protocol) {

        const int безТаблицы = 0;
        const int таблицаПоОбразцам = 1;
        const int таблицаПоПоказателям = 2;
        const int таблицаПоКонтролируемымПараметрам = 3;

        string result = string.Empty;
        // Класс для генерации строки, которая будет содержать табличные данные
        DataClass resultDataClass = new DataClass();

        switch ((int)protocol[Guids.Props.ПараметрВидТаблицы].Value) {
            case безТаблицы:
                break;
            case таблицаПоОбразцам:
                // код для получения выходных данным с вложенного списка "Образцы"
                resultDataClass.Add("Вид испытания");
                resultDataClass.Add("Сводный размер, мм");
                resultDataClass.Add("Предел прочности, МПа");
                resultDataClass.Add("Предел текучести, МПа");
                resultDataClass.Add("Относительное удлинение, %");
                resultDataClass.Add("KCU, Дж/см\u00B2");
                resultDataClass.EndRow();

                foreach (ReferenceObject sample in protocol.GetObjects(Guids.ListsOfObjects.Образцы)) {
                    resultDataClass.Add(sample[Guids.Props.Испытание].Value.ToString());
                    resultDataClass.Add(sample[Guids.Props.СводныйРазмер].Value.ToString());
                    resultDataClass.Add(sample[Guids.Props.ПределПрочности].Value.ToString());
                    resultDataClass.Add(sample[Guids.Props.ПределТекучести].Value.ToString());
                    resultDataClass.Add(sample[Guids.Props.ОтносительноеУдлинение].Value.ToString());
                    resultDataClass.Add(sample[Guids.Props.ПрочностьПриИзгибе].Value.ToString());
                    resultDataClass.EndRow();
                }

                result = resultDataClass.GenerateString();
                break;
            case таблицаПоПоказателям:
                // Код для получения данных с вложенного списка "Показатели"
                resultDataClass.Add("Наименование");
                resultDataClass.Add("Данные по ТУ");
                resultDataClass.Add("Фактические данные");
                resultDataClass.EndRow();

                foreach (ReferenceObject param in protocol.GetObjects(Guids.ListsOfObjects.Показатели)) {
                    resultDataClass.Add(param[Guids.Props.Наименование].Value.ToString());
                    resultDataClass.Add(param[Guids.Props.ДанныеПоТУ].Value.ToString());
                    resultDataClass.Add(param[Guids.Props.ФактическиеДанные].Value.ToString());
                    resultDataClass.EndRow();
                }
                result = resultDataClass.GenerateString();
                break;
            case таблицаПоКонтролируемымПараметрам:
                // код для получения данных с вложенного списка "Контролируемые параметры"
                resultDataClass.Add("Контролируемый параметр");
                resultDataClass.Add("Пункт требований НТД");
                resultDataClass.Add("Норма");
                resultDataClass.Add("Фактические показатели");
                resultDataClass.EndRow();
                
                bool isFirstValue;
                foreach (ReferenceObject controlParam in protocol.GetObjects(Guids.ListsOfObjects.КонтролируемыеПараметры)) {
                    isFirstValue = true;
                    foreach (ReferenceObject factVal in controlParam.GetObjects(Guids.ListsOfObjects.ФактическиеПоказания)) {
                        if (isFirstValue) {
                            resultDataClass.Add(controlParam[Guids.Props.КонтролируемыйПараметр].Value.ToString());
                            resultDataClass.Add(controlParam[Guids.Props.ПунктТребованийНТД].Value.ToString());
                            resultDataClass.Add(controlParam[Guids.Props.Норма].Value.ToString());
                            resultDataClass.Add(string.Format("{0}: {1}",
                                                            factVal[Guids.Props.НаименованиеИзмеряемогоОбъекта].Value.ToString(),
                                                            factVal[Guids.Props.ФактическоеЗначение].Value.ToString()));
                        }
                        else {
                            resultDataClass.Add("---");
                            resultDataClass.Add("---");
                            resultDataClass.Add("---");
                            resultDataClass.Add(string.Format("{0}: {1}",
                                                            factVal[Guids.Props.НаименованиеИзмеряемогоОбъекта].Value.ToString(),
                                                            factVal[Guids.Props.ФактическоеЗначение].Value.ToString()));
                        }
                        isFirstValue = false;
                        resultDataClass.EndRow();
                    }
                }
                result = resultDataClass.GenerateString();
                break;
            default:
                System.Windows.Forms.MessageBox.Show("При попытке передать в отчет табличные данные возникла ошибка. Выбран неизвестный вид таблицы");
                break;
        }

        return result;
    }

    private string GetDataStringForMagnetProtocol(ReferenceObject protocol) {
        
        // Создаем Helper метод, который предназначен для генерации итоговой строки
        DataClass resultDataClass = new DataClass();
        
        // В зависимости от типа исследования, которое выбрано в протоколе, формирует разные итоговые строки
        if ((int)protocol[Guids.Props.ТипИсследования] == 0) {
            // Получаем привязанный к протоколу материал для того, чтобы получить с него допустимые значение и количество требуемых замеров
            ReferenceObject material = protocol.GetObject(Guids.Links.СправочныеМатериалыМагнитнаяЛаборатория);
            // Если материал к протоколу не подключен, передаем пустую строку в генерацию отчета.
            // Сделано это потому, что в данный момент без материала отключается возможность заполнения
            // параметров измерения контрольного образца, а, следовательно, без подключенного материала
            // и отчет не должен формироваться
            if (material == null) {
                Message("Ошибка", "К протоколу не подключен материал");
                return string.Empty;
            }

            // Заполняем шапку таблицы
            // Получаем количество замеров с материала
            int numberOfMeasurments = (int)material[Guids.Props.КоличествоЗамеров].Value;

            // Добавляем первые две колонки, которые будут в коде генерации отчета схлопываться с верхним уровнем заголовка
            resultDataClass.Add(""); // Марка материала 
            resultDataClass.Add(""); // Номер контрольного образца
            
            // Добавляем нужное количество колонок с замерами магнитной индукции
            resultDataClass.Add(material[Guids.Props.Напряженность1].Value.ToString()); // Первый замер добавляем в любом случае
            if (numberOfMeasurments > 1)
                resultDataClass.Add(material[Guids.Props.Напряженность2].Value.ToString()); // Первый замер добавляем в любом случае
            if (numberOfMeasurments > 2)
                resultDataClass.Add(material[Guids.Props.Напряженность3].Value.ToString()); // Первый замер добавляем в любом случае
            if (numberOfMeasurments > 3)
                resultDataClass.Add(material[Guids.Props.Напряженность4].Value.ToString()); // Первый замер добавляем в любом случае
            if (numberOfMeasurments > 4)
                resultDataClass.Add(material[Guids.Props.Напряженность5].Value.ToString()); // Первый замер добавляем в любом случае
            if (numberOfMeasurments > 5)
                resultDataClass.Add(material[Guids.Props.Напряженность6].Value.ToString()); // Первый замер добавляем в любом случае

            // Добавялем заключительную пустую колонку под объединение с верхнем уровнем заголовка (Коэрциальная сила)
            resultDataClass.Add("");
            resultDataClass.EndRow();


            // Добавление регулярной части таблицы
            int counter = 1;
            foreach (ReferenceObject sample in protocol.GetObjects(Guids.ListsOfObjects.ОбразцыМагнитнойЛаборатории)) {
                resultDataClass.Add(sample[Guids.Props.СводноеНаименованиеОбразцаМагнитнойЛаборатории].Value.ToString());
                resultDataClass.Add(counter.ToString()); counter++;

                // Добавление колонок с измерениями магнитной индукции будут зависеть от параметра numberOfMeasurments
                resultDataClass.Add(sample[Guids.Props.МагнитнаяИндукцияОбразца1].Value.ToString());
                if (numberOfMeasurments > 1)
                    resultDataClass.Add(sample[Guids.Props.МагнитнаяИндукцияОбразца2].Value.ToString());
                if (numberOfMeasurments > 2)
                    resultDataClass.Add(sample[Guids.Props.МагнитнаяИндукцияОбразца3].Value.ToString());
                if (numberOfMeasurments > 3)
                    resultDataClass.Add(sample[Guids.Props.МагнитнаяИндукцияОбразца4].Value.ToString());
                if (numberOfMeasurments > 4)
                    resultDataClass.Add(sample[Guids.Props.МагнитнаяИндукцияОбразца5].Value.ToString());
                if (numberOfMeasurments > 5)
                    resultDataClass.Add(sample[Guids.Props.МагнитнаяИндукцияОбразца6].Value.ToString());

                // Добавление последних колонок, которые не не изменяются динамически
                //resultDataClass.Add(sample[Guids.Props.МаксимальнаяПроницаемостьОбразца].Value.ToString());
                resultDataClass.Add(sample[Guids.Props.КоэрцитивнаяСилаОбразца].Value.ToString());
                resultDataClass.EndRow();
            }


            
            // Добавление итоговой части таблицы, которая содержит допустимые значения по ГОСТ
            resultDataClass.Add(string.Format("По {0} {1}",
                        material[Guids.Props.ТипСтандартаМагнитнаяЛаборатория].Value.ToString(),
                        material[Guids.Props.СтандартМагнитнаяЛаборатория].Value.ToString()));
            resultDataClass.Add(string.Empty);

            // Заполнение динамически меняющихся колонок (магнитная индукция)
            resultDataClass.Add(material[Guids.Props.МагнитнаяИндукция1].Value.ToString());
            if (numberOfMeasurments > 1)
                resultDataClass.Add(material[Guids.Props.МагнитнаяИндукция2].Value.ToString());
            if (numberOfMeasurments > 2)
                resultDataClass.Add(material[Guids.Props.МагнитнаяИндукция3].Value.ToString());
            if (numberOfMeasurments > 3)
                resultDataClass.Add(material[Guids.Props.МагнитнаяИндукция4].Value.ToString());
            if (numberOfMeasurments > 4)
                resultDataClass.Add(material[Guids.Props.МагнитнаяИндукция5].Value.ToString());
            if (numberOfMeasurments > 5)
                resultDataClass.Add(material[Guids.Props.МагнитнаяИндукция6].Value.ToString());

            //resultDataClass.Add(material[Guids.Props.МаксимальнаяПроницаемость].Value.ToString());
            resultDataClass.Add(material[Guids.Props.КоэрцитивнаяСила].Value.ToString());
            resultDataClass.EndRow();

            return resultDataClass.GenerateString();
        }
        else {
            // Случай, когда выбрано исследование на удельные потери
            
            resultDataClass.Add(""); // Марка материала
            resultDataClass.Add(""); // Номер контрольного образца

            // Добавляем колонки (удельные потери)
            resultDataClass.Add("P = 1,8/400");
            resultDataClass.Add("P = 2/400");
            resultDataClass.EndRow();

            // Приступаем к заполнению регулярной части
            int counter = 1;
            foreach (ReferenceObject sample in protocol.GetObjects(Guids.ListsOfObjects.ОбразцыМагнитнойЛаборатории)) {
                resultDataClass.Add(sample[Guids.Props.СводноеНаименованиеОбразцаМагнитнойЛаборатории].Value.ToString());
                resultDataClass.Add(counter.ToString()); counter++;

                // Добавление значений удельных потерь для разных значений P
                resultDataClass.Add(sample[Guids.Props.УдельныеПотери1].Value.ToString());
                resultDataClass.Add(sample[Guids.Props.УдельныеПотери2].Value.ToString());
                resultDataClass.EndRow();
            }

            return resultDataClass.GenerateString();
        }
    }

    private string GetDataStringForChemProtocol(ReferenceObject protocol) {
        DataClass resultDataClass = new DataClass();

        // INFO: Данный код я пишу для постоянного количество колонок, которых будет девять (восемь под компоненты)
        int countOfColumns = 9;
        ReferenceObject electrolite = protocol.GetObject(Guids.Links.СправочныеМатериалыХимическаяЛаборатория); // Получаем электролит
        List<ReferenceObject> components = electrolite != null ?
            electrolite.GetObjects(Guids.ListsOfObjects.КомпонентыЭлектролита) :
            new List<ReferenceObject>(); // Получаем компоненты
        int countOfComponents = components.Count; // Получаем количество компонентов (и, следовательно, количество колонок, которые следует заполнять)


        // Заполняем первую строку
        resultDataClass.Add("Электролит");
        // Заполняем название электролита
        string nameOfElectrolite = electrolite != null ? (string)electrolite[Guids.Props.НаименованиеЭлектролита].Value : string.Empty;
        resultDataClass.Add(nameOfElectrolite);
        // Добавляем объединение строк
        for (int i = 0; i < countOfColumns - 2; i++)
            resultDataClass.Add(string.Empty);
        resultDataClass.EndRow();

        // Заполняем вторую строку
        resultDataClass.Add("Компоненты");
        // Получаем названия всех компонентов
        List<string> componentNames = components.Select(comp => (string)comp[Guids.Props.НаименованиеКомпонента].Value).ToList<string>();
        for (int i = 0; i < countOfColumns - 1; i++) {
            if (i < countOfComponents)
                resultDataClass.Add(componentNames[i]);
            else
                resultDataClass.Add("-");
        }
        resultDataClass.EndRow();

        // Заполняем третью строку
        resultDataClass.Add("Содержание г/л");
        // Получаем результаты замеров
        List<string> measurements = new List<string>() {
            ((double)protocol[Guids.Props.СодержаниеКомпонента1].Value).ToString(),
            ((double)protocol[Guids.Props.СодержаниеКомпонента2].Value).ToString(),
            ((double)protocol[Guids.Props.СодержаниеКомпонента3].Value).ToString(),
            ((double)protocol[Guids.Props.СодержаниеКомпонента4].Value).ToString(),
            ((double)protocol[Guids.Props.СодержаниеКомпонента5].Value).ToString(),
            ((double)protocol[Guids.Props.СодержаниеКомпонента6].Value).ToString(),
            ((double)protocol[Guids.Props.СодержаниеКомпонента7].Value).ToString(),
            ((double)protocol[Guids.Props.СодержаниеКомпонента8].Value).ToString()
        };
        for (int i = 0; i < 8; i++) {
            if (i < countOfComponents)
                resultDataClass.Add(measurements[i]);
            else
                resultDataClass.Add("-");
        }
        resultDataClass.EndRow();
        
        // Заполняем четвертую строку
        resultDataClass.Add("Содержание по ТИ");
        List<string> allowedValues = components.Select(comp => ((string)comp[Guids.Props.СводноеСодержаниеКомпонента]).ToString()).ToList<string>();
        for (int i = 0; i < countOfColumns - 1; i++) {
            if (i < countOfComponents)
                resultDataClass.Add(allowedValues[i]);
            else
                resultDataClass.Add("-");
        }
        resultDataClass.EndRow();

        return resultDataClass.GenerateString();
    }

    private class DataClass {
        private string valueSplitter = "^";
        private string rowSplitter = "@";
        private List<string> intermediateResult = new List<string>();
        private List<string> result = new List<string>();

        public DataClass() {
        }

        public DataClass(string valueSplitter, string rowSplitter) {
            this.valueSplitter = valueSplitter;
            this.rowSplitter = rowSplitter;
        }

        public void Add(string value) {
            // Если значение содержит 0, мы не заполняем данную колонку
            if (value == "0")
                value = "-";
            this.intermediateResult.Add(value);
        }

        public void EndRow() {
            result.Add(string.Join(this.valueSplitter, this.intermediateResult));
            this.intermediateResult.Clear();
        }

        public string GenerateString() {
            if (intermediateResult.Count != 0) {
                this.EndRow();
            }
            string result = string.Join(this.rowSplitter, this.result);
            this.result.Clear();
            return result;
        }
    }
    #endregion Получение данных для отчета

    #region Шаблонное заполнение данных

    // Код, который будет заполнять новые объекты в зависимости от того, какой протокол создается
    // TODO

    #endregion Шаблонное заполнение данных

    #region Присвоение дефолтного вида таблицы в зависимости от типа протокола

    // Данный код будет выполняться при любом создании новых объектов в справочнике Архив ЦЗЛ
    public void НазначитьВидТаблицы() {
        ReferenceObject currentObject = Context.ReferenceObject;
        // Вычисляем тип объекта
        if (currentObject.Class.IsInherit(Guids.Types.ПротоколМеталлографическойЛаборатории)) {
            currentObject[Guids.Props.ПараметрВидТаблицы].Value = 0;
        }
        else if (currentObject.Class.IsInherit(Guids.Types.ПротоколФизикомеханическойЛаборатории)) {
            currentObject[Guids.Props.ПараметрВидТаблицы].Value = 1;
        }
        else if (currentObject.Class.IsInherit(Guids.Types.ПротоколСпектральнойЛаборатории)) {
            currentObject[Guids.Props.ПараметрВидТаблицы].Value = 0;
        }
        else if (currentObject.Class.IsInherit(Guids.Types.ПротоколХимическойЛаборатории)) {
            currentObject[Guids.Props.ПараметрВидТаблицы].Value = 0;
        }
        else if (currentObject.Class.IsInherit(Guids.Types.ПротоколГальваническойЛаборатории)) {
            currentObject[Guids.Props.ПараметрВидТаблицы].Value = 0;
        }
        else if (currentObject.Class.IsInherit(Guids.Types.ПротоколМагнитнойЛаборатории)) {
            currentObject[Guids.Props.ПараметрВидТаблицы].Value = 3;
        }
        else if (currentObject.Class.IsInherit(Guids.Types.ПротоколЭлектрическойЛаборатории)) {
            currentObject[Guids.Props.ПараметрВидТаблицы].Value = 0;
        }
        else {
            Error(string.Format("Не получилось определить вариант таблицы для объекта '{0}'", currentObject.ToString()));
        }
    }

    public void ЗаполнитьПолеАвторПротокола() {
        ReferenceObject currentObject = Context.ReferenceObject;
        currentObject[Guids.Props.АвторПротокола].Value = CurrentUser.ToString();
    }

    #endregion Присвоение дефолтного вида таблицы в зависимости от типа протокола
}
