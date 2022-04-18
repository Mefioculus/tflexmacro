using System;
using System.Text;
using System.IO;
using System.Linq;
using System.Collections;
using System.Collections.Generic;
using System.Windows.Forms;

using TFlex.DOCs.Model;
using TFlex.DOCs.Common.Encryption; // Для осуществления подключения к серверу
using TFlex.DOCs.Model.Macros;
using TFlex.DOCs.Model.Macros.ObjectModel;
using TFlex.DOCs.Model.Structure;
using TFlex.DOCs.Model.Search;
using TFlex.DOCs.Model.Classes;
using TFlex.DOCs.Model.References;
using TFlex.DOCs.Model.References.Files;
using TFlex.DOCs.Model.References.Events;

/*
 * Для работы макроса так же потребуется подключение в качестве ссылки библиотеки TFlex.DOCs.Common.dll
*/

public class Macro : MacroProvider {
    public Macro(MacroContext context)
        : base (context) {
        }

    public override void Run() {
        int indent = 5;

        // Названия полей
        string serverNameField = "Имя сервера";
        string frendlyNameField = "Короткое название базы";
        string userNameField = "Имя пользователя";
        string passwordField = "Пароль";
        string pathToSaveDiffField = "Путь для сохранения отчета";
        string directionOfCompareField = "Обратное сравнение";
        string excludedFromComparingFields = "Исключенные поля";
        string excludeReferences = "Исключить справочники из ограничительного списка";

        // Переменные

        string serverName = "TFLEX-DOCS:21324";
        string userName = "Gukovry";
        string frendlyName = "Макет";
        string pathToFile = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "result.txt");
        string password = "123";

        // Перечень полей, по которым идет сравнение
        List<string> allFields = new List<string>() {
            "id",
            "Наименование",
            "Псевдоним",
            "КолСобытий",
            "НаимСобытий",
            "КолОбработчиков",
            "НаимОбработчиков",
            "ДлинаПараметра",
            "ВозможноНулевоеЗначение",
            "ТипПараметра",
            "ЕдИзПараметра",
            "ЗначенияСпискаПараметра",
            "ТипСсылки",
            "КоличествоСтраницДиалогов"
        };
        
        List<Guid> excludedReferences = new List<Guid>() {
            new Guid("21950e00-d388-44e5-a102-f5a3fd1aadbf"), // Справочник ADUSERS
            //new Guid("fe60bf63-8cc3-4a25-bc46-798a0e09dcdb"), // Справочник Cортамент материалов
            //new Guid("22ee7b25-0744-405c-986e-01d62bfed86d"), // Справочник KAT_EDIZ
            //new Guid("c4b21dd6-fea7-4af3-baa1-9addfdd54452"), // Справочник KAT_EDIZ OUT
            //new Guid("748a9ca7-931e-4db3-88da-4f42d9c0157b"), // Справочник KAT_INS
            //new Guid("fed419c0-2106-4951-b609-876bf2428102"), // Справочник KAT_INS OUT
            new Guid("a7b0d67a-9da3-441b-aad4-e6d951a188f4"), // Справочник KAT_IZD
            //new Guid("95532376-2654-4bfb-80ae-1f059420259e"), // Справочник KAT_IZD OUT
            //new Guid("071cc472-1909-43f8-b300-d9f1351d72ca"), // Справочник KAT_IZV
            //new Guid("cb71a894-892a-418b-867a-c085b70c13f7"), // Справочник KAT_IZVP
            //new Guid("dcb2ac46-edf8-4686-9c2b-67c9e54ab18d"), // Справочник KAT_IZVT
            //new Guid("af925d2a-a07d-483e-8b8c-71658d958c2f"), // Справочник KAT_PODR
            //new Guid("4ba54c93-5bd1-4b2a-9761-d8d4cddef8c5"), // Справочник KAT_PODR OUT
            //new Guid("ddd085f9-8c74-4b61-b248-d66cd8a2e35e"), // Справочник KAT_STOL
            //new Guid("5e8b4498-6f42-4456-a71a-ebeb9cf93adf"), // Справочник KAT_VHOD
            new Guid("443d8e71-1385-459d-879e-5e03c19b9d49"), // Справочник KLAS
            //new Guid("4d2872e6-6854-4cb4-ba5c-e0378f169d8a"), // Справочник KLAS OUT
            new Guid("e253835e-260e-45f3-8ae8-600a7a8907a2"), // Справочник KLASM
            //new Guid("31e86a26-0e29-4ee1-b78c-c0bdeb0f4add"), // Справочник KLASM_OUT
            //new Guid("dda8e2f7-67e9-41ea-a8db-cbe6c9b7eb9b"), // Справочник MARCHP
            //new Guid("44a6010a-af40-4008-8431-e1d02e995d2a"), // Справочник MARCHP OUT
            new Guid("7e59b31a-eab5-4da5-938f-812703046345"), // Справочник NORM
            //new Guid("991875e9-7376-4eb0-884f-63dee51a6d66"), // Справочник NORM OUT
            //new Guid("d5368f85-1fb8-45d7-a3ae-5c65669aadc6"), // Справочник OperAem
            //new Guid("e557245a-88eb-4fcd-93d4-7ee47ce9bd07"), // Справочник OSNAST
            //new Guid("1f6be21b-1208-4694-9fb2-394cf94dff8f"), // Справочник OSNAST OUT
            new Guid("610f6a2d-a019-4174-a0f5-769ea0740607"), // Справочник Poluf
            //new Guid("52eac538-cbfc-42ed-9753-b3b4c37e460d"), // Справочник Poluf OUT
            //new Guid("ce452941-d3ea-42b2-a16a-43e048ffe02b"), // Справочник RZ_ALL
            new Guid("c587b002-3be2-46de-861b-c551ed92c4c1"), // Справочник SPEC
            //new Guid("cff14915-9df7-4863-a3a9-0c0a0ad9bffd"), // Справочник SPEC_OUT
            //new Guid("af4a7b46-b290-428b-bc92-00b5189d6a46"), // Справочник Web-сервисы
            //new Guid("f7f43d73-857c-41f9-b449-38ee72caa221"), // Справочник Архив ЦЗЛ
            //new Guid("ad13a201-a1bd-4597-a4e6-90e50cd58e14"), // Справочник База знаний
            //new Guid("5ae00fb9-a666-4a7b-b244-50f23a758fb3"), // Справочник Варианты единиц измерения
            //new Guid("2e6a9419-e008-4f45-901a-30ce2caff481"), // Справочник Виды доставки
            //new Guid("b069af26-b8c5-4fdf-a3eb-ab62f226a8fa"), // Справочник Виды технологических процессов
            //new Guid("bb777c69-0ae9-4b67-8bc3-376daa5f683b"), // Справочник Виды технологической документации
            //new Guid("fd8c8718-f22e-494b-ade5-cdca91704178"), // Справочник Визуальные свойства материалов
            //new Guid("c789eb9e-fb89-4fff-86e3-5b4e24dfa924"), // Справочник Внешние приложения
            //new Guid("07a3cde7-5c60-445c-8e36-c8efad241e9e"), // Справочник Временные комплекты ролей
            new Guid("9450b263-b4e4-4893-97e9-98c3b9735bca"), // Справочник Вспомогательный справочник по ДМ
            //new Guid("f29f7652-310a-4438-acff-2b6c6afb0852"), // Справочник Выгрузки из FoxPro
            //new Guid("79b81852-bd07-4375-b653-bbf1c568a2c7"), // Справочник Генераторы отчётов
            //new Guid("6dcdc95f-993b-4666-8136-7ee9d29b6d13"), // Справочник Глобальные параметры
            //new Guid("90cbc703-658d-47a8-9a2b-c6711cd34e72"), // Справочник Глоссарий
            //new Guid("0952b11f-5f03-4fb5-b48e-6cdbd5aae607"), // Справочник Государства
            //new Guid("0c6cd73f-214a-47e8-933c-68a510556567"), // Справочник Граждане
            //new Guid("19c04ebe-a2ba-4dab-9304-29e82f9f2cf3"), // Справочник Групповые задания
            //new Guid("8ee861f3-a434-4c24-b969-0896730b93ea"), // Справочник Группы и пользователи
            //new Guid("8af2b852-a934-46b9-ad4d-2ab75c65b5d1"), // Справочник Группы рассылки
            //new Guid("423839f0-4705-4252-b8c1-011a938ef21f"), // Справочник Действия
            //new Guid("60e48c95-9d85-4a04-8bed-795bc65b61a4"), // Справочник Действия процессов
            //new Guid("71f942b1-d103-4e9e-baba-de756a6b2040"), // Справочник Договоры
            //new Guid("ac46ca13-6649-4bbb-87d5-e7d570783f26"), // Справочник Документы
            //new Guid("500d4bcf-e02c-4b2e-8f09-29b64d4e7513"), // Справочник Документы ОГТ
            new Guid("ad862a39-3cf4-4ac5-8b12-a4a3b286c52b"), // Справочник Должность
            //new Guid("5759a3ba-11fa-418f-9407-78af48bdfbb2"), // Справочник Дополнительные технологические параметры
            //new Guid("894fc487-7c77-4e51-902b-9037aed251d5"), // Справочник Допуски и отклонения
            //new Guid("01c51d4c-e07d-4f31-9346-5697399a09fb"), // Справочник Единицы измерения
            //new Guid("b57b3b42-be7f-4e1a-baed-4d41306c384b"), // Справочник Единицы измерения ТМЦ
            //new Guid("6c0a813b-a382-4267-975e-41c769d7299e"), // Справочник Журнал выгрузок в FoxPro
            //new Guid("86226d08-6fee-49b7-b170-f92d1d125eab"), // Справочник Журнал выполнения макросов
            //new Guid("124debf1-82b8-4f01-af45-04dd8b821b19"), // Справочник Журнал обмена данными
            //new Guid("b9acb7eb-13b5-49b7-94d6-987b4adc54d8"), // Справочник Журнал ознакомления с бумажными документами
            //new Guid("1389fba1-73ee-483f-a12c-733dd51d0dba"), // Справочник Журнал передачи бумажных документов
            //new Guid("0b60a94f-ead6-4770-b1d2-e6f39c7f0716"), // Справочник Журнал регистрации
            //new Guid("df54c55b-6381-40a4-9122-bf5a2a03e97f"), // Справочник Журнал синхронизации проектов
            //new Guid("d615a9f7-f90c-425c-b92b-4bd9780db3d1"), // Справочник Зависимости замечаний
            //new Guid("e13cee45-39fa-43ff-ba2a-957294d975bf"), // Справочник Зависимости проектов
            //new Guid("8d727772-d7e5-4058-b7e1-046c510e7f76"), // Справочник Заказы на оснастку
            //new Guid("a34c879d-2855-480e-8fd2-24e96f5db1c6"), // Справочник Замечания
            new Guid("9e21f099-73df-4948-b8c2-22ffd731a5e8"), // Справочник Заявки на НСИ
            //new Guid("17fa3249-60a9-4b82-9b22-68881ceca78b"), // Справочник Значения опций изделий
            //new Guid("4853c5ce-6b8a-48ac-94bf-e80cdbcb4c1b"), // Справочник Извещения об изменениях
            //new Guid("85482ef8-6ed6-42a2-adb7-36db7c9a06ff"), // Справочник Изготовление оснастки
            //new Guid("c9a4bb1b-cacb-4f2d-b61a-265f1bfc7fb9"), // Справочник Изменения
            //new Guid("98c4aa69-72f4-4af2-8520-865146f5e67d"), // Справочник Изменения НД
            //new Guid("b0eb4a0c-3f96-4cd4-a08a-eccc64541265"), // Справочник Изменения ОГТ
            //new Guid("db7b5def-dd0e-4d48-996e-ea0ed12ab5ea"), // Справочник Изображения
            //new Guid("f1d28c5d-6b08-4061-b81f-cde490088e1f"), // Справочник Инвентарная книга
            //new Guid("9e67a0c7-b01b-459b-ab82-5d978557b099"), // Справочник Инструкции
            //new Guid("0c470782-0064-470d-b05d-7bb4aea4c6d6"), // Справочник Исполнители операции
            //new Guid("3459a8fb-6bca-47ca-971a-1572b684e92e"), // Справочник Используемые ресурсы
            //new Guid("39b52d1c-23b6-4715-bd86-d7bcd39f3500"), // Справочник История синхронизации справочников
            //new Guid("4b79cfef-89a8-4cb0-8857-20691579c14e"), // Справочник Календари
            //new Guid("5cc572a7-12f8-402e-97c9-1318b6d68f9e"), // Справочник Каталог инструмента
            //new Guid("ddb18820-96b7-4344-8e12-70c6dce0fcae"), // Справочник Каталог оборудования
            //new Guid("a13d8c5a-c76c-4df6-91df-fc5bbf8db86d"), // Справочник Каталог оснащения
            //new Guid("71384391-a828-4861-926e-c2070b1b5bb2"), // Справочник Категории
            //new Guid("b97d5ae6-2abb-494e-a31d-5abb4a443303"), // Справочник Категории изделий
            //new Guid("7e5a07a0-35c0-49d6-b85e-f73d438e8cd1"), // Справочник Классификатор ЕСКД
            //new Guid("07e0dfaa-305e-468b-acd7-b6dad340da8f"), // Справочник Классификатор изделий
            //new Guid("5aa3157d-c41b-47e3-8455-4dbd0d13b8eb"), // Справочник Классификатор оснастки
            //new Guid("6683b47f-9a62-4c61-9147-3310843892cb"), // Справочник Классификаторы
            //new Guid("bcdde5db-eea4-4917-9b08-646bc82b1a4f"), // Справочник Кодификатор
            //new Guid("454c9856-189f-4a53-a2d5-0691dc34c85e"), // Справочник Комплекты документов
            //new Guid("3e4dcbd5-2b16-49e8-b4ca-1ad5c9202b06"), // Справочник Комплекты ролей
            //new Guid("e814922d-10e3-49ee-9920-aaea2e19cc4d"), // Справочник Контакты
            //new Guid("f7e437a0-0c69-4ed0-be0d-4aa0b46914bf"), // Справочник Контексты проектирования
            new Guid("4b5d1fdf-2e04-488f-9043-98d0d09f31d0"), // Справочник Контракты
            //new Guid("aa29ada3-0bb1-416b-a760-ea8bb0aa9c56"), // Справочник Кэширующие файловые серверы
            //new Guid("c78e85d7-1c95-4572-85d8-468f9fd886de"), // Справочник Магнитная лаборатория
            //new Guid("3e6df4d0-b1d8-4375-978c-4da676604cca"), // Справочник Макросы
            //new Guid("29d98631-8d3e-4db1-b81e-eb37435a0c80"), // Справочник Марки материалов
            //new Guid("c5e7ae00-90f2-49e9-a16c-f51ed087752a"), // Справочник Материалы
            //new Guid("ff9a96d0-043c-4dd5-932a-7a561964eeb5"), // Справочник Материалы ТП
            //new Guid("1cd3d615-f117-47f7-a208-b65f73a50ce8"), // Справочник Межоперационное время
            //new Guid("d063cd9a-b772-40b9-851f-cfd84470517c"), // Справочник Модули конвертации файлов
            //new Guid("4d1d356f-5ff3-481f-8f99-0922eae42298"), // Справочник Наладки
            //new Guid("96045113-118c-4efd-90c6-96e49a5bcd04"), // Справочник Наряды
            //new Guid("371762b7-9ccf-4ce0-924d-6eb25f863d09"), // Справочник Населённые пункты
            //new Guid("36d114a1-dc45-483d-ba98-215627a63443"), // Справочник Настройки адресной книги
            //new Guid("769f1922-3937-4430-94d7-232790e3cecd"), // Справочник Настройки конфигурации
            //new Guid("16175b87-d5d6-4a22-a67f-710ea07c49a6"), // Справочник Настройки нумерации
            //new Guid("0ba877d8-b0c8-4ca3-b25b-361c5d0707c2"), // Справочник Номенклатура дел
            //new Guid("221ea415-75fc-458a-aa52-2144225fca43"), // Справочник Нормативные документы
            //new Guid("3bdbd32f-20be-4495-951b-0d3205376d94"), // Справочник Нумератор инвентарной книги
            //new Guid("487ca4c5-9ebc-4273-9a99-c5fef7f5f33e"), // Справочник Оборудование
            //new Guid("2b163e5d-1ac2-4d28-becb-c391282a4c0d"), // Справочник Оборудование_0820
            //new Guid("0f1b81d7-61e4-4d6d-a485-32a9b6bd2d90"), // Справочник Опции изделий
            //new Guid("b1e53a83-9d59-43c8-9715-594eb8cd0df6"), // Справочник Организации
            //new Guid("cc38c226-7264-4813-9900-331ce325d517"), // Справочник Оснащение ТП
            //new Guid("d3396686-2cb9-44ff-994b-d446c0a42515"), // Справочник Отчёты
            //new Guid("84451595-40ca-41c0-88ad-dda393b28fe7"), // Справочник Офисные документы
            //new Guid("93e60269-1440-4371-a734-430115e0df00"), // Справочник Очередь заданий конвертации
            //new Guid("a08bf795-c2ac-4306-9028-b35ae7d4f788"), // Справочник Панель навигации
            //new Guid("9b952f86-c77d-4ce1-85fa-9266efa9cc42"), // Справочник Папки
            //new Guid("cb825bbe-a119-4712-a610-e86a59ad1450"), // Справочник Папки поручений
            //new Guid("a787c3db-c9c5-4a53-8ce7-07e923d08481"), // Справочник Параметры заготовок
            new Guid("46565316-3446-4f65-aa8c-fa6bdef4fdbd"), // Справочник Параметры заявки
            //new Guid("ab97c7bb-f3d1-44c6-88c4-7aaaeec64016"), // Справочник Параметры КТЭ
            //new Guid("e8c275b0-c9cb-43e1-a99d-46a1f8b30672"), // Справочник Параметры режимов обработки
            //new Guid("10ba3bc0-deed-4700-8f25-9e38efa7481b"), // Справочник Партии
            //new Guid("d2dc7c26-47b1-4b45-9389-e320987642e5"), // Справочник Переменные данные ТП
            new Guid("9dd79ab1-2f40-41f3-bebc-0a8b4f6c249e"), // Справочник Подключения
            //new Guid("a8bdff2f-d493-456d-b48e-489fa473d631"), // Справочник Позиции
            //new Guid("41470eff-ccd0-41cc-86c7-c52ff7c33108"), // Справочник Поисковые запросы
            //new Guid("c2f8d2b9-3ad1-46f9-bf26-02cbd69b2e75"), // Справочник Покрытия
            //new Guid("a752d598-799c-4df7-9c7d-41e8b7912842"), // Справочник Пользовательские диалоги
            //new Guid("23b96dac-400b-44cd-b7f0-36512062db01"), // Справочник Поручения
            //new Guid("ad3f4e41-959f-41ea-8b07-b90384119491"), // Справочник Поставщики
            //new Guid("5282f201-f483-449e-9a89-918e32f2cb2b"), // Справочник Правила именования ревизий
            //new Guid("67892771-adc2-4e78-b052-302ad7d10915"), // Справочник Правила настройки интеграции приложений
            //new Guid("212a5ec8-3f36-4501-bb46-082d200ba05f"), // Справочник Правила обмена данными
            //new Guid("805e0647-cd5d-4e55-bb3a-61069ab6d226"), // Справочник Правила подбора припуска
            //new Guid("84c51d0b-3700-409b-9b81-912b548cbadf"), // Справочник Правила подключения
            //new Guid("c9a48b8e-8af4-483e-8c92-f0f39f3df65f"), // Справочник Применяемость в изделиях
            new Guid("48ccc498-f374-4551-9512-deae08716ec3"), // Справочник применяемость_test
            //new Guid("17ebbca7-fbb0-4e96-bd1b-130d7e8451dc"), // Справочник Принтеры
            //new Guid("d5207973-c0be-4bff-80bf-9c828532ab52"), // Справочник Припуски на наружное точение
            //new Guid("8bc6f591-576a-40ec-a3c6-b9bb70b15d4d"), // Справочник Производственные заказы
            //new Guid("fae849c0-9e9c-4a1f-95bd-23da8a254d16"), // Справочник Производственные операции
            //new Guid("bb93bf1b-11fb-4c0e-a227-836047bbc4fc"), // Справочник Производственные планы
            //new Guid("2b688455-be77-408e-a7b8-efd6d840a36f"), // Справочник Пространства планирования
            //new Guid("af24cba8-f57f-4464-9906-857b50253318"), // Справочник Профессии
            //new Guid("61d922d0-4d60-49f1-a6a0-9b6d22d3d8b3"), // Справочник Процедуры
            //new Guid("e0c70b3c-bee1-4321-87b0-c44f3d8b5f68"), // Справочник Процессы
            //new Guid("54f02e16-0a3e-458c-8e40-d2920105b6f5"), // Справочник Рабочие центры
            //new Guid("b2d224d0-7b08-4fe9-b3ae-2a7f0a107d29"), // Справочник Разделы покупных изделий
            //new Guid("80bf50ce-c4f8-4e15-8f54-254174d2b4eb"), // Справочник Разделы спецификаций
            //new Guid("fed09770-fff0-4575-840d-f5292a732a93"), // Справочник Разделы файлового хранилища
            //new Guid("4c4383c3-8c9e-474f-a54b-eba92e4452d6"), // Справочник Распечатки
            //new Guid("ae95aab2-269a-4f4a-bbc6-23cfada9d2c6"), // Справочник Расписание рабочего времени
            //new Guid("35ab455a-7e71-4315-b09a-093fe7256164"), // Справочник Расчёты
            //new Guid("80831dc7-9740-437c-b461-8151c9f85e59"), // Справочник Регистрационно-контрольные карточки
            //new Guid("3595d4b3-8272-4229-83b0-1b0281685800"), // Справочник Реестр обозначений ЕСКД
            //new Guid("ac650fc9-6cae-42af-bc6a-a15ef3dd8fff"), // Справочник Реестр обозначений технологической документации
            //new Guid("d7d0c60d-f0d4-4c94-9501-33a917086805"), // Справочник Реестр счетчиков кодификатора
            //new Guid("fe80ab68-01e1-4a95-96cf-602ec877ff19"), // Справочник Ресурсы
            //new Guid("25655114-e17d-4df4-a121-3c746a201fa4"), // Справочник Роли
            //new Guid("0564ff25-ed71-477b-83e3-ac94be36d43f"), // Справочник Рубрикатор
            //new Guid("afd37b0d-a14e-423a-98c6-a0b30d79e3d3"), // Справочник Сертификаты
            //new Guid("f3074bf9-eea0-489d-b3f1-b00aebf6593f"), // Справочник Синхронизатор справочников
            //new Guid("36f16a0b-7315-4d6c-865a-103059b492d7"), // Справочник Словарь
            new Guid("fd3e2720-a307-4fb0-bf95-bd9c1c13d0af"), // Справочник События
            //new Guid("7f18c233-87ce-4966-af6a-e061189d3e76"), // Справочник События
            //new Guid("c285ad92-877f-47dc-ad9c-5b3c987f307a"), // Справочник Согласование
            //new Guid("c4f3d78b-bda4-4110-9eaf-256e0ff4d0b4"), // Справочник Соединение с 1С
            new Guid("c9d26b3c-b318-4160-90ae-b9d4dd7565b6"), // Справочник Список номенклатуры
            //new Guid("904fc7da-77df-4763-94b9-1ada11aafa4a"), // Справочник Средства технологического оснащения
            //new Guid("a4515680-da88-4693-ba85-69c10f8a0896"), // Справочник Стили отображения
            //new Guid("ff1ec76c-d338-4325-a031-be7efa230cf5"), // Справочник Стили элементов схем
            //new Guid("ebacc3dc-877d-40ec-b097-d78cafe2262b"), // Справочник Структура изделий
            //new Guid("2b610882-d4a3-41b3-859d-8545191f8671"), // Справочник Структурированные документы
            //new Guid("f6cc44ca-886b-404f-89e2-593e1754d278"), // Справочник Текущие действия процессов
            new Guid("fcac4fee-7fc6-48a6-acbf-60aa629cd272"), // Справочник Телефонный справочник
            //new Guid("3787cf43-578d-45a9-8494-5005f863b9e8"), // Справочник Темы и задачи
            //new Guid("eb9dfcdc-21f5-46a0-863d-3917d9002867"), // Справочник Территориальное деление
            new Guid("b467e977-da3d-4a58-89ee-73ce46fafc00"), // Справочник Тест_26012022
            //new Guid("e0ebbdf1-9a92-4b54-b490-7d37fb1cd51e"), // Справочник Технические требования к материалам
            //new Guid("94af857b-0898-4eb8-95ed-37f94a4ae723"), // Справочник Технические условия
            //new Guid("244bec06-cbb2-4688-8aad-be87a6f6a2b2"), // Справочник Технические условия на материалы
            //new Guid("fdad2725-b568-428a-b04d-08eb86ed80ac"), // Справочник Технологические переделы (виды обработки)
            //new Guid("353a49ac-569e-477c-8d65-e2c49e25dfeb"), // Справочник Технологические процессы
            //new Guid("ba44f819-a4be-4260-a0b1-a1a3f2d5e067"), // Справочник Технологические элементы
            //new Guid("91dd3751-4630-4a22-b008-9301f13befbf"), // Справочник Типы заготовок
            //new Guid("b6e2f4e4-1167-478b-94b2-deb0dded4e29"), // Справочник Типы структур изделий
            //new Guid("7f4be921-ca17-405b-b62b-8db3cfa5ff49"), // Справочник Типы технологических операций
            //new Guid("28029923-4832-4875-a274-dfd3628a9157"), // Справочник Требования к сортаменту
            //new Guid("ae32c9f4-40cd-47ca-9bb6-b2434cf6cb6f"), // Справочник Указания в извещении об изменениях
            //new Guid("e489dcb6-c03b-4a46-8a07-97fbd5739be4"), // Справочник Управление вариантами
            new Guid("86ef25d4-f431-4607-be15-f52b551022ff"), // Справочник Управление проектами
            //new Guid("165b6e04-3b90-4b47-9bda-d221b11b74fc"), // Справочник Упрощённый фильтр
            //new Guid("04f0baa6-de57-4f83-aecb-b6b22dd111b0"), // Справочник Учёт отработки ТД
            //new Guid("a0fcd27d-e0f2-4c5a-bba3-8a452508e6b3"), // Справочник Файлы
            //new Guid("fb3d1a16-cbd4-44f4-8e86-68e6deafcbd9"), // Справочник Физическая структура изделий
            //new Guid("17816113-feb7-44a1-9024-7052eb3d9b52"), // Справочник Форматы
            //new Guid("f4cf09ef-eb95-4711-ad13-7590d8d479cc"), // Справочник Характеристики изделий
            //new Guid("8c3bf2d7-78e9-47bd-91b6-0f4d8aba5541"), // Справочник Химическая лаборатория
            //new Guid("6c67d2ff-1fad-45c9-852c-695ce236e84b"), // Справочник Шаблоны значений параметров ДСЕ
            //new Guid("f6c95c62-663f-40ac-aca0-cd4665b9effc"), // Справочник Шаблоны имен исполнений
            //new Guid("1db15b40-2f0b-4628-a8bb-afadb8f92e2d"), // Справочник Шаблоны кодов
            //new Guid("fa7cf0eb-c729-4d56-b783-dcbcbd9f7937"), // Справочник Шаблоны сообщений
            //new Guid("59ceb1ae-8be7-47f5-9cea-3f43518f121d"), // Справочник Шаблоны текстов переходов
            //new Guid("4c6d35c9-82a3-4198-9de6-a389d5486991"), // Справочник Шероховатости
            new Guid("fd4ccacc-8252-477f-91eb-c65f60dfb743"), // Справочник Штатное расписание
            //new Guid("b8245afd-eb33-4eb8-a39a-e2699617fbee"), // Справочник Экземпляры изделий
            //new Guid("853d0f07-9632-42dd-bc7a-d91eae4b8e83"), // Справочник Электронная структура изделий
            //new Guid("2ac850d9-5c70-45c2-9897-517ab571b213"), // Справочник Электронные компоненты
            //new Guid("9c5d65f6-d5d6-4495-b09d-2c8a27df5598"), // Справочник Элементы обозначения материалов
            new Guid("31bf2b36-aa53-4d21-8d71-b499e00aee8d"), // Справочник Этапы контрактов
            //new Guid("d3cc8ac0-1e16-4a81-ae3e-1455897640b3"), // Справочник Этапы проекта
        };

        InputDialog dialog = new InputDialog(this.Context, "Укажите параметры для подключения к базе данных");
        dialog.AddString(serverNameField, serverName);
        dialog.AddString(frendlyNameField, frendlyName);
        dialog.AddString(userNameField, userName);
        dialog.AddString(passwordField, password);
        dialog.AddString(pathToSaveDiffField, pathToFile);
        dialog.AddMultiselectFromList(excludedFromComparingFields, allFields);
        dialog.AddFlag(directionOfCompareField, false);
        dialog.AddFlag(excludeReferences, true);

        if (dialog.Show()) {
            serverName = dialog[serverNameField];
            userName = dialog[userNameField];
            frendlyName = dialog[frendlyNameField];
            pathToFile = dialog[pathToSaveDiffField];
            password = dialog[passwordField];

            // Приступаем к чтению данных
            StDataBase structure = new StDataBase(Context.Connection, "Текущая база", indent, dialog[excludeReferences] ? excludedReferences : new List<Guid>());
            StDataBase otherStructure = StDataBase.CreateFromRemote(userName, serverName, frendlyName, indent, dialog[excludeReferences] ? excludedReferences : new List<Guid>());

            if (dialog[excludedFromComparingFields] != null) {
                foreach (object key in dialog[excludedFromComparingFields]) {
                    structure.ExcludeKey((string)key);
                    otherStructure.ExcludeKey((string)key);
                }
            }

            if (dialog[directionOfCompareField] == false) {
                structure.Compare(otherStructure);
                File.WriteAllText(pathToFile, structure.PrintDifferences());
            }
            else {
                otherStructure.Compare(structure);
                File.WriteAllText(pathToFile, otherStructure.PrintDifferences());
            }

            // Прозводим открытие файла в блокноте
            System.Diagnostics.Process notepad = new System.Diagnostics.Process();
            notepad.StartInfo.FileName = "notepad.exe";
            notepad.StartInfo.Arguments = pathToFile;
            notepad.Start();

        }
        else
            return;
    }

    public void ПолучитьСписокВсехСправочников() {
        string result = "List<Guid> excludedReferences = new List<Guid>() {\n";
        foreach (ReferenceInfo refInfo in Context.Connection.ReferenceCatalog.GetReferences().OrderBy(info => info.Name)) {
            result += $"    //new Guid(\"{refInfo.Guid}\"), // Справочник {refInfo.Name}\n";
        }
        result += "};";

        Message("Все справочники", result);
    }


    public abstract class StBaseNode {
        // Основные параметры
        public Guid Guid { get; set; }

        // Структурные параметры
        public StBaseNode Parent { get; set; }
        public StDataBase Root { get; set; }
        public Dictionary<Guid, StBaseNode> ChildNodes { get; set; } = new Dictionary<Guid, StBaseNode>();
        private Dictionary<string, string> Properties = new Dictionary<string, string>();
        public Dictionary<string, string>.KeyCollection Keys => this.Properties.Keys;

        // Статус и отображение различия
        public CompareResult Status { get; set; }
        public TypeOfNode Type { get; set; } = TypeOfNode.UDF;
        public List<string> Differences { get; set; } = new List<string>();
        public List<string> MissedNodes { get; set; } = new List<string>();

        public bool IsVisibleForDiffReport { get; set; } = false;

        // Декоративные параметры
        public int Indent { get; set; }
        public int Level { get; set; }
        public string IndentString { get; set; }
        public string StringRepresentation => $"({this.Type.ToString()}) {this["Наименование"]} (ID: {this["id"]}; Guid: {this.Guid.ToString()})";

        public StBaseNode(StDataBase root, StBaseNode parent, int indent, TypeOfNode type) {
            this.Root = root != null ? root : (StDataBase)this;
            this.Parent = parent;
            //this.Indent = this.Root != null ? this.Root.Indent : indent;
            this.Indent = 5;
            this.Type = type;
            this.Level = this.Parent == null ? 0 : this.Parent.Level + 1;
            this.IndentString = GetIndentString();
            this.Status = CompareResult.NTP;
        }

        public StBaseNode(StDataBase root, StBaseNode parent, TypeOfNode type) : this(root, parent, 5, type) {
        }

        public string this[string key] {
            get {
                try {
                    return this.Properties[key];
                }
                catch (Exception e) {
                    throw new Exception($"При попытке получения значения по ключу '{key}' возникла ошибка:\n{e.Message}");
                }
            }
            set {
                this.Properties[key] = value;
            }
        }

        public void CompareWith(StBaseNode otherNode) {
            if (this.Guid != otherNode.Guid) {
                throw new Exception($"Невозможно сравнить объекты с разными Guid: {this.Guid.ToString()} => {otherNode.Guid.ToString()}");
            }

            if (this.Type != otherNode.Type) {
                throw new Exception($"Невозможно произвести сравнение объектов с разными типами: {this.Type} => {otherNode.Type}. ({this["Наименование"]})");
            }

            foreach (string key in this.Keys) {
                if ((this[key] != otherNode[key]) && (!this.Root.ExcludedKeys.ContainsKey(this.Type) || !this.Root.ExcludedKeys[this.Type].Contains(key)))
                    this.Differences.Add($"[{key}]: {this[key]} => {otherNode[key]}");
            }

            if (this.Differences.Count == 0)
                this.Status = CompareResult.EQL;
            else
                this.Status = CompareResult.DIF;
        }

        public void Compare(StBaseNode otherNode) {
            this.CompareWith(otherNode);

            // Сравниваем все справочники
            List<Guid> allReferenceGuidsInCurrentStructure = this.ChildNodes.Select(kvp => kvp.Key).ToList<Guid>();
            List<Guid> allReferenceGuidsInOtherStructure = otherNode.ChildNodes.Select(kvp => kvp.Key).ToList<Guid>();

            var missingInOther = allReferenceGuidsInCurrentStructure.Except(allReferenceGuidsInOtherStructure);
            var missingInCurrent = allReferenceGuidsInOtherStructure.Except(allReferenceGuidsInCurrentStructure);
            var existInBoth = allReferenceGuidsInCurrentStructure.Intersect(allReferenceGuidsInOtherStructure);

            // Обрабатываем "Отсутствующие" позиции
            foreach (Guid guid in missingInCurrent) {
                this.MissedNodes.Add(otherNode.ChildNodes[guid].StringRepresentation);
            }

            // Обрабатываем "Новые" позиции
            foreach (Guid guid in missingInOther) {
                this.ChildNodes[guid].Status = CompareResult.NEW;
                this.ChildNodes[guid].SetAllChildNodesStatus(CompareResult.NEW);
                //this.ChildNodes[guid].SetAllChildNodesVisible(); // - Данную строку следует распомментировать, если нужно сделать видимыми все дочерние элементы нового объекта (которые тоже соответственно будут новыми)
                this.ChildNodes[guid].SetAllParentsToVisible();
            }

            // Обрабатываем одинаковые позиции
            // Запускаем сравнение справочников, которые есть и в первой и во второй структуре
            foreach (Guid guid in existInBoth) {
                this.ChildNodes[guid].Compare(otherNode.ChildNodes[guid]);
            }

            if (this.HaveDifference())
                this.SetAllParentsToVisible();
        }

        private void SetAllChildNodesStatus(CompareResult status) {
            this.Status = status;
            foreach (StBaseNode node in this.ChildNodes.Select(kvp => kvp.Value)) {
                node.SetAllChildNodesStatus(status);
            }
        }

        private void SetAllChildNodesVisible() {
            this.IsVisibleForDiffReport = true;
            foreach (StBaseNode node in this.ChildNodes.Select(kvp => kvp.Value)) {
                node.SetAllChildNodesVisible();
            }
        }

        public void SetAllParentsToVisible() {
            this.IsVisibleForDiffReport = true;
            StBaseNode currentType = this.Parent;

            while (true) {
                if ((currentType == null) || (currentType.IsVisibleForDiffReport == true))
                    break;
                currentType.IsVisibleForDiffReport = true;
                currentType = currentType.Parent;
            }
        }

        private bool HaveDifference() {
            if (this.Differences.Count != 0)
                return true;
            if (this.MissedNodes.Count != 0)
                return true;
            if (this.Status == CompareResult.NEW)
                return true;
            return false;
        }

        public string GetIndentString() {
            return this.Level != 0 ? new string(' ', (int)(this.Level * this.Indent)) : string.Empty;
        }

        public string PrintDifferences() {
            string tree = $"{this.IndentString}({this.Type.ToString()})"; // Строковое представление входимости объектов и их тип
            string name = this["Наименование"]; // Имя объекта
            string stat = this.Status.ToString(); // Статус объекта
            string diff = this.Differences.Count.ToString(); // Количество отличий
            string nw = this.ChildNodes.Where(kvp => kvp.Value.Status == CompareResult.NEW).Count().ToString(); // Количество новых
            string miss = this.MissedNodes.Count.ToString(); // Количество отсутствующих
            string childElements = string.Join(string.Empty, this.ChildNodes.OrderBy(kvp => kvp.Value.Status).Where(kvp => kvp.Value.IsVisibleForDiffReport).Select(kvp => kvp.Value.PrintDifferences())); // Информация по входяхим элементам

            // Получаем детальную информацию о позиции
            StringBuilder details = new StringBuilder();
            if ((this.Differences.Count != 0) || this.MissedNodes.Count != 0)
                details.AppendLine();
            foreach (string difference in this.Differences)
                details.AppendLine($"{"--DETAILS--  ", 26}- (diff){difference}");
            foreach (string missed in this.MissedNodes)
                details.AppendLine($"{"--DETAILS--  ", 26}- (miss){missed}");


            return $"{tree, -25} {name,-80} (stat: {stat, 3}, diff: {diff, 2}, new: {nw, 2}, miss: {miss, 2}){details}\n{childElements}";
        }

        public override string ToString() {
            return $"{this.IndentString}({this.Type.ToString()}) {this["Наименование"]} ({this.Status})\n{string.Join(string.Empty, this.ChildNodes.OrderBy(kvp => kvp.Value.Status).Select(kvp => kvp.Value.ToString()))}";
        }
    
    }

    public class StDataBase : StBaseNode {
        // Поля, свойственные только корню
        public string FrendlyName { get; set; }
        public Dictionary<Guid, List<StBaseNode>> AllNodes { get; private set; } = new Dictionary<Guid, List<StBaseNode>>();
        public List<Guid> AllGuids { get; private set; } = new List<Guid>();

        // Словарь с исключениями
        public Dictionary<TypeOfNode, List<string>> ExcludedKeys { get; private set; } = new Dictionary<TypeOfNode, List<string>>();
        public List<Guid> ExcludedReferences { get; private set; }

        public StDataBase(ServerConnection connection, string frendlyName, int indent, List<Guid> excludedReferences = null) : base(null, null, indent, TypeOfNode.SRV) {
            this.FrendlyName = frendlyName;
            this.ExcludedReferences = excludedReferences != null ? excludedReferences : new List<Guid>();

            this.Guid = new Guid("00000000-0000-0000-0000-000000000000");
            this["id"] = "0";
            this["Наименование"] = connection.ServerName;
            this["Псевдоним"] = this.FrendlyName;

            foreach (ReferenceInfo refInfo in connection.ReferenceCatalog.GetReferences()) {
                // Пропускаем те справочники, которые были переданы в качестве исключения из обработки
                if (this.ExcludedReferences.Contains(refInfo.Guid))
                    continue;
                this.AllGuids.Add(refInfo.Guid);
                StReference reference = new StReference(refInfo, this, this);
                this.ChildNodes.Add(reference.Guid, reference);
                if (!this.AllNodes.ContainsKey(reference.Guid))
                    this.AllNodes.Add(reference.Guid, new List<StBaseNode>() { reference });
                else
                    this.AllNodes[reference.Guid].Add(reference);
            }
        }

        public static StDataBase CreateFromRemote(string userName, string serverAddress, string dataBaseFrendlyName, int indent, List<Guid> excludedReferences = null) {
            using (ServerConnection connection = ServerConnection.Open(userName, "123", serverAddress)) {
                if (!connection.IsConnected)
                    throw new Exception("Не удалось подключиться к внешнему серверу");
                else
                    return new StDataBase(connection, dataBaseFrendlyName, indent, excludedReferences);
            }
        }

        public void ExcludeKey(string key) {
            foreach (string name in Enum.GetNames(typeof(TypeOfNode)).Skip(1).ToArray<string>()) {
                this.ExcludeKey(key, (TypeOfNode)Enum.Parse(typeof(TypeOfNode), name));
            }
        }

        public void ExcludeKey(string key, TypeOfNode type) {
            if (type == TypeOfNode.UDF)
                throw new Exception("При добавлении исключения не поддерживается неопределенный тип объекта");

            if (this.ExcludedKeys.ContainsKey(type))
                this.ExcludedKeys[type].Add(key);
            else
                this.ExcludedKeys[type] = new List<string>() { key };
        }

        public StBaseNode GetInfoAboutReference(Guid guid) {
            if (this.AllNodes.ContainsKey(guid))
                return this.ChildNodes[guid];
            return null;
        }

        public void AddGuid(Guid guid) {
            this.AllGuids.Add(guid);
        }

        public void AddNode(StBaseNode node) {
            if (!this.AllNodes.ContainsKey(node.Guid))
                this.AllNodes.Add(node.Guid, new List<StBaseNode>() { node });
            else
                this.AllNodes[node.Guid].Add(node);
        }

    }

    public class StReference : StBaseNode {

        public StReference(ReferenceInfo referenceInfo, StDataBase root, StBaseNode parent) : base(root, parent, TypeOfNode.REF) {
            this.Guid = referenceInfo.Guid;
            this["id"] = referenceInfo.Id.ToString();
            this["Наименование"] = referenceInfo.Name;

            foreach (ClassObject classObject in referenceInfo.Classes.AllClasses.Where(cl => cl is ClassObject)) {
                this.Root.AddGuid(classObject.Guid);
                StClass structureClass = new StClass(classObject, this.Root, this);
                this.ChildNodes.Add(structureClass.Guid, structureClass);
                this.Root.AddNode(structureClass);
            }
        }

    }

    public class StClass : StBaseNode {

        public StClass(ClassObject classObject, StDataBase root, StBaseNode parent) : base(root, parent, TypeOfNode.TYP) {
            this.Guid = classObject.Guid;
            this["id"] = classObject.Id.ToString();
            this["Наименование"] = classObject.Name;
            this["КолСобытий"] = classObject.Events.GetUserEvents().Count.ToString();
            this["НаимСобытий"] = string.Join("; ", classObject.Events.GetUserEvents().Select(ev => ev.Name).OrderBy(name => name));
            this["КолОбработчиков"] = classObject.Events.Handlers.Count.ToString();
            this["НаимОбработчиков"] = string.Join("; ", classObject.Events.Handlers.Select(handler => handler.HandlerName).OrderBy(name => name));
            this["КоличествоСтраницДиалогов"] = classObject.Dialog != null ?
                classObject.Dialog.Groups.Select(group => group.Pages != null ? group.Pages.Count : 0).Sum().ToString() :
                "Отсутствует";


            // Производим поиск групп параметров справочника
            foreach (ParameterGroup group in classObject.GetAllGroups()) {
                this.Root.AddGuid(group.Guid);

                if (group.IsLinkGroup) {
                    StLink link = new StLink(group, this.Root, this);
                    this.ChildNodes.Add(link.Guid, link);
                    this.Root.AddNode(link);
                }
                else {
                    StParameterGroup parameterGroup = new StParameterGroup(group, this.Root, this);
                    this.ChildNodes.Add(parameterGroup.Guid, parameterGroup);
                    this.Root.AddNode(parameterGroup);
                }
            }

        }

    }

    public class StParameterGroup : StBaseNode {

        public StParameterGroup(ParameterGroup group, StDataBase root, StBaseNode parent) : base(root, parent, TypeOfNode.GRP) {
            this.Guid = group.Guid;
            this["id"] = group.Id.ToString();
            this["Наименование"] = group.Name;

            // Получаем параметры из группы параметров
            foreach (ParameterInfo parameterInfo in group.Parameters) {
                this.Root.AddGuid(parameterInfo.Guid);
                StParameter parameter = new StParameter(parameterInfo, this.Root, this);
                this.ChildNodes.Add(parameter.Guid, parameter);
                this.Root.AddNode(parameter);
            }
        }

    }

    public class StParameter : StBaseNode {

        public StParameter(ParameterInfo parameterInfo, StDataBase root, StBaseNode parent) : base(root, parent, TypeOfNode.PRM) {
            this.Guid = parameterInfo.Guid;
            this["id"] = parameterInfo.Id.ToString();
            this["Наименование"] = parameterInfo.Name;
            this["ДлинаПараметра"] = parameterInfo.Length.ToString();
            this["ВозможноНулевоеЗначение"] = parameterInfo.Nullable.ToString();
            this["ТипПараметра"] = parameterInfo.TypeName != null ? parameterInfo.TypeName : "null";
            this["ЕдИзПараметра"] = parameterInfo.Unit != null ? parameterInfo.Unit.Name : "null";
            this["ЗначенияСпискаПараметра"] = parameterInfo.ValueList != null ? string.Join("; ", parameterInfo.ValueList.Select(item => $"{item.Name} - {item.Value}")) : "null";
        }

    }

    public class StLink : StBaseNode {

        public StLink(ParameterGroup group, StDataBase root, StBaseNode parent) : base(root, parent, TypeOfNode.LNK) {
            this.Guid = group.Guid;
            this["id"] = group.Id.ToString();
            this["Наименование"] = group.Name;
            this["ТипСсылки"] = group.LinkType.ToString();
        }

    }

    public enum CompareResult {
        NTP, //Не обработано
        NEW,  //Новый (отсутствует в сравниваемом справочнике)
        DIF, //Есть разница 
        EQL //Соответствует
    }

    public enum TypeOfNode {
        UDF, //Не определено
        SRV, //Сервер
        REF, //Справочник
        TYP, //Тип
        GRP, //Группа
        PRM, //Параметр
        LNK  //Связь
    }
}
