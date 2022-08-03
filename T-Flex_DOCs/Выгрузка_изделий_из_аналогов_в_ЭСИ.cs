using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using TFlex.DOCs.Model;
using TFlex.DOCs.Model.Macros;
using TFlex.DOCs.Model.References;
using TFlex.DOCs.Model.Parameters;
using TFlex.DOCs.Model.Macros.ObjectModel;
using TFlex.DOCs.Model.Classes;
using TFlex.DOCs.Model.Structure;
using TFlex.DOCs.Model.Search;
using FoxProShifrsNormalizer;
using TFlex.DOCs.Model.Desktop;
using TFlex.DOCs.Model.References.Nomenclature;
using TFlex.DOCs.Model.References.Files;


public class Macro : MacroProvider
{
    private RefGuidData СписокНоменклатуры { get; set; }
    private RefGuidData Подключения { get; set; }
    private RefGuidData ЭСИ { get; set; }
    private RefGuidData Документы { get; set; }
    private RefGuidData ЭлектронныеКомпоненты { get; set; }
    private RefGuidData Материалы { get; set; }
    private RefGuidData ЖурналВыгрузокИзFoxPro { get; set; }
    private Log log_save_ref = new Log();
                                                                            
    // Временные настройки
    private bool applyCreation = false;
                                         

    private string ДиректорияДляЛогов { get; set; }

    // Для лога
    private string TimeStamp { get; set; }
    private string NameOfImport { get; set; }
    private string StringForLog => $"{NameOfImport} ({TimeStamp})";

    // Для ревизий
    private string RevisionNameTemplate { get; set; }

    public Macro(MacroContext context)
        : base(context) {
     

        #if DEBUG
        System.Diagnostics.Debugger.Launch();
        System.Diagnostics.Debugger.Break();
        #endif

        // Инициализируем переменную, в которой будет храниться шаблон для наименования ревизий
        RevisionNameTemplate = "ФОКС-";

        // Инициализируем объекты с уникальными идентификаторами
        // Справочник "Журнал выгрузок из FoxPro"
        ЖурналВыгрузокИзFoxPro = new RefGuidData(context, new Guid("85ed8179-1714-4b63-9030-b41c82451cc0"));
        ЖурналВыгрузокИзFoxPro.AddParam("number", new Guid("5e20f92d-c433-4f99-bde9-63a7a843058a")); // int
        ЖурналВыгрузокИзFoxPro.AddParam("shifrIzd", new Guid("562c1053-94a5-4268-ac78-b3fcf1f896b2")); // string
        ЖурналВыгрузокИзFoxPro.AddParam("time_work", new Guid("a0bd05d3-d698-482b-b19d-8abfbfe0bd09")); //int
        ЖурналВыгрузокИзFoxPro.AddParam("count_object", new Guid("e5223c20-f2e5-4726-9f31-b14605bf71dd")); //int
        ЖурналВыгрузокИзFoxPro.AddParam("count_error", new Guid("9d79c7c8-3870-4feb-b301-a53a9bbe9b13")); //int
        ЖурналВыгрузокИзFoxPro.AddParam("create_object", new Guid("2e355047-8b38-4995-bfe8-b7a15c54f0f2")); //int
        ЖурналВыгрузокИзFoxPro.AddLink("Файл выгрузки", new Guid("422ae8ea-318b-4973-b4fd-e9732eb331cf")); // string
        
        // Справочник "Список номенклатуры FoxPro"
        СписокНоменклатуры = new RefGuidData(Context, new Guid("c9d26b3c-b318-4160-90ae-b9d4dd7565b6"));
        // Параметры
        СписокНоменклатуры.AddParam("Обозначение", new Guid("1bbb1d78-6e30-40b4-acde-4fc844477200")); // string
        СписокНоменклатуры.AddParam("Тип номенклатуры", new Guid("3c7a075f-0b53-4d68-8242-9f76ca7b2e97")); // int
        СписокНоменклатуры.AddParam("Наименование", new Guid("c531e1a8-9c6e-4456-86aa-84e0826c7df7")); // string
        СписокНоменклатуры.AddParam("ГОСТ", new Guid("0f48ff0a-36c0-4ae5-ae4c-482f2728181f")); // string
        // Связи
        СписокНоменклатуры.AddLink("Связь на ЭСИ", new Guid("ec9b1e06-d8d5-4480-8a5c-4e6329993eac")); // string

        // Справочник "Подключения"
        Подключения = new RefGuidData(Context, new Guid("9dd79ab1-2f40-41f3-bebc-0a8b4f6c249e"));
        // Параметры
        Подключения.AddParam("Сборка", new Guid("4a3cb1ca-6a4c-4dce-8c25-c5c3bd13a807")); // string
        Подключения.AddParam("Комплектующая", new Guid("7d1ac031-8c7f-49b5-84b8-c5bafa3918c2")); // string
        Подключения.AddParam("Сводное обозначение", new Guid("05ffddba-74e9-4637-b249-90cec5953295")); // string
        Подключения.AddParam("Позиция", new Guid("b05be213-7646-4edb-9d56-391509b48c2a")); // int
        Подключения.AddParam("Количество", new Guid("fa56458a-e817-4e6d-85a0-e64dad032c5f")); // double
        Подключения.AddParam("ОКЕИ", new Guid("19d31f8c-06d2-402b-85ee-bda3f5111e8c")); // int
        Подключения.AddParam("Код Ediz", new Guid("d85db0fe-6c97-4664-9c16-a82695a40984")); // int
        Подключения.AddParam("Единица измерения", new Guid("94158439-cf0a-470b-872c-d783d8ebbd60")); // string
        Подключения.AddParam("Единица измерения сокращенно", new Guid("d485a313-6228-4bbf-b40e-b29e82adbb68")); // string
        Подключения.AddParam("Возвратные отходы", new Guid("d9e79828-12d8-4a8a-b77e-9626cedeb307")); // double
        Подключения.AddParam("Площадь покрытия", new Guid("8b12a1f1-0478-4e31-b05f-7205ae683f38")); // double
        Подключения.AddParam("Потери", new Guid("dd57da68-ebb4-4e43-ab83-05fa621895aa")); // double
        Подключения.AddParam("Толщина покрытия", new Guid("2475329d-3128-4d2d-82a0-af6c35961753")); // double
        Подключения.AddParam("Чистый вес", new Guid("e8200590-255d-4a51-9826-686d21c5f2b6")); // double

        // Справочник "Электронная структура изделия"
        ЭСИ = new RefGuidData(Context, new Guid("853d0f07-9632-42dd-bc7a-d91eae4b8e83"));
        // Параметры
        ЭСИ.AddParam("Обозначение", new Guid("ae35e329-15b4-4281-ad2b-0e8659ad2bfb")); // string
        ЭСИ.AddParam("Наименование", new Guid("45e0d244-55f3-4091-869c-fcf0bb643765")); // string
        ЭСИ.AddParam("Код ОКП", new Guid("b39cc740-93cc-476d-bfed-114fe9b0740c")); // string
        ЭСИ.AddParam("Наименование ревизии", new Guid("8d69bd40-0fe0-4bb1-9d1a-2e728f6cdc68")); // string
        ЭСИ.AddParam("Guid логического объекта", new Guid("49c7b3ec-fa35-4bb1-92a5-01d4d3a40d16")); // guid
        // Типы
        ЭСИ.AddType("Материальный объект", new Guid("0ba28451-fb4d-47d0-b8f6-af0967468959"));
        // Параметры подключений
        ЭСИ.AddHlink("Количество", new Guid("3f5fc6c8-d1bf-4c3d-b7ff-f3e636603818")); // double
        ЭСИ.AddHlink("Позиция", new Guid("ab34ef56-6c68-4e23-a532-dead399b2f2e")); // int
        ЭСИ.AddHlink("Возвратные отходы", new Guid("a750a217-32f2-4438-bb32-60a547da50df")); // double
        ЭСИ.AddHlink("Потери", new Guid("3245305b-d5b6-419c-ad4d-17b317357272")); // double
        ЭСИ.AddHlink("Чистый вес", new Guid("67732ea3-d9a3-448f-bdf9-6b2dee0c2403")); // double
        ЭСИ.AddHlink("Площадь покрытия", new Guid("45df5f0b-dc53-494c-ae2e-54dfb8d64bd9")); // double
        ЭСИ.AddHlink("Толщина покрытия", new Guid("c024de94-d58c-4a15-b9a9-ccb2e57b246e")); // double

        // Справочник "Документы"
        Документы = new RefGuidData(Context, new Guid("ac46ca13-6649-4bbb-87d5-e7d570783f26"));
        // Параметры
        Документы.AddParam("Обозначение", new Guid("b8992281-a2c3-42dc-81ac-884f252bd062")); // string
        Документы.AddParam("Наименование", new Guid("7e115f38-f446-40ce-8301-9b211e6ce5fd")); // string
        Документы.AddParam("Код ОКП", new Guid("45ead73a-1773-4156-bafd-48795f844cfb")); // string
        // Типы
        Документы.AddType("Объект состава изделия", new Guid("f89e9648-c8a0-43f8-82bb-015cfe1486a4"));

        // Справочник "Электронные компоненты"
        ЭлектронныеКомпоненты = new RefGuidData(Context, new Guid("2ac850d9-5c70-45c2-9897-517ab571b213"));
        // Параметры
        ЭлектронныеКомпоненты.AddParam("Обозначение", new Guid("65e0e04a-1a6f-4d21-9eb4-dfe5a135ec3b")); // string
        ЭлектронныеКомпоненты.AddParam("Наименование", new Guid("01184891-8364-4a5c-bf05-2163e1f3d460")); // string
        ЭлектронныеКомпоненты.AddParam("Код ОКП", new Guid("72f18ec6-d471-45c7-b1df-26f8ccd89af3")); // string

        // Справочник "Материалы"
        Материалы = new RefGuidData(Context, new Guid("c5e7ae00-90f2-49e9-a16c-f51ed087752a"));
        // Параметры
        Материалы.AddParam("Обозначение", new Guid("d0441280-01ea-43b5-8726-d2d02e4d996f")); // string
        Материалы.AddParam("Сводное наименование", new Guid("23cfeee6-57f3-4a1e-9cf0-9040fed0e90c")); // string
        Материалы.AddParam("Наименование", new Guid("23cfeee6-57f3-4a1e-9cf0-9040fed0e90c")); // string
        Материалы.AddParam("Код ОКП", new Guid("d0441280-01ea-43b5-8726-d2d02e4d996f")); // string

        // Создаем директорию для ведения логов
        ДиректорияДляЛогов = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "Логи выгрузки из аналогов в ЭСИ");

        // Сохраняем текущее время для использования в логах
        TimeStamp = DateTime.Now.ToString("yyyy.MM.dd HH-mm");

        // Даем временное название текущей выгрузке для использования в логах
        NameOfImport = "Временное название выгрузки";

        if (!Directory.Exists(ДиректорияДляЛогов))
            Directory.CreateDirectory(ДиректорияДляЛогов);

    }

    public override void Run() {
        
        // Производим загрузку всей необходимой информации
        Dictionary<string, ReferenceObject> номенклатура = GetNomenclature();
        Dictionary<string, List<ReferenceObject>> подключения = GetLinks();

        // Запрашиваем у пользователя перечень изделий, по которым нужно произвести выгрузку
        List<string> изделияДляВыгрузки = GetShifrsFromUserToImport(номенклатура);
        //Test(изделияДляВыгрузки[0]);
        log_save_ref.timeStart = DateTime.Now;
        ///*
        // Определяем позиции справочника "Список номенклатуры FoxPro", которые необходимо обрабатывать во время выгрузки
        HashSet<ReferenceObject> номенклатураДляСоздания = GetNomenclatureToProcess(номенклатура, подключения, изделияДляВыгрузки);

        // Производим поиск и (при необходимости) создание объектов в ЭСИ и смежных справочниках
        List<ReferenceObject> созданныеДСЕ = FindOrCreateNomenclatureObjects(номенклатураДляСоздания);

        // Производим соединение созданных ДСЕ в иерархию при помощи подключений
        ConnectCreatedObjects(созданныеДСЕ, подключения);
        //*/       
        
        SaveLogtoRef(изделияДляВыгрузки);

        Message("Информация", "Работа макроса завершена");
    }

    /// <summary>
    /// Создает каждого объекта в списке создает отдельный лог.
    /// </summary>
    public void RunOne()
    {
        
        // Производим загрузку всей необходимой информации
        Dictionary<string, ReferenceObject> номенклатура = GetNomenclature();
        Dictionary<string, List<ReferenceObject>> подключения = GetLinks();

        // Запрашиваем у пользователя перечень изделий, по которым нужно произвести выгрузку
        List<string> изделияДляВыгрузки = GetShifrsFromUserToImport(номенклатура);
        //Test(изделияДляВыгрузки[0]);
        //log_save_ref.messageOff = true;
       
        ///*
        foreach (var изделиеПоОдному in изделияДляВыгрузки)
        {
            log_save_ref.timeStart = DateTime.Now;

            List<string> изделиеВыгружаемое = new List<string>() { изделиеПоОдному };
            SetNameOfExport(изделиеВыгружаемое);
            // Определяем позиции справочника "Список номенклатуры FoxPro", которые необходимо обрабатывать во время выгрузки
            HashSet<ReferenceObject> номенклатураДляСоздания = GetNomenclatureToProcess(номенклатура, подключения, изделиеВыгружаемое);

            // Производим поиск и (при необходимости) создание объектов в ЭСИ и смежных справочниках
            List<ReferenceObject> созданныеДСЕ = FindOrCreateNomenclatureObjects(номенклатураДляСоздания);

            // Производим соединение созданных ДСЕ в иерархию при помощи подключений
            ConnectCreatedObjects(созданныеДСЕ, подключения);
            //*/       

            SaveLogtoRef(изделиеВыгружаемое);
        }
        Message("Информация", "Работа макроса завершена");
    }

    /// <summary>
    /// Создает новую запись в справочнике "Журнал выгрузок из FoxPro"
    /// </summary>
    private void SaveLogtoRef(List<string> изделияДляВыгрузки)
    {
        log_save_ref.getparam(ЖурналВыгрузокИзFoxPro);
        log_save_ref.Shifr_izd = String.Join("\n",изделияДляВыгрузки);
        int second_work = ((log_save_ref.timeStop - log_save_ref.timeStart).Hours * 60 * 60) + ((log_save_ref.timeStop - log_save_ref.timeStart).Minutes * 60) + (log_save_ref.timeStop - log_save_ref.timeStart).Seconds;
        var createdClassObject = ЖурналВыгрузокИзFoxPro.Ref.Classes.Find("Журнал выгрузок из FoxPro");
        ReferenceObject refereceObject = ЖурналВыгрузокИзFoxPro.Ref.CreateReferenceObject(createdClassObject);
        refereceObject[ЖурналВыгрузокИзFoxPro.Params["number"]].Value = log_save_ref.Number + 1;
        refereceObject[ЖурналВыгрузокИзFoxPro.Params["shifrIzd"]].Value = log_save_ref.Shifr_izd;
        refereceObject[ЖурналВыгрузокИзFoxPro.Params["time_work"]].Value = second_work;
        refereceObject[ЖурналВыгрузокИзFoxPro.Params["count_object"]].Value = log_save_ref.count_object;
        refereceObject[ЖурналВыгрузокИзFoxPro.Params["count_error"]].Value = log_save_ref.count_error;
        refereceObject[ЖурналВыгрузокИзFoxPro.Params["create_object"]].Value = log_save_ref.create_object;
        LoadFilesLog(refereceObject);
        refereceObject.EndChanges();
    }

    /// <summary>
    /// Прикрепляет файлы логов к записе в справочнике "Журнал выгрузок из FoxPro"
    /// </summary>
    private void LoadFilesLog(ReferenceObject refereceObject)
    {
        try
        {
            var fileReference = new FileReference(Context.Connection);
            var parentFolder = fileReference.FindByRelativePath("Логи выгрузки из аналогов в ЭСИ") as TFlex.DOCs.Model.References.Files.FolderObject;
            var file1 = fileReference.AddFile(log_save_ref.file_name_error_log, parentFolder);
            var file2 = fileReference.AddFile(log_save_ref.file_name_tree, parentFolder);
            Desktop.CheckIn(file1, "Объект создан", false);
            Desktop.CheckIn(file2, "Объект создан", false);

            refereceObject.AddLinkedObject(ЖурналВыгрузокИзFoxPro.Links["Файл выгрузки"], file1);
            refereceObject.AddLinkedObject(ЖурналВыгрузокИзFoxPro.Links["Файл выгрузки"], file2);
        }
        catch (Exception e)
        {
            // Выводим сообщение об ошибке
            Сообщение("Ошибка", e.Message);
        }

    }

                  
    public void Test()
    {
        /*Dictionary<string, ReferenceObject> номенклатура = GetNomenclature();
        Dictionary<string, List<ReferenceObject>> подключения = GetLinks();
        List<string> изделияДляВыгрузки = GetShifrsFromUserToImport(номенклатура);
        GetFilterRefObj(изделияДляВыгрузки, ЭСИ.RefGuid, ЭСИ.Params["Обозначение"]);
        */
        ReferenceObject currentObject = ЭСИ.Ref.Find(new Guid("a57c6880-bba8-4f52-9788-e5e5cf86b51a"));
        MoveShifrToOKP(currentObject, "7594105844", TypeOfObject.СтандартноеИзделие, new List<string> { "" });
        //MoveShifrToOKP(ReferenceObject resultObject, string designation, List<string> messages)
        //MoveShifrToOKP(currentObject, "7594105844", new List<string> { "" });




        //var result = GetFilterRefObj(изделияДляВыгрузки, Материалы.RefGuid, Материалы.Params["Обозначение"]);
    }

    public void TestGukov() {
        // Создаем новый объект в справочнике "Документы";
        ReferenceObject newDocument = null;
        try {
            newDocument = Документы.Ref.CreateReferenceObject(GetClassFrom(TypeOfObject.Другое, Документы.Ref));
            newDocument[Документы.Params["Обозначение"]].Value = "28072022";
            newDocument[Документы.Params["Наименование"]].Value = "28072022";
            newDocument.SystemFields.RevisionName = "ФОКС-1";
            newDocument.EndChanges();
        }
        catch (Exception e) {
            Message("Ошибка", $"При создании нового документа возникла ошибка:{Environment.NewLine}{e.Message}");
        }

        // Создаем для него номенклатурный объект
        ReferenceObject newNomenclature = (ЭСИ.Ref as NomenclatureReference).CreateNomenclatureObject(newDocument);
    }

    /// <summary>
    /// Метод для проведения тестирования отдельно первой стадии
    /// </summary>
    public void TestFirstStage() {
        // Список шифров объектов для тестирования
        List<string> shifrs  = new List<string>() {
            "6С8220232",
        };

        List<string> messages = new List<string>();

        // Получаем соответствующие объекты в справочнике
        List<ReferenceObject> findedObjects = new List<ReferenceObject>();
        
        foreach (string shifr in shifrs) {
            findedObjects.Add(СписокНоменклатуры.Ref.Find(СписокНоменклатуры.Params["Обозначение"], shifr).FirstOrDefault());
        }

        foreach (ReferenceObject nomenclature in findedObjects) {
            if (nomenclature == null)
                messages.Add("Нулевой объект");
            else {
                try {
                    ReferenceObject result = ProcessFirstStageFindOrCreate(nomenclature, (string)nomenclature[СписокНоменклатуры.Params["Обозначение"]].Value, GetTypeOfObjectFrom(nomenclature), messages);
                    CheckAndLink(nomenclature, result, messages);
                }
                catch (Exception e) {
                    messages.Add($"Error: {e.Message}");
                }
            }
        }

        Message("Результат теста", string.Join(Environment.NewLine, messages));
    }

    /// <summary>
    /// Метод для тестирования второй стадии
    /// </summary>
    public void TestSecondStage() {
        // Список шифров объектов для тестирования
        List<string> shifrs = new List<string>() {
            "7592214458",
            "2253190300",
            "6341268135",
            "6331229015",
            "7594602237",
            "7592229430",
            "8А6.672.028",
            "УЯИС.711351.083",
            "УЯИС.303811.125",
            "7594602237",
        };

        List<string> messages = new List<string>();

        // Получаем соответствующие объекты
        List<ReferenceObject> findedObjects = shifrs
            .Select(shifr => СписокНоменклатуры.Ref.Find(СписокНоменклатуры.Params["Обозначение"], shifr).FirstOrDefault())
            .ToList<ReferenceObject>();

        foreach (ReferenceObject nomenclature in findedObjects) {
            if (nomenclature == null)
                messages.Add("Нулевой объект");
            else {
                string name = (string)nomenclature[СписокНоменклатуры.Params["Обозначение"]].Value;
                messages.Add($"{Environment.NewLine}ОБРАБОТКА {name}:");
                try {
                    ReferenceObject result = ProcessSecondStageFindOrCreate(nomenclature, (string)nomenclature[СписокНоменклатуры.Params["Обозначение"]].Value, GetTypeOfObjectFrom(nomenclature), messages);
                }
                catch (Exception e) {
                    messages.Add($"ERROR: {e.Message}");
                }
            }
        }

        Message("Результат теста", string.Join(Environment.NewLine, messages));
    }

    /// <summary>
    /// Метод для тестирования третьей стадии
    /// </summary>
    public void TestThirdStage() {
        // Создаем тестовые объекты
        List<string> testShifrs = new List<string>();
        for (int i = 1; i < 10; i++)
            testShifrs.Add($"testObject-{i.ToString()}");

        // Создаем записи в справочнике
        List<ReferenceObject> nomRecords = new List<ReferenceObject>();
        for (int i = 0; i < 9; i++) {
            ReferenceObject newRecord = СписокНоменклатуры.Ref.CreateReferenceObject();
            newRecord[СписокНоменклатуры.Params["Обозначение"]].Value = testShifrs[i];
            // Задаем типы этих объектов
            if (i < 3)
                newRecord[СписокНоменклатуры.Params["Тип номенклатуры"]].Value = GetIntFrom(TypeOfObject.Другое);
            else if ((3 <= i) && (i < 6))
                newRecord[СписокНоменклатуры.Params["Тип номенклатуры"]].Value = GetIntFrom(TypeOfObject.Материал);
            else
                newRecord[СписокНоменклатуры.Params["Тип номенклатуры"]].Value = GetIntFrom(TypeOfObject.ЭлектронныйКомпонент);
            newRecord.EndChanges();
            nomRecords.Add(newRecord);
        }

        // Создание тестовых объектов в смежных справочниках
        List<ReferenceObject> newTestObjects = new List<ReferenceObject>();
        for (int i = 0; i < 9; i++) {
            if (i < 3)
                newTestObjects.Add(TestCreateOneObject(nomRecords[i]));
            else if ((3 <= i) && (i < 6))
                newTestObjects.AddRange(TestCreateMultipleRevision(nomRecords[i]));
            else
                newTestObjects.AddRange(TestCreateInMultipleReferences(nomRecords[i]));
        }

        // Контейнер для логов
        List<string> messages = new List<string>();

        foreach (ReferenceObject nomenclature in nomRecords) {
            if (nomenclature == null)
                messages.Add("Нулевой объект");
            else {
                string name = (string)nomenclature[СписокНоменклатуры.Params["Обозначение"]].Value;
                messages.Add($"{Environment.NewLine}ОБРАБОТКА {name}:");
                try {
                    ReferenceObject result = ProcessThirdStageFindOrCreate(nomenclature, (string)nomenclature[СписокНоменклатуры.Params["Обозначение"]].Value, GetTypeOfObjectFrom(nomenclature), messages);
                }
                catch (Exception e) {
                    messages.Add($"ERROR: {e.Message}");
                }
            }
        }

        Message("Результат теста", string.Join(Environment.NewLine, messages));

        // Удаляем тестовые записи из справочника "Список номенклатуры FoxPro"
        for (int i = 0; i < nomRecords.Count; i++)
            nomRecords[i].Delete();

        // Отменяем создание объектов в смежных справочниках
        for (int i = 0; i < newTestObjects.Count; i++) {
            newTestObjects[i].BeginChanges();
            newTestObjects[i].CancelChanges();
            newTestObjects[i].EndChanges();
        }
    }

    /// <summary>
    /// Метод для тестирования последней стадии
    /// </summary>
    public void TestFinalStage() {
        // Создаем записи в справочнике
        List<ReferenceObject> nomRecords = CreateTestNomenclatureRecords(10);

        // Контейнер для логов
        List<string> messages = new List<string>();
        List<ReferenceObject> createdStageObjects = new List<ReferenceObject>();

        foreach (ReferenceObject nomenclature in nomRecords) {
            if (nomenclature == null)
                messages.Add("Нулевой объект");
            else {
                string name = (string)nomenclature[СписокНоменклатуры.Params["Обозначение"]].Value;
                messages.Add($"{Environment.NewLine}ОБРАБОТКА {name}:");
                try {
                    createdStageObjects.Add(ProcessFinalStageFindOrCreate(nomenclature, (string)nomenclature[СписокНоменклатуры.Params["Обозначение"]].Value, GetTypeOfObjectFrom(nomenclature), messages));
                }
                catch (Exception e) {
                    messages.Add($"ERROR: {e.Message}");
                }
            }
        }

        Message("Результат теста", string.Join(Environment.NewLine, messages));

        CleanTestObjects(nomRecords);
        CleanTestObjects(createdStageObjects);
    }

    private ReferenceObject TestCreateOneObject(ReferenceObject nom, TypeOfReference refType = TypeOfReference.Другое) {
        // Определяем тип объекта
        TypeOfObject type = GetTypeOfObjectFrom(nom);

        // Определяем тип справочника, в котором нужно создать объект
        if (refType == TypeOfReference.Другое)
            refType = GetTypeOfReferenceFrom(type);

        RefGuidData refData = GetRefGuidDataFrom(refType);

        ReferenceObject newObject = null;
        switch (refType) {
            case TypeOfReference.Документы:
                newObject = refData.Ref.CreateReferenceObject(GetClassFrom(TypeOfObject.Другое, Документы.Ref));
                break;
            case TypeOfReference.Материалы:
                newObject = refData.Ref.CreateReferenceObject(GetClassFrom(TypeOfObject.Материал, Материалы.Ref));
                break;
            case TypeOfReference.ЭлектронныеКомпоненты:
                newObject = refData.Ref.CreateReferenceObject(GetClassFrom(TypeOfObject.ЭлектронныйКомпонент, ЭлектронныеКомпоненты.Ref));
                break;
            default:
                newObject = refData.Ref.CreateReferenceObject(GetClassFrom(type, refData.Ref));
                break;
        }
        newObject[refData.Params["Обозначение"]].Value = nom[СписокНоменклатуры.Params["Обозначение"]].Value;
        newObject[refData.Params["Наименование"]].Value = nom[СписокНоменклатуры.Params["Обозначение"]].Value;
        newObject[refData.Params["Код ОКП"]].Value = nom[СписокНоменклатуры.Params["Обозначение"]].Value;
        newObject.EndChanges();

        return newObject;
    }

    private List<ReferenceObject> TestCreateMultipleRevision(ReferenceObject nom) {
        List<ReferenceObject> result = new List<ReferenceObject>();
        result.Add(TestCreateOneObject(nom));

        // Создаем несколько дополнительных ревизий объекта
        for (int i = 0; i < 3; i++)
            result.Add(result.First().CreateRevision(null, null, false));

        return result;
    }

    private List<ReferenceObject> TestCreateInMultipleReferences(ReferenceObject nom) {
        List<ReferenceObject> result = new List<ReferenceObject>();
        List<TypeOfReference> references = new List<TypeOfReference>() {
            TypeOfReference.ЭлектронныеКомпоненты,
            TypeOfReference.Материалы,
            TypeOfReference.Документы,
        };

        foreach (TypeOfReference refType in references)
            result.Add(TestCreateOneObject(nom, refType));
        return result;
    }

    /// <summary>
    /// Метод для удаление созданных во время тестирования объектов
    /// </summary>
    /// <param name="testObjects">Список тектовых объектов, которые нужно подчистить после произведения тестирования</param>
    private void CleanTestObjects(List<ReferenceObject> testObjects) {
        if (testObjects == null)
            throw new Exception($"{nameof(CleanTestObjects)}: в метод передан null");
        List<string> errors = new List<string>();

        foreach (ReferenceObject testObject in testObjects) {
            if (testObject == null)
                continue; // Обработка случая, когда стадия ничего не вернула
            // Получаем тип справочника
            TypeOfReference typeOfRef = GetTypeOfReferenceFrom(testObject);
            switch (GetTypeOfReferenceFrom(testObject)) {
                case TypeOfReference.СписокНоменклатуры:
                    try {
                        testObject.Delete(); // Производим обычное удаление объекта
                    }
                    catch (Exception e) {
                        errors.Add($"{testObject.ToString()}: '{e.Message}'");
                    }
                    break;
                case TypeOfReference.ЭСИ:
                case TypeOfReference.Документы:
                case TypeOfReference.Материалы:
                case TypeOfReference.ЭлектронныеКомпоненты:
                    foreach (ReferenceObject revision in GetAllRevisionOf(testObject)) {
                        try {
                            Desktop.UndoCheckOut(revision); // Отметяем создание новых объектов
                        }
                        catch (Exception e) {
                            errors.Add($"{testObject.ToString()}: '{e.Message}'");
                        }
                    }
                    break;
                default:
                    errors.Add($"{testObject}: 'Метод не работает с объектами справочника {testObject.Reference.ToString()}'");
                    break;
            }
        }

        if (errors.Count != 0) {
            string message = string.Join(Environment.NewLine, errors);
            Message("Информация", $"{nameof(CleanTestObjects)}: в процессе работы метода возникли следующие ошибки:{Environment.NewLine}{message}");
        }
    }

    /// <summary>
    /// Функция для создания тестовых объектов в справочнике "Список номенклатуры FoxPro"
    /// </summary>
    /// <param name="amount">Количество объектов для создания</param>
    /// <returns>Список созданных объектов</returns>
    private List<ReferenceObject> CreateTestNomenclatureRecords(int amount) {
        List<ReferenceObject> result = Enumerable
            .Range(0, amount)
            .Select(index => СписокНоменклатуры.Ref.CreateReferenceObject())
            .ToList<ReferenceObject>();

        int i = 0;
        // Производим заполнение параметров
        foreach (ReferenceObject testObject in result) {
            // Задаем основные параметры
            testObject[СписокНоменклатуры.Params["Обозначение"]].Value = $"TestObject-{i + 1}";
            testObject[СписокНоменклатуры.Params["Наименование"]].Value = $"TestObject-{i + 1}";
            // Указываем тип для объекта
            switch (i % 3) {
                case 0:
                    testObject[СписокНоменклатуры.Params["Тип номенклатуры"]].Value = GetIntFrom(TypeOfObject.Другое);
                    break;
                case 1:
                    testObject[СписокНоменклатуры.Params["Тип номенклатуры"]].Value = GetIntFrom(TypeOfObject.Материал);
                    break;
                case 2:
                    testObject[СписокНоменклатуры.Params["Тип номенклатуры"]].Value = GetIntFrom(TypeOfObject.ЭлектронныйКомпонент);
                    break;
                default:
                    throw new Exception($"{nameof(CreateTestNomenclatureRecords)}: Ошибка при присвоении типа для объекта 'TestObject-{i + 1}'");
            }
            testObject.EndChanges();
            i++;
        }

        return result;
    }

    /// <summary>
    /// Функция возвращает словарь с изделиями
    /// </summary>
    private Dictionary<string, ReferenceObject> GetNomenclature()
    {
        var ListNum = СписокНоменклатуры.Ref.Objects;
        var dictListNum = ListNum.ToDictionary(objref => (objref[СписокНоменклатуры.Params["Обозначение"]].Value.ToString()));
        return dictListNum;
    }


    /// <summary>
    /// Функция возвращает словарь с подключениями, сгруппированными по параметру 'Сборка'
    /// </summary>
    private Dictionary<string, List<ReferenceObject>> GetLinks()
    {
        var RefConnectNum = Подключения.Ref.Objects;
        Dictionary<string, List<ReferenceObject>> dict = new Dictionary<string, List<ReferenceObject>>(300000);
        foreach (var item in RefConnectNum)
        {
            string shifr = (String)item[Подключения.Params["Сборка"]].Value;
            if (dict.ContainsKey(shifr))
            {
                dict[shifr].Add(item);
            }
            else
            {
                dict.Add(shifr, new List<ReferenceObject>() { item });
            }
        }
        return dict;
    }

    /// <summary>
    /// Получает объекты справочника по условию если parametr содердит строку str
    /// </summary>
    public List<ReferenceObject> GetFilterRefObj(String str, Guid guidref, Guid parametr)
    {
        ReferenceInfo info = Context.Connection.ReferenceCatalog.Find(guidref);
        Reference reference = info.CreateReference();
        ParameterInfo parameterInfo = reference.ParameterGroup[parametr];
        List<ReferenceObject> result = reference.Find(parameterInfo, ComparisonOperator.Equal, str);
        return result;
    }

    /// <summary>
    /// Функция запрашивает у пользователя, какие изделия необходимо выгрузить.
    /// Если введенное пользователем изделие отсутствует среди всех изделий, должно выдаваться об этом сообщение
    /// Так же должна быть возможность снова ввести данные.
    /// </summary>
    /// <param name="nomenclature">Принимает индексированные по обозначениям данные из справочника "Список номенклатуры FoxPro"</param>
    /// <returns>Возвращает список изделий, для которых необходимо произвести выгрузку</returns>
    private List<string> GetShifrsFromUserToImport(Dictionary<string, ReferenceObject> nomenclature)
    {
        List<string> result = new List<string>();
        /*
        string shifr = "УЯИС.731353.037";
        var filter =  GeFiltertRefObj(shifr, СписокНоменклатуры.RefGuid, СписокНоменклатуры.Params["Обозначение"]);
        if (filter != null)
        {
            result = (filter.Select(objref => (objref[СписокНоменклатуры.Params["Обозначение"]].Value.ToString()))).ToList();
        }
        */

        ДиалогВвода диалог = СоздатьДиалогВвода("Введите значения");
        диалог.ДобавитьСтроковое("Введите обозначение изделия", "УЯИС.731353.038\nУЯИС.731353.037\nУЯИС794711004", многострочное: true, количествоСтрок: 10);
        ДиалогВыбораОбъектов диалог2 = СоздатьДиалогВыбораОбъектов("Список номенклатуры FoxPro");
        диалог2.Заголовок = "Выбор изделий для импорта";
        диалог2.Вид = "Список";
        диалог2.МножественныйВыбор = true;
        диалог2.ВыборФлажками = true;
        диалог2.ПоказатьПанельКнопок = false;

        if (диалог.Показать())
        {
            string oboz = диалог["Введите обозначение изделия"];
            Normalizer normalizer = new Normalizer();
            normalizer.setprefix = false;

            var list_oboz = oboz.Split(new string[] { "\n" }, StringSplitOptions.RemoveEmptyEntries);
            var list_oboz_insert_dot = list_oboz.Select(oboz => normalizer.NormalizeShifrsFromFox(oboz.Replace(".", "")));
            if (list_oboz.Length == 1)
            {
                oboz = normalizer.NormalizeShifrsFromFox(oboz.Replace(".", ""));
                диалог2.Фильтр = $"[Обозначение] начинается с '{oboz}'";
            }
            else
            {
                var filter = String.Join(",", list_oboz_insert_dot.Distinct());
                диалог2.Фильтр = $"[Обозначение] Входит в список '{filter}'";
            }
            if (диалог2.Показать())
            {
                var selectobj = диалог2.SelectedObjects;

                //var s2 = (IEnumerable<ReferenceObject>)selectobj;
                //result = (List<string>)s2.Select(objref => (objref[СписокНоменклатуры.Params["Обозначение"]].Value.ToString()));
                //result = (List<string>)selectobj.Select(objref => (objref["Обозначение"].ToString()));

                foreach (var item in selectobj)
                {
                    result.Add(item["Обозначение"].ToString());
                }

            }
        }
        SetNameOfExport(result);
        return result;
    }

    /// <summary>
    /// Вспомогательная функция для определения названия импотра для использования этих данных в названиях
    /// лог файлов.
    /// Функция производит присвоение строки определенного формата переменной NameOfImport
    /// </summary>
    /// <param name="nomenclature">Список изделий, переданных на выгрузку</param>
    private void SetNameOfExport(List<string> nomenclature) {
     
        if (nomenclature.Count == 0)
            return;

        string addition = nomenclature.Count > 1 ? $" (+{nomenclature.Count - 1})" : string.Empty;
        NameOfImport = $"{nomenclature[0]}{addition}";
    }

    /// <summary>
    /// Функция для определения перечня ДСЕ, которые необходимо обработать во время загрузки данных из FoxPro
    /// </summary>
    /// <param name="nomenclature">Словарь со всеми объектами справочника "Список номенклатуры FoxPro", проиндексированный по обозначениям</param>
    /// <param name="links">Словарь со всеми объектами справочника "Подключения", проиндексированный и сгруппированный по родительсвой сборке</param>
    /// <param name="shifrs">Список с обозначениями изделий, которые необходимо выгрузить</param>
    /// <returns>Объект типа HashSet, содержащий в себе перечень всех объектов справочника "Список номенклатуры FoxPro, выгрузку которых необходимо произвести"</returns>
    private HashSet<ReferenceObject> GetNomenclatureToProcess(Dictionary<string, ReferenceObject> nomenclature, Dictionary<string, List<ReferenceObject>> links, List<string> shifrs) {
     
        // Для каждого шифра создаем объект, реализующий интерфейс ITree, получаем входящие объекты, добавляем их в HashSet (для исключения дубликатов)
        // В конце пишем лог, в котором записываем информацию о сгенерированном дереве, количестве входящих объектов и их структуре
        if (nomenclature == null || links == null || shifrs == null)
            throw new Exception($"{nameof(GetNomenclatureToProcess)}: На вход функции GetNomenclatureToProcess были поданы отсутствующие значения");
        HashSet<ReferenceObject> result = new HashSet<ReferenceObject>();

        // Для лога
        string log = string.Empty;
        string pathToLogFile = Path.Combine(ДиректорияДляЛогов, $"Деревья для {StringForLog}.txt");

        List<string> errors = new List<string>();
        // Формируем деревья и получаем все объекты
        foreach (string shifr in shifrs) {
         
            NomenclatureTree tree = new NomenclatureTree(nomenclature, links, shifr, СписокНоменклатуры, Подключения);
            result.UnionWith(tree.AllReferenceObjects);
            log += $"Дерево изделия {shifr}:\n\n{tree.GenerateLog()}\n\n";
            if (tree.HaveErrors)
                errors.Add(tree.ErrString);
        }

        // Пишем лог
        File.WriteAllText(pathToLogFile, log);
        log_save_ref.file_name_tree = pathToLogFile;

        // Выдаем пользователю сообщение об ошибках в дереве и спрашиваем, продолжать ли выгрузку
        if (errors.Count != 0)
            if (Question(string.Join("\n\n", errors) + "\n\nПродолжать выгрузку?"))
                return result;
            else
                return new HashSet<ReferenceObject>();
        return result;
    }

    /// <summary>
    /// Функция производит поиск выгружаемого объекта во всех справочниках, в которых он может быть.
    /// Если объект находится, то он возвращается, если нет - создается новый и возвращается уже он.
    /// Функция поделена на четыре стадии для удобства, в качестве входных параметров принимает объекты справочника "Список номенклатуры FoxPro", которые необходимо выгрузить
    /// </summary>
    /// <param name="nomenclature">HashSet с перечнем объектов справочника "Список номенклатуры FoxPro", для которых необходимо произвести выгрузку</param>
    /// <returns>Список найденных или созданных объектов справочника ЭСИ</returns>
    private List<ReferenceObject> FindOrCreateNomenclatureObjects(HashSet<ReferenceObject> nomenclature) {
     
        // Функция принимает записи справочника "Список номенклатуры FoxPro" для создания объектов с справочнике ЭСИ и смежных справочников
        // Функция возвращает найденные или созданные записи справочника ЭСИ
        //
        // Необходимо реализовать:
        // - Поиск объектов в справочнике ЭСИ
        // - Создание объектов, если они не были найдены
        // - Подключение созданных или найденных объектов к соответствующим записям справочника "Список номенклатуры FoxPro"
        // - Возврат всех найденных/созданных объектов для последующей с ними работы
        // - Вывод лога о всех произведенных действиях

        // ВАЖНО: при проведении поиска нужно проверять на то, что найденный объект единственный. Если он не единственный, тогда нужно выдать ошибку для принятия решения по поводу обработки данного случая
        List<ReferenceObject> result = new List<ReferenceObject>();
        string pathToLogFile = Path.Combine(ДиректорияДляЛогов, $"Поиск позиций для {StringForLog}.txt");
        List<string> messages = new List<string>();

        // Счетчики для статистики
        int countStage1 = 0;
        int countStage2 = 0;
        int countStage3 = 0;
        int countStage4 = 0;
        int countErrors = 0;

        foreach (ReferenceObject nom in nomenclature) {
         
            // Получаем обозначение текущего объекта и его тип
            ReferenceObject resultObject;
            string nomDesignation = (string)nom[СписокНоменклатуры.Params["Обозначение"]].Value; // обозначение
            TypeOfObject nomType = GetTypeOfObjectFrom(nom); // тип
            // Пишем информацию в лог
            messages.Add(new String('-', 30));
            messages.Add($"{nomDesignation}:");




            try {
             
                // СТАДИЯ 1: Пытаемся получить объект по связи на справочник ЭСИ
                resultObject = ProcessFirstStageFindOrCreate(nom, nomDesignation, nomType, messages);
                if (resultObject != null) {
                    CheckAndLink(nom, resultObject, messages); // Подключаем объект к списку номенклатуры
                    result.Add(resultObject); // Добавляем объект в результат
                    countStage1 += 1; // Обновляем счетчик успешного выполнения стадии
                    continue;
                }

                // СТАДИЯ 2: Пытаемся найти объект в справочнике ЭСИ
                resultObject = ProcessSecondStageFindOrCreate(nom, nomDesignation, nomType, messages);
                if (resultObject != null) {
                    CheckAndLink(nom, resultObject, messages); // Подключаем объект к списку номенклатуры
                    result.Add(resultObject); // Добавляем объект в результат
                    countStage2 += 1; // Обновляем счетчик успешного выполнения стадии
                    continue;
                }

                // СТАДИЯ 3: Пытаемся найти объект в смежных справочниках
                resultObject = ProcessThirdStageFindOrCreate(nom, nomDesignation, nomType, messages);
                if (resultObject != null) {
                    CheckAndLink(nom, resultObject, messages); // Подключаем объект к списку номенклатуры
                    result.Add(resultObject); // Добавляем объект в результат
                    countStage3 += 1; // Обновляем счетчик успешного выполнения стадии
                    continue;
                }

                // СТАДИЯ 4: Создаем объект исходя из того, какой был определен тип в справочнике "Список номенклатуры FoxPro"
                resultObject = ProcessFinalStageFindOrCreate(nom, nomDesignation, nomType, messages);
                if (resultObject != null) {
                    CheckAndLink(nom, resultObject, messages); // Подключаем объект к списку номенклатуры
                    result.Add(resultObject); // Добавляем объект в результат
                    countStage4 += 1; // Обновляем счетчик успешного выполнения стадии
                }
            }
            catch (Exception e) {
             
                messages.Add($"ERROR: {e.Message}");
                countErrors += 1;
                continue;
            }
        }

        // Вывод статистики и генерация лога
        // Генерируем строку с статистикой по данному методу
        string statistics = 
                $"Всего объектов передано на поиск/создание - {nomenclature.Count.ToString()}. Из них:\n" +
                $"Получено по связи - {countStage1.ToString()} шт\n" +
                $"Найдено в ЭСИ - {countStage2.ToString()} шт\n" +
                $"Найдено в смежных справочниках - {countStage3.ToString()} шт\n" +
                $"Создано - {countStage4.ToString()} шт\n" +
                $"Возникли ошибки в процессе обработки - {countErrors.ToString()} шт\n";

        // Получаем текст всех ошибок, которые были перехвачены
        log_save_ref.count_error = countErrors;
        log_save_ref.count_object = nomenclature.Count;
        log_save_ref.create_object = countStage4;
        log_save_ref.timeStop = DateTime.Now;

        string errors = string.Join("\n\n", messages.Where(message => message.StartsWith("Error")));





        // Если ошибки есть, то сначала выводим статистику и спрашиваем у пользователя, отобразить ли ошибки
        // Иначе просто выводим статистику
        if (!log_save_ref.messageOff)
        {
            if (errors != string.Empty)
            {

                statistics += "\n\nОтобразить ошибки?";
                if (Question(statistics))
                    Message("Ошибки", $"В процессе поиска и создания номенклатурных объектов возникли следующие ошибки\n{errors}");
            }
            else
                Message("Статистика", statistics);
        }
        // Пишем лог
        File.WriteAllText(pathToLogFile, string.Join("\n", messages));
        log_save_ref.file_name_error_log = pathToLogFile;
        return result;
    }

    /// <summary>
    /// Функция предназначена для переноса параметра обозначение в параметр код ОКП.
    /// Код обрабатывает записи, которые относятся к типам, для которых SHIFR является не полем "Обозначение" а полем "ОКП".
    /// </summary>
    /// <param name="resultObject">Объект, для которого нужно произвести перенос данных полей "Обозначение" и "Код ОКП"</param>
    /// <param name="designation">Обозначение объекта</param>
    /// <param name="nomtype">Объект типа TypeOfObject, представляющий собой тип номенклатурного объекта</param>
    /// <param name="messages">Список строк, представляющий собой лог ведения выгрузки</param>
    /// <returns>Возвращает resultObject с произведенными над ним изменениями</returns>
    private ReferenceObject MoveShifrToOKP(ReferenceObject resultObject, string designation, TypeOfObject nomtype, List<string> messages)
    {
        
        //resultObject = ProcessThirdStageFindOrCreate(nom, nomDesignation, nomType, messages);
        RefGuidData referData = GetRefGuidDataFrom(resultObject);
        TypeOfObject esiType = GetTypeOfObjectFrom(resultObject); // тип
       
        var obozEsi = resultObject[referData.Params["Обозначение"]].Value.ToString();

        if (esiType == TypeOfObject.СтандартноеИзделие || esiType == TypeOfObject.Материал || esiType == TypeOfObject.ЭлектронныйКомпонент
           || (esiType == TypeOfObject.Другое && (nomtype == TypeOfObject.СтандартноеИзделие || nomtype == TypeOfObject.Материал || nomtype == TypeOfObject.ЭлектронныйКомпонент)))
        {
            var okpEsi = resultObject[referData.Params["Код ОКП"]].Value.ToString();

            if (!designation.Equals(okpEsi))
            {
                if (designation.Equals(obozEsi))
                {
                    try
                    {
                        if (!resultObject.IsCheckedOut)
                            resultObject.CheckOut();
                        resultObject.BeginChanges();
                        resultObject[referData.Params["Код ОКП"]].Value = obozEsi;
                        resultObject[referData.Params["Обозначение"]].Value = " ";
                        resultObject.EndChanges();
                        //linkedObject.CheckIn("Перенос Обозначения в поле код ОКП");
                        Desktop.CheckIn(resultObject, "Перенос Обозначения в поле код ОКП", false);
                        messages.Add($"{nameof(MoveShifrToOKP)}: Перенос обозначения в код ОКП прошёл успешно");
                    }
                    catch (Exception e)
                    {
                        throw new Exception($"{nameof(MoveShifrToOKP)}: ошибка во время переноса 'Обозначения' в 'ОКП':\n{e}");
                    }
                }
                else
                    throw new Exception($"{nameof(MoveShifrToOKP)}: у привязанного объекта отличается код ОКП (указан: '{okpEsi}', должен быть: {designation})");

            }
            return resultObject;
        }
        else
        {
            if (!designation.Equals(obozEsi))
                throw new Exception($"{nameof(MoveShifrToOKP)}: У привязанного объекта отличается обозначенние (указано: '{obozEsi}', должно быть: {designation})");
            return resultObject;
        }

    }

    /// <summary>
    /// Вспомогательный код для функции FindOrCreateNomenclatureObjects.
    /// Данный код производит пробует получить объект по связи, а так же проверить корректность полученного объекта, если таковой имеется.
    /// </summary>
    /// <param name="nom">Объект справочника "Список номенклатуры FoxPro", для которого производится поиск связанного объекта</param>
    /// <param name="designation">Обозначение объекта справочника "Список номенклатуры FoxPro"</param>
    /// <param name="nomtype">Объект типа TypeOfObject, представляющий собой тип номенклатурного объекта</param>
    /// <param name="messages">Список строк, представляющий собой лог ведения выгрузки</param>
    /// <returns>Возвращает ReferenceObject, если объект получилось найти, возвращает null, если объект не получилось найти, выбрасывает ошибку, если возникли проблемы в процессе</returns>
    private ReferenceObject ProcessFirstStageFindOrCreate(ReferenceObject nom, string designation, TypeOfObject type, List<string> messages) {
     
        // Получаем объект по связи и, если он есть, производим его проверку
        ReferenceObject linkedObject = nom.GetObject(СписокНоменклатуры.Links["Связь на ЭСИ"]);

        if (linkedObject != null)
        {
            messages.Add($"{nameof(ProcessFirstStageFindOrCreate)}: Обнаружен связанный объект");
            // Производим проверку подключенного объекта.
            ReferenceObject resultObject = null;
            if (IsFoxRevisionCorrect(linkedObject, designation)) {
                // Если он имеет корректную ревизию, производим дальнейшие корректировки и возвращаем его
                messages.Add($"{nameof(ProcessFirstStageFindOrCreate)}: объект имеет корректную ревизию ({linkedObject.SystemFields.RevisionName})");
                resultObject = linkedObject;
            }
            else {
                // На основе подключенного объекта создаем новую ревизию, производим соответствующие корректировки и возвращаем его
                messages.Add($"{nameof(ProcessFirstStageFindOrCreate)}: Ревизия подключенного объекта ({linkedObject.SystemFields.RevisionName}) не соответствует корректному наименованию.");
                resultObject = FindFoxRevision(linkedObject, designation, messages); // Пробуем найти уже существующую корректную ревизию
                // Пробуем найти нужную ревизию, и если нам это не удается, создаем новую
                if (resultObject == null) {
                    resultObject = CreateFoxRevision(linkedObject, designation, messages);
                    messages.Add($"{nameof(ProcessFirstStageFindOrCreate)}: создание ревизии");
                }
                else
                    messages.Add($"{nameof(ProcessFirstStageFindOrCreate)}: переподключение корректной ревизии");
            }

            MoveShifrToOKP(resultObject, designation, type, messages);
            SyncronizeTypes(nom, resultObject, messages);
            return resultObject;
   
        }
        else
            messages.Add($"{nameof(ProcessFirstStageFindOrCreate)}: Связанный объект отсутствовал");
        return null;
    }

    /// <summary>
    /// Вспомогательный код для функции FindOrCreateNomenclatureObjects.
    /// Данный код запускается в том случае, если не удалось получить объект по связи и производит поиск объекта в справочнике ЭСИ, и если таковой имеется, производит проверку его типа.
    /// </summary>
    /// <param name="nom">Объект справочника "Список номенклатуры FoxPro", для которого производится поиск связанного объекта</param>
    /// <param name="shifr">Обозначение объекта справочника "Список номенклатуры FoxPro"</param>
    /// <param name="nomtype">Объект типа TypeOfObject, представляющий собой тип номенклатурного объекта</param>
    /// <param name="messages">Список строк, представляющий собой лог ведения выгрузки</param>
    /// <returns>Возвращает ReferenceObject, если объект получилось найти, возвращает null, если объект не получилось найти, выбрасывает ошибку, если возникли проблемы в процессе</returns>
    private ReferenceObject ProcessSecondStageFindOrCreate(ReferenceObject nom, string shifr, TypeOfObject type, List<string> messages) {
        
        // Производим преобразование шифра FoxPro, который может содержать информацию о варианте, в очищенное от лишней информации обозначение T-Flex
        string designation = GetDesignationFrom(shifr);
     
        // -- Начало поиска
        HashSet<ReferenceObject> allFinded = new HashSet<ReferenceObject>(); // контейнер для хранения результатов поиска

        // поиск по обозначению
        allFinded.UnionWith(
                ЭСИ.Ref
                .Find(ЭСИ.Params["Обозначение"], designation) // Производим поиск по всему справочнику
                .Where(finded => finded.Class.IsInherit(ЭСИ.Types["Материальный объект"])) // Отфильтровываем только те объекты, которые наследуются от 'Материального объекта'
                );

        // поиск по коду ОКП
        allFinded.UnionWith(
                ЭСИ.Ref
                .Find(ЭСИ.Params["Код ОКП"], designation) // Производим поиск по всему справочнику
                .Where(finded => finded.Class.IsInherit(ЭСИ.Types["Материальный объект"])) // Отфильтровываем только те объекты, которые наследуются от 'Материального объекта'
                );

        // -- Окончание поиска
        
        // -- Начало анализа полученных данных

        ReferenceObject foxRevision = null;

        switch (allFinded.Count) {
            case 0:
                // Возвращаем null
                messages.Add($"{nameof(ProcessSecondStageFindOrCreate)}: не найдено ни одного совпадения");
                return null;
            case 1:
                // Проверяем объект на то, является ли он ревизией T-Flex, и если не является, создаем ревизию и возвращаем ее
                messages.Add($"{nameof(ProcessSecondStageFindOrCreate)}: найдено одно совпадение");
                foxRevision = FindFoxRevision(allFinded.First(), designation, messages);
                if (foxRevision != null) {
                    messages.Add($"{nameof(ProcessSecondStageFindOrCreate)}: найденный объект является корректной ревизией FoxPro");
                }
                else {
                    messages.Add($"{nameof(ProcessSecondStageFindOrCreate)}: найденный объект не является корректной ревизией FoxPro");
                    foxRevision = CreateFoxRevision(allFinded.First(), shifr, messages);
                }
                break;
            default:
                // Проверяем объекты на то, что они являются ревизиями одного объекта (иначе выбрасываем ошибку).
                // Eсли так, ищем среди них ревизию Фокс, и если не находим, создаем ее
                messages.Add($"{nameof(ProcessSecondStageFindOrCreate)}: найдено {allFinded.Count} объектов");
                if (IsOneLogicalObject(allFinded)) {
                    messages.Add($"{nameof(ProcessSecondStageFindOrCreate)}: все нейденные объекты являются ревизиями одного объекта");
                    foxRevision = FindFoxRevision(allFinded.First(), designation, messages);
                    // Создаем новую ревизию, если ее не удалось найти
                    if (foxRevision == null)
                        foxRevision = CreateFoxRevision(allFinded.First(), shifr, messages);
                    break;
                }
                else {
                    throw new Exception(
                            $"{nameof(ProcessSecondStageFindOrCreate)}: найденные объекты относятся к разным логическим объектам:{Environment.NewLine}" +
                            string.Join(Environment.NewLine, allFinded.Select(finded => $"- {finded.ToString()} ({finded.SystemFields.LogicalObjectGuid.ToString()})"))
                            );
                }

        }

        MoveShifrToOKP(foxRevision, designation, type, messages);
        SyncronizeTypes(nom, foxRevision, messages);
        return foxRevision;

        // -- Окончание анализа полученных данных
    }

    /// <summary>
    /// Вспомогательный код для функции FindOrCreateNomenclatureObjects.
    /// Данный код запускается в том случае, если не удалось найти объект в ЭСИ и производит поиск объекта в смежных справочниках, и если такой имеется, производит проверку типа и создание объекта ЭСИ.
    /// </summary>
    /// <param name="nom">Объект справочника "Список номенклатуры FoxPro", для которого производится поиск связанного объекта</param>
    /// <param name="shifr">Обозначение объекта, полученное из FoxPro (может содержать информацию о варианте)</param>
    /// <param name="nomtype">Объект типа TypeOfObject, представляющий собой тип номенклатурного объекта</param>
    /// <param name="messages">Список строк, представляющий собой лог ведения выгрузки</param>
    /// <returns>Возвращает ReferenceObject, если объект получилось найти, возвращает null, если объект не получилось найти, выбрасывает ошибку, если возникли проблемы в процессе</returns>
    private ReferenceObject ProcessThirdStageFindOrCreate(ReferenceObject nom, string shifr, TypeOfObject type, List<string> messages) {

        string designation = GetDesignationFrom(shifr);
     
        // Производим поиск по смежным справочникам
        HashSet<ReferenceObject> allFinded = new HashSet<ReferenceObject>(); // Контейнер для результатов поиска

        // Справочник "Документы"
        allFinded.UnionWith(new HashSet<ReferenceObject>(
                    Документы.Ref
                        .Find(Документы.Params["Обозначение"], designation)
                        .Where(finded => finded.Class.IsInherit(Документы.Types["Объект состава изделия"]))
                    ));
        allFinded.UnionWith(new HashSet<ReferenceObject>(
                    Документы.Ref
                        .Find(Документы.Params["Код ОКП"], designation)
                        .Where(finded => finded.Class.IsInherit(Документы.Types["Объект состава изделия"]))
                    ));

        // Справочник "Электронные компоненты"
        allFinded.UnionWith(new HashSet<ReferenceObject>(
                    ЭлектронныеКомпоненты.Ref
                        .Find(ЭлектронныеКомпоненты.Params["Обозначение"], designation)
                    ));
        allFinded.UnionWith(new HashSet<ReferenceObject>(
                    ЭлектронныеКомпоненты.Ref
                        .Find(ЭлектронныеКомпоненты.Params["Код ОКП"], designation)
                    ));

        // Справочник "Материалы"
        allFinded.UnionWith(new HashSet<ReferenceObject>(
                    Материалы.Ref
                        .Find(Материалы.Params["Обозначение"], designation)
                    ));
        allFinded.UnionWith(new HashSet<ReferenceObject>(
                    Материалы.Ref
                        .Find(Материалы.Params["Код ОКП"], designation)
                    ));

        // Обрабатываем результаты поиска
        ReferenceObject targetObject = null;

        switch (allFinded.Count) {
            case 0:
                messages.Add($"{nameof(ProcessThirdStageFindOrCreate)}: Объект не найден в смежных справочниках");
                return null;
            case 1:
                messages.Add($"{nameof(ProcessThirdStageFindOrCreate)}: Объект '{allFinded.First().ToString()}' найден в справочнике {GetTypeOfReferenceFrom(allFinded.First()).ToString()}");
                // Передаем объект для дальнейшей обработки
                targetObject = allFinded.First();
                break;
            default:
                // TODO: Реализовать обработку множественного случая
                // Случай, когда получилось найти много объектов
                messages.Add($"{nameof(ProcessThirdStageFindOrCreate)}: было найдено несколько объектов");
                // производим проверку результатов поиска и отсеиваем заведомо ложные случаи
                if (!IsOneReference(allFinded))
                    throw new Exception($"{nameof(ProcessThirdStageFindOrCreate)}: найденные объекты принадлежат разным справочникам");
                if (!IsOneLogicalObject(allFinded))
                    throw new Exception($"{nameof(ProcessThirdStageFindOrCreate)}: найденный объекты не являются ревизиями одного логического объекта");
                // Передаем объект для дальнейшей обработки. Здесь можно предусмотреть более сложный алгоритм выбора, но в данный момент просто передается первый объект
                targetObject = allFinded.First();
                break;
        }

        // Производим обработку найденного объекта. TODO: Учесть, что данные корректировки будут производиться только для
        // одной выбранной ревизии из всех найденных. Возможно такой вариант редактирования не будет подходить для финальной версии макроса
        MoveShifrToOKP(targetObject, designation, type, messages);
        SyncronizeTypes(nom, targetObject, messages);
        ReferenceObject newNomObject = ConnectRefObjectToESI(targetObject, designation);
        
        // Возвращаем созданную ревизию Фокс
        return CreateFoxRevision(newNomObject, shifr, messages);
    }


    /// <summary>
    /// Вспомогательный код для функции FindOrCreateNomenclatureObjects.
    /// Данный код запускается только в том случае, если объет найти не получилось и производит создание объекта сначала в смежном справчонике, а затем в ЭСИ.
    /// </summary>
    /// <param name="nom">Объект справочника "Список номенклатуры FoxPro", для которого производится поиск связанного объекта</param>
    /// <param name="designation">Обозначение объекта справочника "Список номенклатуры FoxPro"</param>
    /// <param name="nomtype">Объект типа TypeOfObject, представляющий собой тип номенклатурного объекта</param>
    /// <param name="messages">Список строк, представляющий собой лог ведения выгрузки</param>
    /// <returns>Возвращает ReferenceObject, если объект получилось найти, возвращает null, если объект не получилось найти, выбрасывает ошибку, если возникли проблемы в процессе</returns>
    private ReferenceObject ProcessFinalStageFindOrCreate(ReferenceObject nom, string designation, TypeOfObject type, List<string> messages) {
     
        // Производим создание объекта
        string nomName = (string)nom[СписокНоменклатуры.Params["Наименование"]].Value;
        //string nomTip = nom[СписокНоменклатуры.Params["Тип номенклатуры"]].Value.ToString();
        //var createDocument = CreateDocumentObject(nomName, designation, type.ToString());
        ReferenceObject createDocument = null;
        string nomTip;
        try {
         
            nomTip = GetStringFrom(type);
        }
        catch {
         
            throw new Exception($"{nameof(ProcessFinalStageFindOrCreate)}: Ошибка при создании объекта. Невозможно создать объект типа '{type.ToString()}'");
        }

         
        createDocument = CreateRefObject(nom, nomName, designation, nomTip, GetRefGuidDataFrom(type));
         
                                                       
         
                                                                                                    
         
                                                                               
         
                                                                                                                            
         
            
         
                                                                                                    
         

        // var createDocument2 = CreateRefObject(nomName, designation, type.ToString(), ЭлектронныеКомпоненты.Ref,ЭлектронныеКомпоненты.Params["Наименование"], ЭлектронныеКомпоненты.Params["Код ОКП"];
        // var createDocument = CreateRefObject(nomName, designation, type.ToString(),Документы.Ref,Документы.Params["Наименование"],Документы.Params["Обозначение"]);

        if (createDocument != null)
        {
            messages.Add($"{nameof(ProcessFinalStageFindOrCreate)}: В справочнике {createDocument.Reference.Name.ToString()} создан объект {designation}");
            return createDocument;
        }
        return null;
    }

    /// <summary>
    /// Метод для создание ревизии ФОКС для объекта
    /// </summary>
    /// <param name="initialObject">Исходный объект, на основе которого будет создаваться новая ревизия</param>
    /// <param name="shifr">Обозначение изделия из Fox, для которого нужно создать новую ревизию</param>
    /// <param name="messages">Список строк, представляющий собой лог ведения выгрузки</param>
    /// <returns>Созданная ревизия</returns>
    private ReferenceObject CreateFoxRevision(ReferenceObject initialObject, string shifr, List<string> messages) {
        // -- Начало валидации
        // Производим проверку входного объекта
        if (GetTypeOfReferenceFrom(initialObject) != TypeOfReference.ЭСИ)
            throw new Exception(
                    $"{nameof(CreateFoxRevision)}: в функцию передан объект " +
                    $"неподдерживаемого справочника '{GetTypeOfReferenceFrom(initialObject).ToString()}.{Environment.NewLine}'" +
                    "Для корректной работы метода передавать в него следует объект справочника ЭСИ"
                    );

        // Производим проверку на то, что данной ревизии у объекта еще нет
        string nameOfRevision = GetRevisionName(shifr);
        if (!IsRevisionNameFree(initialObject, nameOfRevision))
            throw new Exception($"{nameof(CreateFoxRevision)}: У объекта {initialObject.ToString()} уже есть ревизия с названием {nameOfRevision}");
        // -- Окончание валидации

        // -- Начало создания ревизии
        // Создаем ревизию
        var level = initialObject.Class.RevisionNamingRule.RevisionLevels[1]; // Получаем уровень мажорной ревизии для ее повышения
        ReferenceObject newRevision = null;

        // Пытаемся создать новую ревизию
        try {
            newRevision = (initialObject as NomenclatureObject).CreateRevision(revisionLevel: level, recursive: false); // Создаем новую ревизию, не подключая к ней связанных объектов
        }
        catch (Exception e) {
            throw new Exception($"{nameof(CreateFoxRevision)}: {e.Message}");
        }
        messages.Add($"{nameof(CreateFoxRevision)}: произведено создание новой ревизии");

        // Переименовываем ревизию в номенклатурном объекте
        newRevision.BeginChanges();
        newRevision.SystemFields.RevisionName = nameOfRevision;
        newRevision.EndChanges();
        // Переименовываем ревизию в связанном объекте
        ReferenceObject linkedObject = (newRevision as NomenclatureObject).LinkedObject;
        linkedObject.BeginChanges();
        linkedObject.SystemFields.RevisionName = nameOfRevision;
        linkedObject.EndChanges();
        messages.Add($"{nameof(CreateFoxRevision)}: произведено переименование новой ревизии");

        // Автоматическое применение изменений, если потребуется
        if (applyCreation)
            Desktop.CheckIn(newRevision, "Создание ревизии FoxPro", false);
        // -- Окончание создания ревизии
        return newRevision;
    }

    /// <summary>
    /// Метод для получения всех ревизий, существующих для конкретного объекта
    /// </summary>
    /// <param name="sourceObject">Объект, для которого нужно получить все связанные ревизии</param>
    /// <returns>Список ревизий</returns>
    private List<ReferenceObject> GetAllRevisionOf(ReferenceObject sourceObject) {
        return ЭСИ.Ref.Find(ЭСИ.Params["Guid логического объекта"], sourceObject.SystemFields.LogicalObjectGuid);
    }

    /// <summary>
    /// Функция для проверки на то, что все объекты являются ревизиями одного объекта
    /// </summary>
    /// <param name="findedObjects">Список объектов, которые проверяются на принадлежность одному логическому объекту</param>
    /// <returns>true - если все объекты являются одним логическим объектом, false - если нет</returns>
    public bool IsOneLogicalObject(IEnumerable<ReferenceObject> findedObjects) {
        // -- Начало валидации данных
        if ((findedObjects == null) || (findedObjects.Count() == 0))
            throw new Exception($"{nameof(IsOneLogicalObject)}: Параметр {nameof(findedObjects)} не содержит данных");
        if (findedObjects.Count() == 1)
            throw new Exception($"{nameof(IsOneLogicalObject)}: Параметр {nameof(findedObjects)} содержит только одно значение (должно быть как минимум два)");
        // -- Окончание валидации данных

        // Проверка
        Guid logicalGuid = findedObjects.First().SystemFields.LogicalObjectGuid;
        foreach (ReferenceObject finded in findedObjects.Skip(1))
            if (finded.SystemFields.LogicalObjectGuid != logicalGuid)
                return false;

        return true;
    }

    /// <summary>
    /// Функция для проверки на то, что все объекты принадлежат одному справочнику
    /// </summary>
    /// <param name="findedObjects"></param>
    /// <returns>true - если все объекты принадлежат к одному справочнику, false - если нет</returns>
    public bool IsOneReference(IEnumerable<ReferenceObject> findedObjects) {
        // -- Начало валидации данных
        if ((findedObjects == null) || (findedObjects.Count() == 0))
            throw new Exception($"{nameof(IsOneLogicalObject)}: Параметр {nameof(findedObjects)} не содержит данных");
        if (findedObjects.Count() == 1)
            throw new Exception($"{nameof(IsOneLogicalObject)}: Параметр {nameof(findedObjects)} содержит только одно значение (должно быть как минимум два)");
        // -- Окончание валидации данных

        // Проверка
        TypeOfReference reference = GetTypeOfReferenceFrom(findedObjects.First());
        foreach (ReferenceObject finded in findedObjects.Skip(1))
            if (GetTypeOfReferenceFrom(finded) != reference)
                return false;

        return true;
    }

    /// <summary>
    /// Метод для поиска корректной ревизии у объекта на основании обозначения, полученного из FoxPro
    /// </summary>
    /// <param name="initialObject">Объект, для которого производится поиск ревизии</param>
    /// <param name="shifr">Обозначение, переданное из FoxPro, которое в себе может содержать название варианта</param>
    /// <returns>ReferenceObject, если получилось найти искомый объект, null, если не получилось его найти, Exception, если совпадений оказалось больше одного</returns>
    private ReferenceObject FindFoxRevision(ReferenceObject initialObject, string shifr, List<string> messages) {
        // На основе обозначения, полученного из FoxPro определяем, какая ревизия должна быть у объекта
        string revisionName = GetRevisionName(shifr);
        List<ReferenceObject> findedRevisions = GetAllRevisionOf(initialObject)
            .Where(rev => rev.SystemFields.RevisionName == revisionName)
            .ToList<ReferenceObject>();

        switch (findedRevisions.Count) {
            case 1:
                messages.Add($"{nameof(FindFoxRevision)}: Фокс ревизия найдена");
                return findedRevisions[0];
            case 0:
                messages.Add($"{nameof(FindFoxRevision)}: Фокс ревизия отсутствовала");
                return null;
            default:
                throw new Exception($"{nameof(FindFoxRevision)}: у объекта '{initialObject.ToString()}' было обнаружено {findedRevisions.Count} ревизий с названием '{revisionName}'");
        }
    }

    /// <summary>
    /// Метод для определения, свободно ли данное название ревизии
    /// </summary>
    /// <param name="sourceObject">Исходный объект, для которого проверяется доступность названия ревизии</param>
    /// <param name="revisionName">Проверяемое имя</param>
    /// <returns>Возвращает true, если название свободно, false - если название уже используется</returns>
    private bool IsRevisionNameFree(ReferenceObject sourceObject, string revisionName) {
        List<ReferenceObject> revisions = GetAllRevisionOf(sourceObject).Where(rev => rev.SystemFields.RevisionName.ToLower() == revisionName.ToLower()).ToList<ReferenceObject>();
        switch (revisions.Count) {
            case 0:
                return true;
            case 1:
                return false;
            default:
                throw new Exception($"{nameof(IsRevisionNameFree)}: Обнаружено больше одной ревизии с одним названием (объект: {sourceObject.ToString()})");
        }
    }

    /// <summary>
    /// Функция для определения корректности названия ревизии по известному обозначению
    /// (по большей части данная функция просто по обозначению определяет, какой номер у данной ревизии должен быть)
    /// и сравнивает обозначение ревизии текущего объекта с той, которой она должна быть
    /// </summary
    /// <param name="sourceObject">Исходный объект, который мы проверяем на то, является ли он корректной ревизией для данного обозначения ДСЕ</param>
    /// <param name="shifr">Обозначение ДСЕ, полученное из FoxPro, для которого мы определяем корректное название ревизии</param>
    /// <returns>Возвращает true, если наименование ревизии sourceObject правильное, или false - если не правильное</returns>
    private bool IsFoxRevisionCorrect(ReferenceObject sourceObject, string shifr) {
        return sourceObject.SystemFields.RevisionName == GetRevisionName(shifr);
    }

    // <summary>
    // Метод возвращает название ревизии в зависимости от полученного шифра объекта из FoxPro
    // </summary>
    // <param name="shifr">Обозначение объекта в FoxPro</param>
    // <returns>Наименование соответствующей ревизии</returns>
    private string GetRevisionName(string shifr) {
        return $"{RevisionNameTemplate}{GetNumRevisionFrom(shifr)}";
    }

    /// <summary>
    /// Метод для определения номера ревизии из обозначения объекта
    /// </summary>
    /// <param name="shifr">Строка с обозначением, полученным из FoxPro, на основании которого следует сделать предположение о номере ревизии</param>
    /// <returns>Номер ревизии</returns>
    private int GetNumRevisionFrom(string shifr) {
        // -- Начало верификации входных данных
        if (string.IsNullOrWhiteSpace(shifr))
            throw new Exception($"{nameof(GetNumRevisionFrom)}: переданное обозначение не содержит значения");
        // -- Окончание верификации входных данных

        int result;
        shifr = shifr.ToUpper();
        try {
            if (!shifr.Contains("ВАР"))
                result = 1;
            else {
                string num = shifr.Replace("ВАР", "~").Split('~')[1];
                result = string.IsNullOrWhiteSpace(num) ? 1 : int.Parse(num);
            }
        }
        catch (Exception e) {
            throw new Exception($"{nameof(GetNumRevisionFrom)}: в процессе определения номера ревизии для обозначение '{shifr}' возникла ошибка:{Environment.NewLine}{e.Message}");
        }
        return result;
    }

    /// <summary>
    /// Метод предназначен для удаления из обозначения не относящейся к обозначению части (это относится к названию вариантов изделия)
    /// </summary>
    /// <param name="shifr">Обозначение из FoxPro</param>
    /// <returns>Возвращает обозначение для работы с изделиями в T-Flex</returns>
    private string GetDesignationFrom(string shifr) {
        return shifr.Replace("ВАР", "~").Split('~')[0].TrimEnd(new Char[] {'-'});
    }

    /// <summary>
    /// Метод для проверки подключения данного объекта к соответствующей записи в справочнике 'Список номенклатуры FoxPro'
    /// </summary>
    /// <param name="nom">Объект справочника "Список номенклатуры FoxPro", для которого производится связывание</param>
    /// <param name="findedOrCreated">Найденный или созданный объект, который необходимо подключить к "Списку номерклатуры FoxPro"</param>
    /// <param name="messages">Список строк, представляющий собой лог ведения выгрузки</param>
    private void CheckAndLink(ReferenceObject nom, ReferenceObject findedOrCreated, List<string> messages) {
        // -- Начало верификации информации

        // Проверяем корректность объекта nom
        if (GetTypeOfReferenceFrom(nom) != TypeOfReference.СписокНоменклатуры)
            throw new Exception(
                    $"{nameof(CheckAndLink)}: объект, переданный в качестве аргумента '{nameof(nom)}' должен принадлежать справочнику 'Список номенклатуры FoxPro'" +
                    $"(передан: {GetTypeOfReferenceFrom(nom).ToString()})"
                    );
        
        // Проверяем корректность объекта findedOrCreated
        if (GetTypeOfReferenceFrom(findedOrCreated) != TypeOfReference.ЭСИ)
            throw new Exception(
                    $"{nameof(CheckAndLink)}: объект, переданный в качестве аргумента '{nameof(findedOrCreated)}' должен принадлежать справочнику 'ЭСИ' " +
                    $"(передан: {GetTypeOfReferenceFrom(findedOrCreated).ToString()})"
                    );

        // Проверяем, что данный объект уже не подключен
        ReferenceObject linkedObject = nom.GetObject(СписокНоменклатуры.Links["Связь на ЭСИ"]);
        if ((linkedObject != null) && (linkedObject.Guid == findedOrCreated.Guid)) {
            messages.Add($"{nameof(CheckAndLink)}: Подключение объекта к 'Списку номенклатуры FoxPro' не потребовалось");
            return;
        }
        // -- Конец верификации информации

        // -- Начало подключения объекта
        try {
            nom.BeginChanges();
            nom.SetLinkedObject(СписокНоменклатуры.Links["Связь на ЭСИ"], findedOrCreated);
            nom.EndChanges();
        }
        catch (Exception e) {
            throw new Exception($"{nameof(CheckAndLink)}: В процессе присоединения объекта '{findedOrCreated.ToString()}' к списку номенклатуры возникла ошибка:{Environment.NewLine}{e.ToString()}");
        }
        messages.Add($"{nameof(CheckAndLink)}: Произведено подключение объекта к 'Списку номенклатуры FoxPro'");
        // -- Конец подключения объекта
    }

    /// <summary>
    /// Метод для преобразования перечисления TypeOfObject в его строковое представление, понятное DOCs
    /// На вход принимает объект TypeOfObject, текстовую репрезентакию которого необходимо получить, на выходе - строку
    /// </summary>
    private string GetStringFrom(TypeOfObject type)
    {

        switch (type)
        {
            case TypeOfObject.СборочнаяЕдиница:
                return "Сборочная единица";
            case TypeOfObject.СтандартноеИзделие:
                return "Стандартное изделие";
            case TypeOfObject.ПрочееИзделие:
                return "Прочее изделие";
            case TypeOfObject.Изделие:
                return "Изделие";
            case TypeOfObject.Деталь:
                return "Деталь";
            case TypeOfObject.ЭлектронныйКомпонент:
                return "Электронный компонент";
            case TypeOfObject.Материал:
                return "Материал";
            case TypeOfObject.Другое:
                return "Другое";
            default:
                throw new Exception($"{nameof(GetStringFrom)}: Ошибка при определении типа объекта справочника ");

        }
    }

   
    /// <summary>
    /// Создаёт объект в справочнике <refName>, в ЭСИ содаётся на него номенклатурный объект. 
    /// </summary>
    /// <param name="nom">Объект из справочника "Список номенклатуры"</param>
    /// <param name="name">Наименование создаваемой ДСЕ</param>
    /// <param name="oboz">Обозначение создаваемой ДСЕ</param>
    /// <param name="refname">Исходный справочник, в котором нужно создавать объект</param>
    /// <returns>Созданный объект в справочнике ЭСИ</returns>
    private ReferenceObject CreateRefObject(ReferenceObject nom, string name, string oboz, string classObjectName, RefGuidData refname) //, Reference refName, Guid guidName, Guid guidShifr)
    {

        try
        {
            var type = nom[СписокНоменклатуры.Params["Тип номенклатуры"]].Value.ToString();
            //string nomTip = GetStringFrom(type);
            name = NormalizeNameObject(name);
            var refName = refname.Ref;
            var createdClassObject = refName.Classes.Find(classObjectName);
            ReferenceObject refereceObject = refName.CreateReferenceObject(createdClassObject);
            
            if (classObjectName.Equals("Материал"))
                refereceObject[refname.Params["Сводное наименование"]].Value = name;
            else
                refereceObject[refname.Params["Наименование"]].Value = name;

            if (classObjectName.Equals("Стандартное изделие") || classObjectName.Equals("Электронный компонент") || classObjectName.Equals("Материал"))
                refereceObject[refname.Params["Код ОКП"]].Value = oboz;
            else
                refereceObject[refname.Params["Обозначение"]].Value = oboz;
            refereceObject.EndChanges();

            //Desktop.CheckIn(refereceObject, "Объект создан", false);
            /*            NomenclatureReference nomReference;
                        nomReference = ЭСИ.Ref as NomenclatureReference;
                        ReferenceObject newNomenclature = nomReference.CreateNomenclatureObject(refereceObject);*/

            //Desktop.CheckIn(newNomenclature, "Объект в создан", false);
            var newNomenclature = ConnectRefObjectToESI(refereceObject,oboz);
            return newNomenclature;
        }
        catch (Exception e)
        {
            string arguments = $"nom: {nom.ToString()}\nname: {name}\noboz: {oboz}\nclassObjectName: {classObjectName}\nrefname: {refname.Ref.Name}";
            throw new Exception($"{nameof(CreateRefObject)}: ошибка при создании объекта:\n{e}\n\nАргументы функции:\n{arguments}");
        }
    }

    private String NormalizeNameObject(string name)
    {
        if (string.IsNullOrEmpty(name))
            return name;
        while (name.Contains("  "))
        {
            name = name.Replace("  ", " ");
        }
        var nameArr = name.Split(' ');        
        nameArr[0] = nameArr[0].ToLower();
        nameArr[0] = char.ToUpper(nameArr[0][0]) + nameArr[0].Substring(1);
        return String.Join(" ", nameArr);
    }


    /// <summary>
    /// Подключаем найденый объект в смежных справочниках к ЭСИ    
    /// </summary>
    private ReferenceObject ConnectRefObjectToESI(ReferenceObject refereceObject, string designation)
    {
        try
        {
            NomenclatureReference nomReference;
            nomReference = ЭСИ.Ref as NomenclatureReference;
            ReferenceObject newNomenclature = nomReference.CreateNomenclatureObject(refereceObject);
            
            //Desktop.CheckIn(newNomenclature, "Объект в создан", false);

            //string designation = newNomenclature[ЭСИ.Params["Обозначение"]].Value.ToString();
            List<ReferenceObject> findedObjectInSpisokNom = СписокНоменклатуры.Ref
                .Find(СписокНоменклатуры.Params["Обозначение"], designation);

            // связываем объект списка номенклатуры с объектом ЭСИ
            /*
            if (findedObjectInSpisokNom != null)
            {
                var firstFindObj = findedObjectInSpisokNom.First();
                firstFindObj.BeginChanges();
                firstFindObj.SetLinkedObject(СписокНоменклатуры.Links["Связь на ЭСИ"], newNomenclature);
                firstFindObj.EndChanges();
            }
            */
            return newNomenclature;
        }
        catch (Exception e)
        {
            throw new Exception($"{nameof(ConnectRefObjectToESI)}: ошибка при подключении {refereceObject} в ЭСИ:\n{e}");
        }
    }
     
    /// <summary>
    /// Метод для проверки соответствия типов объекта справочника 'Список номенклатуры FoxPro' и объектов остальных участвующих в выгрузке справочников
    /// При нахождении разницы в типах, или при нахождении неопределенных типов данный метод пробует в автоматическом режиме сделать предположение о правильном типа,
    /// и если ему это удается - меняет тип у одного из объектов.
    /// Если в автоматическом режиме изменить тип не удается, выбрасывается исключение с целью предупредить пользователя
    /// </summary>
    /// <param name="nomenclatureRecord">Объект справочника "Список номенклатуры FoxPro"</param>
    /// <param name="findedObject">Найденный объект справочника ЭСИ</param>
    /// <param name="messages">Список строк, представляющий собой лог ведения выгрузки</param>
    private void SyncronizeTypes(ReferenceObject nomenclatureRecord, ReferenceObject findedObject, List<string> messages) {
     
        // Производим верификацию входных данных
        // Проверка параметра nomenclatureRecord
        if (nomenclatureRecord.Reference.Name != СписокНоменклатуры.Ref.Name)
            throw new Exception($"{nameof(SyncronizeTypes)}: Параметр nomenclatureRecord поддерживает только объекты справочника {СписокНоменклатуры.Ref.Name}");

        // Проверка параметра findedObject
        List<TypeOfReference> supportedReferences = new List<TypeOfReference>() {
            TypeOfReference.ЭСИ,
            TypeOfReference.Документы,
            TypeOfReference.Материалы,
            TypeOfReference.ЭлектронныеКомпоненты,
                                        
        };

        if (!supportedReferences.Contains(GetTypeOfReferenceFrom(findedObject)))
            throw new Exception($"{nameof(SyncronizeTypes)}: неправильное использование метода SyncronizeTypes. Параметр findedObject не поддерживает объекты справочника {findedObject.Reference.Name}");

        // Определяем указанный тип и тип найденной записи
        TypeOfObject typeOfNom = GetTypeOfObjectFrom(nomenclatureRecord);
        TypeOfObject typeOfFinded = GetTypeOfObjectFrom(findedObject);

        // Если типы не соответствуют друг другу, пытаемся привести их в соответствие
        if (typeOfNom != typeOfFinded) {
         
            messages.Add($"{nameof(SyncronizeTypes)}: Обнаружено несоответствие типов: запись - {typeOfNom.ToString()}, найденный объект - {typeOfFinded.ToString()}");

            // 1 СЛУЧАЙ
            // Случай, когда в справочнике 'Номенклатура FoxPro' указан более общий тип, чем тип у привязанного/найденного объекта.
            // В этом случае необходимо поправить значение типа в поле записи номенклатурного объекта
            if ((typeOfNom == TypeOfObject.НеОпределено) || (typeOfNom == TypeOfObject.Другое)) {
             
                nomenclatureRecord.BeginChanges();
                nomenclatureRecord[СписокНоменклатуры.Params["Тип номенклатуры"]].Value = GetIntFrom(typeOfFinded);
                nomenclatureRecord.EndChanges();
                messages.Add($"{nameof(SyncronizeTypes)}: Произведена синхронизация типов. В поле 'Тип номенклатуры' вписан тип '{typeOfFinded.ToString()}'");
                return;
            }

            // 2 СЛУЧАЙ
            // Случай, когда в найденный объект принадлежит к более общему типу, нежели указанный тип в справочнике "Номенклатура FoxPro"
            if (typeOfFinded == TypeOfObject.Другое) {
             
                // Начальный тип принадлежит справочнику "Документы". В этом случае мы просто производим смену типа
                if ((typeOfNom != TypeOfObject.Материал) && (typeOfNom != TypeOfObject.ЭлектронныйКомпонент)) {
                 
                    // Производим смену типа
                    NomenclatureObject castObject = findedObject as NomenclatureObject;
                    // Если объект является объектом справочника "ЭСИ", то работаем с ним как с номенклатурным объектом
                    try {
                     
                        if (castObject != null) {
                         
                            if (!castObject.IsCheckedOut)
                                castObject.CheckOut();
                            findedObject = castObject.BeginChanges(GetClassFrom(typeOfNom, castObject.Reference));
                            findedObject.EndChanges();
                            if (findedObject == null)
                                throw new Exception($"{nameof(SyncronizeTypes)}: Возникла ошибка в процессе смены типа объекта");
                        }
                        else {
                         
                            if (!findedObject.IsCheckedOut)
                                findedObject.CheckOut();
                            findedObject = findedObject.BeginChanges(GetClassFrom(typeOfNom, findedObject.Reference));
                            findedObject.EndChanges();
                        }
                    }
                    catch (Exception e) {
                     
                        throw new Exception($"{nameof(SyncronizeTypes)}: Ошибка при смене типа:\n{e}");
                    }

                    // Пишем лог о результатах
                    messages.Add($"{nameof(SyncronizeTypes)}: Произведена синхронизация типов. Тип привязанного объекта с 'Другое' изменен на '{typeOfNom.ToString()}'");
                    return;
                }
                // Начальный тип принадлежит справочнику Материалы или Электронные компоненты.
                // Это значит, что вместо смены типа нужно перенести объект из одного справочника, где он был создан по ошибке,
                // в другой.
                else {
                 
                    findedObject = MoveObjectToAnotherReference(findedObject, typeOfNom);
                    messages.Add($"{nameof(SyncronizeTypes)}: Для смены типов требуется перенос объекта из одного справочника в другой");
                    //messages.Add($"Произведена синхронизация типов. Тип привязанного объекта с 'Другое' изменен на '{typeOfNom.ToString()}'");
                    return;
                }
            }

            // 3 СЛУЧАЙ
            // У данной разницы в типах нет однозначного варианта для автоматической синхронизации, поэтому для каждого возникающего случая
            // нужно задавать вопрос пользователю, чтобы он принял решение о том, какой тип в итоге должен остаться.

            // TODO: Реализовать код для опроса пользователя о том, какой тип нужно выбрать.
            // И так же реализовать тип, который будет производить соответствующие изменения

            throw new Exception(
                    $"{nameof(SyncronizeTypes)}: обнаружено несоответствие типов объекта {nomenclatureRecord.ToString()} и {findedObject.ToString()}" +
                    " которое не удалось устранить в автоматическом режиме" +
                    $" ({typeOfNom.ToString()} и {typeOfFinded.ToString()} соответственно)"
                    );
        }
        else
            messages.Add($"{nameof(SyncronizeTypes)}: Синхронизация типов не потребовалась");
            return;
    }

    /// <summary>
    /// Функция производит изменение типа в том случае, если тип назначения принадлежит справочнику, отличному от исходного
    ///
    /// Аргументы:
    /// initialObject - исходный объект, может принадлежать справочникам "ЭСИ", "Документы", "Материалы", "Электронные компоненты"
    /// targetType - тип, на который нужно изменить исходный объект
    ///
    /// Функция возвращает новый объект
    /// </summary>
    private ReferenceObject MoveObjectToAnotherReference(ReferenceObject initialObject, TypeOfObject targetType) {
        // Получаем информацию о типах справочников, которые будут использоваться в текущем вызове функции
        TypeOfReference initialReference = GetTypeOfReferenceFrom(initialObject);
        TypeOfReference targetReference = GetTypeOfReferenceFrom(targetType);

        // Получаем информацию о типе исходного объекта
        TypeOfObject initialType = GetTypeOfObjectFrom(initialObject);

        // -- Начало верификации --
        // Проверяем корректность переданного initialObject. Он должен принадлежать к ЭСИ, или исходным справочникам
        List<TypeOfReference> supportedReferences = new List<TypeOfReference>() {
            TypeOfReference.ЭСИ,
            TypeOfReference.Документы,
            TypeOfReference.Материалы,
            TypeOfReference.ЭлектронныеКомпоненты
        };

        if (!supportedReferences.Contains(initialReference))
            throw new Exception(
                    $"{nameof(MoveObjectToAnotherReference)}: ошибка при попытке изменения типа.\n" +
                    $"Метод {nameof(MoveObjectToAnotherReference)} работает только с объектами справочников, которые относятся к ЭСИ" +
                    $"В метод передан объект '{initialObject.ToString()}', который относится к справочнику '{initialObject.Reference.Name}'."
                    );

        // Проверяем, что начальные справочники исходного объекта и типа назначения - разные
        TypeOfReference initialObjectSourceReference = GetTypeOfReferenceFrom(GetTypeOfObjectFrom(initialObject));
        if (targetReference == initialObjectSourceReference) {
            throw new Exception(
                    "Ошибка при попытке изменения типа.\n" +
                    $"Метод {nameof(MoveObjectToAnotherReference)} работает только с объектами разных исходных справочников" +
                    $"В метод передан объект '{initialObject.ToString()}' с исходным справочником '{initialObjectSourceReference.ToString()}' для перемещения в справочник '{targetReference.ToString()}'."
                    );
        }
        // -- Окончание верификации --
        
        // -- Начало сбора исходной информации --
        // Получаем RefGuidData для исходного справочника и справочника назначения
        RefGuidData initialRefGuidData = GetRefGuidDataFrom(initialObject);
        RefGuidData targetRefGuidData = GetRefGuidDataFrom(targetReference);

        // Копируем основные параметры номенклатурного объекта
        object designation = initialObject[initialRefGuidData.Params["Обозначение"]].Value; // Обозначение
        object name = initialObject[initialRefGuidData.Params["Наименование"]].Value; // Наименование
        object okp = initialObject[initialRefGuidData.Params["Код ОКП"]].Value; // Код ОКП

        List<HLinkTransferData> parentLinks = null;
        List<HLinkTransferData> childLinks = null;
        // Если переданный объект - объект справочника ЭСИ, то так же производим копирование подключений
        if (initialReference == TypeOfReference.ЭСИ) {
            // Получаем информацию о всех родительских подключениях
            parentLinks = new List<HLinkTransferData>();
            foreach (ComplexHierarchyLink link in initialObject.Parents.GetHierarchyLinks()) {
                HLinkTransferData linkData = new HLinkTransferData();
                linkData.LinkedObject = link.ParentObject;

                foreach (Parameter param in link) {
                    if (!param.IsReadOnly)
                        linkData.Parameters.Add(param.ParameterInfo.Guid, param.Value);
                }
                parentLinks.Add(linkData);
            }

            // Получаем информацию о всех дочерних подключениях
            childLinks = new List<HLinkTransferData>();
            foreach (ComplexHierarchyLink link in initialObject.Children.GetHierarchyLinks()) {
                HLinkTransferData linkData = new HLinkTransferData();
                linkData.LinkedObject = link.ChildObject;

                foreach (Parameter param in link) {
                    if (!param.IsReadOnly)
                        linkData.Parameters.Add(param.ParameterInfo.Guid, param.Value);
                }
                childLinks.Add(linkData);
            }
        }

        // -- Окончание сбора исходной информации --

        // -- Начало пересоздания объекта
        // Создаем новый объект
        ReferenceObject newObject = targetRefGuidData.Ref.CreateReferenceObject(GetClassFrom(targetType, targetRefGuidData.Ref));

        newObject[targetRefGuidData.Params["Наименование"]].Value = name;
        
        // Производим проверку на то, не сменилось ли место расположения поля SHIFR у нового типа
        bool needSwap = IsShifrOKP(initialType) == IsShifrOKP(initialType);
        newObject[targetRefGuidData.Params["Обозначение"]].Value = needSwap ?  okp : designation;
        newObject[targetRefGuidData.Params["Код ОКП"]].Value = needSwap ? designation : okp;

        newObject.EndChanges();
        Desktop.CheckIn(newObject, "Создание объекта при смене типа", false);



        if (initialReference == TypeOfReference.ЭСИ) {
            // Производим создание нового номенклатурного объекта
            NomenclatureObject newNomObject = ((NomenclatureReference)initialRefGuidData.Ref).CreateNomenclatureObject(newObject);
            // Переносим на объект подключения со старого объекта
            // Завершаем перенос родительских подключений
            List<string> errors = new List<string>();
            foreach (HLinkTransferData linkData in parentLinks) {
                try {
                    ComplexHierarchyLink newLink = newNomObject.CreateParentLink(linkData.LinkedObject);
                    foreach (KeyValuePair<Guid, object> kvp in linkData.Parameters)
                        newLink[kvp.Key].Value = kvp.Value;
                    newLink.EndChanges();
                }
                catch (Exception e) {
                    errors.Add($"ошибка подключения к родительскому объекту ({e.Message})");
                }
            }
            // Завершаем перенос дочерних подключений
            foreach (HLinkTransferData linkData in childLinks) {
                try {
                    ComplexHierarchyLink newLink = newNomObject.CreateChildLink(linkData.LinkedObject);
                    foreach (KeyValuePair<Guid, object> kvp in linkData.Parameters)
                        newLink[kvp.Key].Value = kvp.Value;
                    newLink.EndChanges();
                }
                catch (Exception e) {
                    errors.Add($"ошибка подключения к дочернему объекту ({e.Message})");
                }
            }

            if (errors.Count != 0)
                Message("Предупреждение", $"В процессе изменения типа объекта {designation} возникли ошибки. Часть данных было утрачено в процессе:\n{string.Join("\n", errors)}");

            Desktop.CheckIn(newNomObject, "Создание номенклатурного объекта при смене типа", false);
            newObject = newNomObject as ReferenceObject;
        }

        if (newObject == null)
            throw new Exception($"{nameof(MoveObjectToAnotherReference)}: Переменная 'newObject' содержит null");

        // Производим удаление объекта в момент, когда новый объект полностью создан
        initialObject.CheckOut(true); // Берем объект на изменение с целью удаления
        Desktop.CheckIn(initialObject, "Удаление объекта при смене типа", false); // Применение удаления объекта
        Desktop.ClearRecycleBin(initialObject);

        return newObject;
    }


    /// <summary>
    /// Функция для создания связей между созданными в процессе выгрузки номенклатурными объектами
    /// </summary>
    /// <param name="createdObjects">Список созданных объектов</param>
    /// <param name="links">Перечень всех подключений в виде словаря по ключу с обознчением ДСЕ</param>
    private void ConnectCreatedObjects(List<ReferenceObject> createdObjects, Dictionary<string, List<ReferenceObject>> links) {
     
        // Функция принимает созданные номенклатурный объекты, а так же объекты справочника "Подключения"
        // 
        // Необходимо реализовать:
        // - Анализ полученных объектов (у них уже могут быть связи, так как там будут и найденные позиции)
        // - Создание/Корректировка/Удаление связей при необходимости
        // - Вывод лога о всех произведенных действиях
    }

    /// <summary>
    /// Метод для определения, относиться ли данный тип объекта к объекту, у которого SHIFR - Код ОКП
    /// </summary>
    /// <param name="type">Тип объекта, для которого нужно определить, является ли его обозначение кодом ОКП</param>
    /// <returns>Возаращает true, если в качестве обозначения для этого типа используется код ОКП, false - в противном случае</returns>
    private bool IsShifrOKP(TypeOfObject type) {
        switch(type) {
            case TypeOfObject.Материал:
            case TypeOfObject.ЭлектронныйКомпонент:
            case TypeOfObject.СтандартноеИзделие:
                return true;
            case TypeOfObject.Изделие:
            case TypeOfObject.Деталь:
            case TypeOfObject.СборочнаяЕдиница:
            case TypeOfObject.ПрочееИзделие:
            case TypeOfObject.Другое:
                return false;
            default:
                throw new Exception($"{nameof(IsShifrOKP)}: функция не предназначена для обработки типа {type.ToString()}");
        }

    }

    /// <summary>
    /// Метод для определения, относится ли данный тип объекта к объекту, у которого SHIFR - обозначение
    /// </summary>
    /// <param name="type">Тип объекта, для которого нужно определить, является ли его обозначение кодом ОКП</param>
    /// <returns>Возаращает false, если в качестве обозначения для этого типа используется код ОКП, true - в противном случае</returns>
    private bool IsShifrDesignation(TypeOfObject type) {
        return !IsShifrOKP(type);
    }

    /// <summary>
    /// Метод для определения типа переданного объекта
    /// </summary>
    /// <param name="nomenclature">Объект ReferenceObject, который должен принадлежать к одному из используемых для работы с номенклатурой справочников</param>
    /// <returns>Возвращает тип объекта или выбрасывает ошибку, если передан некорректный объект справочника</returns>
    private TypeOfObject GetTypeOfObjectFrom(ReferenceObject nomenclature) {
        TypeOfReference refType = GetTypeOfReferenceFrom(nomenclature);
                                                       

        // Разбираем случай, если в метод был передан объект справочника 'Список номенклатуры FoxPro'
        if (refType == TypeOfReference.СписокНоменклатуры) {
         
            int intType = (int)nomenclature[СписокНоменклатуры.Params["Тип номенклатуры"]].Value;
            return (TypeOfObject)intType;
        }

        // Разбираем случай, если в метод был передан объект справочника 'Электронная структура изделий'
        if ((refType == TypeOfReference.ЭСИ) || (refType == TypeOfReference.Документы) || (refType == TypeOfReference.ЭлектронныеКомпоненты) || (refType == TypeOfReference.Материалы)) {
         
            string typeName = nomenclature.Class.Name;
            switch (typeName) {
             
                case "Сборочная единица":
                    return TypeOfObject.СборочнаяЕдиница;
                case "Стандартное изделие":
                    return TypeOfObject.СтандартноеИзделие;
                case "Прочее изделие":
                    return TypeOfObject.ПрочееИзделие;
                case "Изделие":
                    return TypeOfObject.Изделие;
                case "Деталь":
                    return TypeOfObject.Деталь;
                case "Электронный компонент":
                    return TypeOfObject.ЭлектронныйКомпонент;
                case "Материал":
                    return TypeOfObject.Материал;
                case "Другое":
                    return TypeOfObject.Другое;
                default:
                    throw new Exception($"{nameof(GetTypeOfObjectFrom)}: ошибка при определении типа объекта справочника '{nomenclature.Reference.Name}'. Разбор типа объекта '{typeName}' не предусмотрен");
            }
        }

        throw new Exception($"{nameof(GetTypeOfObjectFrom)}: функция не работает с объектами справочника '{nomenclature.Reference.Name}'");
    }

    /// <summary>
    /// Метод для получения из заданного перечисления типа значение int, соответствующее в списке значений поля "Тип номенклатуры"
    /// </summary>
    /// <param name="type">Тип, для которого нужно получить его числовое представление</param>
    /// <returns>Числовое представление типа</returns>
    private int GetIntFrom(TypeOfObject type) {
     
        return (int)type;
    }

    /// <summary>
    /// Метод для получения из заданного перечисления типа объекта ClassObject, представляющего собой тип объекта в T-Flex DOCs
    /// </summary>
    /// <param name="type">Искомый тип объекта, заданный перечислением</param>
    /// <param name="reference">Справочник, для которого необходимо получить ClassObject для искомого типа</param>
    /// <returns>Возвращает тип объекта в виде ClassObject данного справочника. В случае, если объекта не получилось найти, выбрасывается исключение</returns>
    private ClassObject GetClassFrom(TypeOfObject type, Reference reference) {
     
        ClassObject result = null;
        switch (type) {
         
            case TypeOfObject.Изделие:
                result = reference.Classes.Find("Изделие");
                break;
            case TypeOfObject.СборочнаяЕдиница:
                result = reference.Classes.Find("Сборочная единица");
                break;
            case TypeOfObject.СтандартноеИзделие:
                result = reference.Classes.Find("Стандартное изделие");
                break;
            case TypeOfObject.ПрочееИзделие:
                result = reference.Classes.Find("Прочее изделие");
                break;
            case TypeOfObject.Деталь:
                result = reference.Classes.Find("Деталь");
                break;
            case TypeOfObject.ЭлектронныйКомпонент:
                result = reference.Classes.Find("Электронный компонент");
                break;
            case TypeOfObject.Материал:
                result = reference.Classes.Find("Материал");
                break;
            case TypeOfObject.Другое:
                result = reference.Classes.Find("Другое");
                break;
            default:
                throw new Exception(
                        $"{nameof(GetClassFrom)}: при определении соответствующего типа ClassObject для TypeOfObject возникла ошибка." +
                        $" '{type.ToString()}' не поддерживается"
                        );
        }

        if (result == null)
            throw new Exception(
                    $"{nameof(GetClassFrom)}: при определении соответствующего типа ClassObject для TypeOfObject возникла ошибка." +
                    $"Не удалось найти объект типа ClassObject в справочнике '{reference.Name}' для типа '{type.ToString()}'"
                    );

        return result;
    }

    /// <summary>
    /// Функция для определения типа справочника из переданного объета ReferenceObject
    /// </summary>
    /// <param name="referenceObject">Объект справочника, для которого нужно получить тип справочника, к которому он принадлежит</param>
    /// <returns>Тип справочника, представленный в виде перечисления</returns>
    private TypeOfReference GetTypeOfReferenceFrom(ReferenceObject referenceObject) {
     
        switch (referenceObject.Reference.Name.ToLower()) {
         
            case "документы":
                return TypeOfReference.Документы;
            case "электронные компоненты":
                return TypeOfReference.ЭлектронныеКомпоненты;
            case "материалы":
                return TypeOfReference.Материалы;
            case "список номенклатуры foxpro":
                return TypeOfReference.СписокНоменклатуры;
            case "подключения":
                return TypeOfReference.Подключения;
            default:
                if (referenceObject.Reference.Name.ToLower().Contains("электронная структура изделий"))
                    return TypeOfReference.ЭСИ;
                else
                    return TypeOfReference.Другое;
        }
    }

    /// <summary>
    /// Перегрузка функции для определения типа справочника на основе типа объекта.
    /// Стоит отметить, что данная функция работает только для номенклатурных объектов и
    /// по сути может вернуть только справочники 'Документы', 'Материалы', 'Электронные компоненты'
    /// </summary>
    /// <param name="type">Тип в виде перечисления</param>
    /// <returns>Тип справочника, к которому относится переданный тип</returns>
    private TypeOfReference GetTypeOfReferenceFrom(TypeOfObject type) {
        switch (type) {
            case TypeOfObject.Материал:
                return TypeOfReference.Материалы;
            case TypeOfObject.ЭлектронныйКомпонент:
                return TypeOfReference.ЭлектронныеКомпоненты;
            case TypeOfObject.Изделие:
            case TypeOfObject.СборочнаяЕдиница:
            case TypeOfObject.СтандартноеИзделие:
            case TypeOfObject.ПрочееИзделие:
            case TypeOfObject.Деталь:
            case TypeOfObject.Другое:
                return TypeOfReference.Документы;
            default:
                throw new Exception($"{nameof(GetTypeOfReferenceFrom)}: Функция не предназначена для работы с '{type.ToString()}'");
        }
    }

    /// <summary>
    /// Функция для получения соответствующего переданному typeOfReference RefGuidData объекта
    /// </summary>
    /// <param name="typeOfReference">Тип справочинка в виде перечисления</param>
    /// <returns>Соотвутствующий объект RefGuidData, который инкапсулирует в себе всю потребную информацию по справочнику</returns>
    private RefGuidData GetRefGuidDataFrom(TypeOfReference typeOfReference) {
     
        switch (typeOfReference) {
         
            case TypeOfReference.Документы:
                return Документы;
            case TypeOfReference.Материалы:
                return Материалы;
            case TypeOfReference.ЭлектронныеКомпоненты:
                return ЭлектронныеКомпоненты;
            case TypeOfReference.ЭСИ:
                return ЭСИ;
            case TypeOfReference.СписокНоменклатуры:
                return СписокНоменклатуры;
            case TypeOfReference.Подключения:
                return Подключения;
            default:
                throw new Exception($"{nameof(GetRefGuidDataFrom)}: для типа '{typeOfReference.ToString()}' не предусмотрена обработка");
        }
    }

    /// <summary>
    /// Перегрузка метода GetRefGuidDataFrom, реализованная для получения типа RefGuidData из типа текущего объекта.
    /// Особое внимание стоит обратить на то, что данный метод возвращает только исходные справочники типов, т.е.
    /// он не будет возвращать справочник ЭСИ, или справочники, которые не относятся к типам GetTypeOfObjectFrom
    /// (к примеру 'Подключения' или 'Список номенклатуры FoxPro')
    /// </summary>
    /// <param name="type">Тип номенклатурного объекта, для которого нужно вывести RefGuidData</param>
    /// <returns>Соотвутствующий объект RefGuidData, который инкапсулирует в себе всю потребную информацию по справочнику</returns>
    private RefGuidData GetRefGuidDataFrom(TypeOfObject type) {
        return GetRefGuidDataFrom(GetTypeOfReferenceFrom(type));
    }

    /// <summary>
    /// Перегрузка метода, которая вместо типа справочника принимает объект справочника, на основе которого вычисляется нужный RefGuidData объект
    /// </summary>
    /// <param name="referenceObject">Объект справочника, для которого нужно вывести RefGuidData</param>
    /// <returns>Соотвутствующий объект RefGuidData, который инкапсулирует в себе всю потребную информацию по справочнику</returns>
    private RefGuidData GetRefGuidDataFrom(ReferenceObject referenceObject) {
     
        return GetRefGuidDataFrom(GetTypeOfReferenceFrom(referenceObject));
    }

    // Интерфейсы
    public interface ITree {
     
        // Название изделия
        string NameProduct { get; } // Шифр корневой ноды дерева
        INode RootObject { get; } // Корневая нода дерева
        HashSet<ReferenceObject> AllReferenceObjects { get; set; } // Уникальные объекты дерева
        RefGuidData Nomenclature { get; }
        RefGuidData Connections { get; }

        // Обработка исключений
        bool HaveErrors { get; }
        string ErrString { get; }

        // Создание сообщение для лога с деревом изделия
        string GenerateLog();
        void AddError(string error);
    }

    public interface INode {
     
        string Name { get; }
        int Level { get; }
        ITree Tree { get; }
        INode Parent { get; }
        List<INode> Children { get; }
        ReferenceObject NomenclatureObject { get; }

        string GetTail(int quantityOfNodes);
    }

    // Перечисления

    // Нумерация задана в соответствии с списком значений поля "Тип номенклатуры".
    // При любом добавлении новых типов необходимо производить соответствующие изменения в данном перечислении
    private enum TypeOfObject {
     
        НеОпределено = 0,
        Изделие = 4,
        СборочнаяЕдиница = 1,
        СтандартноеИзделие = 2,
        ПрочееИзделие = 3,
        Деталь = 5,
        ЭлектронныйКомпонент = 6,
        Материал = 7,
        Другое = 8
    }

    // Перечисление для работы с справочниками
    private enum TypeOfReference {
     
        Документы,
        Материалы,
        ЭлектронныеКомпоненты,
        ЭСИ,
        СписокНоменклатуры,
        Подключения,
        Другое
    }

    

    // Классы
    private class NomenclatureTree : ITree {
     
        public string NameProduct { get; private set; }
        public INode RootObject { get; private set; }
        public HashSet<ReferenceObject> AllReferenceObjects { get; set; }
        private List<string> Errors { get; set; }
        public RefGuidData Nomenclature { get; set; }
        public RefGuidData Connections { get; set; }

        public bool HaveErrors => this.Errors.Count > 0;
        public string ErrString => this.HaveErrors ? $"Ошибки в дереве {this.NameProduct}:\n{string.Join("\n", this.Errors)}" : string.Empty;

        public NomenclatureTree(Dictionary<string, ReferenceObject> nomenclature, Dictionary<string, List<ReferenceObject>> links, string shifr, RefGuidData nom, RefGuidData conn) {
         
            if (nomenclature == null || links == null || shifr == null)
                throw new Exception($"{nameof(NomenclatureTree)}: в конструктор были переданы отсутствующие значения");

            this.Nomenclature = nom;
            this.Connections = conn;

            this.Errors = new List<string>();
            this.NameProduct = shifr;
            this.AllReferenceObjects = new HashSet<ReferenceObject>();
            this.RootObject = new NomenclatureNode(this, null, nomenclature, links, shifr);
        }

        public string GenerateLog() {
         
            return RootObject.ToString();
        }

        public void AddError(string error) {
         
            this.Errors.Add(error);
        }
    }

    private class NomenclatureNode : INode {
     
        public string Name { get; private set; }
        public int Level { get; private set; }
        public ITree Tree { get; private set; }
        public INode Parent { get; private set; }
        public List<INode> Children { get; private set; }
        public ReferenceObject NomenclatureObject { get; private set; }

        /// <summary>
        /// Конструктор новой ноды дерева состава.
        /// Аргуметны:
        /// tree - корневой объект
        /// parent - родительский объект
        /// nomenclature - словарь с объектами справочника "Список номенклатуры FoxPro", отсортированный по обозначениям
        /// links - словарь с объектами справочника "Подключения", отсортированный по обозначениям
        /// shifr - обозначение изделия
        /// </summary>
        public NomenclatureNode (ITree tree, INode parent, Dictionary<string, ReferenceObject> nomenclature, Dictionary<string, List<ReferenceObject>> links, string shifr) {
         

            // Валидация входной информации
            List<string> errors = new List<string>();

            if (tree == null)
                errors.Add("Параметр 'tree' имеет значение null");
            if (nomenclature == null)
                errors.Add("Параметр 'nomenclature' имеет значение null");
            if (links == null)
                errors.Add("Параметр 'links' имеет значение null");
            if (shifr == null)
                errors.Add("Параметр 'shifr' имеет значение null");
            if (errors.Count != 0)
                throw new Exception($"{nameof(NomenclatureNode)}: в процессе создания объекта возникла ошибка:\n{string.Join("\n", errors)}");

            this.Tree = tree;
            this.Parent = parent;

            this.Level = parent == null ? 0 : parent.Level + 1;

            if (nomenclature.ContainsKey(shifr)) {
             
                this.NomenclatureObject = nomenclature[shifr];
                if (this.NomenclatureObject == null)
                    throw new Exception(
                            $"{nameof(NomenclatureNode)}: при построении дерева для изделия {this.Tree.NameProduct} возникла ошибка:\n" +
                            "Переданный номенклатурный объект для формирования дочерней ноды содержит null.\n" +
                            $"Родительская нода - {parent.Name}"
                            );
                // Подключаем номенклатурный объект в список осех объектов дерева
                this.Tree.AllReferenceObjects.Add(this.NomenclatureObject);
            }
            else
                throw new Exception($"{nameof(NomenclatureNode)}: в дереве изделия {this.Tree.NameProduct} отсутствует '{shifr}'");

            // Получаем название объекта
            this.Name = (string)this.NomenclatureObject[Tree.Nomenclature.Params["Обозначение"]].Value;

            // Создаем список дочерний нод
            this.Children = new List<INode>();
            
            // Рекурсивно получаем потомков
            // Отключаем рекурсию при достижении большого уровня вложенности
            if (this.Level > 100) {
             
                throw new Exception($"{nameof(NomenclatureNode)}: превышена предельная глубина дерева в 100 уровней, возможна бесконечная рекурсия (Последние четыре элемента бесконечной ветки: {this.GetTail(4)})");
            }

            if (links.ContainsKey(shifr))
                foreach (string childShifr in links[shifr].Select(link => (string)link[Tree.Connections.Params["Комплектующая"]].Value)) {
                 
                    try {
                     
                        this.Children.Add(new NomenclatureNode(this.Tree, this, nomenclature, links, childShifr));
                    }
                    catch (Exception e) {
                     
                        this.Tree.AddError(e.Message);
                    }
                }
        }

        public string GetTail(int quantityOfNodes) {
         
            string result = this.Name;
            INode currentNode = this;
            while (true) {
             
                currentNode = currentNode.Parent;
                if (currentNode != null)
                    result = $"{currentNode.Name} -> {result}";
                else
                    break;
                quantityOfNodes -= 1;
                if (quantityOfNodes < 2)
                    break;
            }

            return result;
        }


        public override string ToString() {
         
            string prefix = string.Empty;
            if (this.Level == 1)
                prefix = "└";
            if (this.Level > 1)
                prefix = new string('│', this.Level - 1) + "└";
            return $"{prefix}{this.Name}\n{string.Join("", Children.Select(child => child.ToString()))}";
        }
    }

    public class RefGuidData {
     
        public Reference Ref { get; private set; }
        public Guid RefGuid { get; private set; }
        public Container Types { get; private set; }
        public Container Params { get; private set; }
        public Container Links { get; private set; }
        public Container Hlinks { get; private set; }
        public Container Objects { get; private set; }

        // Конструктор
        public RefGuidData (MacroContext context, Guid guidOfReference) {
         
            this.RefGuid = guidOfReference;
            ReferenceInfo refInfo = context.Connection.ReferenceCatalog.Find(guidOfReference);
            // Проверка на то, что удалось найти в базе справочник с таким Guid
            if (refInfo == null)
                throw new Exception($"{nameof(RefGuidData)}: ошибка поиска справочника с уникальным идентификатором {guidOfReference} при инициализации нового объекта ReferenceData");
            Ref = context.Connection.ReferenceCatalog.Find(guidOfReference).CreateReference();

            // Инициализируем контейнеры для уникальных идентификаторов
            this.Types = new Container(this, "Types");
            this.Params = new Container(this, "Params");
            this.Links = new Container(this, "Links");
            this.Hlinks = new Container(this, "Hlinks");
            this.Objects = new Container(this, "Objects");
        }

        // Методы для добавления Guid
        public void AddType(string name, Guid guid) =>
            this.Types.Add(name, guid);

        public void AddLink(string name, Guid guid) =>
            this.Links.Add(name, guid);

        public void AddHlink(string name, Guid guid) =>
            this.Hlinks.Add(name, guid);

        public void AddParam(string name, Guid guid) =>
            this.Params.Add(name, guid);

        public void AddObject(string name, Guid guid) =>
            this.Objects.Add(name, guid);

        public class Container {
         
            public string Name { get; private set; }
            private Dictionary<string, Guid> Storage { get; set; }
            private RefGuidData Parent { get; set; }

            public Container(RefGuidData parent, string name) {
             
                this.Parent = parent;
                this.Name = name;
                this.Storage = new Dictionary<string, Guid>();
            }

            public int Count => this.Storage.Count;

            public bool ContainsKey(string key) => this.Storage.ContainsKey(key);

            public void Add(string key, Guid guid) {
             
                string lowerKey = key.ToLower();
                if (this.Storage.ContainsKey(lowerKey))
                    throw new Exception($"{nameof(Container)}: хранилище '{this.Name}' объекта RefGuidData для справочника '{Parent.Ref.Name}' уже содержит ключ '{key}'");
                this.Storage.Add(lowerKey, guid);
            }

            public Guid this[string key] {
             
                get {
                 
                    string lowerKey = key.ToLower();
                    if (!this.Storage.ContainsKey(lowerKey))
                        throw new Exception($"{nameof(Container)}: хранилище {this.Name} объекта RefGuidData для справочника {Parent.Ref.Name} не содержит ключа '{key}'");
                    return this.Storage[lowerKey];
                }

                set {
                 
                    throw new Exception($"{nameof(Container)}: для добавления объекта используйте методы класса RefGuidData");
                }
            }
        }
    }

    public class HLinkTransferData {
        public ReferenceObject LinkedObject { get; set; } = null;
        public Dictionary<Guid, object> Parameters { get; set; } = new Dictionary<Guid, object>();
    }

    public class Log
    {
        public int Number { get; private set; }
        public string Shifr_izd { get; set; }
        public string file_name_tree { get; set; }
        public string file_name_error_log { get; set; }
        public DateTime timeStart { get; set; }
        public DateTime timeStop { get; set; }
        public int count_object { get; set; }
        public int create_object { get; set; }
        public int count_error { get; set; }
        public bool messageOff { get; set; } = false;

        public void getparam(RefGuidData refdata)
        {
            var logRefObj = refdata.Ref.Objects;
            int lognumberMax = logRefObj.Max(x => (Int32)x.GetObjectValue("Номер выгрузки").Value);
            this.Number = lognumberMax;
        }
    }
}
