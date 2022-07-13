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
                                                                            
                                         

    private string ДиректорияДляЛогов { get; set; }

    // Для лога
    private string TimeStamp { get; set; }
    private string NameOfImport { get; set; }
    private string StringForLog => $"{NameOfImport} ({TimeStamp})";

    public Macro(MacroContext context)
        : base(context) {
     

        #if DEBUG
        System.Diagnostics.Debugger.Launch();
        System.Diagnostics.Debugger.Break();
        #endif

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

                  
    private void Test(string изделияДляВыгрузки)
    {
        var result = GetFilterRefObj(изделияДляВыгрузки, Материалы.RefGuid, Материалы.Params["Обозначение"]);
    }



    /// <summary>
    /// Метод для проведения тестирования
    /// </summary>
    public void TestGukov() {
        //ReferenceObject currentObject = Context.ReferenceObject;
        ReferenceObject currentObject = ЭСИ.Ref.Find(new Guid("ae6e8d15-cec5-4849-aee5-f5ffd8c80b56")); 
        MoveObjectToAnotherReference(currentObject, TypeOfObject.Материал);
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
    /// Метод ничего не принимает и не возвращает.
    /// Он производит присвоение строки определенного формата переменной NameOfImport
    /// </summary>
    private void SetNameOfExport(List<string> nomenclature) {
     
        if (nomenclature.Count == 0)
            return;

        string addition = nomenclature.Count > 1 ? $" (+{nomenclature.Count - 1})" : string.Empty;
        NameOfImport = $"{nomenclature[0]}{addition}";
    }

    /// <summary>
    /// Функция для определения перечня ДСЕ, которые необходимо обработать во время загрузки данных из FoxPro
    /// На вход принимает:
    /// nomenclature - словарь с всеми объектами справочника 'Список номеклатуры FoxPro' проиндексированный по обозначению
    /// links - словарь с всеми объектами справочника 'Подключения', проиндексированный и сгруппированный по родительской сборке
    /// shifs - список с обозначениями изделий, которые необходимо выгрузить
    /// На выход предоставляет HashSet объектов справочника 'Список номенклатуры FoxPro', выгрузку которых необходимо произвести
    /// </summary>
    private HashSet<ReferenceObject> GetNomenclatureToProcess(Dictionary<string, ReferenceObject> nomenclature, Dictionary<string, List<ReferenceObject>> links, List<string> shifrs) {
     
        // Для каждого шифра создаем объект, реализующий интерфейс ITree, получаем входящие объекты, добавляем их в HashSet (для исключения дубликатов)
        // В конце пишем лог, в котором записываем информацию о сгенерированном дереве, количестве входящих объектов и их структуре
        if (nomenclature == null || links == null || shifrs == null)
            throw new Exception("На вход функции GetNomenclatureToProcess были поданы отсутствующие значения");
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
             
                messages.Add($"Error: {e.Message}");
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
        if (errors != string.Empty) {
         
            statistics += "\n\nОтобразить ошибки?";
            if (Question(statistics))
                Message("Ошибки", $"В процессе поиска и создания номенклатурных объектов возникли следующие ошибки\n{errors}");
        }
        else
            Message("Статистика", statistics);
        
        // Пишем лог
        File.WriteAllText(pathToLogFile, string.Join("\n", messages));
        log_save_ref.file_name_error_log = pathToLogFile;
        return result;
    }



    /// <summary>
    /// Функция предназначена для переноса параметра обозначение в параметр код ОКП
    /// </summary>
    private ReferenceObject MoveShifrToOKP(ReferenceObject resultObject, string designation, List<string> messages)
    {
        RefGuidData referData = GetRefGuidDataFrom(resultObject);
        TypeOfObject esiType = GetTypeOfObjectFrom(resultObject); // тип
       
        var obozEsi = resultObject[referData.Params["Обозначение"]].Value.ToString();

        if (esiType == TypeOfObject.СтандартноеИзделие || esiType == TypeOfObject.Материал || esiType == TypeOfObject.ЭлектронныйКомпонент)
        {
            var okpEsi = resultObject[referData.Params["Код ОКП"]].Value.ToString();

            if (!designation.Equals(okpEsi))
            {
                if (designation.Equals(obozEsi))
                {
                    try
                    {
                        resultObject.CheckOut();
                        resultObject.BeginChanges();
                        resultObject[referData.Params["Код ОКП"]].Value = obozEsi;
                        resultObject[referData.Params["Обозначение"]].Value = "";
                        resultObject.EndChanges();
                        //linkedObject.CheckIn("Перенос Обозначения в поле код ОКП");
                        Desktop.CheckIn(resultObject, "Перенос Обозначения в поле код ОКП", false);
                    }
                    catch (Exception e)
                    {
                        throw new Exception($"Во время переноса 'Обозначения' в 'ОКП':\n{e}");
                    }
                }
                else
                    throw new Exception($"У привязанного объекта отличается код ОКП (указан: '{okpEsi}', должен быть: {designation})");

            }
            return resultObject;
        }
        else
        {
            if (!designation.Equals(obozEsi))
                throw new Exception($"У привязанного объекта отличается обозначенние (указано: '{obozEsi}', должно быть: {designation})");
            return resultObject;
        }

    }

    /// <summary>
    /// Вспомогательный код для функции FindOrCreateNomenclatureObjects.
    /// Данный код производит пробует получить объект по связи, а так же проверить корректность полученного объекта, если таковой имеется.
    /// Аргументы:
    /// nom - объект справочника "Список номенклатуры FoxPro"
    /// designation - обозначение объекта
    /// messages - список сообщений, который пойдет в общий лог
    /// Функия может завершиться одним из трех вариантов: вернуть ReferenceObject, null и выбросить исключение.
    /// ReferenceObject возвращается если объект был получен и он полностью корректен.
    /// null возвращается, если объект не был получен и требуется произвети поиск при помощи следующих стадий.
    /// Исключение выбрасывается если объект был найден, но при его проверке возникли проблемы (об этом нужно оповестить пользователя и пока что не включать объект в выгрузку).
    /// </summary>
    private ReferenceObject ProcessFirstStageFindOrCreate(ReferenceObject nom, string designation, TypeOfObject type, List<string> messages) {
     
        // Получаем объект по связи и, если он есть, производим его проверку
        ReferenceObject linkedObject = nom.GetObject(СписокНоменклатуры.Links["Связь на ЭСИ"]);

        if (linkedObject != null)
        {
            messages.Add("Объект был получен по связи");
            SyncronizeTypes(nom, linkedObject, messages);
            return MoveShifrToOKP(linkedObject, designation, messages);
   
        }
        else
            messages.Add("Связанный объект отсутствовал");
        return null;
    }

    /// <summary>
    /// Вспомогательный код для функции FindOrCreateNomenclatureObjects.
    /// Данный код запускается в том случае, если не удалось получить объект по связи и производит поиск объекта в справочнике ЭСИ, и если таковой имеется, производит проверку его типа.
    /// Аргументы:
    /// nom - объект справочника "Список номенклатуры FoxPro"
    /// designation - обозначение объекта
    /// messages - список сообщений, который пойдет в общий лог
    /// Функия может завершиться одним из трех вариантов: вернуть ReferenceObject, null и выбросить исключение.
    /// ReferenceObject возвращается если объект был получен и он полностью корректен.
    /// null возвращается, если объект не был получен и требуется произвети поиск при помощи следующих стадий.
    /// </summary>
    private ReferenceObject ProcessSecondStageFindOrCreate(ReferenceObject nom, string designation, TypeOfObject type, List<string> messages) {
     

        List<ReferenceObject> findedObjects = ЭСИ.Ref
            .Find(ЭСИ.Params["Обозначение"], designation) // Производим поиск по всему справочнику
            .Where(finded => finded.Class.IsInherit(ЭСИ.Types["Материальный объект"])) // Отфильтровываем только те объекты, которые наследуются от 'Материального объекта'
            .ToList<ReferenceObject>();

        List<ReferenceObject> findedObjectsOKP = ЭСИ.Ref
            .Find(ЭСИ.Params["Код ОКП"], designation) // Производим поиск по всему справочнику
            .Where(finded => finded.Class.IsInherit(ЭСИ.Types["Материальный объект"])) // Отфильтровываем только те объекты, которые наследуются от 'Материального объекта'
            .ToList<ReferenceObject>();

        // Проверяем совпадают ли по гуиду объекты findedObjects и findedObjectsOKP
        if (findedObjects.Count > 1 || findedObjectsOKP.Count > 1)
            throw new Exception($"В ЭСИ найдено более одного совпадения по данному обозначению:\n{string.Join("\n", findedObjects.Select(obj => obj.ToString()))}");
        else if ((findedObjects.Count !=0 && findedObjectsOKP.Count != 0)) {
         
            if (findedObjects[0].SystemFields.Guid != findedObjectsOKP[0].SystemFields.Guid)
                findedObjects.AddRange(findedObjectsOKP);
        }
        else if ((findedObjects.Count == 0 && findedObjectsOKP.Count != 0)) {
         
            findedObjects.AddRange(findedObjectsOKP);
        }

        if (findedObjects.Count == 1) {
         
            messages.Add("Объект найден в ЭСИ");
            SyncronizeTypes(nom, findedObjects[0], messages);
            ReferenceObject resultObject = MoveShifrToOKP(findedObjects[0], designation, messages);
            /*
            // Производим привязываение объекта к "Списку номенклатуры FoxPro"
            nom.BeginChanges();
            nom.SetLinkedObject(СписокНоменклатуры.Links["Связь на ЭСИ"], resultObject);
            nom.EndChanges();
            */
            // Возвращаем результат
            return resultObject;
        }
        else if (findedObjects.Count > 1) {
         
            messages.Add("Объект найден в ЭСИ");
            throw new Exception($"В ЭСИ найдено более одного совпадения по данному обозначению:\n{string.Join("\n", findedObjects.Select(obj => obj.ToString()))}");
        }
        else {
         
            messages.Add("Объект не найден в ЭСИ");
            return null;
        }
    }

    /// <summary>
    /// Вспомогательный код для функции FindOrCreateNomenclatureObjects.
    /// Данный код запускается в том случае, если не удалось найти объект в ЭСИ и производит поиск объекта в смежных справочниках, и если такой имеется, производит проверку типа и создание объекта ЭСИ.
    /// Аргументы:
    /// nom - объект справочника "Список номенклатуры FoxPro"
    /// designation - обозначение объекта
    /// type - перечисление с типом изделия, которое требуется создать
    /// messages - список сообщений, который пойдет в общий лог
    /// Функия может завершиться одним из трех вариантов: вернуть ReferenceObject, null и выбросить исключение.
    /// ReferenceObject возвращается если объект был получен и он полностью корректен.
    /// null возвращается, если объект не был получен и требуется произвети поиск при помощи следующих стадий.
    /// </summary>
    private ReferenceObject ProcessThirdStageFindOrCreate(ReferenceObject nom, string designation, TypeOfObject type, List<string> messages) {
     
        // Производим поиск по смежным справочникам
        List<ReferenceObject> findedObjects = new List<ReferenceObject>();

        List<ReferenceObject> tempResult;

        // Производим поиск по справочнику "Документы"

        tempResult = Документы.Ref
            .Find(Документы.Params["Обозначение"], designation)
            .Where(finded => finded.Class.IsInherit(Документы.Types["Объект состава изделия"]))
            .ToList<ReferenceObject>();
        if (type == TypeOfObject.СтандартноеИзделие)
        {
            var okptmpResult = Документы.Ref
            .Find(Документы.Params["Код ОКП"], designation)
            .Where(finded => finded.Class.IsInherit(Документы.Types["Объект состава изделия"]))
            .ToList<ReferenceObject>();
            if (tempResult.Count != 0 && tempResult != null && okptmpResult.Count != 0)
            {
                throw new Exception($"В справочнике документы найдены 2 объктка {designation}");
            }
            else
            {
                tempResult = okptmpResult;
            }
        }

        if (tempResult != null && tempResult.Count!=0)
            findedObjects.AddRange(tempResult);

        // Производим поиск по справочнику "Электронные компоненты"
        tempResult = ЭлектронныеКомпоненты.Ref.Find(ЭлектронныеКомпоненты.Params["Обозначение"], designation);
        if (tempResult != null && tempResult.Count != 0)
            findedObjects.AddRange(tempResult);

        // Производим поиск по справочнику "Материалы"
        tempResult = Материалы.Ref.Find(Материалы.Params["Обозначение"], designation);
        if (tempResult != null && tempResult.Count != 0)
            findedObjects.AddRange(tempResult);

        switch (findedObjects.Count) {
         
            case 0:
                messages.Add("Объект не найден в смежных справочниках");
                return null;
            case 1:
                // Производим синхронизацию типов
                messages.Add("Объект найден в смежных справочниках");
                SyncronizeTypes(nom, findedObjects[0], messages);
                                                                                     
                ConnectRefObjectToESI(findedObjects[0], designation);
                return MoveShifrToOKP(findedObjects[0], designation, messages);
            default:
                messages.Add("Объект найден в смежных справочниках");
                throw new Exception($"Было найдено несколько совпадений:\n{string.Join("\n", findedObjects.Select(res => $"{res.ToString()} (Справочник: {res.Reference.Name})"))}");
        }
    }


    /// <summary>
    /// Вспомогательный код для функции FindOrCreateNomenclatureObjects.
    /// Данный код запускается только в том случае, если объет найти не получилось и производит создание объекта сначала в смежном справчонике, а затем в ЭСИ.
    /// В качестве исходных данных принимает объект справочника "Список номенклатуры FoxPro", обозначение и тип объекта, а так же список строк messages, который пойдет в лог.
    /// Функия может завершиться одним из трех вариантов: вернуть ReferenceObject, null и выбросить исключение.
    /// ReferenceObject возвращается если объект был получен и он полностью корректен.
    /// null возвращается, если объект не был получен и требуется произвети поиск при помощи следующих стадий.
    /// </summary>
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
         
            throw new Exception($"Ошибка при создании объекта. Невозможно создать объект типа '{type.ToString()}'");
        }

         
        createDocument = CreateRefObject(nom, nomName, designation, nomTip, GetRefGuidDataFrom(type));
         
                                                       
         
                                                                                                    
         
                                                                               
         
                                                                                                                            
         
            
         
                                                                                                    
         

        // var createDocument2 = CreateRefObject(nomName, designation, type.ToString(), ЭлектронныеКомпоненты.Ref,ЭлектронныеКомпоненты.Params["Наименование"], ЭлектронныеКомпоненты.Params["Код ОКП"];
        // var createDocument = CreateRefObject(nomName, designation, type.ToString(),Документы.Ref,Документы.Params["Наименование"],Документы.Params["Обозначение"]);

        if (createDocument != null)
        {
            messages.Add($"В справочнике {createDocument.Reference.Name.ToString()} создан объект {designation}");
            return createDocument;
        }
        return null;
    }

    /// <summary>
    /// Метод для проверки подключения данного объекта к соответствующей записи в справочнике 'Список номенклатуры FoxPro'
    /// Аргументы:
    /// nom - объект справочника 'Список номенклатуры FoxPro'
    /// findedOrCreated - объект справочника 'ЭСИ' который необходимо подключить к 'Списку номенклатуры'
    /// </summary>
    private void CheckAndLink(ReferenceObject nom, ReferenceObject findedOrCreated, List<string> messages) {
        // -- Начало верификации информации

        // Проверяем корректность объекта nom
        if (GetTypeOfReferenceFrom(nom) != TypeOfReference.СписокНоменклатуры)
            throw new Exception(
                    $"Ошибка в процессе работы метода '{nameof(CheckAndLink)}':\n" +
                    $"объект, переданный в качестве аргумента '{nameof(nom)}' должен принадлежать справочнику 'Список номенклатуры FoxPro'" +
                    $"(передан: {GetTypeOfReferenceFrom(nom).ToString()})"
                    );
        
        // Проверяем корректность объекта findedOrCreated
        if (GetTypeOfReferenceFrom(findedOrCreated) != TypeOfReference.ЭСИ)
            throw new Exception(
                    $"Ошибка в процессе работы метода '{nameof(CheckAndLink)}':\n" +
                    $"объект, переданный в качестве аргумента '{nameof(findedOrCreated)}' должен принадлежать справочнику 'ЭСИ' " +
                    $"(передан: {GetTypeOfReferenceFrom(findedOrCreated).ToString()})"
                    );

        // Проверяем, что данный объект уже не подключен
        ReferenceObject linkedObject = nom.GetObject(СписокНоменклатуры.Links["Связь на ЭСИ"]);
        if ((linkedObject != null) && (linkedObject.Guid == findedOrCreated.Guid)) {
            messages.Add("Подключение объекта к 'Списку номенклатуры FoxPro' не потребовалось");
            return;
        }
        // -- Конец верификации информации

        // -- Начало подключения объекта
        findedOrCreated.BeginChanges();
        findedOrCreated.SetLinkedObject(СписокНоменклатуры.Links["Связь на ЭСИ"], nom);
        findedOrCreated.EndChanges();
        messages.Add("Произведено подключение объекта к 'Списку номенклатуры FoxPro'");
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
                throw new Exception($"Ошибка при определении типа объекта справочника ");

        }
    }

   
    /// <summary>
    /// Создаёт объект в справочнике <refName>, в ЭСИ содаётся на него номенклатурный объект. 
    /// Добавляет связь созданого объекта в ЭСИ на объект из справочника Список номенклатуры FoxPro
    /// nom, Объект из справочника "Список номенклатуры"
    /// name, "Наименование"
    /// oboz, "Обозначение"
    /// refname "Справочник в котором создается объект"
    /// </summary>
    private ReferenceObject CreateRefObject(ReferenceObject nom, string name, string oboz, string classObjectName, RefGuidData refname) //, Reference refName, Guid guidName, Guid guidShifr)
    {
        try
        {
            var type = nom[СписокНоменклатуры.Params["Тип номенклатуры"]].Value.ToString();
            //string nomTip = GetStringFrom(type);
            var refName = refname.Ref;
            var createdClassObject = refName.Classes.Find(classObjectName);
            ReferenceObject refereceObject = refName.CreateReferenceObject(createdClassObject);
            
            if (classObjectName.Equals("Материал"))
                refereceObject[refname.Params["Сводное обозначение"]].Value = name;
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
            return refereceObject;
        }
        catch (Exception e)
        {
            throw new Exception($"Ошибка при создании объекта:\n{e}");
        }
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
            throw new Exception($"Ошибка при подключении {refereceObject} в ЭСИ:\n{e}");
        }
    }
     
    /// <summary>
    /// Метод для проверки соответствия типов объекта справочника 'Список номенклатуры FoxPro' и объектов остальных участвующих в выгрузке справочников
    ///
    /// Аргументы:
    /// nomenclatureRecord - объект справочника 'Список номенклатуры FoxPro'
    /// findedObject - объект справочников ЭСИ, Документы, Материалы, ЭлектронныеКомпоненты
    ///
    /// При нахождении разницы в типах, или при нахождении неопределенных типов данный метод пробует в автоматическом режиме сделать предположение о правильном типа,
    /// и если ему это удается - меняет тип у одного из объектов.
    /// Если в автоматическом режиме изменить тип не удается, выбрасывается исключение с целью предупредить пользователя
    /// </summary>
    private void SyncronizeTypes(ReferenceObject nomenclatureRecord, ReferenceObject findedObject, List<string> messages) {
     
        // Производим верификацию входных данных
        // Проверка параметра nomenclatureRecord
        if (nomenclatureRecord.Reference.Name != СписокНоменклатуры.Ref.Name)
            throw new Exception($"Неправильное использование метода SyncronizeTypes. Параметр nomenclatureRecord поддерживает только объекты справочника {СписокНоменклатуры.Ref.Name}");

        // Проверка параметра findedObject
        List<TypeOfReference> supportedReferences = new List<TypeOfReference>() {
            TypeOfReference.ЭСИ,
            TypeOfReference.Документы,
            TypeOfReference.Материалы,
            TypeOfReference.ЭлектронныеКомпоненты,
                                        
        };

        if (!supportedReferences.Contains(GetTypeOfReferenceFrom(findedObject)))
            throw new Exception($"Неправильное использование метода SyncronizeTypes. Параметр findedObject не поддерживает объекты справочника {findedObject.Reference.Name}");

        // Определяем указанный тип и тип найденной записи
        TypeOfObject typeOfNom = GetTypeOfObjectFrom(nomenclatureRecord);
        TypeOfObject typeOfFinded = GetTypeOfObjectFrom(findedObject);

        // Если типы не соответствуют друг другу, пытаемся привести их в соответствие
        if (typeOfNom != typeOfFinded) {
         
            messages.Add($"Обнаружено несоответствие типов: запись - {typeOfNom.ToString()}, найденный объект - {typeOfFinded.ToString()}");

            // 1 СЛУЧАЙ
            // Случай, когда в справочнике 'Номенклатура FoxPro' указан более общий тип, чем тип у привязанного/найденного объекта.
            // В этом случае необходимо поправить значение типа в поле записи номенклатурного объекта
            if ((typeOfNom == TypeOfObject.НеОпределено) || (typeOfNom == TypeOfObject.Другое)) {
             
                nomenclatureRecord.BeginChanges();
                nomenclatureRecord[СписокНоменклатуры.Params["Тип номенклатуры"]].Value = GetIntFrom(typeOfFinded);
                nomenclatureRecord.EndChanges();
                messages.Add($"Произведена синхронизация типов. В поле 'Тип номенклатуры' вписан тип '{typeOfFinded.ToString()}'");
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
                         
                            castObject.CheckOut();
                            findedObject = castObject.BeginChanges(GetClassFrom(typeOfNom, castObject.Reference));
                            findedObject.EndChanges();
                            if (findedObject == null)
                                throw new Exception("Возникла ошибка в процессе смены типа объекта");
                        }
                        else {
                         
                            findedObject.CheckOut();
                            findedObject = findedObject.BeginChanges(GetClassFrom(typeOfNom, findedObject.Reference));
                            findedObject.EndChanges();
                        }
                    }
                    catch (Exception e) {
                     
                        throw new Exception($"Ошибка при смене типа:\n{e}");
                    }

                    // Пишем лог о результатах
                    messages.Add($"Произведена синхронизация типов. Тип привязанного объекта с 'Другое' изменен на '{typeOfNom.ToString()}'");
                    return;
                }
                // Начальный тип принадлежит справочнику Материалы или Электронные компоненты.
                // Это значит, что вместо смены типа нужно перенести объект из одного справочника, где он был создан по ошибке,
                // в другой.
                else {
                 
                    // TODO: Реализовать код по смене типа между справочниками
                    findedObject = MoveObjectToAnotherReference(findedObject, typeOfNom);
                    messages.Add($"Для смены типов требуется перенос объекта из одного справочника в другой");
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
                    $"Обнаружено несоответствие типов объекта {nomenclatureRecord.ToString()} и {findedObject.ToString()}" +
                    " которое не удалось устранить в автоматическом режиме" +
                    $" ({typeOfNom.ToString()} и {typeOfFinded.ToString()} соответственно)"
                    );
        }
        else
            messages.Add("Синхронизация типов не потребовалась");
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
                    "Ошибка при попытке изменения типа.\n" +
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
            throw new Exception($"При выполнении метода {nameof(MoveObjectToAnotherReference)} возникла ошибка. Переменная 'newObject' содержит null");

        // Производим удаление объекта в момент, когда новый объект полностью создан
        initialObject.CheckOut(true); // Берем объект на изменение с целью удаления
        Desktop.CheckIn(initialObject, "Удаление объекта при смене типа", false); // Применение удаления объекта
        Desktop.ClearRecycleBin(initialObject);

        return newObject;
    }


    /// <summary>
    /// Функция для создания связей между созданными в процессе выгрузки номенклатурными объектами
    /// Входные данные:
    /// createdObjects - список созданных объектов
    /// links - словарь с всеми подключениями в виде словаря по ключу с обозначением ДСЕ
    /// </summary>
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
    /// Аргументы:
    /// type - тип проверяемого объекта
    /// </summary>
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
                throw new Exception($"Функция '{nameof(IsShifrOKP)}' не предназначена для обработки типа {type.ToString()}");
        }

    }

    /// <summary>
    /// Метод для определения, относится ли данный тип объекта к объекту, у которого SHIFR - обозначение
    /// Аргументы:
    /// type - тип проверяемого объекта
    /// </summary>
    private bool IsShifrDesignation(TypeOfObject type) {
        return !IsShifrOKP(type);
    }

    /// <summary>
    /// Метод для определения типа переданного объекта
    /// На вход передается ReferenceObject, который может относиться к справочникам:
    /// Список номенклатуры FoxPro, ЭСИ, Документы, Материалы, Электронные компоненты
    /// Если в метод передается объект другого справочника, выдается ошибка
    /// </summary>
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
                    throw new Exception($"Ошибка при определении типа объекта справочника '{nomenclature.Reference.Name}'. Разбор типа объекта '{typeName}' не предусмотрен");
            }
        }

        throw new Exception($"Метод '{nameof(GetTypeOfObjectFrom)}' не работает с объектами справочника '{nomenclature.Reference.Name}'");
    }

    /// <summary>
    /// Метод для получения из заданного перечисления типа значение int, соответствующее в списке значений поля "Тип номенклатуры"
    /// Аргументы:
    /// type - перечисление TypeOfObject, которое необходимо конвертировать в соответствующее число.
    /// На выход поступает его цифровое представление.
    /// </summary>
    private int GetIntFrom(TypeOfObject type) {
     
        return (int)type;
    }

    /// <summary>
    /// Метод для получения из заданного перечисления типа объекта ClassObject, представляющего собой тип объекта в T-Flex DOCs
    /// Аргументы:
    /// type - перечисление TypeOfObject, которое необходимо конвертировать в соответствующий ClassObject.
    /// reference - справочник, в котором нужно производить поиск соответствующего типа
    /// На выход поступает найденный объект ClassObject. Если объект не удалось найти, выбрасывается сообщение об ошибке.
    /// </summary>
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
                        "При определении соответствующего типа ClassObject для TypeOfObject возникла ошибка." +
                        $" '{type.ToString()}' не поддерживается"
                        );
        }

        if (result == null)
            throw new Exception(
                    "При определении соответствующего типа ClassObject для TypeOfObject возникла ошибка." +
                    $"Не удалось найти объект типа ClassObject в справочнике '{reference.Name}' для типа '{type.ToString()}'"
                    );

        return result;
    }

    /// <summary>
    /// Функция для определения типа справочника из переданного объета ReferenceObject
    /// </summary>
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
    ///
    /// Аргументы:
    /// type - тип объекта, из которого нужно определить тип справочника
    /// </summary>
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
                throw new Exception($"Метод {nameof(GetTypeOfReferenceFrom)} не предназначен для работы с '{type.ToString()}'");
        }
    }

    /// <summary>
    /// Функция для получения соответствующего переданному typeOfReference RefGuidData объекта
    ///
    /// Аргументы:
    /// typeOfReference - перечисление, в котором есть все основные справочники, которые используются данным макросом
    ///
    /// Функция возвращает RefGuidData
    /// </summary>
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
                throw new Exception($"Для типа '{typeOfReference.ToString()}' не предусмотрена обработка в методе GetRefGuidDataFrom");
        }
    }

    /// <summary>
    /// Перегрузка метода GetRefGuidDataFrom, реализованная для получения типа RefGuidData из типа текущего объекта.
    /// Особое внимание стоит обратить на то, что данный метод возвращает только исходные справочники типов, т.е.
    /// он не будет возвращать справочник ЭСИ, или справочники, которые не относятся к типам GetTypeOfObjectFrom
    /// (к примеру 'Подключения' или 'Список номенклатуры FoxPro')
    ///
    /// Аргументы:
    /// type - тип объекта, из которого нужно вывести объект RefGuidData
    /// </summary>
    private RefGuidData GetRefGuidDataFrom(TypeOfObject type) {
        return GetRefGuidDataFrom(GetTypeOfReferenceFrom(type));
    }

    /// <summary>
    /// Перегрузка метода, которая вместо типа справочника принимает объект справочника, на основе которого вычисляется нужный RefGuidData объект
    /// </summary>
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
                throw new Exception("В конструктор создания NomenclatureTree были поданы отсутствующие значения");

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
                throw new Exception($"Во время создания объекта NomenclatureNode возникла ошибка:\n{string.Join("\n", errors)}");

            this.Tree = tree;
            this.Parent = parent;

            this.Level = parent == null ? 0 : parent.Level + 1;

            if (nomenclature.ContainsKey(shifr)) {
             
                this.NomenclatureObject = nomenclature[shifr];
                if (this.NomenclatureObject == null)
                    throw new Exception(
                            $"При построении дерева для изделия {this.Tree.NameProduct} возникла ошибка:\n" +
                            "Переданный номенклатурный объект для формирования дочерней ноды содержит null.\n" +
                            $"Родительская нода - {parent.Name}"
                            );
                // Подключаем номенклатурный объект в список осех объектов дерева
                this.Tree.AllReferenceObjects.Add(this.NomenclatureObject);
            }
            else
                throw new Exception($"В дереве изделия {this.Tree.NameProduct} отсутствует '{shifr}'");

            // Получаем название объекта
            this.Name = (string)this.NomenclatureObject[Tree.Nomenclature.Params["Обозначение"]].Value;

            // Создаем список дочерний нод
            this.Children = new List<INode>();
            
            // Рекурсивно получаем потомков
            // Отключаем рекурсию при достижении большого уровня вложенности
            if (this.Level > 100) {
             
                throw new Exception($"Превышена предельная глубина дерева в 100 уровней, возможна бесконечная рекурсия (Последние четыре элемента бесконечной ветки: {this.GetTail(4)})");
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
                throw new Exception($"Ошибка поиска справочника с уникальным идентификатором {guidOfReference} при инициализации нового объекта ReferenceData");
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
                    throw new Exception($"Хранилище '{this.Name}' объекта RefGuidData для справочника '{Parent.Ref.Name}' уже содержит ключ '{key}'");
                this.Storage.Add(lowerKey, guid);
            }

            public Guid this[string key] {
             
                get {
                 
                    string lowerKey = key.ToLower();
                    if (!this.Storage.ContainsKey(lowerKey))
                        throw new Exception($"Хранилище {this.Name} объекта RefGuidData для справочника {Parent.Ref.Name} не содержит ключа '{key}'");
                    return this.Storage[lowerKey];
                }

                set {
                 
                    throw new Exception("Для добавления объекта используйте методы класса RefGuidData");
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
        


        public void getparam(RefGuidData refdata)
        {
            var logRefObj = refdata.Ref.Objects;
            int lognumberMax = logRefObj.Max(x => (Int32)x.GetObjectValue("Номер выгрузки").Value);
            this.Number = lognumberMax;
        }
    }
}
