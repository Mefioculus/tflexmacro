using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using TFlex.DOCs.Model.Macros;
using TFlex.DOCs.Model.References;
using TFlex.DOCs.Model.Macros.ObjectModel;
using TFlex.DOCs.Model;
using TFlex.DOCs.Model.Structure;
using TFlex.DOCs.Model.Search;
using FoxProShifrsNormalizer;
using TFlex.DOCs.Model.Desktop;
using TFlex.DOCs.Model.References.Nomenclature;

public class Macro : MacroProvider
{
    private RefGuidData СписокНоменклатуры { get; set; }
    private RefGuidData Подключения { get; set; }
    private RefGuidData ЭСИ { get; set; }
    private RefGuidData Документы { get; set; }
    private RefGuidData ЭлектронныеКомпоненты { get; set; }
    private RefGuidData Материалы { get; set; }
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
        ЭСИ.AddParam("Код ОКП", new Guid("b39cc740-93cc-476d-bfed-114fe9b0740c")); // string
        // Типы
        ЭСИ.AddType("Материальный объект", new Guid("0ba28451-fb4d-47d0-b8f6-af0967468959"));

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

        // Определяем позиции справочника "Список номенклатуры FoxPro", которые необходимо обрабатывать во время выгрузки
        HashSet<ReferenceObject> номенклатураДляСоздания = GetNomenclatureToProcess(номенклатура, подключения, изделияДляВыгрузки);

        // Производим поиск и (при необходимости) создание объектов в ЭСИ и смежных справочниках
        List<ReferenceObject> созданныеДСЕ = FindOrCreateNomenclatureObjects(номенклатураДляСоздания);

        // Производим соединение созданных ДСЕ в иерархию при помощи подключений
        ConnectCreatedObjects(созданныеДСЕ, подключения);

        Message("Информация", "Работа макроса завершена");
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
        диалог.ДобавитьСтроковое("Введите обозначение изделия", "УЯИС.731353.038\r\nУЯИС.731353.037", многострочное: true, количествоСтрок: 10);
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

            var list_oboz = oboz.Split(new string[] { "\r\n" }, StringSplitOptions.RemoveEmptyEntries);
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
            TypeOfObject nomType = DefineTypeOfObject(nom); // тип
            // Пишем информацию в лог
            messages.Add(new String('-', 30));
            messages.Add($"{nomDesignation}:");




            try {
                // СТАДИЯ 1: Пытаемся получить объект по связи на справочник ЭСИ
                resultObject = ProcessFirstStageFindOrCreate(nom, nomDesignation, nomType, messages);
                if (resultObject != null) {
                    result.Add(resultObject);
                    countStage1 += 1;
                    continue;
                }

                // СТАДИЯ 2: Пытаемся найти объект в справочнике ЭСИ
                resultObject = ProcessSecondStageFindOrCreate(nom, nomDesignation, nomType, messages);
                if (resultObject != null) {
                    result.Add(resultObject);
                    countStage2 += 1;
                    continue;
                }

                // СТАДИЯ 3: Пытаемся найти объект в смежных справочниках
                resultObject = ProcessThirdStageFindOrCreate(nom, nomDesignation, nomType, messages);
                if (resultObject != null) {
                    result.Add(resultObject);
                    countStage3 += 1;
                    continue;
                }

                // СТАДИЯ 4: Создаем объект исходя из того, какой был определен тип в справочнике "Список номенклатуры FoxPro"
                resultObject = ProcessFinalStageFindOrCreate(nom, nomDesignation, nomType, messages);
                if (resultObject != null) {
                    countStage4 += 1;
                    result.Add(resultObject);
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

        return result;
    }




    private ReferenceObject MoveShifrtoOKP(ReferenceObject resultObject, string designation, List<string> messages)
    {
        TypeOfObject esiType = DefineTypeOfObject(resultObject); // тип
        var okpEsi = resultObject[ЭСИ.Params["Код ОКП"]].Value.ToString();
        var obozEsi = resultObject[ЭСИ.Params["Обозначение"]].Value.ToString();

        if (esiType == TypeOfObject.СтандартноеИзделие || esiType == TypeOfObject.Материал || esiType == TypeOfObject.ЭлектронныйКомпонент)
        {
            if (!designation.Equals(okpEsi))
            {
                if (designation.Equals(obozEsi))
                {
                    try
                    {
                        resultObject.CheckOut();
                        resultObject.BeginChanges();
                        resultObject[ЭСИ.Params["Код ОКП"]].Value = obozEsi;
                        resultObject[ЭСИ.Params["Обозначение"]].Value = "";
                        resultObject.EndChanges();
                        //linkedObject.CheckIn("Перенос Обозначения в поле код ОКП");
                        Desktop.CheckIn(resultObject, "Перенос Обозначения в поле код ОКП", false);
                    }
                    catch (Exception e)
                    {
                        messages.Add($"Error: При переносе Обозначения в ОКП  {e.Message}");
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
    /// В качестве исходных данных принимает объект справочника "Список номенклатуры FoxPro", обозначение и тип объекта, а так же список строк messages, который пойдет в лог.
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
            SyncronizeTypes(nom, linkedObject);
            return MoveShifrtoOKP(linkedObject, designation, messages);
   
        }
        else
            messages.Add("Связанный объект отсутствовал");
        return null;
    }

    /// <summary>
    /// Вспомогательный код для функции FindOrCreateNomenclatureObjects.
    /// Данный код запускается в том случае, если не удалось получить объект по связи и производит поиск объекта в справочнике ЭСИ, и если таковой имеется, производит проверку его типа.
    /// В качестве исходных данных принимает объект справочника "Список номенклатуры FoxPro", обозначение и тип объекта, а так же список строк messages, который пойдет в лог.
    /// Функия может завершиться одним из трех вариантов: вернуть ReferenceObject, null и выбросить исключение.
    /// ReferenceObject возвращается если объект был получен и он полностью корректен.
    /// null возвращается, если объект не был получен и требуется произвети поиск при помощи следующих стадий.
    /// </summary>
    private ReferenceObject ProcessSecondStageFindOrCreate(ReferenceObject nom, string designation, TypeOfObject type, List<string> messages) {
        List<ReferenceObject> findedObjectInEsi = ЭСИ.Ref
            .Find(ЭСИ.Params["Обозначение"], designation) // Производим поиск по всему справочнику
            .Where(finded => finded.Class.IsInherit(ЭСИ.Types["Материальный объект"])) // Отфильтровываем только те объекты, которые наследуются от 'Материального объекта'
            .ToList<ReferenceObject>();

        if (findedObjectInEsi.Count == 1) {
            messages.Add("Объект найден в ЭСИ");
            SyncronizeTypes(nom, findedObjectInEsi[0]);
            return MoveShifrtoOKP(findedObjectInEsi[0], designation, messages);
        }
        else if (findedObjectInEsi.Count > 1) {
            messages.Add("Объект найден в ЭСИ");
            throw new Exception($"В ЭСИ найдено более одного совпадения по данному обозначению:\n{string.Join("\n", findedObjectInEsi.Select(obj => obj.ToString()))}");
        }
        else {
            messages.Add("Объект не найден в ЭСИ");
            return null;
        }
    }

    /// <summary>
    /// Вспомогательный код для функции FindOrCreateNomenclatureObjects.
    /// Данный код запускается в том случае, если не удалось найти объект в ЭСИ и производит поиск объекта в смежных справочниках, и если такой имеется, производит проверку типа и создание объекта ЭСИ.
    /// В качестве исходных данных принимает объект справочника "Список номенклатуры FoxPro", обозначение и тип объекта, а так же список строк messages, который пойдет в лог.
    /// Функия может завершиться одним из трех вариантов: вернуть ReferenceObject, null и выбросить исключение.
    /// ReferenceObject возвращается если объект был получен и он полностью корректен.
    /// null возвращается, если объект не был получен и требуется произвети поиск при помощи следующих стадий.
    /// </summary>
    private ReferenceObject ProcessThirdStageFindOrCreate(ReferenceObject nom, string designation, TypeOfObject type, List<string> messages) {
        // Производим поиск по смежным справочникам
        List<ReferenceObject> result = new List<ReferenceObject>();

        List<ReferenceObject> tempResult;

        // Производим поиск по справочнику "Документы"
        tempResult = Документы.Ref
            .Find(Документы.Params["Обозначение"], designation)
            .Where(finded => finded.Class.IsInherit(Документы.Links["Объект состава изделия"]))
            .ToList<ReferenceObject>();
        if (tempResult != null)
            result.AddRange(tempResult);

        // Производим поиск по справочнику "Электронные компоненты"
        tempResult = ЭлектронныеКомпоненты.Ref.Find(ЭлектронныеКомпоненты.Params["Обозначение"], designation);
        if (tempResult != null)
            result.AddRange(tempResult);

        // Производим поиск по справочнику "Материалы"
        tempResult = Материалы.Ref.Find(Материалы.Params["Обозначение"], designation);
        if (tempResult != null)
            result.AddRange(tempResult);

        switch (result.Count) {
            case 0:
                messages.Add("Объект не найден в смежных справочниках");
                return null;
            case 1:
                messages.Add("Объект найден в смежных справочниках");
                return MoveShifrtoOKP(result[0], designation, messages);
            default:
                messages.Add("Объект найден в смежных справочниках");
                throw new Exception($"Было найдено несколько совпадений:\n{string.Join("\n", result.Select(res => $"{res.ToString()} (Справочник: {res.Reference.Name})"))}");
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
        string nomTip = getTypeString(type);
        if (type == TypeOfObject.СтандартноеИзделие)
        {
            createDocument = CreateRefObject(nom, nomName, designation, nomTip, Документы.Ref,
                                                    Документы.Params["Наименование"], Документы.Params["Код ОКП"]);
        }
        else if (type == TypeOfObject.Материал)
        {
            createDocument = CreateRefObject(nom, nomName, designation, nomTip, Материалы.Ref,
                                Материалы.Params["Сводное наименование"], Материалы.Params["Обозначение"]);
        }
        else if (type == TypeOfObject.ЭлектронныйКомпонент)
        {
            createDocument = CreateRefObject(nom, nomName, designation, nomTip, ЭлектронныеКомпоненты.Ref,
                                            ЭлектронныеКомпоненты.Params["Наименование"], ЭлектронныеКомпоненты.Params["Код ОКП"]);
        }
        else
        {
            createDocument = CreateRefObject(nom, nomName, designation, nomTip, Документы.Ref,
                                                       Документы.Params["Наименование"], Документы.Params["Обозначение"]);
        }

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
    /// Метод для преобразования перечисления TypeOfObject в его строковое представление, понятное DOCs
    /// На вход принимает объект TypeOfObject, текстовую репрезентакию которого необходимо получить, на выходе - строку
    /// </summary>
    private string getTypeString(TypeOfObject type)
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
    /// </summary>
    private ReferenceObject CreateRefObject(ReferenceObject nom, string name, string oboz, string classObjectName, Reference refName, Guid guidName, Guid guidShifr)
    {
        try
        {
            var createdClassObject = refName.Classes.Find(classObjectName);
            ReferenceObject refereceObject = refName.CreateReferenceObject(createdClassObject);
            refereceObject[guidName].Value = name;
            refereceObject[guidShifr].Value = oboz;
            refereceObject.EndChanges();
            Desktop.CheckIn(refereceObject, "Объект создан", false);
            NomenclatureReference nomReference;
            nomReference = ЭСИ.Ref as NomenclatureReference;
            ReferenceObject newNomenclature = nomReference.CreateNomenclatureObject(refereceObject);
            Desktop.CheckIn(newNomenclature, "Объект в создан", false);
            nom.BeginChanges();
            nom.SetLinkedObject(СписокНоменклатуры.RefGuid, newNomenclature);
            nom.EndChanges();
            return refereceObject;
        }
        catch (Exception e)
        {
            throw new Exception($"Ошибка при создании объекта {e}");
        }

        return null;
    }


    /// <summary>
    /// Метод для проверки соответствия типов объекта справочника 'Список номенклатуры FoxPro' и объектов остальных участвующих в выгрузке справочников
    /// Входные параметры:
    /// nomenclatureRecord - объект справочника 'Список номенклатуры FoxPro'
    /// findedObject - объект справочников ЭСИ, Документы, Материалы, ЭлектронныеКомпоненты
    ///
    /// При нахождении разницы в типах, или при нахождении неопределенных типов данный метод пробует в автоматическом режиме сделать предположение о правильном типа,
    /// и если ему это удается - меняет тип у одного из объектов.
    /// Если в автоматическом режиме изменить тип не удается, выбрасывается исключение с целью предупредить пользователя
    /// </summary>
    private void SyncronizeTypes(ReferenceObject nomenclatureRecord, ReferenceObject findedObject) {
        // Производим верификацию входных данных
        // Проверка параметра nomenclatureRecord
        if (nomenclatureRecord.Reference.Name != СписокНоменклатуры.Ref.Name)
            throw new Exception($"Неправильное использование метода SyncronizeTypes. Параметр nomenclatureRecord поддерживает только объекты справочника {СписокНоменклатуры.Ref.Name}");

        // Проверка параметра findedObject
        List<string> supportedReferences = new List<string>() {
            Документы.Ref.Name,
            ЭлектронныеКомпоненты.Ref.Name,
            Материалы.Ref.Name,
        };

        if (!supportedReferences.Contains(findedObject.Reference.Name) && (!findedObject.Reference.Name.Contains("Электронная структура изделий")))
            throw new Exception($"Неправильное использование метода SyncronizeTypes. Параметр findedObject не поддерживает объекты справочника {findedObject.Reference.Name}");

        TypeOfObject typeOfNom = DefineTypeOfObject(nomenclatureRecord);
        TypeOfObject typeOfFinded = DefineTypeOfObject(findedObject);

        if (typeOfNom != typeOfFinded)
            // TODO: Реализовать код, который будет пытаться привести типы к соответствию, и в том случае, если ему это не будет удаваться. будет выдавать исключения
            throw new Exception(
                    $"Обнаружено несоответствие типов объекта {nomenclatureRecord.ToString()} и {findedObject.ToString()}" +
                    $" ({typeOfNom.ToString()} и {typeOfFinded.ToString()} соответственно)"
                    );
        else
            return;
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
    /// Метод для определения типа переданного объекта
    /// На вход передается ReferenceObject, который может относиться к справочникам:
    /// Список номенклатуры FoxPro, ЭСИ, Документы, Материалы, Электронные компоненты
    /// Если в петод передается объект другого справочника, выдается ошибка
    /// </summary>
    private TypeOfObject DefineTypeOfObject(ReferenceObject nomenclature) {
        string reference = nomenclature.Reference.Name;

        // Разбираем случай, если в метод был передан объект справочника 'Список номенклатуры FoxPro'
        if (reference == СписокНоменклатуры.Ref.Name) {
            int intType = (int)nomenclature[СписокНоменклатуры.Params["Тип номенклатуры"]].Value;
            switch (intType) {
                case 0: // Не определено
                    return TypeOfObject.НеОпределено;
                case 1: // Сборочная единица
                    return TypeOfObject.СборочнаяЕдиница;
                case 2: // Стандартное изделие
                    return TypeOfObject.СтандартноеИзделие;
                case 3: // Прочее изделие
                    return TypeOfObject.ПрочееИзделие;
                case 4: // Изделие
                    return TypeOfObject.Изделие;
                case 5: // Деталь
                    return TypeOfObject.Деталь;
                case 6: // Электронный компонент
                    return TypeOfObject.ЭлектронныйКомпонент;
                case 7: // Материал
                    return TypeOfObject.Материал;
                case 8: // Другое
                    return TypeOfObject.Другое;
                default:
                    throw new Exception($"Ошибка при определении типа объекта справочника '{reference}'. Номера {intType.ToString()} не предусмотрено при разборе");
            }
        }

        // Разбираем случай, если в метод был передан объект справочника 'Электронная структура изделий'
        if ((reference.Contains("Электронная структура изделий")) || (reference == Документы.Ref.Name) || (reference == ЭлектронныеКомпоненты.Ref.Name) || (reference == Материалы.Ref.Name)) {
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
                    throw new Exception($"Ошибка при определении типа объекта справочника '{reference}'. Разбрт типа объекта '{typeName}' не предусмотрен");
            }
        }

        throw new Exception($"Метод 'DefineTypeOfObject' не работает с объектами справочника '{reference}'");
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

    private enum TypeOfObject {
        НеОпределено,
        Изделие,
        СборочнаяЕдиница,
        СтандартноеИзделие,
        ПрочееИзделие,
        Деталь,
        ЭлектронныйКомпонент,
        Материал,
        Другое,
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

        public NomenclatureNode (ITree tree, INode parent, Dictionary<string, ReferenceObject> nomenclature, Dictionary<string, List<ReferenceObject>> links, string shifr) {
            this.Tree = tree;
            this.Parent = parent;

            this.Level = parent == null ? 0 : parent.Level + 1;

            if (nomenclature.ContainsKey(shifr)) {
                this.NomenclatureObject = nomenclature[shifr];
                // Подключаем номенклатурный объект в список осех объектов дерева
                this.Tree.AllReferenceObjects.Add(this.NomenclatureObject);
            }
            else {
                this.NomenclatureObject = null;
                this.Tree.AddError($"В дереве изделия {this.Tree.NameProduct} отсутствует '{shifr}'");
            }

            // Получаем название объекта
            this.Name = (string)this.NomenclatureObject[Tree.Nomenclature.Params["Обозначение"]].Value;

            // Создаем список дочерний нод
            this.Children = new List<INode>();
            
            // Рекурсивно получаем потомков
            // Отключаем рекурсию при достижении большого уровня вложенности
            if (this.Level > 100) {
                this.Tree.AddError($"Превышена предельная глубина дерева в 100 уровней, возможна бесконечная рекурсия (Последние четыре элемента бесконечной ветки: {this.GetTail(4)})");
                return;
            }

            if (links.ContainsKey(shifr))
                foreach (string childShifr in links[shifr].Select(link => (string)link[Tree.Connections.Params["Комплектующая"]].Value))
                    this.Children.Add(new NomenclatureNode(this.Tree, this, nomenclature, links, childShifr));
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
            this.Objects = new Container(this, "Objects");
        }

        // Методы для добавления Guid
        public void AddType(string name, Guid guid) =>
            this.Types.Add(name, guid);

        public void AddLink(string name, Guid guid) =>
            this.Links.Add(name, guid);

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
                    throw new Exception($"Хранилище {this.Name} объекта RefGuidData для справочника {Parent.Ref.Name} уже содержит ключ '{key}'");
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

}
