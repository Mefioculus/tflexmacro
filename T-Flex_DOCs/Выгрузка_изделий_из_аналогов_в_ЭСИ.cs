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

public class Macro : MacroProvider
{
    private Reference СписокНоменклатурыСправочник { get; set; }
    private Reference ПодключенияСправочник { get; set; }
    private Reference ЭсиСправочник { get; set; }
    private Reference ДокументыСправочник { get; set; }
    private Reference ЭлектронныеКомпонентыСправочник { get; set; }
    private Reference МатериалыСправочник { get; set; }
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

        // Получаем экземпляры справочников для работы
        СписокНоменклатурыСправочник = Context.Connection.ReferenceCatalog.Find(Guids.References.СписокНоменклатурыFoxPro).CreateReference();
        ПодключенияСправочник = Context.Connection.ReferenceCatalog.Find(Guids.References.Подключения).CreateReference();
        ЭсиСправочник = Context.Connection.ReferenceCatalog.Find(Guids.References.ЭСИ).CreateReference();
        ЭлектронныеКомпонентыСправочник = Context.Connection.ReferenceCatalog.Find(Guids.References.ЭлектронныеКомпоненты).CreateReference();
        МатериалыСправочник = Context.Connection.ReferenceCatalog.Find(Guids.References.Материалы).CreateReference();
        ДокументыСправочник = Context.Connection.ReferenceCatalog.Find(Guids.References.Документы).CreateReference();
        
        // Создаем директорию для ведения логов
        ДиректорияДляЛогов = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "Логи выгрузки из аналогов в ЭСИ");

        // Сохраняем текущее время для использования в логах
        TimeStamp = DateTime.Now.ToString("yyyy.MM.dd HH-mm");

        // Даем временное название текущей выгрузке для использования в логах
        NameOfImport = "Временное название выгрузки";

        if (!Directory.Exists(ДиректорияДляЛогов))
            Directory.CreateDirectory(ДиректорияДляЛогов);

    }

    private static class Guids {
        public static class References {
            public static Guid СписокНоменклатурыFoxPro = new Guid("c9d26b3c-b318-4160-90ae-b9d4dd7565b6");
            public static Guid Подключения = new Guid("9dd79ab1-2f40-41f3-bebc-0a8b4f6c249e");
            public static Guid ЭСИ = new Guid("853d0f07-9632-42dd-bc7a-d91eae4b8e83");
            public static Guid ЭлектронныеКомпоненты = new Guid("2ac850d9-5c70-45c2-9897-517ab571b213");
            public static Guid Материалы = new Guid("c5e7ae00-90f2-49e9-a16c-f51ed087752a");
            public static Guid Документы = new Guid("ac46ca13-6649-4bbb-87d5-e7d570783f26");
        }

        public static class Parameters {
            // Параметры справочника "Список номерклатуры FoxPro"
            public static Guid НоменклатураОбозначение = new Guid("1bbb1d78-6e30-40b4-acde-4fc844477200"); // string
            public static Guid НоменклатураТипНоменклатуры = new Guid("3c7a075f-0b53-4d68-8242-9f76ca7b2e97"); // int
            public static Guid НоменклатураНаименование = new Guid("c531e1a8-9c6e-4456-86aa-84e0826c7df7"); // string
            public static Guid НоменклатураГОСТ = new Guid("0f48ff0a-36c0-4ae5-ae4c-482f2728181f"); // string

            // Параметры справчоника "Подключения"
            public static Guid ПодключенияСборка = new Guid("4a3cb1ca-6a4c-4dce-8c25-c5c3bd13a807"); // string
            public static Guid ПодключенияКомплектующая = new Guid("7d1ac031-8c7f-49b5-84b8-c5bafa3918c2"); // string
            public static Guid ПодключенияСводноеОбозначение = new Guid("05ffddba-74e9-4637-b249-90cec5953295"); // string
            public static Guid ПодключенияПозиция = new Guid("b05be213-7646-4edb-9d56-391509b48c2a"); // int
            public static Guid ПодключенияКоличество = new Guid("fa56458a-e817-4e6d-85a0-e64dad032c5f"); // double
            public static Guid ПодключенияОКЕИ = new Guid("19d31f8c-06d2-402b-85ee-bda3f5111e8c"); // int
            public static Guid ПодключенияКодEdiz = new Guid("d85db0fe-6c97-4664-9c16-a82695a40984"); // int
            public static Guid ПодключенияЕдиницаИзмерения = new Guid("94158439-cf0a-470b-872c-d783d8ebbd60"); // string
            public static Guid ПодключенияЕдиницаИзмеренияСокр = new Guid("d485a313-6228-4bbf-b40e-b29e82adbb68"); // string
            public static Guid ПодключенияВозвратныеОтходы = new Guid("d9e79828-12d8-4a8a-b77e-9626cedeb307"); // double
            public static Guid ПодключенияПлощадьПокрытия = new Guid("8b12a1f1-0478-4e31-b05f-7205ae683f38"); // double
            public static Guid ПодключенияПотери = new Guid("dd57da68-ebb4-4e43-ab83-05fa621895aa"); // double
            public static Guid ПодключенияТолщинаПокрытия = new Guid("2475329d-3128-4d2d-82a0-af6c35961753"); // double
            public static Guid ПодключенияЧистыйВес = new Guid("e8200590-255d-4a51-9826-686d21c5f2b6"); // double

            // Параметры справочника "ЭСИ"
            public static Guid ЭсиОбозначение = new Guid("ae35e329-15b4-4281-ad2b-0e8659ad2bfb");
            public static Guid ЭсиКодОКП = new Guid("b39cc740-93cc-476d-bfed-114fe9b0740c");

            // Параметры справочника "Документы"
            public static Guid ДокументыОбозначение = new Guid("b8992281-a2c3-42dc-81ac-884f252bd062");
            public static Guid ДокументыКодОКП = new Guid("45ead73a-1773-4156-bafd-48795f844cfb");

            // Параметры справочника "Электронные компоненты"
            public static Guid ЭкОбозначение = new Guid("65e0e04a-1a6f-4d21-9eb4-dfe5a135ec3b");
            public static Guid ЭкКодОКП = new Guid("72f18ec6-d471-45c7-b1df-26f8ccd89af3");

            // Параметры справочника "Материалы"
            public static Guid МатериалыОбозначение = new Guid("d0441280-01ea-43b5-8726-d2d02e4d996f");
            
        }

        public static class Links {
            public static Guid СвязьСпискаНоменклатурыНаЭСИ = new Guid("ec9b1e06-d8d5-4480-8a5c-4e6329993eac");
        }

        public static class Types {
            // Тип справочника ЭСИ
            public static Guid МатериальныйОбъект = new Guid("0ba28451-fb4d-47d0-b8f6-af0967468959");
            // Тип справочника Документы
            public static Guid ОбъектСоставаИзделия = new Guid("f89e9648-c8a0-43f8-82bb-015cfe1486a4");
        }
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
        var ListNum = СписокНоменклатурыСправочник.Objects;
        var dictListNum = ListNum.ToDictionary(objref => (objref[Guids.Parameters.НоменклатураОбозначение].Value.ToString()));
        return dictListNum;
    }

    /// <summary>
    /// Функция возвращает словарь с подключениями, сгруппированными по параметру 'Сборка'
    /// </summary>
    private Dictionary<string, List<ReferenceObject>> GetLinks()
    {
        var RefConnectNum = ПодключенияСправочник.Objects;
        Dictionary<string,List<ReferenceObject>> dict = new Dictionary<string,List<ReferenceObject>>(300000);
        foreach (var item in RefConnectNum)
        {
            string shifr = (String)item[Guids.Parameters.ПодключенияСборка].Value;
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
        var filter =  GeFiltertRefObj(shifr, Guids.References.СписокНоменклатурыFoxPro, Guids.Parameters.НомерклатураОбозначение);
        if (filter != null)
        {
            result = (filter.Select(objref => (objref[Guids.Parameters.НомерклатураОбозначение].Value.ToString()))).ToList();
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
                //result = (List<string>)s2.Select(objref => (objref[Guids.Parameters.НомерклатураОбозначение].Value.ToString()));
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
            NomenclatureTree tree = new NomenclatureTree(nomenclature, links, shifr);
            result.UnionWith(tree.AllReferenceObjects);
            log += $"Дерево изделия {shifr}:\n\n{tree.GenerateLog()}\n\n";
            if (tree.HaveErrors)
                errors.Add(tree.ErrString);
        }

        if (errors.Count != 0)
            Message("Ошибки в процессе построения деревьев", string.Join("\n\n", errors));

        // Пишем лог
        File.WriteAllText(pathToLogFile, log);

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

        foreach (ReferenceObject nom in nomenclature) {
            // Получаем обозначение текущего объекта и его тип
            ReferenceObject resultObject;
            string nomDesignation = (string)nom[Guids.Parameters.НоменклатураОбозначение].Value; // обозначение
            TypeOfObject nomType = DefineTypeOfObject(nom); // тип
            // Пишем информацию в лог
            messages.Add(new String('-', 30));
            messages.Add($"{nomDesignation}:");

            try {
                // СТАДИЯ 1: Пытаемся получить объект по связи на справочник ЭСИ
                resultObject = ProcessFirstStageFindOrCreate(nom, nomDesignation, nomType, messages);
                if (resultObject != null) {
                    result.Add(resultObject);
                    continue;
                }

                // СТАДИЯ 2: Пытаемся найти объект в справочнике ЭСИ
                resultObject = ProcessSecondStageFindOrCreate(nom, nomDesignation, nomType, messages);
                if (resultObject != null) {
                    result.Add(resultObject);
                    continue;
                }

                // СТАДИЯ 3: Пытаемся найти объект в смежных справочниках
                resultObject = ProcessThirdStageFindOrCreate(nom, nomDesignation, nomType, messages);
                if (resultObject != null) {
                    result.Add(resultObject);
                    continue;
                }

                // СТАДИЯ 4: Создаем объект исходя из того, какой был определен тип в справочнике "Список номенклатуры FoxPro"
                resultObject = ProcessFinalStageFindOrCreate(nom, nomDesignation, nomType, messages);
                if (resultObject != null) {
                    result.Add(resultObject);
                }
            }
            catch (Exception e) {
                messages.Add($"Error: {e.Message}");
                continue;
            }
        }

        // Проверяем, были ли ошибки в процессе выполнения данного метода.
        // Если ошибки были, выдаем сообщение пользователю
        string errors = string.Join("\n\n", messages.Where(message => message.StartsWith("Error")));
        if (errors != string.Empty)
            Message("Ошибка", $"В процессе поиска и создания номенклатурных объектов возникли следующие ошибки\n{errors}");

        // Пишем лог
        File.WriteAllText(pathToLogFile, string.Join("\n", messages));

        return result;
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
        ReferenceObject linkedObject = nom.GetObject(Guids.Links.СвязьСпискаНоменклатурыНаЭСИ);

        if (linkedObject != null)
        {
            // TODO: Реализовать код произведения проверки объекта

            TypeOfObject esiType = DefineTypeOfObject(nom); // тип

            
            if (type == esiType)
                if (esiType == TypeOfObject.СтандартноеИзделие || esiType == TypeOfObject.Материал || esiType == TypeOfObject.ЭлектронныйКомпонент)
                {
                    
                    var okp = linkedObject[Guids.Parameters.ЭсиКодОКП].Value.ToString();
                    if (designation.Equals(okp))
                        return linkedObject;
                }
                else
                {
                    var oboz = linkedObject[Guids.Parameters.ЭсиОбозначение].Value.ToString();
                    if (designation.Equals(oboz))
                        return linkedObject;
                }
            else
            {
                messages.Add($"Ошибка! У объекта {designation} тип в foxpro {type} в ЭСИ {esiType}");
            }


        }
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
        List<ReferenceObject> findedObjectInEsi = ЭсиСправочник
            .Find(Guids.Parameters.ЭсиОбозначение, designation) // Производим поиск по всему справочнику
            .Where(finded => finded.Class.IsInherit(Guids.Types.МатериальныйОбъект)) // Отфильтровываем только те объекты, которые наследуются от 'Материального объекта'
            .ToList<ReferenceObject>();

        if (findedObjectInEsi.Count == 1) {
            messages.Add("Объект найден в ЭСИ");
            SyncronizeTypes(nom, findedObjectInEsi[0]);
            return findedObjectInEsi[0];
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
        tempResult = ДокументыСправочник
            .Find(Guids.Parameters.ДокументыОбозначение, designation)
            .Where(finded => finded.Class.IsInherit(Guids.Types.ОбъектСоставаИзделия))
            .ToList<ReferenceObject>();
        if (tempResult != null)
            result.AddRange(tempResult);

        // Производим поиск по справочнику "Электронные компоненты"
        tempResult = ЭлектронныеКомпонентыСправочник.Find(Guids.Parameters.ЭкОбозначение, designation);
        if (tempResult != null)
            result.AddRange(tempResult);
        
        // Производим поиск по справочнику "Материалы"
        tempResult = МатериалыСправочник.Find(Guids.Parameters.МатериалыОбозначение, designation);
        if (tempResult != null)
            result.AddRange(tempResult);

        switch (result.Count) {
            case 0:
                messages.Add("Объект не найден в смежных справочниках");
                return null;
            case 1:
                messages.Add("Объект найден в смежных справочниках");
                return result[0];
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
        if (nomenclatureRecord.Reference.Name != СписокНоменклатурыСправочник.Name)
            throw new Exception($"Неправильное использование метода SyncronizeTypes. Параметр nomenclatureRecord поддерживает только объекты справочника {СписокНоменклатурыСправочник.Name}");

        // Проверка параметра findedObject
        List<string> supportedReferences = new List<string>() {
            ЭсиСправочник.Name,
            ДокументыСправочник.Name,
            ЭлектронныеКомпонентыСправочник.Name,
            МатериалыСправочник.Name,
        };
        if (!supportedReferences.Contains(findedObject.Reference.Name)) {
            throw new Exception($"Неправильное использование метода SyncronizeTypes. Параметр findedObject не поддерживает объекты справочника {findedObject.Reference.Name}");
        }

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
        if (reference == СписокНоменклатурыСправочник.Name) {
            int intType = (int)nomenclature[Guids.Parameters.НоменклатураТипНоменклатуры].Value;
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
        if ((reference == ЭсиСправочник.Name) || (reference == ДокументыСправочник.Name) || (reference == ЭлектронныеКомпонентыСправочник.Name) || (reference == МатериалыСправочник.Name)) {
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
        public bool HaveErrors => this.Errors.Count > 0;
        public string ErrString => this.HaveErrors ? $"Ошибки в дереве {this.NameProduct}:\n{string.Join("\n", this.Errors)}" : string.Empty;

        public NomenclatureTree(Dictionary<string, ReferenceObject> nomenclature, Dictionary<string, List<ReferenceObject>> links, string shifr) {
            if (nomenclature == null || links == null || shifr == null)
                throw new Exception("В конструктор создания NomenclatureTree были поданы отсутствующие значения");

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
            this.Name = (string)this.NomenclatureObject[Guids.Parameters.НоменклатураОбозначение].Value;

            // Создаем список дочерний нод
            this.Children = new List<INode>();
            
            // Рекурсивно получаем потомков
            // Отключаем рекурсию при достижении большого уровня вложенности
            if (this.Level > 100) {
                this.Tree.AddError($"Превышена предельная глубина дерева в 100 уровней, возможна бесконечная рекурсия (node: {this.Name}; parent node: {this.Parent.Name})");
                return;
            }

            if (links.ContainsKey(shifr))
                foreach (string childShifr in links[shifr].Select(link => (string)link[Guids.Parameters.ПодключенияКомплектующая].Value))
                    this.Children.Add(new NomenclatureNode(this.Tree, this, nomenclature, links, childShifr));
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

}
