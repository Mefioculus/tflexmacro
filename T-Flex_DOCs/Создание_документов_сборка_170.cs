using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using TFlex.DOCs.Model.Classes;
using TFlex.DOCs.Model.Desktop;
using TFlex.DOCs.Model.Macros;
using TFlex.DOCs.Model.Macros.ObjectModel;
using TFlex.DOCs.Model.References;
using TFlex.DOCs.Model.References.Documents;
using TFlex.DOCs.Model.References.Files;
using TFlex.DOCs.Model.References.Nomenclature;
using TFlex.DOCs.Model.Search;
using TFlex.DOCs.Model.Structure;
using FolderObject = TFlex.DOCs.Model.References.Files.FolderObject;

public class MacroAEMCreateDocuments : MacroProvider
{
    /// <summary>
    /// Строка для записи лога ошибок
    /// </summary>
    private readonly StringBuilder _logRemarks = new StringBuilder();

    public MacroAEMCreateDocuments(MacroContext context)
        : base(context)
    {
        if (Вопрос("Хотите запустить в режиме отладки?"))
        {
            System.Diagnostics.Debugger.Launch();
            System.Diagnostics.Debugger.Break();
        }
    }

    public override void Run()
    {
    }

    /// <summary>
    /// Создание технологических документов
    /// </summary>
    public void СоздатьТехнологическиеДокументы()
    {
        СоздатьДокументы(ContextImport.DocumentTypes.Technologist);
    }

    /// <summary>
    /// Создание конструкторских документов
    /// </summary>
    public void СоздатьКонструкторскиеДокументы()
    {

        СоздатьДокументы(ContextImport.DocumentTypes.Constructor);
    }


    public void СоздатьДокументы(ContextImport.DocumentTypes typeDocuments)
    {
        ДиалогОжидания.Показать("Подождите, идет обработка", false);

        AddTextLog(new string('-', 72));

        List<FileReferenceObject> fileReferenceObjects = new List<FileReferenceObject>();

        FolderObject rootFolderObject = null;

        GetSelectFilesAndFolder(out rootFolderObject, out fileReferenceObjects);

        //Находим и заполняем все настройки, такие как справочник документов и т.д.
        var contextImport = CreateContextImport(rootFolderObject, typeDocuments);

        int current = 0;

        //Перебираем все объекты 
        foreach (var fileReferenceObject in fileReferenceObjects)
        {
            current++;
            if (!ДиалогОжидания.СледующийШаг(
                string.Format("Подождите идет обработка объектов: {0} из {1}", current, fileReferenceObjects.Count)))
            {
                return;
            }

            ProcessFileObject(fileReferenceObject, contextImport);
        }

        /*
        ReferenceObjectSaveSet.RunWithSaveSet(set =>
        {
            //Пакетное сохранение объектов на сервер
            foreach (ReferenceObject saveObj in contextImport.SaveObjects)
            {
                if(saveObj.Changing)
                set.Add(saveObj);
            }

            set.EndChanges();
        });*/

        Reference.EndChanges(contextImport.SaveObjects.Where(obj => obj.Changing));

        //Применение изменений объектов
        if (Вопрос("Применить изменения созданных объектов?"))
            Desktop.CheckIn(contextImport.SaveObjects, "Подключение электронного архива из файлов", false);

        AddTextLog(Environment.NewLine + "Завершение процесса создания документов");

        //Запись файла во временную папку
        string logFilePath = string.Format("{0}{1} {2:d.MM.yyyy HH-mm-ss}.txt", Path.GetTempPath(), "Загрузка файлов в DOCs", DateTime.Now);
        File.AppendAllText(logFilePath, Environment.NewLine + _logRemarks);


        FileReference fileReference = rootFolderObject.Reference;
        //Загрузка файла в докс
        FileObject logObjPath = fileReference.AddFile(logFilePath, rootFolderObject);

        ДиалогОжидания.Скрыть();
        Сообщение("Создание завершено", "Создание объектов документов было завершено. Лог можно найти по пути: " + logObjPath.Path);

    }

    private void GetSelectFilesAndFolder(out FolderObject rootFolderObject, out List<FileReferenceObject> fileReferenceObjects)
    {
        var selectedObjects = Context.GetSelectedObjects();
        rootFolderObject = null;
        fileReferenceObjects = new List<FileReferenceObject>();

        //Если были выбрано несколько объектов
        if (selectedObjects.Length > 1)
        {
            //Получаем родительскую папку
            rootFolderObject = Context.ReferenceObject.Parent as FolderObject;

            foreach (var selectedObject in selectedObjects)
            {
                //Приводим выбранные объекты к типу файл
                var fileReferenceObject = selectedObject as FileReferenceObject;
                if (fileReferenceObject != null)
                {
                    fileReferenceObjects.Add(selectedObject as FileReferenceObject);
                }
            }
        }
        else
        {
            //Если выбран один объект
            var fileReferenceObject = Context.ReferenceObject as FileReferenceObject;
            if (fileReferenceObject != null)
            {
                //Если выбранный объект папка
                if (fileReferenceObject.IsFolder)
                {
                    //Добавляем дочерние объекты итд
                    rootFolderObject = fileReferenceObject as FolderObject;
                    fileReferenceObjects.AddRange(fileReferenceObject.Children);
                }
                else
                {
                    //Если файл, то должны файл записать в коллекцию, чтобы его обработать и получить папку
                    rootFolderObject = fileReferenceObject.Parent;

                    fileReferenceObjects.Add(fileReferenceObject);
                }
            }
            else
                Ошибка("Выбранный объект не является частью справочника 'Файлы'");
        }
    }

    private void ProcessFileObject(FileReferenceObject fileReferenceObject, ContextImport contextImport)
    {
        //Получаем наименование файла, без расширения. т.е. Name^Code.grb => Name^Code
        string fileObjectName = Path.GetFileNameWithoutExtension(fileReferenceObject.Name);

        //Проверка, что объект является файлом, иначе пишем в лог
        if (!fileReferenceObject.IsFile)
        {
            AddTextLog(string.Format("Объект {0}, не является файлом", fileObjectName));
            return;
        }

        string docName = string.Empty;
        string docDenotation = string.Empty;

        if (contextImport.DocumentType == ContextImport.DocumentTypes.Constructor)
        {
            //Проверяем есть ли символ ^ так что возвращает true если все хорошо.
            //Записывает наименование обозначение в переменные docName, docCode
            if (!TryParseFileName(fileObjectName, out docName, out docDenotation))
                return;
        }

        if (contextImport.DocumentType == ContextImport.DocumentTypes.Technologist)
        {
            if (!TryParseFileName(fileObjectName, out docName, out docDenotation))
                docName = fileObjectName;
        }

        //Получаем связанные документы Settings.DocumentFiles это гуид связи
        var linkDoсuments = fileReferenceObject.GetObjects(Settings.ФайлыДокумента);

        ReferenceObject foundLinkedDocument = linkDoсuments.FirstOrDefault(docObj =>
            (docObj[Settings.НаименованиеДокумента].GetString() == docName) &&
            (docObj[Settings.ОбозначениеДокумента].GetString() == docDenotation));

        //Если нашли такой документ, то подключать не нужно, пишем в лог что такой документ уже подключен
        if (foundLinkedDocument != null)
        {
            AddTextLog(string.Format("Файл: '{0}' уже подключен к документу с наименование: '{1}' и обозначением: '{2}'",
                fileObjectName, docName, docDenotation));
            return;
        }

        //Находим документ по наименование и обозначению
        var documentReferenceObject = FindDocumentReferenceObject(contextImport, docName, docDenotation);
        if (documentReferenceObject == null)
        {
            //Разбирает обозначение и получает у него тип объекта
            string typeObject = string.Empty;
            if (!TryGetDocumentTypeObject(contextImport, docDenotation, out typeObject))
            {
                AddTextLog(string.Format("Невозможно работать тип объекта из файла: {0}", fileObjectName));
                return;
            }

            //Создаем объект справочника документы с наименованием обозначением и указанным типом
            documentReferenceObject = CreateDocumentObject(contextImport, docName, docDenotation, typeObject);
        }

        //Берем в редактирование файл
        if (!TryBeginChangesReferenceObject(fileReferenceObject))
        {
            AddTextLog(string.Format("Невозможно взять в редактирование объект '{0}'", fileReferenceObject));
            return;
        }

        //Добавляем документ по связи к файлу
        fileReferenceObject.AddLinkedObject(Settings.ФайлыДокумента, documentReferenceObject);

        contextImport.SaveObjects.Add(fileReferenceObject);
    }

    private bool TryGetDocumentTypeObject(ContextImport contextImport, string docDenotation, out string typeObject)
    {
        typeObject = string.Empty;

        if (contextImport.DocumentType == ContextImport.DocumentTypes.Constructor)
        {
            typeObject = ParseTypeOfName(docDenotation);
            return true;
        }

        if (contextImport.DocumentType == ContextImport.DocumentTypes.Technologist)
        {
            typeObject = Settings.ДокументТипТехнологическийДокумент;
            return true;
        }

        return false;
    }

    /// <summary>
    /// Преобразует из строки тип, возвращает строку с типом, если удалось разобрать
    /// Если не удалось разобрать, то возвращает тип "Деталь"
    /// </summary>
    /// <param name="docCode"></param>
    /// <returns></returns>
    private string ParseTypeOfName(string docCode)
    {
        // Проверяем не сборочный чертеж ли это.
        // Есть такая строка "АБВГ.002003.003 СБ"
        // надо узнать оканчивается ли она на "СБ"
        if (docCode.EndsWith("СБ"))
            return Settings.ДокументТипЧертеж;

        //Дальше проверяем не сборка ли это.
        //Надо найти число которое стоит за первой точкой т.е. "АБВГ.3" нам нужна 3
        //Первое условие, находим точку в строке
        if (docCode.Contains('.'))
        {
            //Разделяем строку на 2 составляющие, получаем вторую т.е. 301002
            string secondSplitStirng = docCode.Split('.')[1];

            //Если указанная строка начинается с "3" то это тип Сборочная единица
            if (secondSplitStirng.StartsWith("3"))
                return Settings.ДокументТипСборка;
        }

        //Во всех остальных случаях это "Деталь"
        return Settings.ДокументТипДеталь;
    }

    /// <summary>
    /// Создает объект в справочнике документы, тип по умолчанию Деталь
    /// </summary>
    /// <param name="contextImport"></param>
    /// <param name="name"></param>
    /// <param name="code"></param>
    /// <param name="classObjectName"></param>
    /// <returns></returns>
    private ReferenceObject CreateDocumentObject(ContextImport contextImport, string name, string code, string classObjectName = "Деталь")
    {
        //Получаем наименование папки которое является текущим объектом
        var folderName = contextImport.FileFolderObject.Name;

        //Находим есть ли в контексте уже папка документов
        if (contextImport.DocumentFolderReferenceObject == null)
        {
            //Если нет папки, то мы находим папку в документах по наименованию
            contextImport.DocumentFolderReferenceObject = FindDocumentReferenceObject(contextImport, folderName);

            //Если не нашли папку в справочнике, то создаем папку
            if (contextImport.DocumentFolderReferenceObject == null)
                contextImport.DocumentFolderReferenceObject = CreateDocumentFolder(contextImport, folderName);
        }

        //Находим указанный тип
        ClassObject createdClassObject = contextImport.DocumentReference.Classes.Find(classObjectName);
        if (createdClassObject == null)
        {
            AddTextLog(string.Format(
                "Не найден тип '{0}' в справочнике документы",
                folderName));
            return null;
        }

        //Создаем объект в справочнике документы
        var createDocumentObject = contextImport.DocumentReference.CreateReferenceObject
            (contextImport.DocumentFolderReferenceObject, createdClassObject);

        //Заполняем параметры
        createDocumentObject[Settings.НаименованиеДокумента].Value = name;
        createDocumentObject[Settings.ОбозначениеДокумента].Value = code;

        //Заполняем параметр Вид документа
        if (contextImport.DocumentType == ContextImport.DocumentTypes.Technologist)
            createDocumentObject[Settings.ВидДокументаВДокументах].Value = contextImport.TypeOfDocumentValue;

        createDocumentObject.EndChanges();

        if (contextImport.DocumentType == ContextImport.DocumentTypes.Constructor)
            AddLinkNomenclature(contextImport, createDocumentObject);

        //Добавляем объект для применения изменений объектов
        contextImport.SaveObjects.Add(createDocumentObject);

        return createDocumentObject;
    }

    /// <summary>
    /// Подключаем объект документа к номенклатуре
    /// </summary>
    /// <param name="contextImport"></param>
    /// <param name="createDocumentObject"></param>
    private void AddLinkNomenclature(ContextImport contextImport, ReferenceObject createDocumentObject)
    {
        //Находим папку, в номенклатуре, для указания родителя
        //if (contextImport.NomenclatureFolderObject == null)
        //    contextImport.NomenclatureFolderObject = CreateNomenclatureFolderObject(contextImport);

        //Создаем объект номенклатуры
        var createdNomenclatureObject = contextImport.NomenclatureReference.CreateNomenclatureObject
            (createDocumentObject);

        if (createdNomenclatureObject.Changing)
            createdNomenclatureObject.EndChanges();

        AddLinkNomenclatureEquipment(contextImport, createdNomenclatureObject);
    }

    /// <summary>
    /// 
    /// </summary>
    /// <param name="contextImport"></param>
    /// <param name="createdNomenclatureObject"></param>
    private void AddLinkNomenclatureEquipment(ContextImport contextImport, NomenclatureObject createdNomenclatureObject)
    {
        ReferenceObject foundEquipmentreferenceObject =
            contextImport.CatalogEquipmentReference.FindOne(contextImport.ParameterNameEquipment, createdNomenclatureObject.Name);
        if (foundEquipmentreferenceObject != null)
        {
            if (!TryBeginChangesReferenceObject(foundEquipmentreferenceObject))
            {
                AddTextLog(string.Format("Невозможно взять в редактирование объект '{0}'", foundEquipmentreferenceObject));
                return;
            }

            foundEquipmentreferenceObject.AddLinkedObject(Settings.ОснащениеСвязанныйНоменклатурныйОбъект,
                createdNomenclatureObject);
            foundEquipmentreferenceObject.EndChanges();
        }
    }

    /// <summary>
    /// Создает папку в справочнике номенклатура
    /// </summary>
    /// <param name="contextImport"></param>
    /// <returns></returns>
    private NomenclatureReferenceObject CreateNomenclatureFolderObject(ContextImport contextImport)
    {
        var folder = FindObject(Settings.СправочникНоменклатура.ToString(), "Наименование", "Импорт");

        if (folder == null)
        {
            folder = CreateObject(Settings.СправочникНоменклатура.ToString(), "Папка");
            folder["Наименование"] = "Импорт";
            folder.Save();
        }

        contextImport.SaveObjects.Add((ReferenceObject)folder);

        return (NomenclatureReferenceObject)folder;
    }

    /// <summary>
    /// Находит объект в справочнике Документы, по совпадению наименование обозначение
    /// </summary>
    /// <param name="contextImport"></param>
    /// <param name="name"></param>
    /// <param name="code"></param>
    /// <returns></returns>
    private ReferenceObject FindDocumentReferenceObject(ContextImport contextImport, string name, string code = "")
    {
        //Формируем базовую строку фильтра по наименованию
        string filterString = string.Format("[{0}] = '{1}'", "Наименование", name);

        //Если указанно обозначение, его так же добавляем в фильтр
        if (!string.IsNullOrEmpty(code))
            filterString += string.Format(" И [{0}] = '{1}'", "Обозначение", code);

        Filter filter;

        //Преобразуем строку фильтра в фильтр для справочника
        if (!Filter.TryParse(filterString, contextImport.DocumentReference.ParameterGroup, out filter))
            return null;

        //Находим и возвращаем первый найденный объект
        return contextImport.DocumentReference.Find(filter).FirstOrDefault();
    }

    /// <summary>
    /// Преобразует строку наименования файла в две строки наименование и обозначение
    /// </summary>
    /// <param name="fileName"></param>
    /// <param name="name"></param>
    /// <param name="denotation"></param>
    /// <returns>Возвращает true если преобразование произошло</returns>
    private bool TryParseFileName(string fileName, out string name, out string denotation)
    {
        name = string.Empty;
        denotation = string.Empty;

        //Проверка на наличие спец символа
        if (!fileName.Contains(Settings.СтрокаРазделитель))
        {
            AddTextLog(string.Format(
                "Невозможно разобрать наименование объекта {0}{1}В наименовании отсутствует символ: '{2}'",
                fileName, Environment.NewLine, Settings.СтрокаРазделитель));
            return false;
        }

        //Разбиваем строку на массив строк
        var splitFileName = fileName.Split(new string[] { Settings.СтрокаРазделитель }, StringSplitOptions.RemoveEmptyEntries);

        //Проверяем что при разбитии строк получилось 2 значения
        if (splitFileName.Length > 2)
        {
            AddTextLog(string.Format
            ("Невозможно разобрать наименование объекта {0}{1}В наименовании присутствуют больше двух символов: '{2}'",
                fileName, Environment.NewLine, Settings.СтрокаРазделитель));
            return false;
        }

        //Записываем значения из массива строк
        name = splitFileName[1];
        denotation = splitFileName[0];

        return true;
    }

    private ReferenceObject CreateDocumentFolder(ContextImport contextImport, string folderName)
    {
        var createdDocumentFolderReferenceObject = contextImport.DocumentReference.CreateReferenceObject(contextImport.DocumentFolderClassObject);

        createdDocumentFolderReferenceObject[Settings.НаименованиеДокумента].Value = folderName;
        createdDocumentFolderReferenceObject.EndChanges();

        return createdDocumentFolderReferenceObject;
    }

    private void AddTextLog(string text)
    {
        _logRemarks.AppendLine(string.Format("{0:d.MM.yyyy HH-mm-ss}: {1}{2}", DateTime.Now, text, Environment.NewLine));
    }

    private ContextImport CreateContextImport(FolderObject folderObject, ContextImport.DocumentTypes DocumentType)
    {
        var documentReference = (DocumentReference)FindReference(Settings.СправочникДокументы);

        var nomenclatureReference = (NomenclatureReference)FindReference(Settings.СправочникНоменклатура);

        var сatalogEquipmentReference = FindReference(Settings.СправочникКаталогОснащения);

        var documentTypeParameterInfo = FindParameterInfo(documentReference, Settings.ВидДокументаВДокументах);

        var equipmentName = FindParameterInfo(сatalogEquipmentReference, Settings.НаименованиеОснащения);

        //Надо разобрать список значений параметра документа
        int typeDocumentValue = 0;

        //Находим элемент списка значений, по наименованию
        //Переводим в нижний регистр, чтобы не было регистро устойчивости
        var foundTypeDocumentInValue = documentTypeParameterInfo.ValueList
            .FirstOrDefault(parValue =>
                string.Equals(parValue.Name, folderObject.Name.ToString(), StringComparison.CurrentCultureIgnoreCase));

        if (foundTypeDocumentInValue != null)
            typeDocumentValue = (int)foundTypeDocumentInValue.Value;


        var classObject = documentReference.Classes.Find("Папка");

        var contextImport = new ContextImport()
        {
            DocumentReference = documentReference,
            FileFolderObject = folderObject,
            DocumentFolderClassObject = classObject,
            SaveObjects = new List<ReferenceObject>(),
            NomenclatureReference = nomenclatureReference,
            DocumentType = DocumentType,
            TypeOfDocumentValue = typeDocumentValue,
            CatalogEquipmentReference = сatalogEquipmentReference,
            ParameterNameEquipment = equipmentName
        };

        return contextImport;
    }

    private Reference FindReference(Guid referenceGuid)
    {
        var foundReferenceInfo = Context.Connection.ReferenceCatalog.Find(referenceGuid);
        if (foundReferenceInfo == null)
            Ошибка(string.Format("Не найден справочник с гуидом '{0}'", referenceGuid));

        return foundReferenceInfo.CreateReference();
    }

    private ParameterInfo FindParameterInfo(Reference reference, Guid parameterGuid)
    {
        var foundParameterInfo = reference.ParameterGroup.OneToOneParameters.Find(parameterGuid);
        if (foundParameterInfo == null)
            Ошибка(string.Format("В справочнике '{0}' не найден параметр '{1}'", reference, parameterGuid));

        return foundParameterInfo;
    }

    private static bool TryBeginChangesReferenceObject(ReferenceObject referenceObject)
    {
        if (referenceObject.IsCheckedOut)
        {
            if (referenceObject.IsCheckedOutByCurrentUser)
            {
                if (referenceObject.CanEdit)
                {
                    referenceObject.BeginChanges();
                    return true;
                }

                return referenceObject.Changing;
            }
        }
        else if (referenceObject.CanCheckOut)
        {
            referenceObject.CheckOut(false);

            if (referenceObject.CanEdit)
            {
                referenceObject.BeginChanges();
                return true;
            }
        }
        return false;
    }

    public class ContextImport
    {
        public enum DocumentTypes
        {
            Constructor,
            Technologist
        };

        public DocumentTypes DocumentType;

        public Reference DocumentReference;
        public DocumentType DocumentFolderClassObject;

        public int TypeOfDocumentValue;

        public NomenclatureReference NomenclatureReference;

        public Reference CatalogEquipmentReference;

        public ParameterInfo ParameterNameEquipment;

        //public NomenclatureReferenceObject NomenclatureFolderObject;

        /// <summary>
        /// Папка в справочнике документы в которую будут создаваться новые документы
        /// </summary>
        public ReferenceObject DocumentFolderReferenceObject;

        /// <summary>
        /// ТкущийОбъект только с переопределенным типом
        /// </summary>
        public FileReferenceObject FileFolderObject;

        /// <summary>
        /// Для пакетного сохранения объектов
        /// </summary>
        public List<ReferenceObject> SaveObjects;
    }

    private static class Settings
    {
        public static Guid СправочникДокументы = new Guid("ac46ca13-6649-4bbb-87d5-e7d570783f26");

        public static Guid НаименованиеДокумента = new Guid("7e115f38-f446-40ce-8301-9b211e6ce5fd");
        public static Guid ОбозначениеДокумента = new Guid("b8992281-a2c3-42dc-81ac-884f252bd062");
        public static Guid ФайлыДокумента = new Guid("9eda1479-c00d-4d28-bd8e-0d592c873303");
        public static Guid ВидДокументаВДокументах = new Guid("0a50c800-161a-4f6a-8738-e712693a02b3");

        public static string ДокументТипДеталь = "7c41c277-41f1-44d9-bf0e-056d930cbb14";
        public static string ДокументТипСборка = "dd2cb8e8-48fa-4241-8cab-aac3d83034a7";
        public static string ДокументТипЧертеж = "d6324424-a39e-4112-a207-7e96a1971852";
        public static string ДокументТипТехнологическийДокумент = "9745b167-0e66-43c2-91c0-899a0149b19e";

        public static Guid СправочникНоменклатура = new Guid("853d0f07-9632-42dd-bc7a-d91eae4b8e83");

        public static Guid СправочникКаталогОснащения = new Guid("a13d8c5a-c76c-4df6-91df-fc5bbf8db86d");

        public static Guid НаименованиеОснащения = new Guid("9188798b-0df9-49e3-83cd-04d77d5838c2");

        public static Guid ОснащениеСвязанныйНоменклатурныйОбъект = new Guid("cf883d14-919e-49ba-b9ba-44a732c06924");

        public static string СтрокаРазделитель = "^";
    }
}
