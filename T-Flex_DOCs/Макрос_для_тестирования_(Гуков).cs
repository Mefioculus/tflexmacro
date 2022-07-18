using System;
using System.IO;
using System.Linq;
using System.Collections;
using System.Collections.Generic;
using TFlex.DOCs.Model;
using TFlex.DOCs.Model.Classes;
using TFlex.DOCs.Model.Macros;
using TFlex.DOCs.Model.Macros.ObjectModel;
using TFlex.DOCs.Model.References;
using TFlex.Model.Technology.References.SetOfDocuments;
using TFlex.DOCs.Model.References.Files;
using TFlex.DOCs.Model.References.Nomenclature;
using TFlex.DOCs.Model.References.Users;
using TFlex.DOCs.Model.References.Revisions;
using TFlex.DOCs.Model.Desktop;

// Макрос для тестовых задач
// Для работы данного макроса так же потребуется добавление ссылки TFlex.Model.Technology.dll

public class Macro : MacroProvider {
    public Macro (MacroContext context) : base(context) {
    }

    public static class Guids {
        public static class Directories {
            public static Guid ЛичнаяПапка = new Guid("61c60c06-71bd-4aeb-b67d-ef42d8ed04a7");
        }

        public static class Objects {
            public static Guid КомплектДокументов = new Guid("21fa01fa-74d6-4960-80ef-fc5d67adfb51");
        }

        public static class References {
            public static Guid КомплектыДокументов = new Guid ("454c9856-189f-4a53-a2d5-0691dc34c85e");
            public static Guid ЭСИ = new Guid("853d0f07-9632-42dd-bc7a-d91eae4b8e83");
        }

        public static class Parameters {
            public static Guid ОбозначениеЭСИ = new Guid("ae35e329-15b4-4281-ad2b-0e8659ad2bfb");
        }
    }

    // Код тестового макроса
    
    public override void Run() {
        АнализОбъектовСправочникаФайлов();
    }
    
    public void ImportFileInFileReference() {
        // Для начала получаем объект setOfDocuments
        ReferenceInfo referenceInfo = Context.Connection.ReferenceCatalog.Find(Guids.References.КомплектыДокументов);
        Reference reference = referenceInfo.CreateReference();

        TechnologicalSet setOfDocuments = reference.Find(Guids.Objects.КомплектДокументов) as TechnologicalSet;

        // Получаем с этого объекта нужную директорию
        TFlex.DOCs.Model.References.Files.FolderObject folder = setOfDocuments.Folder as TFlex.DOCs.Model.References.Files.FolderObject;

        string pathToFile = @"C:\Users\gukovry\AppData\Local\Temp\testPdf.pdf";

        if (!File.Exists(pathToFile)) {
            Сообщение("Ошибка", "Файл не был найден");
            return;
        }

        if (folder == null) {
            Сообщение("Ошибка", "Не удалось найти директорию для сохранения");
            return;
        }

        FileReference fileReference = new FileReference(Context.Connection);
        FileObject file = fileReference.AddFile(pathToFile, folder);

        if (file == null) {
            Сообщение("Ошибка", "Объект файл не был создан");
        }

    }

    public void ПоискПользователяВГруппахИПользователях() {
        List<User> users = Context.Connection.References.Users.GetAllUsers();
        Message("Количество пользователей", users.Count);
        Message("Найденные пользователи", string.Join("\n", users.Select(user => (string)user[new Guid("42c81c2b-7354-46aa-9547-0f1a93e9d4e1")].Value).OrderBy(login => login)));
    }

    // Метод для получения списка обозначений ДСЕ, которые создала Сергеева (так как они могут считаться эталонами в плане проставления точек в обозначениях)
    public void ПолучитьОбозначенияДокументовСергеевой() {

        Reference nomenclatureReference = Context.Connection.ReferenceCatalog.Find(Guids.References.ЭСИ).CreateReference();
        //nomenclatureReference.Load(new ObjectIterator(10000));
        //List<ReferenceObject>allObjects = nomenclatureReference.GetLoadedObjects();

        List<string> shifrs = nomenclatureReference.Objects.GetAllTreeNodes()
            .Where(rec => rec.SystemFields.Editor.ToString() == "Сергеева Елена Алексеевна")
            .Where(rec => rec.Class.IsInherit(new Guid("0ba28451-fb4d-47d0-b8f6-af0967468959")))
            .Select(rec => ((string)rec[Guids.Parameters.ОбозначениеЭСИ].Value).Replace(" ", string.Empty))
            .Where(rec => rec != string.Empty)
            .Distinct()
            .OrderBy(rec => rec)
            .ToList<string>();

        // Пишем полученные значения в файл
        File.WriteAllText(
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "Обозначения изделий Сергеевой.txt"),
                string.Join("\n", shifrs)
                );
    }

    public void ТестированиеРаботыССтруктурами() {
        NomenclatureReference nomReference = new NomenclatureReference(Context.Connection);
        NomenclatureObject parent1 = nomReference.Find(new Guid("c1529b46-3e69-4236-91a5-ac7ada3ff56e")) as NomenclatureObject;
        NomenclatureObject parent2 = nomReference.Find(new Guid("0adceb5f-1194-4277-9873-2d71cfd6c225")) as NomenclatureObject;
        NomenclatureObject child1 = nomReference.Find(new Guid("f3d1a9b7-1cc2-44fd-a7b4-6e54f50e6b4b")) as NomenclatureObject;
        NomenclatureObject child2 = nomReference.Find(new Guid("2c54e674-c50b-4454-9f13-8adfa748e027")) as NomenclatureObject;

        Message(
                "Информация",
                string.Format(
                    "{0}\n\n{1}\n\n{2}\n\n{3}",
                    GetAllLinks(parent1),
                    GetAllLinks(parent2),
                    GetAllLinks(child1),
                    GetAllLinks(child2)
                    ));
    }

    private string GetAllLinks(NomenclatureObject dse) {
        string result =
            $"Объект {dse.ToString()}" +
            $"\nРодительские подключения: {string.Join("; ", dse.Parents.GetHierarchyLinks().Select(link => GetInfoAboutLink(link)))}" +
            $"\nДочерние подключения: {string.Join("; ", dse.Children.GetHierarchyLinks().Select(link => GetInfoAboutLink(link)))}";


        return result;
    }

    private string GetInfoAboutLink(ComplexHierarchyLink hLink) {
        string name = $"{(string)hLink.ParentObject[new Guid("ae35e329-15b4-4281-ad2b-0e8659ad2bfb")].Value} -> {(string)hLink.ChildObject[new Guid("ae35e329-15b4-4281-ad2b-0e8659ad2bfb")].Value}";
        string connectionsToStructures = $"{hLink.GetObjects(new Guid("77726357-b0eb-4cea-afa5-182e21eb6373")).Count.ToString()}";
        return $"{name} {connectionsToStructures}";
    }

    public void АнализОбъектовСправочникаФайлов() {
        FileReference fileReference = new FileReference(Context.Connection);
        List<ReferenceObject> files = fileReference.Objects.GetAllTreeNodes();

        string dir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "Анализ файлового справочника");
        if (!Directory.Exists(dir))
            Directory.CreateDirectory(dir);

        Dictionary<string, int> quantity = new Dictionary<string, int>();
        List<string>errors = new List<string>();
        foreach (ClassObject classObject in fileReference.Classes) {
            try {
                quantity.Add($"{classObject.Name} {classObject.Guid.ToString()}", ПолучитьВсеФайлыТипа(files, classObject, dir));
            }
            catch (Exception e) {
                throw e;
                errors.Add($"{classObject.Name}");
            }
        }

        if (errors.Count > 0)
            Message("Информация", $"Возникли ошибки при обработке классов:\n{string.Join("\n", errors)}");
        Message("Информация", $"Работа макроса завершена:\n\n{string.Join("\n", quantity.Select(kvp => $"{kvp.Key}: {kvp.Value} шт."))}");

    }

    public void ПолучитьАттрибутыТипа() {
        FileReference fileReference = new FileReference(Context.Connection);
        ClassObject testClass = fileReference.Classes.Find(new Guid("a477f9ed-37b5-4b70-a968-a89f8af1b37d"));
        if (testClass == null)
            Message("", "Не удалось найти класс");
            return;
        if (testClass.Attributes == null) {
            Message("", "Тестовый класс не содержит аттрибутов");
            return;
        }
        Message("Информация", string.Join("\n", testClass.Attributes.Select(attr => $"{attr.Name}: {attr.Value.ToString()}")));
    }

    private int ПолучитьВсеФайлыТипа(List<ReferenceObject> files, ClassObject classObject, string pathToDir) {

        List<string> paths = files
            .Where(refObj => refObj.Class.IsInherit(classObject))
            //.Select(refObj => (FileObject)refObj)
            //.Select(file => file.Path.ToString())
            //.Select(refObj => refObj.ToString())
            .Select(refObj => refObj is FileObject ? $"{((FileObject)refObj).Path.ToString()}" : refObj.ToString())
            .ToList<string>();

        string extension = classObject.Attributes != null ?
            classObject.Attributes.Contains("Extension") ?
                classObject.Attributes["Extension"] != null ?
                    Convert.ToString(classObject.Attributes["Extension"]).Replace("Расширение: ", string.Empty) :
                    string.Empty :
                string.Empty :
            string.Empty;

        string pathToFile = Path.Combine(pathToDir, $"{extension} - ({classObject.Guid.ToString()}).txt");

        try {
            File.WriteAllText(
                    Path.Combine(pathToDir, $"{extension} - {classObject.Name.Replace(@"\", "").Replace(@"/", "").Replace("\"", "'")} ({classObject.Guid.ToString()}).txt"),
                    string.Join("\n", paths)
                    );
        }
        catch (Exception e) {
            throw new Exception($"При создании файла по пути {pathToFile} возникла ошибка {e.Message}");
        }

        return paths.Count();
    }

    public void ПолучениеАтрибутовКлассов() {
        FileReference fileReference = new FileReference(Context.Connection);

        List<string> resultsWithAttr = new List<string>();
        List<string> resultsWithoutAttr = new List<string>();


        foreach (ClassObject classObject in fileReference.Classes) {
            if (classObject.Attributes != null)
                resultsWithAttr.Add($"Класс {classObject.Name} содержит аттрибуты: {string.Join("; ", classObject.Attributes.Select(attr => $"{attr.Name}: {attr.Value}"))}");
            else
                resultsWithoutAttr.Add($"Класс {classObject.Name} не содержит никаких аттрибутов");
        }

        if (Question($"Типов с аттрибутами: {resultsWithAttr.Count}\nТипов без аттрибутов: {resultsWithoutAttr.Count}\nОтобразить результаты?"))
            Message("Результаты", $"{string.Join("\n", resultsWithAttr)}\n\n{string.Join("\n", resultsWithoutAttr)}");
    }

    public void ТестированиеСозданияРевизии() {
        InputDialog dialog = new InputDialog(Context, "Изменение ревизии");
        string guidField = "Введите GUID объекта";
        string nameRevisionField = "Целевое название ревизии";
        dialog.AddString(guidField);
        dialog.AddString(nameRevisionField);

        if (dialog.Show()) {
            ReferenceObject currentObject = Context.Connection.ReferenceCatalog
                .Find(new Guid("853d0f07-9632-42dd-bc7a-d91eae4b8e83"))
                .CreateReference()
                .Find(new Guid(dialog[guidField]));
            if (currentObject == null) {
                Message("Ошибка", $"Не удалось найти номенклатурный объект с уникальным идентификатором '{dialog[guidField]}'");
                return;
            }

            NomenclatureObject nomObject = currentObject as NomenclatureObject;
            
            // Производим смену названия ревизии
            if (!nomObject.IsCheckedOut)
                nomObject.CheckOut();
            nomObject.BeginChanges();
            nomObject.SystemFields.RevisionName = dialog[nameRevisionField];
            nomObject.EndChanges();
            Desktop.CheckIn(nomObject, "Изменение названия ревизии", false);

            ReferenceObject linkedObject = nomObject.LinkedObject;
            if (!linkedObject.IsCheckedOut)
                linkedObject.CheckOut();
            linkedObject.BeginChanges();
            linkedObject.SystemFields.RevisionName = dialog[nameRevisionField];
            linkedObject.EndChanges();
            Desktop.CheckIn(linkedObject, "Изменение названия ревизии", false);
            
        }
    }

    public void ПолучитьНазванияВсехРевизийДляОбъекта() {
        InputDialog dialog = new InputDialog(Context, "Выбор объекта");
        string guidField = "Укажите Guid объекта";
        dialog.AddString(guidField);

        if (dialog.Show()) {
            ReferenceObject currentObject = Context.Connection.ReferenceCatalog
                .Find(new Guid("853d0f07-9632-42dd-bc7a-d91eae4b8e83"))
                .CreateReference()
                .Find(new Guid(dialog[guidField]));
            if (currentObject == null)
                throw new Exception($"Во время работы метода {nameof(ПолучитьНазванияВсехРевизийДляОбъекта)} возникла ошибка. Не получилось найти объект с Guid {dialog[guidField]}");

            // Выводим список всех имен ревизий
            Message($"Ревизии объекта {currentObject.ToString()}", string.Join(Environment.NewLine, currentObject.GetExistingRevisionNames()));
        }
    }

    public void ПолучитьВсеОбъектыПоИмениРевизии() {
        InputDialog dialog = new InputDialog(Context, "Введите название ревизии");
        string revisionNameField = "Укажите название ревизии";
        dialog.AddString(revisionNameField);

        if (dialog.Show()) {
            List<ReferenceObject> findedObjects = Context.Connection.ReferenceCatalog
                .Find(new Guid("853d0f07-9632-42dd-bc7a-d91eae4b8e83"))
                .CreateReference()
                .Find(new Guid("8d69bd40-0fe0-4bb1-9d1a-2e728f6cdc68"), dialog[revisionNameField]);

            if (findedObjects.Count == 0)
                Message("Информация", $"По запросу '{dialog[revisionNameField]}' ничего не было найдено");
            else
                Message("Информация", string.Join(Environment.NewLine, findedObjects.Select(obj => $"{obj.ToString()}")));
        }
    }

    public void ВывестиУникальныеИдентификаторыСистемныхПараметров() {
        // Получаем справочник, для которого будет получать системные параметры
        Reference reference = Context.Connection.ReferenceCatalog
            .Find(new Guid("853d0f07-9632-42dd-bc7a-d91eae4b8e83"))
            .CreateReference();

        List<string> messages = new List<string>();
        foreach (var parameterInfo in reference.ParameterGroup.SystemParameters) {
            messages.Add($"Параметр: {parameterInfo.ToString()}; GUID: {parameterInfo.Guid.ToString()}");
        }

        Message("Системные параметры справочника ЭСИ", string.Join(Environment.NewLine, messages));
    }

    public void ПолучитьВсеРевизииНаОбъект() {
        InputDialog dialog = new InputDialog(Context, "Поиск ревизий объекта");
        string guidField = "Введите guid объекта";
        dialog.AddString(guidField);

        if (dialog.Show()) {
            Reference reference = Context.Connection.ReferenceCatalog
                .Find(new Guid("853d0f07-9632-42dd-bc7a-d91eae4b8e83"))
                .CreateReference();

            ReferenceObject initialObject = reference.Find(new Guid(dialog[guidField]));
            if (initialObject == null)
                throw new Exception($"Для переданного guid ({dialog[guidField]}) не было найдено совпадения в справочнике ЭСИ");

            // Пробуем получить все ревизии
            List<ReferenceObject> allRevision = initialObject.Reference.Find(new Guid("49c7b3ec-fa35-4bb1-92a5-01d4d3a40d16"), initialObject.SystemFields.LogicalObjectGuid);

            Message("Информация", string.Join(Environment.NewLine, allRevision.Select(rev => rev.ToString())));

        }
    }
}

