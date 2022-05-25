using System;
using System.IO;
using System.Linq;
using System.Collections;
using System.Collections.Generic;
using TFlex.DOCs.Model;
using TFlex.DOCs.Model.Macros;
using TFlex.DOCs.Model.References;
using TFlex.Model.Technology.References.SetOfDocuments;
using TFlex.DOCs.Model.References.Files;
using TFlex.DOCs.Model.References.Nomenclature;
using TFlex.DOCs.Model.References.Users;

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
        ТестированиеРаботыССтруктурами();
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
}


