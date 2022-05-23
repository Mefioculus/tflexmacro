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
    
    public void ImportFileInFileReference() {
        // Для начала получаем объект setOfDocuments
        ReferenceInfo referenceInfo = Context.Connection.ReferenceCatalog.Find(Guids.References.КомплектыДокументов);
        Reference reference = referenceInfo.CreateReference();

        TechnologicalSet setOfDocuments = reference.Find(Guids.Objects.КомплектДокументов) as TechnologicalSet;

        // Получаем с этого объекта нужную директорию
        FolderObject folder = setOfDocuments.Folder as FolderObject;


        




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
}

