using System;
using System.Collections;
using System.Collections.Generic;
using System.IO.Compression; // Для работы с zip архивами
using System.Web; // Данное пространство имен подключено для того, чтобы произвести декодирование данных с кодировки Url
using Forms = System.Windows.Forms; // Для отображения диалогового окна выбора файла для открытия и для сохранения
// Пространства имен DOCs
using TFlex.DOCs.Model.Macros;
using TFlex.DOCs.Model.Macros.ObjectModel; // пространство имен для отображения диалога ожидания 
using TFlex.DOCs.Model.References;
using TFlex.DOCs.Model.References.Files;
using TFlex.DOCs.Model.References.Nomenclature;
using TFlex.DOCs.Model.DataExchange; // пространстпо имен для экспорта данных

// Для работы макроса так же необходимо подключить следующие библиотеки
// System.Web.dll
// System.IO.Compression.dll
// System.IO.Compression.ZipFile.dll
// System.IO.Compression.FileSystem.dll

public class Macro : MacroProvider {
    public Macro(MacroContext context)
        : base(context)
    {
    }

    #region Guids

    private static class Guids {
        public static class References {
            public static Guid Files = new Guid("a0fcd27d-e0f2-4c5a-bba3-8a452508e6b3");
            public static Guid Nomenclature = new Guid("853d0f07-9632-42dd-bc7a-d91eae4b8e83");
        }

        public static class Links {
            public static Guid DocumentsToFiles = new Guid("89e45926-0f0f-4c36-b649-3784d274e348");
        }

        public static class Types {
            
        }

        public static class Properties {
            public static Guid RelativePath = new Guid("adda774c-dbdf-48ba-bcf6-87bb42a67e90");
        }
    }

    #endregion Guids

    #region Fields and Properties

    string simpleTemplate = "{0}\n";

    #endregion Fields and Properties

    #region Entry points

    #region Method Run

    public override void Run() {
        // Получаем от пользователя список изделий, для которого нужно произвести экспорт
        List<ReferenceObject> listOfReferenceObjects = GetReferenceObjects();
        if (listOfReferenceObjects.Count == 0)
            return;

        // Инициалилзируем список, который будет хранить все объекты, для которых необхоимо произвести экспорт
        List<ReferenceObject> listAllObjectsOnExport= new List<ReferenceObject>();

        // Получаем для выбранных изделий все входящие в них объекты
        foreach (ReferenceObject element in listOfReferenceObjects) {
            listAllObjectsOnExport.AddRange(GetAllChildObjects(element));
        }

        string message = string.Empty;
        foreach (ReferenceObject element in listAllObjectsOnExport) {
            message += string.Format(simpleTemplate, element.ToString());
        }
        Message("Список экспортируемых объектов", message);

        listAllObjectsOnExport.AddRange(GetAttachmentFiles(listAllObjectsOnExport));
        //List<ReferenceObject> AllFiles = GetAttachmentFiles(listAllObjectOnExport);
        ExportWithoutLinks(listAllObjectsOnExport);


        Message("Информация", "Работа по экспорту изделий завершена");

    }

    #endregion Method Run

    #region Method ЭкспортироватьВсеВложенияЭкспортФайла

    public void ЭкспортироватьВсеВложенияЭкспортФайла() {

        string pathToExportFile = GetPathToOpenFile();
        if (pathToExportFile == string.Empty)
            return;

        List<string> listOfFilesOnExport = GetAllExportedFiles(pathToExportFile);
        if (listOfFilesOnExport.Count == 0) {
            Message("Информация", "В выбранном экспортируемом файле не были обнаружены объекты файлового справочника для экспорта");
            return;
        }

        ExportWithoutLinks(RelativePathToReferenceObject(listOfFilesOnExport));
        Message("Информация", "Экспорт завершен");

    }

    #endregion Method ЭкспортироватьВсеВложенияЭкспортФайла

    #region Method ЭкспортироватьВсеВложенияИзделия
    
    // Метод, который будет формировать файл эскпорта по всем вложениям, которые относятся к конкретному изделию
    public void ЭкспортироватьВсеВложенияИзделия() {
        

        //TODO Реализовать метод, который получает список всех файлов в виде List<string> с относительными путями файлов

        // 
    }

    #endregion Method ЭкспортироватьВсеВложенияИзделия

    #endregion Entry points


    #region Method GetReferenceObject
    
    // Метод для запроса изделий у пользователя для проведения экспорта
    private List<ReferenceObject> GetReferenceObjects() {
        List<ReferenceObject> result = new List<ReferenceObject>();

        SelectObjectsDialog dialog = CreateSelectObjectsDialog("Электронная структура изделий");
        dialog.MultipleSelect = true;
        if (dialog.Show()) {
            foreach (var refObj in dialog.SelectedObjects) {
                result.Add((ReferenceObject)refObj);
            }
        }

        return result;
    }

    #endregion Method GetReferenceObject

    #region Method GetAllChildObjects
    
    // Метод для рекурсивного получения списка всех входящих объектов в изделие
    private List<ReferenceObject> GetAllChildObjects(ReferenceObject parent) {
        // Метод, который получает все объекты, подключенные к номенклатуре
        List<ReferenceObject> result = parent.Children.RecursiveLoad();
        return result;
    }

    #endregion Method GetAllChildObjects

    #region Method GetAttachmentFiles

    // Метод для получения всех файлов, которые подключены к изделию
    private List<ReferenceObject> GetAttachmentFiles(List<ReferenceObject> nomenclature) {
        List<ReferenceObject> result = new List<ReferenceObject>();
        List<FileObject> files = new List<FileObject>();
        
        foreach (ReferenceObject element in nomenclature) {
            // Производим проверку на то, есть ли у данного объекта связь на файловый справочник
            NomenclatureObject nomObj = element as NomenclatureObject;
            if (nomObj != null) {
                files.AddRange(nomObj.LinkedObject.GetAllLinkedFiles());
            }
        }

        string message = string.Empty;
        foreach (FileObject file in files) {
            result.Add(file as ReferenceObject);
            message += string.Format(simpleTemplate, file[Guids.Properties.RelativePath]);
        }
        Message("Файлы, которые будут выгружены", message);

        return result;
    }

    #endregion Method GetAttachmentFiles

    #region Get list of files from archive with exported data

    // Метод для получения списка файлов из файла-экспорта
    private List<string> GetAllExportedFiles(string pathToExportFile) {
        // Создаем список, в котором будем хранить относительные пути ко всем файлам, которые находятся в файле экспорта
        List<string> listOfFiles = new List<string>();

        // Извлекаем список всех файлов, которые находятся в файле экспорта
        using (ZipArchive archive = ZipFile.Open(pathToExportFile, ZipArchiveMode.Read)) {
            foreach (ZipArchiveEntry entry in archive.Entries) {
                string nameOfFile = HttpUtility.UrlDecode(entry.FullName);
                if (nameOfFile.StartsWith("Files")) {
                    listOfFiles.Add(nameOfFile.Remove(0, 6));
                }
            }
        }

        return listOfFiles;
    }

    #endregion Get list of files from archive with exported data

    #region Export data

    private List<ReferenceObject> RelativePathToReferenceObject(List<string> data) {

        // Создаем новый экземпляр справочника
        FileReference fileReference = new FileReference(Context.Connection);
        List<ReferenceObject> result = new List<ReferenceObject>();
        string pathNotFound = string.Empty;
        
        // Производим поиск по справочнику
        int countFindingFiles = 0;
        int countNotFindingFiles = 0;
        int allFilesCount = data.Count;

        foreach (string filePath in data) {
            string fPath = filePath.Replace("/", "\\");
            ReferenceObject file = fileReference.FindByRelativePath(fPath) as ReferenceObject;

            if (file != null) {
                countFindingFiles++;
                result.Add(file);
            }
            else {
                countNotFindingFiles++;
                pathNotFound += string.Format(simpleTemplate, fPath);
            }
        }

        string template = "Всего файлов на экспорт - {0}\nИз них найдено - {1} шт., не найдено - {2} шт.";
        Message("Информация", string.Format(template, allFilesCount, countFindingFiles, countNotFindingFiles));
        if (pathNotFound != string.Empty)
            Message("Список ненайденных позиций", pathNotFound);

        template = "Начать экспорт {0} найденный позиций из {1} позиций всего?";
        if (!Question(string.Format(template, countFindingFiles, allFilesCount)))
            return null;

        return result;

    }

    // Функция для экспорта файлов
    private void ExportWithoutLinks(List<ReferenceObject>objectToExport) {

        if ((objectToExport == null) | (objectToExport.Count == 0)) {
            Message("Ошибка", "Отсутствуют файлы для экспорта");
            return;
        }
        
        // Написать диалог сохранения файла при помощи WinForms
        string pathToExportFile = GetPathToSaveData();
        if (pathToExportFile == string.Empty) {
            return;
        }

        // Экспорт файлов
        ExportOptions options = new ExportOptions();

        // Определяем параметры для экспорта
        options.FileName = pathToExportFile;
        // Не экспортировать диалоги
        options.DialogsMode = ExportDialogsMode.None;
        // Не экспортировать связанные справочники
        options.LinkedObjectsMode = ExportLinkedObjectsMode.None;
        // Экспортировать только выбранные объекты справочника
        options.ObjectsMode = ExportObjectsMode.Specified;
        // Не экспортировать доступы
        options.IncludeAccesses = false;
        // Не экспортировать связанные макросы
        options.IncludeMacros = false;
        // Не экспортировать прототипы
        options.IncludePrototypes = false;
        // Не экспортировать структуру и типы
        options.IncludeStructure = false;
        // Не экспортировать виды
        options.IncludeViews = false;
        // Упаковать все в один файл
        options.UsePackage = true;

        DataExchangeGateway.Export(Context.Connection, objectToExport, options);
    }

    #endregion Export data

    #region Methods for set path to opening and saving files

    // Метод для выбора файла для его последующего анализа
    private string GetPathToOpenFile() {
        Forms.OpenFileDialog dialog = new Forms.OpenFileDialog();
        dialog.Filter = "export files (*.ddx)|*.ddx|All files (*.*)|*.*";
        dialog.FilterIndex = 1;
        dialog.RestoreDirectory = true;

        if (dialog.ShowDialog() == Forms.DialogResult.OK)
            return dialog.FileName;
        
        return string.Empty;
    }

    // Функция для получения пути сохранения файла
    private string GetPathToSaveData() {
        Forms.SaveFileDialog dialog = new Forms.SaveFileDialog();
        dialog.Filter = "export files (*.ddx)|*.ddx|All files (*.*)|*.*";
        dialog.FilterIndex = 1;
        dialog.RestoreDirectory = true;
        dialog.AddExtension = true;

        if (dialog.ShowDialog() == Forms.DialogResult.OK)
            return dialog.FileName;
        
        return string.Empty;
    }


    #endregion Methods for set path to opening and saving files
}
