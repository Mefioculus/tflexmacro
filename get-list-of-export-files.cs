using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
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
using Newtonsoft.Json; // Для сохранения данных в строку

// Для работы макроса так же необходимо подключить следующие библиотеки
// System.Web.dll
// System.IO.Compression.dll
// System.IO.Compression.ZipFile.dll
// System.IO.Compression.FileSystem.dll

// А так же подкючить ссылку на
// Newtonsoft.Json.dll

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
    // Поле, которое будет хранить логи экспорта
    private List<ExportLog> listOfExportedLogs = null;

    #endregion Fields and Properties

    #region Entry points

    #region Method Run

    public override void Run() {
    }

    #endregion Method Run

    #region Method ЭкспортироватьВсеВложенияЭкспортФайла

    public void ЭкспортироватьВсеВложенияЭкспортФайла() {

        string pathToExportFile = GetPathToOpenFile("ddx");
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
        // Получаем от пользователя список изделий, для которого нужно произвести экспорт
        List<ReferenceObject> listOfInitialObject = GetReferenceObjects();
        if (listOfInitialObject.Count == 0)
            return;

        // Инициалилзируем список, который будет хранить все объекты, для которых необхоимо произвести экспорт
        List<ReferenceObject> listAllObjectsOnExport= new List<ReferenceObject>();

        // Получаем для выбранных изделий все входящие в них объекты
        listOfExportedLogs = new List<ExportLog>();
        foreach (ReferenceObject element in listOfInitialObject) {
            List<ReferenceObject> allElements = GetAllChildObjects(element);

            // Ведение лога
            ExportLog exportLogOfProduct = new ExportLog();
            exportLogOfProduct.NameOfProduct = element.ToString();
            // Готовим информацию о составе экспортируемых данных
            foreach (ReferenceObject refObj in allElements) {
                exportLogOfProduct.Add(
                        refObj.ToString(),
                        refObj.SystemFields.Guid,
                        refObj.Parent.ToString(),
                        refObj.Parent.SystemFields.Guid);
            }
            listOfExportedLogs.Add(exportLogOfProduct);
            // Конец ведения лога

            // Добавляем элементы в список на экспорт
            listAllObjectsOnExport.AddRange(allElements);
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

    #endregion Method ЭкспортироватьВсеВложенияИзделия

    #region Method ПроверитьЭкспортируемуюСтруктуру

    public void ПроверитьЭкспортируемуюСтруктуру() {
        // Для начала получаем объект, для которого будем проводить проверку
        ReferenceObject product = GetReferenceObject();
        if (product == null) {
            Message("Ошибка", "Не выбрано изделия для проведения сравнения структуры");
            return;
        }
        // Получаем все входящие элементы в текущий объект
        List<ReferenceObject>listOfAllObjects = GetAllChildObjects(product);

        // Далее, открываем сохраненную структуру, с которой будем проводить сравнение
        string pathToLogFile = GetPathToOpenFile("txt");
        ExportLog logOfProduct = ReadLogs(pathToLogFile);
        string message = logOfProduct.Compare(listOfAllObjects);
        Message("Информация", message);
    }

    #endregion Method ПроверитьЭкспортируемуюСтруктуру

    #endregion Entry points

    #region Method GetReferenceObject

    // Метод для запроса изделия у пользователя для проведения сравнения состава
    private ReferenceObject GetReferenceObject() {
        ReferenceObject result = null;

        SelectObjectDialog dialog = CreateSelectObjectsDialog("Электронная структура изделий");
        dialog.MultipleSelect = true;
        if (dialog.Show()) {
            result = dialog.SelectedObjects[0];
        }
        return result;
    }
    
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
        string pathToExportFile = GetPathToSaveData("ddx");
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

        // Сохранение логов экспорта
        if (listOfExportedLogs != null) {
            foreach (var product in listOfExportedLogs) {
                WriteLogs(product, pathToExportFile);
            }
        }
    }

    #endregion Export data

    #region Methods for read and write logs

    private void WriteLogs(ExportLog productLog, string pathToExportFile) {
        string jsonString = JsonConvert.SerializeObject(productLog);

        // Получаем директорию для сохранения и название экспортируемого файла
        string pathToDirectory = Path.GetDirectoryName(pathToExportFile);
        string nameOfExportFile = Path.GetFileNameWithoutExtension(pathToExportFile);

        string nameOfLogFile = string.Format("{0} ({1}).txt", productLog.NameOfProduct, nameOfExportFile);
        string pathToLogFile = Path.Combine(pathToDirectory, nameOfLogFile);

        File.WriteAllText(pathToLogFile, jsonString);
    }

    private ExportLog ReadLogs (string pathToFile) {
        string jsonString = File.ReadAllText(pathToFile);

        return JsonConvert.DeserializeObject<ExportLog>(jsonString);
    }

    #endregion Methods for read and write logs

    #region Methods for set path to opening and saving files

    // Метод для выбора файла для его последующего анализа
    private string GetPathToOpenFile(string typeFile) {
        Forms.OpenFileDialog dialog = new Forms.OpenFileDialog();
        dialog.Filter = "export files (*.ddx)|*.ddx|All files (*.*)|*.*";
        
        switch (typeFile) {
            case "ddx":
                dialog.FilterIndex = 1;
                break;
            case "txt":
                dialog.FilterIndex = 2;
                break;
            default:
                dialog.FilterIndex = 3;
                break;
        }

        dialog.RestoreDirectory = true;

        if (dialog.ShowDialog() == Forms.DialogResult.OK)
            return dialog.FileName;
        
        return string.Empty;
    }

    // Функция для получения пути сохранения файла
    private string GetPathToSaveData(string typeFile) {
        Forms.SaveFileDialog dialog = new Forms.SaveFileDialog();
        dialog.Filter = "export files (*.ddx)|*.ddx|text files (*.txt)|*.txt|All files (*.*)|*.*";

        switch (typeFile) {
            case "ddx":
                dialog.FilterIndex = 1;
                break;
            case "txt":
                dialog.FilterIndex = 2;
                break;
            default:
                dialog.FilterIndex = 3;
                break;
        }

        dialog.RestoreDirectory = true;
        dialog.AddExtension = true;

        if (dialog.ShowDialog() == Forms.DialogResult.OK)
            return dialog.FileName;
        
        return string.Empty;
    }


    #endregion Methods for set path to opening and saving files

    #region Service classes

    // Класс для хранения данных об экспортированной структуре для последующей проверки структуры на
    // полноту проведения экспорта
    private class ExportLog {
        public string NameOfProduct { get; set; }
        public List<LogElement> listOfElements = new List<LogElement>();

        public LogElement this[int index] {
            get {
                return this.listOfElements[index];
            }
            set {
                this.listOfElements[index] = value;
            }
        }

        public void Add(string name, Guid guid, string parentName, Guid parentGuid) {
            listOfElements.Add(new LogElement(name, guid, parentName, parentGuid));
        }

        public string Compare(List<ReferenceObject> targetStructure) {

            string resultMessage = string.Empty;
            int indexOnRemove = -1;

            foreach (LogElement element in this.listOfElements) {
                for (int index = 0; index < targetStructure.Count; index++) {
                    // Производим поиск элемента
                    if (element.GuidOfElement == targetStructure[index].SystemFields.Guid) {
                        // Элемент найден, его можно удалить
                        indexOnRemove = index;
                        break;
                    }
                }

                if (indexOnRemove >= 0) {
                    targetStructure.RemoveAt(indexOnRemove);
                    indexOnRemove = -1;
                }
                else {
                    // Обработка случая, когда у нас не получилось найти элемент изначальной структуры в перенесенной структуре
                    if (resultMessage == string.Empty)
                        resultMessage += "В экспортированной структуре отсутствуют следующие позиции:\n";

                    resultMessage += string.Format("----\nName: {0}\nGuid: {1}\nName of Parent: {2}\nGuid of Parent: {3}\n",
                                                    element.Name,
                                                    element.GuidOfElement.ToString(),
                                                    element.ParentName,
                                                    element.ParentGuid.ToString());
                }
            }

            if (resultMessage == string.Empty)
                resultMessage += "В экспортированной структуре есть все элементы из структуры источника";

            // По логике в списке targetSutructure к данному моменту не должно остаться позиций, так как если они останутся, это будет означать, что в
            // перенесенной структуре появились элементы, которых не было в изначальной структуре.
            // Но на всякий случай обработать этот вариант так же необходимо.
            
            if (targetStructure.Count != 0) {
                resultMessage += "\n\nВ экспортированной структуре обнаружены позиции, которых не было в структуре источника";
                foreach(ReferenceObject refObj in targetStructure) {
                    resultMessage += string.Format("----\nName: {0}\nGuid: {1}\nName of Parent: {2}\nGuid of Parent: {3}\n",
                                                    refObj.ToString(),
                                                    refObj.SystemFields.Guid.ToString(),
                                                    refObj.Parent.ToString(),
                                                    refObj.Parent.SystemFields.Guid.ToString());
                }
            }
            else
                resultMessage += "\n\nВ экспортированной структуре отсутствуют позиции, которых не было в структуре источнике";

            return resultMessage;
        }
    }

    private class LogElement {
        public string Name { get; set; }
        public Guid GuidOfElement { get; set;}

        public string ParentName { get; set; }
        public Guid ParentGuid { get; set; }

        public LogElement (string name, Guid guid, string parentName, Guid parentGuid) {
            this.Name = name;
            this.GuidOfElement = guid;
            this.ParentName = parentName;
            this.ParentGuid = parentGuid;
        }
    }


    #endregion Service classes
}
