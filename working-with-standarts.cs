using System;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Collections;
using System.Collections.Generic;


using TFlex.DOCs.Model.Macros;
using TFlex.DOCs.Model.Macros.ObjectModel;
using TFlex.DOCs.Model.Structure;
using TFlex.DOCs.Model.Search;
using TFlex.DOCs.Model.References;
using TFlex.DOCs.Model.References.Files;

public class Macro : MacroProvider {
    public Macro(MacroContext context)
        : base (context) {
        }


    public override void Run() {
        ЗагрузитьГосты();
        Message("Информация", "Работа макроса завершена");
    }

    public void ЗагрузитьГосты() {
        // Данный метод предназначен для добавления кнопки, которая будет производить импорт pdf из выбранной директории

        DocumentRepository repo = new DocumentRepository(this);
        repo.AskUserAboutFixingErrors();

    }

    public class RegulatoryDocument {
        // Регулярные выражения

        private static class RegexPatterns {
            public static class Types {
                public static Regex ГОСТ = new Regex(@"^[Гг][Оо][Сс][Тт]");
                public static Regex ОСТ = new Regex(@"^[Оо][Сс][Тт]");
                public static Regex ТУ = new Regex(@"^[Тт][Уу]");
                public static Regex ПИ = new Regex(@"^[Пп][Ии]");
                public static Regex СТО = new Regex(@"^[Сс][Тт][Оо]");
                public static Regex СТП = new Regex(@"^[Сс][Тт][Пп]");
                public static Regex Нормали = new Regex(@"[Нн][Оо][Рр][Мм][Аа][Лл]");
                public static Regex Метрология = new Regex(@"[Мм][Ее][Тт][Рр][Оо][Лл][Оо][Гг]");
            }

            public static class Common {
                public static Regex Designation = new Regex(@"(\d{1,5}[\.\-]){1,3}\d{2,4}");
                public static Regex TypeOfDocument = new Regex(@"^([\w]{1,10}\s){1,2}");
                public static Regex TypeAndDesignation = new Regex(@"^([\w]{1,10}\s){1,2}(\d{1,5}[\.\-]){1,3}\d{2,4}");
            }
        }

        private static Dictionary<TypeOfDocument, Regex> TypeRegexPatterns = new Dictionary<TypeOfDocument, Regex> () {
            [TypeOfDocument.ГОСТ] = RegexPatterns.Types.ГОСТ,
            [TypeOfDocument.ОСТ] = RegexPatterns.Types.ОСТ,
            [TypeOfDocument.ТУ] = RegexPatterns.Types.ТУ,
            [TypeOfDocument.ПИ] = RegexPatterns.Types.ПИ,
            [TypeOfDocument.СТО] = RegexPatterns.Types.СТО,
            [TypeOfDocument.СТП] = RegexPatterns.Types.СТП,
            [TypeOfDocument.Нормали] = RegexPatterns.Types.Нормали,
            [TypeOfDocument.Метрология] = RegexPatterns.Types.Метрология
        };



        public string Name { get; private set; }
        public string Designation { get; private set; }
        public Dictionary<string, string> AdditionalField { get; private set; }
        public FileInfo LinkedFile { get; private set; }
        public string FileName { get; private set; }
        public TypeOfDocument Type { get; private set; } = TypeOfDocument.Неизвестно;
        public string TypeString => this.GetStringRepresentationOfType(this.Type);
        public StatusOfDocument Status { get; private set; } = StatusOfDocument.НеОбработан;

        public RegulatoryDocument(string pathToFile, string fileName = null, TypeOfDocument type = TypeOfDocument.Неизвестно) : this(new FileInfo(pathToFile), fileName, type) {
        }

        public RegulatoryDocument(FileInfo file, string fileName = null, TypeOfDocument type = TypeOfDocument.Неизвестно) {

            this.AdditionalField = new Dictionary<string, string>();

            this.LinkedFile = file;
            this.Type = type;
            
            // Если в конструкторе задано имя, отличное от имени файла, тогда присваиваем его
            if (fileName != null)
                this.FileName = fileName;
            else
                this.FileName = this.LinkedFile.Name;

            // Определение и проверка типа документа
            if (this.Type == TypeOfDocument.Неизвестно)
                this.Type = this.TryToDetermineTypeOfDocument();
            else
                CheckType();

            if (this.Type == TypeOfDocument.Неизвестно)
                throw new Exception("Не удалось определить тип документа по названию файла");

            FillFieldsData(this.Type);

            // TODO: Производим смену файла в соответствии с извлеченными из файла мета-данными
            //RenameDocumentFile();
        }
        
        private bool IsTypeFit(TypeOfDocument type) {
            // Проверка на то, что для данного типа есть регулярное выражение
            if (!TypeRegexPatterns.ContainsKey(type))
                return false;

            return TypeRegexPatterns[type].IsMatch(this.FileName);
        }

        private TypeOfDocument TryToDetermineTypeOfDocument() {
            List<TypeOfDocument> types = new List<TypeOfDocument>();

            foreach (TypeOfDocument type in Enum.GetValues(typeof(TypeOfDocument))) {
                if (this.IsTypeFit(type))
                    types.Add(type);
            }

            // Проверка полученного результата перед возвратом
            if (types.Count == 1)
                return types[0];

            return TypeOfDocument.Неизвестно;
        }

        private void CheckType() {
            if (!this.IsTypeFit(this.Type)) {
                this.Type = TypeOfDocument.Неизвестно;
                this.Status = StatusOfDocument.НетПодходящегоТипа;
            }
        }

        private void FillFieldsData(TypeOfDocument type) {
            // В зависимости от типа документа производим различные манипуляции
            switch (type) {
                case TypeOfDocument.ГОСТ:
                    FillFieldsDataForGost();
                    break;
                default:
                    throw new Exception(string.Format("Для типа {0} еще не написана обработка заполнения параметров", type.ToString()));
            }
        }

        private void FillFieldsDataForGost() {
            string fileName = Regex.Replace(this.FileName, @"\.[pP][dD][fF]$", string.Empty);
            
            // Для начала получаем из названия файла тип документа плюс его обозначение
            Match typeAndDesignationOfDocMatch = RegexPatterns.Common.TypeAndDesignation.Match(fileName);
            if (!typeAndDesignationOfDocMatch.Success)
                throw new Exception("Ошибка при получении типа документа и его обозначения");
            this.Name = fileName
                .Replace(typeAndDesignationOfDocMatch.Value, string.Empty)
                .Trim();
            if (string.IsNullOrWhiteSpace(this.Name))
                throw new Exception("Отсутствует название ГОСТ");

            // Получаем обозначение документа
            Match designationMatch = RegexPatterns.Common.Designation.Match(fileName);
            if (!designationMatch.Success)
                throw new Exception("Ошибка при получении обозначения документа");
            this.Designation = designationMatch.Value;

            // Получаем тип объекта
            Match typeMatch = RegexPatterns.Common.TypeOfDocument.Match(fileName);
            if (!designationMatch.Success)
                throw new Exception("Ошибка при получении типа документа");
            
            // Проверяем, есть ли у данного ГОСТа дополнительный тип
            string[] wordsInType = typeMatch.Value.Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
            this.AdditionalField["Тип ГОСТа"] = wordsInType.Length == 2 ? wordsInType[1] : string.Empty;
            
        }

        private string GetStringRepresentationOfType(TypeOfDocument type) {
            switch (type) {
                case TypeOfDocument.Неизвестно:
                    return "Неизвестно";
                case TypeOfDocument.ГОСТ:
                    return "ГОСТ";
                case TypeOfDocument.ОСТ:
                    return "ОСТ";
                case TypeOfDocument.СТО:
                    return "СТО";
                case TypeOfDocument.СТП:
                    return "СТП";
                case TypeOfDocument.ПИ:
                    return "ПИ";
                case TypeOfDocument.ТУ:
                    return "ТУ";
                case TypeOfDocument.Нормали:
                    return "Нормаль";
                case TypeOfDocument.Метрология:
                    return "Метрологический документ";
                default:
                    throw new Exception(string.Format("Переданный в функцию GetStringRepresentationOfType тип - {0} является неизвестным", type.ToString()));
            }
        }

        public void CreateInDocs(Reference nomenclatureReference, FileReference fileReference, FolderObject folder) {
            // TODO: Реализовать метод создания документа в DOCs
        }

        private FileObject UploadFile(FileReference fileReference, FolderObject folder) {
            // TODO: Реализовать метод загрузки файла в файловый справочник DOCs
            //
            // Для начала проверяем, нет ли данного файла в файловом справочнике. Если есть - выводим ошибку
            //
            // Производим загрузку файла
            //
            // Производим переименование файла в соответствии с именем файла, которое указано в поле FileName
            return null;
        }

        private ReferenceObject CreateRecordInReference(Reference nomenclatureReference) {
            // TODO: Реализовать метод создания записи в справочнике нормативных документов
            return null;
        }

        public override string ToString() {
            string template = 
                "Тип документа: {0}\n" +
                "Обозначение документа: {1}\n" +
                "Наименование документа: {2}\n\n";
            return string.Format(template, this.TypeString, this.Designation, this.Name);
        }
    }

    public class DocumentRepository {

        // Данные, получаемые от пользователя
        public string Dir { get; private set; }
        private string SearchPattern { get; set; }
        private TypeOfDocument SearchType { get; set; } = TypeOfDocument.Неизвестно;

        // Данные, получаемые при инициализации нового объекта
        private MacroProvider Provider { get; set; }
        
        // Данные, получаемые в процессе работы конструктора
        public string[] Files { get; private set; }
        private List<RegulatoryDocument> Documents { get; set; }
        private Dictionary<FileInfo, Exception> Errors { get; set; }
        public string ErrorMessage { get; private set; }

        public DocumentRepository(MacroProvider provider) {

            // Проверяем, был ли передан MacroProvider.
            // Без него не получится общаться с пользователем
            if (provider == null)
                throw new Exception("Для корректной работы макроса в класс DocumentRepository необходимо передать экземпляр класса макроса");
            this.Provider = provider;

            // Инициируем основные коллекции класса
            this.Documents = new List<RegulatoryDocument>();
            this.Errors = new Dictionary<FileInfo, Exception>();

            // Запрашиваем файлы у пользователя
            GetInputDataFromUser();
            
            // Производим чтение документов и при возникновении ошибок регистрацию их
            ReadDocuments();

            this.ErrorMessage = GetErrorMessage();
        }

        private void GetInputDataFromUser() {
            // Запросить у пользователя директорию, в которой производить поиск
            // TODO: После тестирования убрать введенный по умолчанию путь (или установить его на рабочий стол пользователя)
            this.Dir = @"D:\ГОСТы";
            //this.Dir = Environment.GetFolderName(Environment.SpecialFolder.Desktop);
            this.SearchPattern = "*.pdf";

            string directory = "Директория";
            string type = "Тип";
            string browse = "Обзор";
            string pattern = "Поисковый запрос";

            InputDialog dialog = new InputDialog(this.Provider.Context, "Укажите директорию");
            dialog.AddString(directory, this.Dir);
            string[] types = Enum.GetNames(typeof(TypeOfDocument));
            dialog.AddButton(
                    browse,
                    (name) => {
                        OpenFolderDialog folderDialog = new OpenFolderDialog(this.Provider.Context, "Выберите директорию");
                        if (folderDialog.Show())
                            dialog[directory] = folderDialog.DirectoryName;
                        },
                    false);
            dialog.AddString(pattern, this.SearchPattern);
            dialog.AddSelectFromList(type, types[0], true, types);
            if (dialog.Show()) {
                this.Dir = (string)dialog[directory];
                this.SearchPattern = (string)dialog[pattern];
                this.SearchType = (TypeOfDocument)Enum.Parse(typeof(TypeOfDocument), (string)dialog[type]);
            }

            this.Files = Directory.GetFiles(this.Dir, this.SearchPattern, SearchOption.AllDirectories);
        }
        
        private void ReadDocuments() {
            // TODO: Предусмотреть указание типа заранее
            foreach (string file in this.Files) {
                try {
                    RegulatoryDocument newDoc = new RegulatoryDocument(file, null, this.SearchType);
                    this.Documents.Add(newDoc);
                }
                catch (Exception exception) {
                    FileInfo fileInfo = new FileInfo(file);
                    this.Errors[fileInfo] = exception;
                }
            }
        }

        public override string ToString() {
            string template =
                "Всего файлов: {0}\n" +
                "Инициализировано документов: {1}\n" +
                "Ошибок в процессе обработки документов: {2}\n";
            return string.Format(template, this.Files.Length, this.Documents.Count, this.Errors.Count);
        }

        private string GetErrorMessage() {
            if (this.Errors.Count == 0)
                return string.Empty;

            string template = "В процессе обработки файлов в директории '{0}' возникли следующие ошибки:\n\n{1}";
            string innerTemplate = "файл '{0}': {1}\n";
            return string.Format(template, this.Dir, string.Join("\n", this.Errors.Select(kvp => string.Format(innerTemplate, kvp.Key.Name, kvp.Value.Message))));
        }

        public void RefreshErrorMessage() {
            this.ErrorMessage = GetErrorMessage();
        }

        public void AskUserAboutFixingErrors() {
            // TODO: Реализовать запрос у пользователя исправления ошибочных названий документов
            if (this.Errors.Count == 0)
                return;
            string template =
                "В директории '{0}' было обнаружено {1} файлов, подходящих под условия поиска. Среди них:\n" +
                "Успешно обработано: {2}\n" +
                "Обработано с ошибкой: {3}\n\n" +
                "Если хотите произвести корректировку сразу, выберите файлы, которые хотите исправить и нажмите 'ОК'\n" +
                "Если вы не хотите производить корректировку названий документов - нажмите 'Отмена'\n" +
                "Тогда будут загружены только те документы, который прошли проверку, остальные документы будут проигнорированы\n";
            string errors = "Файлы c ошибками";
            string quantityOfErrorsOnPage = "Количество одновременно корректируемых записей";
            
            InputDialog dialog = new InputDialog(this.Provider.Context, "Исправление названий файлов");
            dialog.AddText(string.Format(template, this.Dir, this.Files.Length, this.Documents.Count, this.Errors.Count));
            dialog.AddMultiselectFromList(errors, this.Errors.Select(kvp => kvp.Key.Name), true);
            dialog.AddInteger(quantityOfErrorsOnPage, 5);

            // Отображаем диалог
            if (dialog.Show()) {
                // Получаем список файлов на обработку

                // Так как список, получаемый из диалога является динамическим, для того, чтобы произвести
                // все необходимые преобразования, для начала нужно привести его к динамическому списку
                // (иначе не получится воспользоваться Linq методами
                List<string> files = ((IEnumerable<dynamic>)dialog[errors]).Cast<string>().ToList<string>();

                

                this.Provider.Message(string.Empty, string.Join("\n", files));
                // Запускаем корректировку файлов

                List<FileInfo> correctedFiles = FixErrors(
                        this.Errors
                            .Select(kvp => kvp.Key)
                            .Where(fileInfo => files.Contains(fileInfo.Name))
                            .ToList<FileInfo>(),
                        (int)dialog[quantityOfErrorsOnPage]
                        );
                // TODO: Реализовать повторную попытку инициализации документов для исправленных файлов
            }
        }

        private List<FileInfo> FixErrors(List<FileInfo> files, int quantityOnPage) {
            // TODO: Реализовать метод, который будет постранично запрашивать у пользователя исправления файлов
            // Будет создавать в временной папке данные файлы, и возвращать для них список объектов FileInfo
            List<FileInfo> result = new List<FileInfo>();
            return result;
        }
    }

    public enum TypeOfDocument {
        Неизвестно,
        ГОСТ,
        ОСТ,
        СТО,
        СТП,
        ПИ,
        ТУ,
        Нормали,
        Метрология
    }

    public enum StatusOfDocument {
        НеОбработан,
        НетПодходящегоТипа,
        НесколькоПодходящихТипов,
        Обработан
    }

}
