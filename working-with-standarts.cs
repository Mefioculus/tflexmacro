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

public class Macro : MacroProvider {
    public Macro(MacroContext context)
        : base (context) {
        }


    public override void Run() {
        ЗагрузитьГосты();
    }

    public void ЗагрузитьГосты() {
        // Данный метод предназначен для добавления кнопки, которая будет производить импорт pdf из выбранной директории

        DocumentRepository repo = new DocumentRepository(this);
        Message("Информация", repo.ToString());
        Message("Информация", repo.ErrorMessage);

    }

    private string[] GetFilesFromUser() {
        // Запросить у пользователя директорию, в которой производить поиск
        string pathToDirectory = @"D:\ГОСТы";
        string searchPattern = "*.pdf";

        return Directory.GetFiles(pathToDirectory, searchPattern, SearchOption.AllDirectories);
    }

    public class RegulatoryDocument {
        // Регулярные выражения
        private static Dictionary<TypeOfDocument, Regex> TypeRegexPatterns = new Dictionary<TypeOfDocument, Regex> () {
            [TypeOfDocument.ГОСТ] = new Regex(@"^[Гг][Оо][Сс][Тт](\s[\w]{1,6})?\s(\d{1,5}[\.\-]){1,3}\d{2,4}\s")
            //[TypeOfDocument.ОСТ] = new Regex(@""),
            //[TypeOfDocument.ТУ] = new Regex(@""),
            //[TypeOfDocument.ПИ] = new Regex(@""),
            //[TypeOfDocument.СТО] = new Regex(@""),
            //[TypeOfDocument.СТП] = new Regex(@""),
            //[TypeOfDocument.Нормали] = new Regex(@""),
            //[TypeOfDocument.Метрология] = new Regex(@""),
        };

        private static Dictionary<string, Regex> AdditionalRegexPatterns = new Dictionary<string, Regex> () {
            ["designation"] = new Regex(@"(\d{1,5}[\.\-]){1,3}\d{2,4}"),
            ["type of document"] = new Regex(@"^([\w]{1,10}\s){1,2}"),
            ["type and designation"] = new Regex(@"^([\w]{1,10}\s){1,2}(\d{1,5}[\.\-]){1,3}\d{2,4}")
        };



        public string Name { get; private set; }
        public string Designation { get; private set; }
        public Dictionary<string, string> AdditionalField { get; private set; }
        public FileInfo LinkedFile { get; private set; }
        public TypeOfDocument Type { get; private set; } = TypeOfDocument.Неизвестно;
        public string TypeString => this.GetStringRepresentationOfType(this.Type);
        public StatusOfDocument Status { get; private set; } = StatusOfDocument.НеОбработан;

        public RegulatoryDocument(string pathToFile, TypeOfDocument type = TypeOfDocument.Неизвестно) {
            // Проверка на то, что файл существует
            if (!File.Exists(pathToFile)) {
                string template =
                    "При инициализации объекта справочника 'Нормативные документы' возникла ошибка\n" +
                    "Исходный файл по пути '{0}' не был обнаружен";
                throw new Exception(string.Format(template, pathToFile));
            }

            this.AdditionalField = new Dictionary<string, string>();

            this.LinkedFile = new FileInfo(pathToFile);
            this.Type = type;

            // Определение и проверка типа документа
            if (this.Type == TypeOfDocument.Неизвестно)
                this.Type = this.TryToDetermineTypeOfDocument();
            else
                CheckType();

            if (this.Type == TypeOfDocument.Неизвестно)
                throw new Exception(string.Format("Не удалось однозначно определить тип документа по названию файла: {0}", this.LinkedFile.Name));

            FillFieldsData(this.Type);
        }
        
        private bool IsTypeFit(TypeOfDocument type) {
            // Проверка на то, что для данного типа есть регулярное выражение
            if (!TypeRegexPatterns.ContainsKey(type))
                return false;

            return TypeRegexPatterns[type].IsMatch(this.LinkedFile.Name);
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
            string fileName = Regex.Replace(LinkedFile.Name, @"\.[pP][dD][fF]$", string.Empty);
            
            // Для начала получаем из названия файла тип документа плюс его обозначение
            Match typeAndDesignationOfDocMatch = AdditionalRegexPatterns["type and designation"].Match(fileName);
            if (!typeAndDesignationOfDocMatch.Success)
                throw new Exception(string.Format("Ошибка при получении типа документа и его обозначения ({0})", fileName));
            this.Name = fileName
                .Replace(typeAndDesignationOfDocMatch.Value, string.Empty)
                .Trim();

            // Получаем обозначение документа
            Match designationMatch = AdditionalRegexPatterns["designation"].Match(fileName);
            if (!designationMatch.Success)
                throw new Exception(string.Format("Ошибка при получении обозначения документа ({0})", fileName));
            this.Designation = designationMatch.Value;

            // Получаем тип объекта
            Match typeMatch = AdditionalRegexPatterns["type of document"].Match(fileName);
            if (!designationMatch.Success)
                throw new Exception(string.Format("Ошибка при получении типа документа ({0})", fileName));
            
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

        public override string ToString() {
            string template = 
                "Тип документа: {0}\n" +
                "Обозначение документа: {1}\n" +
                "Наименование документа: {2}\n\n";
            return string.Format(template, this.TypeString, this.Designation, this.Name);
        }
    }

    public class DocumentRepository {

        public string Dir { get; private set; }
        public string[] Files { get; private set; }
        private string SearchPattern { get; set; }
        private MacroProvider Provider { get; set; }
        private List<RegulatoryDocument> Documents { get; set; }
        private Dictionary<string, List<Exception>> Errors { get; set; }
        public string ErrorMessage => this.GetErrorMessage();

        public DocumentRepository(MacroProvider provider) {

            // Проверяем, был ли передан MacroProvider.
            // Без него не получится общаться с пользователем
            if (provider == null)
                throw new Exception("Для корректной работы макроса в класс DocumentRepository необходимо передать экземпляр класса макроса");
            this.Provider = provider;

            // Инициируем основные коллекции класса
            this.Documents = new List<RegulatoryDocument>();
            this.Errors = new Dictionary<string, List<Exception>>();

            // Запрашиваем файлы у пользователя
            GetInputDataFromUser();
            
            // Производим чтение документов и при возникновении ошибок регистрацию их
            ReadDocuments();


        }

        private void GetInputDataFromUser() {
            // Запросить у пользователя директорию, в которой производить поиск
            this.Dir = @"D:\ГОСТы";
            this.SearchPattern = "*.pdf";

            // TODO: Предусмотреть указание типа заранее
            this.Files = Directory.GetFiles(this.Dir, this.SearchPattern, SearchOption.AllDirectories);
        }
        
        private void ReadDocuments() {
            // TODO: Предусмотреть указание типа заранее
            foreach (string file in this.Files) {
                try {
                    RegulatoryDocument newDoc = new RegulatoryDocument(file);
                    this.Documents.Add(newDoc);
                }
                catch (Exception e) {
                    string fileName = Path.GetFileNameWithoutExtension(file);
                    if (this.Errors.ContainsKey(fileName))
                        this.Errors[fileName].Add(e);
                    else
                        this.Errors[fileName] = new List<Exception>() { e };
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

            string template = "В процессе обработки файлов в директории '{0}' возникли следующие ошибки:\n{1}";
            string innerTemplate = "У файла '{0}':\n{1}\n";
            return string.Format(template, this.Dir, string.Join("\n", this.Errors.Select(kvp => string.Format(innerTemplate, kvp.Key, string.Join("\n", kvp.Value.Select(err => err.Message))))));

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
