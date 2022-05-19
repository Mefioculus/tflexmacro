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

    private static class Guids {

        public static class References {
            public static Guid Файлы = new Guid("a0fcd27d-e0f2-4c5a-bba3-8a452508e6b3");
            public static Guid НормативныеДокументы = new Guid("221ea415-75fc-458a-aa52-2144225fca43");
        }

        public static class Types {
            public static Guid НормативныйДокумент = new Guid("37ef8098-14a7-4787-8814-14c2eb1b5b6a");
        }

        public static class Objects {
            public static Guid ПапкаАрхивНД = new Guid("3d33548b-3366-4fb6-8126-bce53b0a7d68");
        }

        public static class Parameters {
            public static Guid ОбозначениеДокумента = new Guid("be6cbd27-113a-4e88-aab1-0b8540480603");
            public static Guid НаименованиеДокумента = new Guid("efb1bd41-3cf9-434c-a2b9-a814376da604");
            public static Guid ТипДокумента = new Guid("37d4ad41-424c-4a09-b17d-65356216f59b"); // Общий параметр для всех документов
            public static Guid ТипГоста = new Guid("3f43523e-0ef1-46ad-97bb-6b1ce3e053a4"); // Данный параметр есть только для ГОСТов
            public static Guid НаименованиеФайла = new Guid("63aa0058-4a37-4754-8973-ffbc1b88f576");
        }

        public static class Links {
            public static Guid ФайлНормативногоДокумента = new Guid("04f51b77-d2f8-448d-8778-df3bb97ea625");
        }
    }

    public override void Run() {
        ЗагрузитьДокументы();
        Message("Информация", "Работа макроса завершена");
    }

    public void ЗагрузитьДокументы() {
        // Данный метод предназначен для добавления кнопки, которая будет производить импорт pdf из выбранной директории

        DocumentRepository repo;

        try {
            repo = new DocumentRepository(this);
        }
        catch {
            return;
        }
        // Запрашиваем у пользователя исправления найденных ошибок
        repo.AskUserAboutFixingErrors();
        // Производим процедуру загрузки документов в DOCs
        repo.UploadDocumentsInDocs();

    }

    public class RegulatoryDocument {

        // Поля, относящиеся к стандарту
        public string Name { get; private set; }
        public string Designation { get; private set; }
        public string TflexDesignation => !string.IsNullOrWhiteSpace(this.AdditionalType) ?
            $"{this.GetStringRepresentationOfType(this.Type)} {this.AdditionalType} {this.Designation}" :
            $"{this.GetStringRepresentationOfType(this.Type)} {this.Designation}";
        public string AdditionalType { get; private set; }

        // Поля, относящиеся к файлам (source file и destination file)
        public FileInfo LinkedFile { get; private set; }
        public string FileName { get; private set; }
        public string FileExtension { get; private set; }
        public string FullFileName => this.FileName + this.FileExtension;

        // Поля, относящиеся к объектам T-Flex DOCs
        public ReferenceObject NormativeReferenceObject { get; private set; }
        public FileObject FileReferenceObject { get; private set; }

        public DocumentRepository ParentRepository { get; private set; }

        // Служебные поля: Тип
        public TypeOfDocument Type { get; private set; } = TypeOfDocument.Неизвестно;
        public string TypeString => this.GetStringRepresentationOfType(this.Type);

        // Флаги:
        public bool IsDocumentSearched { get; private set; } = false;
        public bool IsFileSearched { get; private set; } = false;
        public bool IsExcluded { get; private set; } = false;
        public bool? IsDocumentExist { get; private set; }
        public bool? IsFileExist { get; private set; }

        // Служебные поля: Статус
        // TODO: Решить, нужно ли это поле
        public StatusOfDocument Status { get; private set; } = StatusOfDocument.НеОбработан;
        
        // Служебные поля: Перехваченная ошибка
        public Exception Error { get; private set; }
        public bool HasError => this.Error != null;
        public string ErrorMessage => this.HasError ? this.Error.Message : "Ошибка отсутствует";

        public RegulatoryDocument(string file, DocumentRepository repository) {
            // Заполняем ссылку на родительский репозиторий
            this.ParentRepository = repository;
            
            // Пытаемся получить доступ к файлу
            this.LinkedFile = new FileInfo(file);

            // Получаем название документа и его расширение
            this.FileExtension = this.LinkedFile.Extension;
            this.FileName = this.LinkedFile.Name.Replace(this.FileExtension, string.Empty);

            try {
                InitializeObject();
            }
            catch (Exception e) {
                this.SetError(e);
            }
        }

        private void SetError(Exception exception) {
            this.Error = exception;
        }

        private void ClearError() {
            this.Error = null;
        }

        private void InitializeObject(string newFileName = null) {
            // Определяем, под каким именем в T-Flex DOCs будет сохранен файл
            if (newFileName != null)
                this.FileName = newFileName;

            // Производим проверку типа
            if (this.Type == TypeOfDocument.Неизвестно)
                this.Type = this.TryToDetermineTypeOfDocument();
            else
                this.CheckType();

            if (this.Type == TypeOfDocument.Неизвестно)
                throw new Exception("Неизвестный тип");
            
            // Пробуем распрарсить название файла и заполнить полученными данными поля объекта
            this.FillFieldsData(this.Type);
        }

        public void ReinitializeObject(string newFileName) {
            // Если пользователь передал такое же название, реинициализация объекта производиться не будет, так как в этом нет смысла
            if (newFileName == this.FileName)
                return;

            // Если пользователем предпринимается попытка присвоить старое старое или нулевое обозначение, сразу выдавать ошибку
            if (string.IsNullOrWhiteSpace(newFileName))
                throw new Exception("Название документа не может отсутствовать или состоять из пустых символов");

            // Убираем ошибку в начале реинициализации
            this.ClearError();
            
            // Пробуем произвести повторную инициализацию объекта
            try {
                this.InitializeObject(newFileName);
            }
            catch (Exception exception) {
                this.SetError(exception);
                return;
            }
        }
        
        private bool IsTypeFit(TypeOfDocument type) {
            return ParentRepository.Patterns.IsTypeRegexExist(type) ? ParentRepository.Patterns.GetTypeRegex(type).IsMatch(this.FileName) : false;
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
                    FillFieldsDataForType();
                    break;
                case TypeOfDocument.ОСТ:
                    FillFieldsDataForType();
                    break;
                case TypeOfDocument.ТУ:
                    FillFieldsDataForType();
                    break;
                default:
                    throw new Exception(string.Format("Тип {0} пока не поддерживается", type.ToString()));
            }
        }

        private void FillFieldsDataForType() {
            string fileName = this.FileName.TrimEnd(new char[] { '.' });

            // Для начала пытаемся получить тип документа из названия файла
            Match typeMatch = ParentRepository.Patterns.GetTypeRegex(this.Type).Match(fileName);
            if (!typeMatch.Success)
                throw new Exception("Отсутствует тип");
            
            // Заполняем необходимые поля
            string[] wordsInType = typeMatch.Value.Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
            this.AdditionalType = wordsInType.Length != 1 ? string.Join(" ", wordsInType.Skip(1)) : string.Empty;

            fileName = ParentRepository.Patterns.GetTypeRegex(this.Type).Replace(fileName, string.Empty).Trim();

            // Пытаемся получить обозначение документа из названия файла

            Match designationMatch = ParentRepository.Patterns.GetDesignationRegex(this.Type).Match(fileName);
            if (!designationMatch.Success)
                throw new Exception("Отсутствует обозначение");
            this.Designation = designationMatch.Value;
            fileName = ParentRepository.Patterns.GetDesignationRegex(this.Type).Replace(fileName, string.Empty).Trim();

            // Пробуем получить название документа
            if (string.IsNullOrWhiteSpace(fileName))
                throw new Exception("Отсутствует наименование");
            // Производим очистку от возможных лишних пробелов в наименовании
            this.Name = string.Join(" ", fileName.Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries));

            // Обновляем имя файла из произведенного разбора
            this.FileName = $"{typeMatch.Value} {this.Designation} {this.Name}";
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

        public void ProcessDocument() {
            // TODO: Реализовать метод создания документа в DOCs

            // Производим поиск документа в TFlex (или создание нового документа, если найти ничего не удалось)
            this.NormativeReferenceObject = FindOrCreateNormRecordInDocs();
            if (this.NormativeReferenceObject == null)
                return;

            // Производим поиск файла в T-Flex (или создание файла, если найти ничего не удалось)
            this.FileReferenceObject = FindOrUploadFileObjectInDocs();

        }

        private ReferenceObject FindOrCreateNormRecordInDocs() {
            // Производим поиск по обозначению
            ParameterInfo designation = ParentRepository.NDReference.ParameterGroup[Guids.Parameters.ОбозначениеДокумента];
            if (designation == null)
                throw new Exception("Ошибка при проведении поиска в справочнике Нормативных докумнетов");

            List<ReferenceObject> findedObjects = ParentRepository.NDReference.Find(designation, this.TflexDesignation);
            this.IsDocumentSearched = true;

            switch (findedObjects.Count) {
                case 0:
                    this.IsDocumentExist = false;
                    return CreateNormRecordInDocs();
                case 1: 
                    this.IsDocumentExist = true;
                    if (!this.ParentRepository.ProcessExistingDocuments)
                        this.IsExcluded = true;
                    return this.ParentRepository.ProcessExistingDocuments ? findedObjects[0] : null;
                default:
                    this.IsDocumentExist = true;
                    if (!this.ParentRepository.ProcessExistingDocuments)
                        this.IsExcluded = true;
                    return this.ParentRepository.ProcessExistingDocuments ? AskUserSelectFindedObject(findedObjects) : null;
            }
        }

        private ReferenceObject CreateNormRecordInDocs() {
            ReferenceObject newObject = ParentRepository.NDReference.CreateReferenceObject();
            newObject[Guids.Parameters.ОбозначениеДокумента].Value = this.TflexDesignation;
            newObject[Guids.Parameters.НаименованиеДокумента].Value = this.Name;
            newObject[Guids.Parameters.ТипДокумента].Value = this.ParentRepository.TypesToIntDict[this.Type];
            if (this.Type == TypeOfDocument.ГОСТ) {
                newObject[Guids.Parameters.ТипГоста].Value = this.GetTypeOfGost();
            }
            newObject.EndChanges();
            return newObject;
        }

        private int? GetTypeOfGost() {
            if (this.Type != TypeOfDocument.ГОСТ)
                throw new Exception($"Метод {nameof(GetTypeOfGost)} можно использовать только для типа документа 'ГОСТ'");

            switch (this.AdditionalType) {
                case "EN":
                    return 0;
                case "ISO":
                    return 1;
                case "Р":
                    return 2;
                case "Р ЕН":
                    return 3;
                case "Р ИСО":
                    return 4;
                case "Р МЭК":
                    return 5;
                case "РВ":
                    return 6;
                case "В":
                    return 7;
                default:
                    return null;
            }
        }

        private ReferenceObject AskUserSelectFindedObject(List<ReferenceObject> resultOfSearch) {
            if (resultOfSearch.Count == 0)
                throw new Exception("Количество объектов для выбора не может быть нулевым");

            string message = $"При обработке записи {this.TflexDesignation} было обнаружено несколько совпадений:\n\n";

            foreach (ReferenceObject findedRecord in resultOfSearch) {
                ReferenceObject file = findedRecord.GetObject(Guids.Links.ФайлНормативногоДокумента);
                string fileString = file != null ? file.ToString() : "отсутствует";
                message +=
                    $"ID: {findedRecord.Id.ToString()}\nОбозначение: {findedRecord[Guids.Parameters.ОбозначениеДокумента].Value.ToString()}\n" +
                    $"Дата создания: {findedRecord.SystemFields.CreationDate.ToString("dd.MM.yyyy")}\n" +
                    $"Связанный файл: {fileString}\n" +
                    "\n";
            }

            message += "Произведите выбор объекта по его ID, или отмените выбор";

            InputDialog dialog = new InputDialog(ParentRepository.Provider.Context, "Произведите выбор объекта");
            string selectField = "ID:";
            dialog.AddText(message);
            List<int> IDs = resultOfSearch.Select(rec => rec.Id).ToList<int>();
            dialog.AddSelectFromList(selectField, IDs[0], true, IDs);

            return dialog.Show() ? resultOfSearch.FirstOrDefault(rec => rec.Id == (int)dialog[selectField]) : null;
        }

        private FileObject FindOrUploadFileObjectInDocs() {
            if (this.IsDocumentSearched == false)
                throw new Exception($"Метод {nameof(this.FindOrUploadFileObjectInDocs)} не может вызываться до метода {nameof(this.FindOrCreateNormRecordInDocs)}");
            // TODO: Реализовать метод загрузки файла в файловый справочник DOCs

            // Если для данного госта уже есть запись в нормативном справочнике, проверяем, есть ли у него привязанный файл
            FileObject resultFile = (FileObject)this.NormativeReferenceObject.GetObject(Guids.Links.ФайлНормативногоДокумента);

            if (resultFile != null) {
                // Файл существует. Нужно сравнить его с загружаемым и произвести выбор, если есть различия
                CompareAndPick(resultFile);
            }
            else {
                // Создаем новый файл на основе загружаемой pdf
                resultFile = CreateFile();
            }
            return resultFile;
        }

        private FileObject CreateFile(bool setLink = true) {
            // TODO: Реализовать функцию создания файла
            // Создаем файл
            FileObject newFile = ParentRepository.FReference.AddFile(this.LinkedFile.FullName, ParentRepository.NDFolder);
            
            newFile.BeginChanges();
            // Переименовываем файл (this.FullFileName)
            newFile[Guids.Parameters.НаименованиеФайла].Value = this.FullFileName;
            
            // Производим привязывание файла к нормативному документу
            if (setLink == true) {
                this.NormativeReferenceObject.BeginChanges();
                this.NormativeReferenceObject.SetLinkedObject(Guids.Links.ФайлНормативногоДокумента, (ReferenceObject)newFile);
                this.NormativeReferenceObject.EndChanges();
            }

            newFile.EndChanges();
            //newFile.CheckIn();

            return newFile;
        }

        private void CompareAndPick(FileObject linkedFileObject) {
            linkedFileObject.GetHeadRevision();
            FileInfo oldFile = new FileInfo(linkedFileObject.LocalPath);

            int oldFileByte;
            int newFileByte;
            bool isFilesEqual = true;
            
            // Производим сравнение файлов
            if (this.LinkedFile.Length == oldFile.Length) {
                // Если размер файлов одинаковый, сравниваем файлы побайтово
                FileStream fsOld = new FileStream(oldFile.FullName, FileMode.Open, FileAccess.Read);
                FileStream fsNew = new FileStream(this.LinkedFile.FullName, FileMode.Open, FileAccess.Read);

                while (true) {
                    oldFileByte = fsOld.ReadByte();
                    newFileByte = fsNew.ReadByte();

                    // Если мы достигли конца файла, разрываем цикл
                    if (oldFileByte == -1)
                        break;
                    
                    // Если байты не равны, прерываем цикл
                    if (oldFileByte != newFileByte) {
                        isFilesEqual = false;
                        break;
                    }

                }
                fsOld.Close();
                fsNew.Close();
            }
            else {
                isFilesEqual = false;
            }

            if (isFilesEqual)
                return;

            // Спрашиваем у пользователя, какой файл использовать
            string message =
                $"У документа {this.TflexDesignation} уже есть прикрепленный файл, который отличается от загружаемого файла.\n" +
                $"Прикрепленный файл: \n{GetInformationAboutFile(oldFile)}\n\n" +
                $"Загружаемый файл: \n{GetInformationAboutFile(this.LinkedFile)}\n\n" +
                "Нажмите 'Ок' если хотите произвести замену прикрепленного файла загружаемым, или нажмите 'Отмена' чтобы не производить загрузку";

            if (ParentRepository.Provider.Question(message)) {
                if (!linkedFileObject.IsCheckedOut)
                    linkedFileObject.CheckOut();
                linkedFileObject.BeginChanges();
                File.Copy(LinkedFile.FullName, linkedFileObject.LocalPath, true);
                linkedFileObject.EndChanges();
            }

            return;
        }

        private string GetInformationAboutFile(FileInfo file) {
            string info =
                $"Название файла {file.Name}\n" +
                $"Размер файла {file.Length}\n" +
                $"Дата создания {file.CreationTime.ToString("dd.MM.yyyy hh:mm")}\n" +
                $"Дата изменения {file.LastWriteTime.ToString("dd.MM.yyyy hh:mm")}\n";

            return info;
        }

        public override string ToString() {
            string template = "{0}:\nДокумент: {1}\nСтатус: {3}\nФайл: {2}\nСтатус: {4}\n";

            return string.Format(
                    template,
                    this.Designation,
                    this.NormativeReferenceObject != null ? this.NormativeReferenceObject.ToString() : "отсутствует",
                    this.FileReferenceObject != null ? this.FileReferenceObject.ToString() : "отсутствует",
                    this.IsDocumentExist == null ? "поиск не производился" : (bool)this.IsDocumentExist ? "был найден" : "был создан",
                    this.IsFileExist == null ? "поиск не производился" : (bool)this.IsFileExist ? "был найден" : "был загружен"
                    );
        }
    }

    public class DocumentRepository {

        // Контейнер для регулярных выражений
        public RegexPatterns Patterns { get; private set; }
        // Словарь для преобразования типов при создании объектов в их порядковые номера
        public Dictionary<TypeOfDocument, int> TypesToIntDict { get; private set; }

        // Данные, получаемые от пользователя
        public string Dir { get; private set; }
        private string SearchPattern { get; set; }

        // Данные, получаемые при инициализации нового объекта
        public MacroProvider Provider { get; private set; }
        
        // Данные, получаемые в процессе работы конструктора
        public string[] Files { get; private set; }
        private List<RegulatoryDocument> SuccessDocuments { get; set; }
        private List<RegulatoryDocument> ErrorDocuments { get; set; }
        private string ErrorMessage { get; set; }

        // Для работы с DOCs
        public Reference NDReference { get; private set; }
        public FileReference FReference { get; private set; }
        public FolderObject NDFolder { get; private set; }

        // Флаги
        public bool ProcessExistingDocuments { get; private set; }

        public DocumentRepository(MacroProvider provider) {

            // Проверяем, был ли передан MacroProvider.
            // Без него не получится общаться с пользователем
            if (provider == null)
                throw new Exception("Для корректной работы макроса в класс DocumentRepository необходимо передать экземпляр класса макроса");
            this.Provider = provider;

            // Инициализируем контейнер с регулярными выражениями
            this.Patterns = new RegexPatterns();

            // Инициализируем словарь с цифровыми обозначениями типов документов
            this.TypesToIntDict = new Dictionary<TypeOfDocument, int>() {
                [TypeOfDocument.ГОСТ] = 0,
                [TypeOfDocument.ОСТ] = 1,
                [TypeOfDocument.ТУ] = 2,
                [TypeOfDocument.СТО] = 3,
                [TypeOfDocument.СТП] = 4,
                [TypeOfDocument.Нормали] = 5,
                [TypeOfDocument.Метрология] = 6,
                [TypeOfDocument.ПИ] = 7
            };

            // Инициируем основные коллекции класса
            this.SuccessDocuments = new List<RegulatoryDocument>();
            this.ErrorDocuments = new List<RegulatoryDocument>();

            // Запрашиваем файлы у пользователя
            this.Files = GetInputDataFromUser();
            if (this.Files == null)
                throw new Exception("Отмена обработки нормативных документов");
            
            // Производим чтение документов и при возникновении ошибок регистрацию их
            this.ReadDocuments();

            this.ErrorMessage = GetErrorMessage();

            // Получаем основные объекты, необходимые для загрузки объектов в DOCs
            this.NDReference = this.Provider.Context.Connection.ReferenceCatalog.Find(Guids.References.НормативныеДокументы).CreateReference();
            this.FReference = new FileReference(this.Provider.Context.Connection);
            this.NDFolder = this.FReference.Find(Guids.Objects.ПапкаАрхивНД) as FolderObject;

            if (this.NDReference == null)
                throw new Exception("Не удалось получить справочник 'Нормативные документы'");
            if (this.FReference == null)
                throw new Exception("Не удалось получить справочник 'Файлы'");
            if (this.NDReference == null)
                throw new Exception("Не удалось найти папку 'Архив НД'");
        }

        private string[] GetInputDataFromUser() {
            // Запросить у пользователя директорию, в которой производить поиск
            // TODO: После тестирования убрать введенный по умолчанию путь (или установить его на рабочий стол пользователя)
            //this.Dir = @"D:\ГОСТы (тест)";
            this.Dir = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            this.SearchPattern = "*.pdf";

            string directory = "Директория";
            string browse = "Обзор";
            string pattern = "Поисковый запрос";
            string flag = "Обрабатывать существующие записи";

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
            dialog.AddFlag(flag, false);

            // Отображаем диалог
            if (dialog.Show()) {
                this.Dir = (string)dialog[directory];
                this.SearchPattern = (string)dialog[pattern];
                this.ProcessExistingDocuments = dialog[flag];
            }
            else {
                return null;
            }

            return Directory.GetFiles(this.Dir, this.SearchPattern, SearchOption.AllDirectories);
        }
        
        private void ReadDocuments() {
            foreach (string file in this.Files) {
                RegulatoryDocument newDoc = new RegulatoryDocument(file, this);
                if (!newDoc.HasError)
                    this.SuccessDocuments.Add(newDoc);
                else
                    this.ErrorDocuments.Add(newDoc);
            }
        }

        public override string ToString() {
            string template =
                "Всего файлов: {0}\n" +
                "- успешно инициализировано: {1}\n" +
                "- инициализировано с ошибками: {2}\n";
            return string.Format(template, this.Files.Length, this.SuccessDocuments.Count, this.ErrorDocuments.Count);
        }

        private string GetErrorMessage() {
            if (this.ErrorDocuments.Count == 0)
                return string.Empty;

            string template = "В процессе обработки файлов в директории '{0}' возникли следующие ошибки:\n\n{1}";
            string innerTemplate = "файл '{0}': {1}\n";
            return string.Format(template, this.Dir, string.Join("\n", this.ErrorDocuments.Select(doc => string.Format(innerTemplate, doc.LinkedFile.Name, doc.ErrorMessage))));
        }

        public void RefreshErrorMessage() {
            this.ErrorMessage = GetErrorMessage();
        }

        public void AskUserAboutFixingErrors() {
            if (this.ErrorDocuments.Count == 0)
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
            dialog.AddText(string.Format(template, this.Dir, this.Files.Length, this.SuccessDocuments.Count, this.ErrorDocuments.Count));
            dialog.AddMultiselectFromList(errors, this.ErrorDocuments.Select(doc => doc.FileName), true);
            dialog.AddInteger(quantityOfErrorsOnPage, 5);

            // Отображаем диалог
            if (dialog.Show()) {
                // Получаем список файлов на обработку

                // Так как список, получаемый из диалога является динамическим, для того, чтобы произвести
                // все необходимые преобразования, для начала нужно привести его к динамическому списку
                // (иначе не получится воспользоваться Linq методами
                List<string> documents = ((IEnumerable<dynamic>)dialog[errors]).Cast<string>().ToList<string>();

                // Запускаем корректировку файлов
                FixErrors(
                        this.ErrorDocuments
                            .Where(doc => documents.Contains(doc.FileName))
                            .ToList<RegulatoryDocument>(),
                        (int)dialog[quantityOfErrorsOnPage]
                        );

                // После произведенной корректировке переносим исправленные документы из ErrorDocuments в SuccessDocuments
                this.SuccessDocuments.AddRange(this.ErrorDocuments.Where(doc => !doc.HasError));
                this.ErrorDocuments = this.ErrorDocuments.Where(doc => doc.HasError).ToList<RegulatoryDocument>();
            }
        }

        private void FixErrors(List<RegulatoryDocument> documents, int quantityOnPage) {
            // Ограничиваем пользовательский ввод на количество максимально отображаемых записей на корректировку
            if (quantityOnPage < 1)
                quantityOnPage = 1;
            if (quantityOnPage > 8)
                quantityOnPage = 8;

            // Объявление переменных, которые будут использоваться в цикле;
            int count = 0;
            int limit;
            InputDialog dialog;
            Dictionary<int, string> errorFilesDict = new Dictionary<int, string>();

            while (true) {

                errorFilesDict.Clear();

                dialog = new InputDialog(this.Provider.Context, "Произведите корректировку названий файлов");
                dialog.AddText(string.Format("Произведите корректировку следующих файлов ({0}/{1}):", count + 1, documents.Count));

                limit = (count + quantityOnPage) < documents.Count ? (count + quantityOnPage) : documents.Count;
                for (int i = count; i < limit; i++) {
                    count++;
                    errorFilesDict[i] = string.Format("Файл {0}", count.ToString());
                    dialog.AddString(errorFilesDict[i], documents[i].FileName);
                    dialog.AddComment(errorFilesDict[i], documents[i].ErrorMessage); 
                }

                dialog.AddButton(
                        "Проверить",
                        (name) => {
                            foreach (KeyValuePair<int, string> kvp in errorFilesDict) {
                                documents[kvp.Key].ReinitializeObject((string)dialog[kvp.Value]);
                                dialog.AddComment(kvp.Value, documents[kvp.Key].ErrorMessage);
                                // Код ниже добавлен для принудительного пересчитывания диалога
                                dialog[kvp.Value] = string.Empty;
                                dialog[kvp.Value] = documents[kvp.Key].FileName;
                            }
                        },
                        false);

                if (dialog.Show()) {
                    foreach (KeyValuePair<int, string> kvp in errorFilesDict) {
                        documents[kvp.Key].ReinitializeObject((string)dialog[kvp.Value]);
                    }
                }

                // Условие выхода из бесконечного цикла
                if (documents.Count <= count)
                    break;
                
            }
        }

        public void UploadDocumentsInDocs() {

            foreach (RegulatoryDocument doc in this.SuccessDocuments) {
                doc.ProcessDocument();
            }

            string processedDocuments = string.Join("\n", this.SuccessDocuments.Where(doc => !doc.IsExcluded).Select(doc => doc.ToString()));
            this.Provider.Message("Результат работы макроса", string.IsNullOrWhiteSpace(processedDocuments) ? "Все документы уже существовали" : processedDocuments);
        }
    }

    public class RegexPatterns {
        // Словари для хранения регулярных выражений
        private Dictionary<TypeOfDocument, Regex> Types { get; set; }
        private Dictionary<TypeOfDocument, Regex> Designations { get; set; }

        public RegexPatterns() {
            // Стоить учесть, что данные регулярные выражения помимо самого типа стандарта так же должны ключать возможные уточнения
            this.Types = new Dictionary<TypeOfDocument, Regex>() {
                [TypeOfDocument.ГОСТ] = new Regex(@"^[Гг][Оо][Сс][Тт](\s[а-яА-Яa-zA-Z]{1,10}){0,2}"),
                [TypeOfDocument.ОСТ] = new Regex(@"^[Оо][Сс][Тт](\s.{1,2})?"),
                [TypeOfDocument.ТУ] = new Regex(@"^[Тт][Уу]"),
                [TypeOfDocument.ПИ] = new Regex(@"^[Пп][Ии]"),
                [TypeOfDocument.СТО] = new Regex(@"^[Сс][Тт][Оо]"),
                [TypeOfDocument.СТП] = new Regex(@"^[Сс][Тт][Пп]"),
                [TypeOfDocument.Нормали] = new Regex(@"^[Нн][Оо][Рр][Мм][Аа][Лл]"),
                [TypeOfDocument.Метрология] = new Regex(@"^[Мм][Ее][Тт][Рр][Оо][Лл][Оо][Гг]")
            };

            this.Designations = new Dictionary<TypeOfDocument, Regex>() {
                [TypeOfDocument.ГОСТ] = new Regex(@"^(\d{1,5}[\.\-]){1,3}\d{2,4}"),
                [TypeOfDocument.ОСТ] = new Regex(@"^(.{1,2}\.)?(\d{1,5}[\.\-]){1,3}\d{2,4}"),
                [TypeOfDocument.ТУ] = new Regex(@"^(\d{1,8}[\.\-]){1,3}\d{2,4}"),
            };
        }

        public Regex GetTypeRegex(TypeOfDocument type) {
            if (!this.Types.ContainsKey(type))
                throw new Exception($"Тип {type.ToString()} пока не поддерживается (GetTypeRegex)");
            return this.Types[type];
        }

        public Regex GetDesignationRegex(TypeOfDocument type) {
            if (!this.Designations.ContainsKey(type))
                throw new Exception($"Тип {type.ToString()} пока не поддерживается (GetDesignationRexex)");
            return this.Designations[type];
        }

        public bool IsTypeRegexExist(TypeOfDocument type) {
            return this.Types.ContainsKey(type);
        }

        public bool IsDesignationRegexExist(TypeOfDocument type) {
            return this.Designations.ContainsKey(type);
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

