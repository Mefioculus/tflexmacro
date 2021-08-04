using System;
using System.IO;
using System.Linq;
using System.Text; // Для работы с кодировками
using System.Collections;
using System.Collections.Generic;
using System.Windows.Forms;

using TFlex.DOCs.Model.Macros;
using TFlex.DOCs.Model.Macros.ObjectModel;
using TFlex.DOCs.Model.References;
using TFlex.DOCs.Model.References.Files;

using System.Reflection; // Для подключения сторонней библиотеки (иначе подключить ее не получилось)
using Newtonsoft.Json; // Для сериализации информации

// Для работы данного макроса так же потребуется использование дополнительных библиотек
// Ссылки:
// - Newtonsoft.Json.dll
// Библиотеки:
// - DbfDataReader.dll - Данная библиотека подключается через Reflection в коде макроса

public class Macro : MacroProvider {
    public Macro(MacroContext context)
        : base(context)
    {
    }

    #region Fields and Properties
    // Общие поля приложения
    private string pathToSourceDirectoryFoxProDb = @"\\fs\FoxProDB\COMDB\PROIZV";
    private string pathToTempDirectoryFoxProDb = 
        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Temp");
    private string[] arrayOfDbFiles = new string[] {"spec.dbf", "trud.dbf", "trud.dbt", "trud.fpt", "trud.tbk", "trud.cdx"};

    // Поля для хранения классов и методов, необходимых для использования библиотеки
    // через Reflection
    private Type dataReaderType;
    private MethodInfo readMethod;
    private MethodInfo getStringMethod;
    private MethodInfo closeMethod;
    
    // Дебаг
    private bool debug = true;

    // Класс с уникальными идентификаторами, которые используются в данном макросе
    private static class Guids {
        public static class References {
            public static Guid АрхивОГТ = new Guid("500d4bcf-e02c-4b2e-8f09-29b64d4e7513");
            
        }
        public static class Files {
            public static Guid dbfDataReaderLibFile = new Guid("e9f3174e-43d0-46c9-9b57-35b5733a5131");
        }
        public static class Parameters {
            public static Guid АрхивОГТ_обозначениеДеталиУзла = new Guid("c11b5a98-c22c-42bc-8375-be30052ffba2");
            public static Guid АрхивОГТ_сканДокумента = new Guid("5947d0ce-b096-4791-96a4-e3ac03f9c49c");
        }
    }

    // Опции для проведения чтения dbf таблиц
    DbfDataReader.DbfDataReaderOptions dataReaderOptions = new DbfDataReader.DbfDataReaderOptions() {
        SkipDeletedRecords = true,
        Encoding = System.Text.Encoding.GetEncoding(1251)
    };

    #endregion Fields and Properties

    #region Entry Points
    public override void Run() {
        // Копируем все необходимые файлы баз данных
        CopyDataBaseFiles();

        // Инициализируем объекты для чтения базы данных
        // Если во время инициализации не инициализировались все объекты, проинформаровать пользователя
        // и завершить работу макроса
        if (!InitDbfLibraryWithReflection()) {
            Message("Ошибка", "При инициализации библиотеки DbfDataReader произошла ошибка. Обратитесь к администратору для ее устранения");
            return;
        }

        Dictionary<string, List<SpecRow>> reportTables = GetCompositionOfProducts();
        if (reportTables.Count == 0) {
            DeleteTemporaryFiles();
            Message("Информация", "В процессе работы макроса не было сформировано ни одного отчета");
            return;
        }

        FetchInfoAboutMK(reportTables);

        PrintReports(reportTables);

        // Удаляем все скопированные ранее файлы баз данных данных,
        // так как в дальнейшем они не нужны
        Message("Информация", "Работа макроса успешно завершена");
    }
    #endregion Entry Points

    #region Input and Output operations methods
    // Методы для копирования и удаления временных файлов базы данных
    // (Необходимо для того, чтобы не открывать файлы рабочей базы данных,
    // вся работа должна производиться над локальными копиями необходимых файлов)
    private void CopyDataBaseFiles() {
        string message = string.Empty;
        string headTemplate = "Во время копирования файлов базы данных не были обнаружены следующие файлы:\n";
        string regularTemplate = "- {0};\n";

        foreach (string file in arrayOfDbFiles) {
            string source = Path.Combine(pathToSourceDirectoryFoxProDb, file);
            string destination = Path.Combine(pathToTempDirectoryFoxProDb, file);
            
            if (File.Exists(source)) {
                // Перед копированием проводим проверку на то, необходимо ли производить копирование
                if (NeedACopy(source, destination))
                    File.Copy(source, destination, true);
            }
            else {
                if (message == string.Empty) {
                    message += headTemplate;
                }
                message += string.Format(regularTemplate, file);
            }
        }
    }

    private bool NeedACopy (string sourceFile, string targetFile) {
        string fileName = Path.GetFileNameWithoutExtension(sourceFile);
        string message = string.Empty;

        // Если файла в директории назначения не сущестует, возвращаем истину
        if (!File.Exists(targetFile)) {
            message = "В директории с копиями файлов не был обнаружен файл {0} базы данных FoxPro\n";
            message += "Будет произведено копирование данного файла";
            if (debug)
                Message("Информация", string.Format(message,  fileName));
            return true;
        }

        // Если файлы различаются, возвращаем истину
        DateTime sourceFileModification = File.GetLastWriteTime(sourceFile);
        DateTime targetFileModification = File.GetLastWriteTime(targetFile);
        if (sourceFileModification != targetFileModification) {
            message = "У источника и копии файла '{0}' разные даты модификации:\nФайл источник - {1}\nКопия файла - {2}\n";
            message += "Будет произведено копирование данного файла";
            if (debug)
                Message("Информация", string.Format(message, fileName, sourceFileModification, targetFileModification));
            return true;
        }

        return false;
    }

    private void DeleteTemporaryFiles() {
        foreach (string file in arrayOfDbFiles) {
            string pathToTargetFile = Path.Combine(pathToTempDirectoryFoxProDb, file);

            if (File.Exists(pathToTargetFile)) {
                File.Delete(pathToTargetFile);
            }
        }
    }

    #region Method GetDirectory 
    private string GetDirectory() {
        // TODO Реализовать запрос папки для сохранения отчетов от пользователей
        string result = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "TestReports");
        if (!Directory.Exists(result))
            Directory.CreateDirectory(result);

        return result;
    }
    #endregion Method GetDirectory 
    #endregion Input and Output operations methods

    #region Methods for getting composition of product

    #region Method GetPathToDBFile
    private string GetPathToDBFile(string nameOfFile) {
        nameOfFile = nameOfFile.ToLower();
        if (!nameOfFile.EndsWith(".dbf"))
            nameOfFile += ".dbf";
        return Path.Combine(pathToTempDirectoryFoxProDb, nameOfFile);
    }
    #endregion Method GetPathToDBFile

    #region Method InitDbfLibraryWithReflection
    private bool InitDbfLibraryWithReflection() {
        // Данный метод будет производить поиск нужной библиотеки в справочнике и возвращать путь на локальной машине

        FileReference fileReference = new FileReference(Context.Connection);
        
        // Находим файл библиотеки
        FileObject libFile = fileReference.Find(Guids.Files.dbfDataReaderLibFile) as FileObject;

        if (libFile == null) {
            Message("Ошибка", "В файловом справочнике не была обнаружена библиотека для работы с dbf файлами");
            return false;
        }

        libFile.GetHeadRevision();
        string pathToLib = libFile.LocalPath;

        if (!pathToLib.ToLower().Contains("dbfdatareader.dll")) {
            Message("Ошибка", "Проблема при открытии файла библиотеки. Найденный файл не является искомой библиотекой");
            return false;
        }

        Assembly dbfDataReaderAssembly = Assembly.LoadFrom(pathToLib);
        if (dbfDataReaderAssembly == null)
            return false;
        dataReaderType = dbfDataReaderAssembly.GetType("DbfDataReader.DbfDataReader");
        
        if (dataReaderType == null)
            return false;

        readMethod = dataReaderType.GetMethod("Read");
        getStringMethod = dataReaderType.GetMethod("GetString");
        closeMethod = dataReaderType.GetMethod("Close");

        // Если хотя бы один из методов не был загружен, возвращаем false
        if ((readMethod == null) || (getStringMethod == null) || (closeMethod == null))
            return false;

        return true;
    }
    #endregion Method InitDbfLibraryWithReflection

    #region Method GetCompositionOfProducts
    private Dictionary<string, List<SpecRow>> GetCompositionOfProducts() {
        // Запрашиваем у пользователя имя изделия для поиска данных
        string[] namesOfProduct = GetNamesOfProduct();

        // Осуществляем в таблице поиск данного изделия.
        // Если данное изделие не было обнаружено, оповещаем об этом пользователя и завершаем работу макроса
        List<SpecRow> specTable = GetSpecTable();

        Dictionary<string, List<SpecRow>> result = new Dictionary<string, List<SpecRow>>();
        foreach (string name in namesOfProduct) {
            List<SpecRow> reportData = GetCompositionOfProductRecursively(name, specTable);
            if (reportData != null) {
                // Удаляем дубликаты из таблицы и подключаем ее к результату
                result[name] = reportData.Distinct().ToList<SpecRow>();
            }
        }
        
        return result;
    }
    #endregion Method GetCompositionOfProducts

    #region Method GetSpecTable
    // Метод для получения данных из таблицы Spec
    private List<SpecRow> GetSpecTable() {
        List<SpecRow> result = new List<SpecRow>();

        // Получаем путь к таблице dbf
        string pathToDbFile = GetPathToDBFile("spec");

        object reader = Activator.CreateInstance(dataReaderType, new object[] {pathToDbFile, dataReaderOptions});
        
        // Производим чтение до тех пор, пока в базе есть данные 
        try {
            while ((bool)readMethod.Invoke(reader, new object[] {})) {
                SpecRow row = new SpecRow(
                        (string)getStringMethod.Invoke(reader, new object[] {1}),
                        (string)getStringMethod.Invoke(reader, new object[] {2}),
                        (string)getStringMethod.Invoke(reader, new object[] {4})
                        );
                result.Add(row);
            }
        }
        finally {
            // Закрываем reader
            closeMethod.Invoke(reader, new object[] {});
        }

        return result;
    }

    #endregion Method GetSpecTable

    #region Method GetTrudTable
    // Метод для получения данных из таблицы Spec
    private List<TrudRow> GetTrudTable() {
        List<TrudRow> result = new List<TrudRow>();

        // Получаем путь к таблице dbf
        string pathToDbFile = GetPathToDBFile("trud");

        object reader = Activator.CreateInstance(dataReaderType, new object[] {pathToDbFile, dataReaderOptions});
        
        // Производим чтение до тех пор, пока в базе есть данные 
        try {
            while ((bool)readMethod.Invoke(reader, new object[] {})) {
                TrudRow row = new TrudRow() {
                    Shifr = (string)getStringMethod.Invoke(reader, new object[] {0}),
                    Izg = (string)getStringMethod.Invoke(reader, new object[] {1}),
                    Num_op = (string)getStringMethod.Invoke(reader, new object[] {2}),
                    Shifr_op = (string)getStringMethod.Invoke(reader, new object[] {3}),
                    Op_op = (string)getStringMethod.Invoke(reader, new object[] {12}),
                    Naim_st = (string)getStringMethod.Invoke(reader, new object[] {13}),
                    Prof = (string)getStringMethod.Invoke(reader, new object[] {14})
                };
                result.Add(row);

            }
        }
        finally {
            // Закрываем reader
            closeMethod.Invoke(reader, new object[] {});
        }

        return result;
    }

    #endregion Method GetTrudTable
    
    #region Method GetCompositionOfProductRecursively
    // Метод для получения данных о составе изделия рекурсивно
    private List<SpecRow> GetCompositionOfProductRecursively(string nameOfProduct, List<SpecRow> table) {
        List<SpecRow> result = new List<SpecRow>();

        SpecRow currentProduct = table.FirstOrDefault(row => row.Shifr == nameOfProduct);
        if (currentProduct == null) {
            Message("Ошибка", string.Format("Изделие '{0}' не было найдено", nameOfProduct));
            return result;
        }
        else
            result.Add(currentProduct);

        var children = table.Where(row => row.Parent == nameOfProduct);
        if (children.Count() != 0)
            foreach (SpecRow row in children)
                result.AddRange(GetCompositionOfProductRecursively(row.Shifr, table));

        return result;
    }   
    #endregion Method GetCompositionOfProductRecursively

    #region Method GetNamesOfProduct
    // Метод для запроса у пользователя имени изделия
    private string[] GetNamesOfProduct() {
        //TODO Реализовать метод запроса изделия у пользователя
        // Пока что возвращаем тестовое обозначение для ускорения тестирования
        return new string [] {"8А3049047"};
    }
    #endregion Method GetNamesOfProduct

    #endregion Methods for getting composition of product

    #region Fetchng data about mk from Tflex and FoxPro
    #region Main method FetchInfoAboutMK
    private void FetchInfoAboutMK(Dictionary<string, List<SpecRow>> data) {
        foreach (KeyValuePair<string, List<SpecRow>> kvp in data) {
            FetchDataFromTFlex(kvp.Value);
            FetchDataFromFox(kvp.Value);
            SetStatus(kvp.Value);
        }
    }
    #endregion Main method FetchInfoAboutMK

    #region Method FetchDataFromTFlex
    // Метод для получения сведений о наличии МК в архиве ОГТ
    private void FetchDataFromTFlex(List<SpecRow> table) {
        // Получаем справочник архива ОГТ
        Reference archiveOgt = Context.Connection.ReferenceCatalog.Find(Guids.References.АрхивОГТ).CreateReference();
        TFlex.DOCs.Model.Structure.ParameterInfo shifrParam =
            archiveOgt.ParameterGroup.Parameters.Find(Guids.Parameters.АрхивОГТ_обозначениеДеталиУзла);

        foreach (SpecRow record in table) {
            // Производим поиск объектов
            ReferenceObject archiveRecord = archiveOgt.FindOne(shifrParam, record.Shifr);

            // Если получилось найти соотвутствующую запись в архиве, тогда проверяем значение в колонке
            // "АрхивОГТ_сканДокумента"
            if (archiveRecord != null) {
                record.TFlexMK = archiveRecord[Guids.Parameters.АрхивОГТ_сканДокумента].GetString() != string.Empty ? 
                    StatusMKArchive.Exist : StatusMKArchive.NotExist;
            }
            else
                record.TFlexMK = StatusMKArchive.NotFound;
        }

        // Получаем все объекты справочника
    }
    #endregion Method FetchDataFromTFlex

    #region Method FetchDataFromFox
    // Метод для получения сведений о том, какие сведения нормы на данную ДСЕ
    // в FoxPro
    private void FetchDataFromFox(List<SpecRow> specTable) {
        // TODO Реализовать получение данных из таблицы TRUD
        // Получаем таблицу TRUD
        List<TrudRow> trudTable = GetTrudTable();

        foreach (SpecRow sRow in specTable) {

            string errMessage = string.Empty; // аккумулятор ошибок в ходе просмотра технологии
            List<string>listOfEmptyCells = new List<string>(); // аккумулятор ошибок в ходе просмотра одной операции технологии
            List<string>listOfCeh = new List<string>(); // аккумулятор маршрута
            List<string>listOfErrors = new List<string>(); // список всех ошибок

            foreach (TrudRow tRow in trudTable.Where(r => r.Shifr == sRow.Shifr)) {
                // Составляем маршрут
                if (listOfCeh.Count == 0)
                    listOfCeh.Add(tRow.Izg);
                else {
                    if (listOfCeh[listOfCeh.Count - 1] != tRow.Izg)
                        listOfCeh.Add(tRow.Izg);
                }

                // Составляем список замечаний, которые возникли в процессе
                if (string.IsNullOrWhiteSpace(tRow.Op_op))
                    listOfEmptyCells.Add("описание");

                if (string.IsNullOrWhiteSpace(tRow.Naim_st))
                    listOfEmptyCells.Add("оборудование");


                if (string.IsNullOrWhiteSpace(tRow.Prof))
                    listOfEmptyCells.Add("профессия");

                if (listOfEmptyCells.Count > 0) {
                    if (listOfEmptyCells.Count == 3)
                        listOfErrors.Add(string.Format("оп {0}.{1}.{2} пустая", tRow.Izg, tRow.Num_op, tRow.Shifr_op));
                    else
                        listOfErrors.Add(string.Format("в оп {0}.{1}.{2} отсутствуют данные о: {3}, ", tRow.Izg, tRow.Num_op, tRow.Shifr_op, string.Join(", ", listOfEmptyCells)));
                }


                // Обнуление временный переменных
                listOfEmptyCells.Clear();

            }

            sRow.ErrorMessage = string.Join(", ", listOfErrors);
            sRow.TechRoute = string.Join("-", listOfCeh);
            listOfCeh.Clear();
            listOfErrors.Clear();

            // Присвоение статусов
            if (string.IsNullOrWhiteSpace(sRow.TechRoute)) {
                sRow.FoxProMK = StatusMKFoxPro.NotFound;
                continue;
            }

            // Присваиваем статус по наличию ошибок
            sRow.FoxProMK = string.IsNullOrWhiteSpace(sRow.ErrorMessage) ? StatusMKFoxPro.ExistWithoutError : StatusMKFoxPro.ExistWithError;
        }


    }
    #endregion Method FetchDataFromFox

    #region Method SetStatus
    // Установка статуса текущей ДСЕ
    private void SetStatus(List<SpecRow> table) {
        foreach (SpecRow row in table) {
            // В фоксе не были найдены операции
            if (row.FoxProMK == StatusMKFoxPro.NotFound) {
                switch (row.TFlexMK) {
                    // В архиве не была найдено данных по этой ДСЕ
                    case StatusMKArchive.NotFound:
                        row.Status = StatusOfDSE.NeedQualification;
                        break;
                    // В архиве есть ДСЕ, но на нее отсутствует технология
                    case StatusMKArchive.NotExist:
                        row.Status = StatusOfDSE.Creating;
                        break;
                    // В архиве есть ДСЕ и на нее присутствует технология
                    case StatusMKArchive.Exist:
                        row.Status = StatusOfDSE.NeedQualification;
                        break;
                    // Данные о наличии или отсутствии записи не были обработаны
                    default:
                        row.Status = StatusOfDSE.NeedQualification;
                        break;
                }
            }

            // в фоксе была найдена технология
            if (row.FoxProMK == StatusMKFoxPro.ExistWithoutError) {
                switch (row.TFlexMK) {
                    // В архиве отсутствует информация по данной ДСЕ
                    case StatusMKArchive.NotFound:
                        row.Status = StatusOfDSE.NeedQualification;
                        break;
                    // В архиве есть ДСЕ но на нее отсутствует технология
                    case StatusMKArchive.NotExist:
                        row.Status = StatusOfDSE.Reworking;
                        break;
                    // В архиве есть МК на данное изделие
                    case StatusMKArchive.Exist:
                        row.Status = StatusOfDSE.Done;
                        break;
                    default:
                        row.Status = StatusOfDSE.NeedQualification;
                        break;
                }
            }

            // В фоксе была найдена МК с пустыми операциями
            if (row.FoxProMK == StatusMKFoxPro.ExistWithError) {
                row.Status = StatusOfDSE.NeedQualification;
            }
        }
    }
    #endregion Method SetStatus
    #endregion Fetchng data about mk from Tflex and FoxPro

    #region Methods for generate output report file

    private void PrintReports(Dictionary<string,List<SpecRow>> dataOnPrint) {
        string pathToDirectoryToSave = GetDirectory();

        foreach (KeyValuePair<string, List<SpecRow>> kvp in dataOnPrint) {
            PrintReport(pathToDirectoryToSave, kvp.Key, kvp.Value);
        }
    }

    private void PrintReport(string directory, string nameOfProduct, List<SpecRow> data) {
        string text = SpecRow.GetHeader();
        text += string.Join("\n", data);
        string pathToFile = string.Format("{0}.csv", Path.Combine(directory, nameOfProduct));
        File.WriteAllText(pathToFile, text, Encoding.GetEncoding(1251));
    }

    #endregion Methods for generate output report file

    #region serviceClasses

    #region Enums
    private enum StatusMKArchive {
        NotProcessed, // Дефолтный статус
        Exist, // Маршрутная карта существует в Архиве ОГТ
        NotExist, // Маршрутная карту не существует а Архиве ОГТ
        NotFound // Данная ДСЕ не была обнаружена в Архиве ОГТ
    }

    private enum StatusMKFoxPro {
        NotProcessed, // Дефолтный статус
        NotFound, // Технология не найдена
        ExistWithError, // Технология найдена, но обнаружены замечания 
        ExistWithoutError // Технология найдена, замечаний нет
    }

    private enum StatusOfDSE {
        NotProcessed, // Дефолтный статус
        Done, // Технология готова (есть и в Fox и в Архиве)
        NeedQualification, // Требует внимания, так как результат неожидаемый
        Reworking, // Технология есть только в Fox
        Creating // Технологии нет ни в Fox на в Архиве
    }
    #endregion Enums

    #region Class SpecRow
    private class SpecRow {
        public string Shifr { get; set; }
        public string Name { get; set; }
        public string Parent { get; set; }
        public string TechRoute { get; set; }
        public string ErrorMessage { get; set; }
        public StatusMKFoxPro FoxProMK { get; set; }
        public StatusMKArchive TFlexMK { get; set; }
        public StatusOfDSE Status { get; set; }
        public string FoxStatus => GetFoxStatus();
        public string ArchiveStatus => GetArchiveStatus();
        public string DSEStatus => GetDSEStatus();

        public SpecRow (string shifr, string name, string parent) {
            this.Shifr = shifr;
            this.Name = name;
            this.Parent = parent;

            this.TFlexMK = StatusMKArchive.NotProcessed;
            this.FoxProMK = StatusMKFoxPro.NotProcessed;
            this.Status = StatusOfDSE.NotProcessed;
        }

        public override string ToString() {
            return string.Format(
                    "{0};{1};{2};{3};{4};{5};{6};{7}",
                    Shifr,
                    Name,
                    Parent,
                    FoxStatus,
                    ArchiveStatus,
                    DSEStatus,
                    TechRoute,
                    ErrorMessage
                    );
        }

        public static string GetHeader() {
            return "Шифр;Наименование;Родитель;Наличие в FoxPro;Наличие в Архиве ОГТ;Статус;Маршрут;Возникшие замечания\n";
        }

        #region Methods for localization of statuses
        private string GetFoxStatus() {
            string result = string.Empty;
            switch (this.FoxProMK) {
                case StatusMKFoxPro.NotFound:
                    result = "Технология не найдена";
                    break;
                case StatusMKFoxPro.ExistWithError:
                    result = "Технология пустая";
                    break;
                case StatusMKFoxPro.ExistWithoutError:
                    result = "Технология найдена";
                    break;
                default:
                    result = "Не обработано";
                    break;
            }
            return result;
        }

        private string GetArchiveStatus() {
            string result = string.Empty;
            switch (this.TFlexMK) {
                case StatusMKArchive.NotFound:
                    result = "Отсутствует изделие";
                    break;
                case StatusMKArchive.Exist:
                    result = "Технология найдена";
                    break;
                case StatusMKArchive.NotExist:
                    result = "Технология не найдена";
                    break;
                default:
                    result = "Не обработано";
                    break;
            }
            return result;
        }

        private string GetDSEStatus() {
            string result = string.Empty;
            switch (this.Status) {
                case StatusOfDSE.Done:
                    result = "Готово";
                    break;
                case StatusOfDSE.Reworking:
                    result = "Корректируется";
                    break;
                case StatusOfDSE.Creating:
                    result = "Создается";
                    break;
                case StatusOfDSE.NeedQualification:
                    result = "Результаты требуют уточнения";
                    break;
                default:
                    result = "Не обработано";
                    break;
            }
            return result;
        }
        #endregion Methods for localization of statuses
    }
    #endregion Class SpecRow

    #region Class TrudRow
    // Класс для временного хранения выгруженных данных из таблицы Trud FoxPro
    private class TrudRow {
        public string Shifr { get; set; } //0
        public string Izg { get; set; } //1
        public string Num_op { get; set; } //2
        public string Shifr_op { get; set; } //3
        public string Op_op { get; set; } //12
        public string Naim_st { get; set; } //13
        public string Prof { get; set; } //14
    }
    #endregion Class TrudRow
    #endregion serviceClasses
}
