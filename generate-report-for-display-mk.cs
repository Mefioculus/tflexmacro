using System;
using System.IO;
using System.Linq;
using System.Text; // Для работы с кодировками
using System.Collections;
using System.Collections.Generic;
using System.Windows.Forms;
using System.Drawing;

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
    private string[] arrayOfDbFiles = new string[] {"spec.dbf", "marchp.dbf", "trud.dbf", "trud.dbt", "trud.fpt", "trud.tbk", "trud.cdx"};

    // Поля для хранения классов и методов, необходимых для использования библиотеки
    // через Reflection
    private Type dataReaderType;
    private MethodInfo readMethod;
    private MethodInfo getStringMethod;
    private MethodInfo getIntMethod;
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
        if (reportTables == null)
            return;
        if (reportTables.Count == 0) {
            Message("Информация", "В процессе работы макроса не было сформировано ни одного отчета");
            return;
        }

        FetchInfoAboutMK(reportTables);

        PrintReports(reportTables);

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
        string result = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "Выгрузки статистики МК (TFlex)");
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
        getIntMethod = dataReaderType.GetMethod("GetInt32");
        closeMethod = dataReaderType.GetMethod("Close");

        // Если хотя бы один из методов не был загружен, возвращаем false
        if ((readMethod == null) || (getStringMethod == null) || (closeMethod == null) || (getIntMethod == null))
            return false;

        return true;
    }
    #endregion Method InitDbfLibraryWithReflection

    #region Method GetCompositionOfProducts
    private Dictionary<string, List<SpecRow>> GetCompositionOfProducts() {
        // Производим чтение таблицы с данными
        List<SpecRow> specTable = GetSpecTable();
        // Получаем список изделий для того, чтобы пользователь мог выбрать из них те, на которые
        // нужно созать отчеты
        string[] allProducts = specTable.Select(row => row.Parent).Distinct().ToArray();

        // Запрашиваем у пользователя имя изделия для поиска данных
        string[] namesOfProduct = GetNamesOfProduct(allProducts);
        if (namesOfProduct.Length == 0)
            return null;

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
                    Op_op = (string)getStringMethod.Invoke(reader, new object[] {11}),
                    Naim_st = (string)getStringMethod.Invoke(reader, new object[] {12}),
                    Prof = (string)getStringMethod.Invoke(reader, new object[] {13})
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
    
    #region Method GetMarchpTable
    private List<MarchpRow> GetMarchpTable() {
        List<MarchpRow> result = new List<MarchpRow>();

        // Получаем путь к таблице dbf
        string pathToDbFile = GetPathToDBFile("marchp");

        object reader = Activator.CreateInstance(dataReaderType, new object[] {pathToDbFile, dataReaderOptions});
        
        // Производим чтение до тех пор, пока в базе есть данные 
        try {
            int tempNper;
            while ((bool)readMethod.Invoke(reader, new object[] {})) {
                try {
                    tempNper = (int)getIntMethod.Invoke(reader, new object[] {1});
                }
                catch {
                    tempNper = 9999;
                }
                MarchpRow row = new MarchpRow() {
                    Shifr = (string)getStringMethod.Invoke(reader, new object[] {0}),
                    //Nper = (int)getIntMethod.Invoke(reader, new object[] {1}),
                    Nper = tempNper,
                    Izg = (string)getStringMethod.Invoke(reader, new object[] {2})
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
    #endregion Method GetMarchpTable

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
    private string[] GetNamesOfProduct(string[] allProducts) {
        // Пока что возвращаем тестовое обозначение для ускорения тестирования
        SelectProductDialog dialog = new SelectProductDialog(allProducts);
        if (dialog.ShowDialog() == DialogResult.OK) {
            return dialog.SelectedProducts;
        }
        return new string[] {};
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
        // Получаем таблицы trud и marchp
        List<TrudRow> trudTable = GetTrudTable();
        List<MarchpRow> marchpTable = GetMarchpTable();

        foreach (SpecRow sRow in specTable) {
            FetchDataFromTrudTable(sRow, trudTable);
            FetchDataFromMarchpTable(sRow, marchpTable);
        }
    }

    #region Method FetchDataFromTrudTable
    private void FetchDataFromTrudTable(SpecRow sRow, List<TrudRow> trudTable) {
        string errMessage = string.Empty; // аккумулятор ошибок в ходе просмотра технологии
        List<string>listOfMarks = new List<string>(); // аккумулятор отметок о наличии или отсутствии элементов в технологической операции
        List<string>listOfCeh = new List<string>(); // аккумулятор маршрута
        List<string>listOfErrors = new List<string>(); // список всех ошибок
        bool techonogyIsEmpty = true;

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
                listOfMarks.Add("-");
            else
                listOfMarks.Add("+");

            if (string.IsNullOrWhiteSpace(tRow.Naim_st))
                listOfMarks.Add("-");
            else
                listOfMarks.Add("+");

            if (string.IsNullOrWhiteSpace(tRow.Prof))
                listOfMarks.Add("-");
            else
                listOfMarks.Add("+");

            string mark = string.Join("/", listOfMarks);
            listOfMarks.Clear(); // Обнуление переменной для будущих циклов

            if (mark != "+/+/+")
                listOfErrors.Add(string.Format("{0}({1}) {2}", tRow.Num_op, tRow.Izg, mark));
            else
                techonogyIsEmpty = false;
        }
        sRow.ErrorMessage = string.Join(";", listOfErrors);
        if (techonogyIsEmpty && !string.IsNullOrWhiteSpace(sRow.ErrorMessage))
            sRow.ErrorMessage = "Технология пустая";
        techonogyIsEmpty = true; // Обнуление флага для следующего цикла

        sRow.TechRoute = string.Join("-", listOfCeh);
        listOfCeh.Clear();
        listOfErrors.Clear();

        // Присвоение статусов
        if (string.IsNullOrWhiteSpace(sRow.TechRoute)) {
            sRow.FoxProMK = StatusMKFoxPro.NotFound;
            return;
        }

        // Присваиваем статус по наличию ошибок
        sRow.FoxProMK = string.IsNullOrWhiteSpace(sRow.ErrorMessage) ? StatusMKFoxPro.ExistWithoutError : StatusMKFoxPro.ExistWithError;
    }
    #endregion Method FetchDataFromTrudTable

    #region Method FetchDataFromMarchpTable
    private void FetchDataFromMarchpTable(SpecRow sRow, List<MarchpRow> marchpTable) {
        sRow.FoxRoute = string.Join("-", marchpTable.Where(r => r.Shifr == sRow.Shifr).OrderBy(r => r.Nper).Select(r => r.Izg));
    }
    #endregion Method FetchDataFromMarchpTable
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
        // TODO Реализовать печать документа при помощи OpenXML
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
        public string FoxRoute { get; set; }
        public string ErrorMessage { get; set; }
        public StatusMKFoxPro FoxProMK { get; set; }
        public StatusMKArchive TFlexMK { get; set; }
        public StatusOfDSE Status { get; set; }
        public string FoxStatus => GetFoxStatus();
        public string ArchiveStatus => GetArchiveStatus();
        public string DSEStatus => GetDSEStatus();
        public string EqualityOfRouts => TechRoute == FoxRoute ? string.Empty : "Не совпадают";

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
                    "{0};{1};{2};{3};{4};{5};{6};{7};{8};{9}",
                    Shifr,
                    Name,
                    Parent,
                    FoxStatus,
                    ArchiveStatus,
                    DSEStatus,
                    FoxRoute,
                    TechRoute,
                    EqualityOfRouts,
                    ErrorMessage
                    );
        }

        public static string GetHeader() {
            return "Шифр;Наименование;Родитель;Наличие в FoxPro;Наличие в Архиве ОГТ;Статус;Маршрут;Маршрут по МК;Сверка маршрутов;Замечания (Опер(подр) Опис/Обор/Проф)\n";
        }

        #region Methods for localization of statuses
        private string GetFoxStatus() {
            string result = string.Empty;
            switch (this.FoxProMK) {
                case StatusMKFoxPro.NotFound:
                    result = "Не найдена";
                    break;
                case StatusMKFoxPro.ExistWithError:
                    result = "Найдена, есть замечания";
                    break;
                case StatusMKFoxPro.ExistWithoutError:
                    result = "Найдена, все данные";
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
        public string Op_op { get; set; } //11
        public string Naim_st { get; set; } //12
        public string Prof { get; set; } //13
    }
    #endregion Class TrudRow
    
    #region Class MarchpRow
    private class MarchpRow {
        public string Shifr { get; set; }
        public string Izg { get; set; }
        public int Nper { get; set; }

        public override string ToString() {
            return string.Format("{0};{1};{2}",
                    Shifr,
                    Nper.ToString(),
                    Izg
                    );
        }
    }
    #endregion Class MarchpRow

    #region Select product for reporting form
    // Диалоговое окно для выбора изделий, на которые будет формироваться отчет
    private class SelectProductDialog : Form {
        
        #region Fields and Properties
        private string[] allProducts;
        private TextBox inputField;
        private ListBox foundItems;
        private ListBox selectedItems;
        private Label annotationFoundItems;
        private Label annotationSelectedItems;
        private Button buttonOk;
        private Button buttonCancel;
        private Button buttonSearch;
        private Button buttonAdd;
        private Button buttonDelete;

        public string[] SelectedProducts { get; set; }

        #endregion Fields and Properties

        #region Constructors
        public SelectProductDialog(string[] allProducts) {
            // Получаем список всех изделий
            this.allProducts = allProducts;
            this.SuspendLayout(); // Если я правильно понял, мы тормозим отрисовку до тех пор,
            // пока не инициализируем все объекты

            // Производим основные настройки формы
            InitializeComponents();

            this.ResumeLayout(false);
            this.PerformLayout();
        }
        #endregion Constructors

        #region Method InitializeComponents
        private void InitializeComponents() {
            // Инициализация поля для ввода названия изделия
            this.inputField = new TextBox();
            this.inputField.Name = "inputField";
            this.inputField.Location = new Point(30, 21);
            this.inputField.Size = new Size(250, 20);
            this.inputField.TabIndex = 7;

            // Инициализация текста аннотации для списка найденных позиций
            this.annotationFoundItems = new Label();
            this.annotationFoundItems.Name = "annotationFoundItems";
            this.annotationFoundItems.Location = new Point(27, 75);
            this.annotationFoundItems.Size = new Size(253, 23);
            this.annotationFoundItems.TabIndex = 8;
            this.annotationFoundItems.Text = "Найденные позиции";
            this.annotationFoundItems.TextAlign = ContentAlignment.MiddleCenter;

            // Инициализация текста аннотации для списка выбранных позиций
            this.annotationSelectedItems = new Label();
            this.annotationSelectedItems.Name = "annotationSelectedItems";
            this.annotationSelectedItems.Location = new Point(364, 75);
            this.annotationSelectedItems.Size = new Size(253, 23);
            this.annotationSelectedItems.TabIndex = 9;
            this.annotationSelectedItems.Text = "Позиции для формирования отчетов";
            this.annotationSelectedItems.TextAlign = ContentAlignment.MiddleCenter;

            // Инициализация списка найденных позиций
            this.foundItems = new ListBox();
            this.foundItems.Name = "foundItems";
            this.foundItems.FormattingEnabled = true;
            this.foundItems.Location = new Point(30, 101);
            this.foundItems.Size = new Size(250, 303);
            this.foundItems.TabIndex = 0;
            
            // Инициализация списка выбранных позиций
            this.selectedItems = new ListBox();
            this.selectedItems.Name = "selectedItems";
            this.selectedItems.FormattingEnabled = true;
            this.selectedItems.Location = new Point(367, 101);
            this.selectedItems.Size = new Size(250, 303);
            this.selectedItems.TabIndex = 6;

            // Инициализация кнопки подтверждения 
            this.buttonOk = new Button();
            this.buttonOk.Name = "buttonOk";
            this.buttonOk.Location = new Point(461, 415);
            this.buttonOk.Size = new Size(75, 23);
            this.buttonOk.TabIndex = 4;
            this.buttonOk.Text = "Ок";
            this.buttonOk.UseVisualStyleBackColor = true;
            this.buttonOk.Click += new EventHandler(this.ButtonOk_Click);
            
            // Инициализация кнопки отмены 
            this.buttonCancel = new Button();
            this.buttonCancel.Name = "buttonCancel";
            this.buttonCancel.Location = new Point(542, 415);
            this.buttonCancel.Size = new Size(75, 23);
            this.buttonCancel.TabIndex = 5;
            this.buttonCancel.Text = "Отмена";
            this.buttonCancel.UseVisualStyleBackColor = true;
            this.buttonCancel.Click += new EventHandler(this.ButtonCancel_Click);

            // Инициализация кнопки поиска 
            this.buttonSearch = new Button();
            this.buttonSearch.Name = "buttonSearch";
            this.buttonSearch.Location = new Point(286, 21);
            this.buttonSearch.Size = new Size(76, 23);
            this.buttonSearch.TabIndex = 1;
            this.buttonSearch.Text = "Поиск";
            this.buttonSearch.UseVisualStyleBackColor = true;
            this.buttonSearch.Click += new EventHandler(this.ButtonSearch_Click);

            // Инициализация кнопки добавления 
            this.buttonAdd = new Button();
            this.buttonAdd.Name = "buttonAdd";
            this.buttonAdd.Location = new Point(286, 226);
            this.buttonAdd.Size = new Size(75, 23);
            this.buttonAdd.TabIndex = 2;
            this.buttonAdd.Text = "->";
            this.buttonAdd.UseVisualStyleBackColor = true;
            this.buttonAdd.Click += new EventHandler(this.ButtonAdd_Click);

            // Инициализация кнопки удаления
            this.buttonDelete = new Button();
            this.buttonDelete.Name = "buttonDelete";
            this.buttonDelete.Location = new Point(286, 255);
            this.buttonDelete.Size = new Size(75, 23);
            this.buttonDelete.TabIndex = 3;
            this.buttonDelete.Text = "<-";
            this.buttonDelete.UseVisualStyleBackColor = true;
            this.buttonDelete.Click += new EventHandler(this.ButtonDelete_Click);

            // Инициализация формы
            this.AutoScaleDimensions = new SizeF(6F, 13F);
            this.AutoScaleMode = AutoScaleMode.Font;
            this.ClientSize = new Size(647, 450);
            this.Controls.Add(this.annotationFoundItems);
            this.Controls.Add(this.annotationSelectedItems);
            this.Controls.Add(this.inputField);
            this.Controls.Add(this.selectedItems);
            this.Controls.Add(this.foundItems);
            this.Controls.Add(this.buttonOk);
            this.Controls.Add(this.buttonCancel);
            this.Controls.Add(this.buttonSearch);
            this.Controls.Add(this.buttonAdd);
            this.Controls.Add(this.buttonDelete);
            this.Name = "Form1";
            this.Text = "Выбор изделия для формирования отчетов";
        }
        #endregion Method InitializeComponents

        #region Methods from button clics
        private void ButtonOk_Click(object sender, EventArgs e) {
            // Получаем выбранные элементы
            int count = this.selectedItems.Items.Count;
            this.SelectedProducts = new string[count];
            for (int i = 0; i < count; i++) {
                this.SelectedProducts[i] = (string)this.selectedItems.Items[i];
            }
            this.DialogResult = DialogResult.OK;
            this.Close();
        }

        #region ButtonCancel_Click
        private void ButtonCancel_Click(object sender, EventArgs e) {
            this.DialogResult = DialogResult.Cancel;
            this.Close();
        }
        #endregion ButtonCancel_Click
        
        #region ButtonSearch_Click
        private void ButtonSearch_Click(object sender, EventArgs e) {
            // Данное действие будет производиться по нажатию на кнопку "Поиск"
            // и будет производить поиск среди всех изделий только те, которые необходимы
            string searchText = inputField.Text;
            string[] resultOfSearch = null;
            if (!string.IsNullOrWhiteSpace(searchText)) {
                resultOfSearch = this.allProducts.Where(product => product.Contains(searchText)).ToArray();
            }
            inputField.Text = string.Empty;


            // Выводим результаты поиска в ListBox
            this.foundItems.BeginUpdate();
            // Очищаем коллекцию, если в ней до этого были объекты
            this.foundItems.Items.Clear();

            // Добавляем найденные элементы
            foreach (string shifrOfProduct in resultOfSearch) {
                this.foundItems.Items.Add(shifrOfProduct);
            }
            this.foundItems.EndUpdate();
        }
        #endregion ButtonSearch_Click

        #region ButtonAdd_Click
        private void ButtonAdd_Click(object sender, EventArgs e) {
            // Получаем выбранный объект из foundItems
            string nameProd = this.foundItems.SelectedItem != null ? (string)this.foundItems.SelectedItem : string.Empty;

            if (nameProd != string.Empty) {
                this.selectedItems.BeginUpdate();
                this.selectedItems.Items.Add(nameProd);
                this.selectedItems.EndUpdate();
            }
        }
        #endregion ButtonAdd_Click

        #region ButtonDelete_Click
        private void ButtonDelete_Click(object sender, EventArgs e) {
            object productOnRemove = this.selectedItems.SelectedItem;
            
            if (productOnRemove != null) {
                this.selectedItems.BeginUpdate();
                this.selectedItems.Items.Remove(productOnRemove);
                this.selectedItems.EndUpdate();
            }
        }
        #endregion ButtonDelete_Click
        #endregion Methods from button clics
    }
    #endregion Select product for reporting form
    #endregion serviceClasses
}
