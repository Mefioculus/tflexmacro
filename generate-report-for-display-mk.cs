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
    private string[] arrayOfDbFiles = new string[] {"spec.dbf", "mars.dbc", "mars.dct"};

    // Поля для хранения классов и методов, необходимых для использования библиотеки
    // через Reflection
    private Type dataReaderType;
    private MethodInfo readMethod;
    private MethodInfo getStringMethod;
    private MethodInfo closeMethod;

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
        DeleteTemporaryFiles();
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

    private void DeleteTemporaryFiles() {
        foreach (string file in arrayOfDbFiles) {
            string pathToTargetFile = Path.Combine(pathToTempDirectoryFoxProDb, file);

            if (File.Exists(pathToTargetFile)) {
                File.Delete(pathToTargetFile);
            }
        }
    }

    private string GetDirectory() {
        // TODO Реализовать запрос папки для сохранения отчетов от пользователей
        string result = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "TestReports");
        if (!Directory.Exists(result))
            Directory.CreateDirectory(result);

        return result;
    }
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
        FileObject libFile = fileReference.Find(new Guid("e9f3174e-43d0-46c9-9b57-35b5733a5131")) as FileObject;

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
        while ((bool)readMethod.Invoke(reader, new object[] {})) {
            SpecRow row = new SpecRow(
                    (string)getStringMethod.Invoke(reader, new object[] {1}),
                    (string)getStringMethod.Invoke(reader, new object[] {2}),
                    (string)getStringMethod.Invoke(reader, new object[] {4})
                    );
            result.Add(row);
        }
        // Закрываем reader
        closeMethod.Invoke(reader, new object[] {});

        return result;
    }

    #endregion Method GetSpecTable
    
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

    private void FetchInfoAboutMK(Dictionary<string, List<SpecRow>> data) {

    }

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
    private class SpecRow {
        public string Shifr { get; set; }
        public string Name { get; set; }
        public string Parent { get; set; }
        public bool FoxProMK { get; set; }
        public bool TFlexMK { get; set; }
        public string Status { get; set; }

        public SpecRow (string shifr, string name, string parent) {
            this.Shifr = shifr;
            this.Name = name;
            this.Parent = parent;
        }

        public override string ToString() {
            return string.Format("{0};{1};{2}", Shifr, Name, Parent);
        }

        public static string GetHeader() {
            return "Шифр;Наименование;Родитель\n";
        }
    }
    #endregion serviceClasses
}
