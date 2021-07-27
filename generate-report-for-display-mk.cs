using System;
using System.IO;
using System.Linq;
using System.Collections;
using System.Collections.Generic;
using System.Windows.Forms;

using TFlex.DOCs.Model.Macros;
using TFlex.DOCs.Model.Macros.ObjectModel;
using TFlex.DOCs.Model.References;

using System.Data; // Для подключения к базе данных FoxPro
using System.Data.OleDb; // Для подключения к базе данных FoxPro
using Newtonsoft.Json; // Для сериализации информации

// Для работы данного макроса так же потребуется использование дополнительных библиотек
// Ссылки:
// - Newtonsoft.Json.dll
// Библиотеки:
// - Syste.Data.dll - Для подключения к базе данных
// - System.Data.DataSetExtensions.dll - Для преобразования DataTable в Enumerable<DataRow>()

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
    private Dictionary<string, DataTable> dataTables;
    #endregion Fields and Properties

    #region Entry Points
    public override void Run() {
        // Копируем все необходимые файлы баз данных
        CopyDataBaseFiles();

        // Производим чтение всех таблиц
        dataTables = GetDataTables(new string[] {"spec"});

        List<SpecRow> reportTable = GetCompositionOfProduct();
        if (reportTable.Count == 0) {
            DeleteTemporaryFiles();
            Message("Информация", "В процессе работы макроса возникли ошибки, отчет не был сформирован");
            return;
        }

        Message("", string.Join("\n", reportTable));
        
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
    #endregion Input and Output operations methods

    //TODO Реализовать методы чтения базы данных и формирования состава на одно изделие
    #region Methods for getting composition of product

    #region Method GetPathToDBFile
    private string GetPathToDBFile(string nameOfFile) {
        nameOfFile = nameOfFile.ToLower();
        if (!nameOfFile.EndsWith(".dbf"))
            nameOfFile += ".dbf";
        return Path.Combine(pathToTempDirectoryFoxProDb, nameOfFile);
    }
    #endregion Method GetPathToDBFile

    #region Method GetCompositionOfProduct
    private List<SpecRow> GetCompositionOfProduct() {
        List<SpecRow> result = new List<SpecRow>();

        // Запрашиваем у пользователя имя изделия для поиска данных
        string nameOfProduct = GetNameOfProduct();

        // Осуществляем в таблице поиск данного изделия.
        // Если данное изделие не было обнаружено, оповещаем об этом пользователя и завершаем работу макроса
        DataRow product = dataTables["spec"].AsEnumerable().FirstOrDefault(dataRow => ((string)dataRow["shifr"]).Trim() == nameOfProduct);
        if (product != null) {
            result.Add(new SpecRow( 
                    ((string)product["shifr"]).Trim(),
                    ((string)product["naim"]).Trim(),
                    ((string)product["izd"]).Trim()
                    ));
        }
        else {
            Message("Информация", string.Format("Изделие с обозначением '{0}' не было найдено", nameOfProduct));
            return result;
        }
        
        // Производим рекурсивное конструирование дерева
        result.AddRange(GetCompositionOfProductRecursively(nameOfProduct, dataTables["spec"]));

        // Удаляем все дубликаты и возвращаем полученные данные
        return result.Distinct().ToList<SpecRow>();
    }
    #endregion Method GetCompositionOfProduct

    #region Method GetDataTables
    // Метод для чтения данных из DBF
    private Dictionary<string, DataTable> GetDataTables (string[] dbfFiles) {
        //TODO переписать данный метод с учетом работы через базы данных
        Dictionary<string, DataTable> result = new Dictionary<string, DataTable>();

        // Создаем подкючение к базе dbf
        OleDbConnection connection = new OleDbConnection();
        string stringConnection = string.Format("Provider=VFPOLEDB.1;Data Source=\"{0}\"", pathToTempDirectoryFoxProDb);
        string errors = string.Empty;

        connection.ConnectionString = stringConnection;
        connection.Open()

        foreach (string file in dbfFiles) {
            string pathToFile = GetPathToDBFile(file);
            if (File.Exists(pathToFile)) {
                OleDbDataAdapter dataAdapter = new OleDbDataAdapter(string.Format("select * from {0}", Path.GetFileName(pathToFile)), connection);
                DataTable dataTable = new DataTable();
                dataAdapter.Fill(dataTable);

                result["file"] = dataTable;
            }
            else
                errors += string.Format("- {0}\n", file);

            if (errors != string.Empty) {
                errors = string.Format("Во процессе чтения dbf таблиц не были обнаружены следующие файлы:\n", errors);
                Message("Ошибка", errors);
                // Возвращаем пустой словарь для того, чтобы работа макроса завершилась
                return new Dictionary<string, DataTable>();
            }
            
        }
        connection.Close();

        return result;
    }
    #endregion Method GetDataTables
    
    #region Method GetCompositionOfProductRecursively
    // Метод для получения данных о составе изделия рекурсивно
    private List<SpecRow> GetCompositionOfProductRecursively(string nameOfProduct, DataTable table) {
        List<SpecRow> result = new List<SpecRow>();

        var children = table.AsEnumerable().Where(dataRow => ((string)dataRow["shifr"]).Trim() == nameOfProduct);
        if (children.Count() != 0)
            foreach (DataRow row in children) {
                result.Add(new SpecRow(
                        ((string)row["shifr"]).Trim(),
                        ((string)row["naim"]).Trim(),
                        ((string)row["izd"]).Trim()
                        ));
                result.AddRange(GetCompositionOfProductRecursively(result[result.Count - 1].Shifr, table));
            }
        return result;
    }   
    #endregion Method GetCompositionOfProductRecursively

    // Метод для запроса у пользователя имени изделия
    private string GetNameOfProduct() {
        //TODO Реализовать метод запроса изделия у пользователя
        // Пока что возвращаем тестовое обозначение для ускорения тестирования
        return "8A3049047";
    }

    #endregion Methods for getting composition of product


    //TODO Класc WinForms для отображения диалога запроса имени изделия
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
            return string.Format("{0, -25} | {1, -50} | {2, -25}", Shifr, Name, Parent);
        }
    }
    #endregion serviceClasses
}
