using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Collections;
using System.Collections.Generic;
using System.Windows.Forms;
using System.Drawing;

using TFlex.DOCs.Model.Macros;
using TFlex.DOCs.Model.Macros.ObjectModel;
using TFlex.DOCs.Model.References;
using TFlex.DOCs.Model.References.Files;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

/* 
Так же для работы данного макроса потребуется подключение дополнительных библиотек
DbfDataReader.dll для работы с файлами базы FoxPro
System.Data.dll (необходим для работы DbfDataReader, необходимо, чтобы он находился в директории "Служебные файлы"
WindowsBase.dll (так же необходим для работы DbfDataReader)
DocumentFormat.OpenXml для работы с Excel // Уже есть в директории DOCs, подключается через ссылку
*/


public class Macro : MacroProvider {

    #region Constructor

    public Macro(MacroContext context)
        : base(context) {
            // Производим копирование файлов базы данных для последующей работы с ними
            if (!CopyDataBaseFiles())
                throw new Exception("Во время получения файлов базы данных FoxPro возникла ошибка");
        }

    #endregion Constructor

    #region Fields and Properties

    private static string pathToSourceDirectoryFoxProDb = @"\\fs\FoxProDB\COMDB\PROIZV";
    private static string pathToTempDirectoryFoxProDb = 
        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Temp");
    // Список файлов, которые нужно грузить в кэш директорию
    private static string[] arrayOfDbFiles =
        new string[] {
            "spec.dbf",
            "klas.dbf",
            "klas.cdx",
            "klasm.dbf",
            "kat_ediz.dbf",
            "norm.dbf",
            "norm.cdx",
            "marchp.dbf",
            "marchp.cdx"
        };
    // Список подразделений, которые относятся к предприятию (Следовательно не являются покупными)
    private static string[] arrayOfUnits = new string[] {
            "001",
            "002",
            "004",
            "005",
            "006",
            "016",
            "017",
            "022",
            "023",
            "024",
            "032",
            "100",
            "101",
            "102",
            "103",
            "104",
            "105",
            "106",
            "335",
            "338"
    };

    #endregion Fields and Properties

    #region EntryPoints

    #region Run()

    public override void Run () {
    }

    #endregion Run()

    #region ВыгрузитьДеревоИзделия()

    public void ВыгрузитьДеревоИзделия() {
        
        #region Производим чтение всех необходимых таблиц

        string message = string.Empty;

        Table specTable = new Table("spec", pathToTempDirectoryFoxProDb);
        message += specTable.Status;

        Table marchpTable = new Table("marchp", pathToTempDirectoryFoxProDb);
        message += marchpTable.Status;
        
        Message("Чтение таблиц FoxPro", message);

        #endregion Производим чтение всех необходимых таблиц

        #region Формируем дерево изделия

        // Передаем таблицы, необходимые для формирования дерева
        TreeOfProduct.SpecTable = specTable;
        TreeOfProduct.MarchpTable = marchpTable;

        string[] listOfSelectedProducts = GetNamesOfProductsFromUser(specTable);

        foreach (string product in listOfSelectedProducts) {
            // Формируем деревья для выбранных изделий
            TreeOfProduct tree = TreeOfProduct.GenerateTree(product);
            tree.GetInfoAboutPurchaseProducts();
            tree.GetInfoAboutRoutes();

            // Создаем таблицу и заполняем ее
            ExcelTableOptions options = new ExcelTableOptions() { UseAutoFilter = true, NameOfSheet = "Выгрузка состава изделия" };
            ExcelTable table = new ExcelTable(
                    GetDirectory(),
                    string.Format("{0} (состав изделия)", product),
                    options
                    );

            tree.FillDataForTreeOfProduct(table);

            table.Columns.Add("Обозначение", 20);
            table.Columns.Add("Наименование", 20);
            table.Columns.Add("Применяемость", 20);
            table.Columns.Add("Маршрут", 20);
            table.Columns.Add("Признак покупного", 20);

            table.Generate(new string[] {
                    "Обозначение",
                    "Наименование",
                    "Применяемость",
                    "Маршрут",
                    "Признак покупного"
                    });
        }
        
        Message("Информация", "Выгрузка произведена");
        #endregion Формируем дерево изделия
    }

    #endregion ВыгрузитьДеревоИзделия()

    #region ВыгрузитьСтандартныеИзделия()

    public void ВыгрузитьСтандартныеИзделия() {

        #region Производим чтение всех необходимых таблиц

        string message = string.Empty;

        Table specTable = new Table("spec", pathToTempDirectoryFoxProDb);
        message += specTable.Status;

        Table klasTable = new Table("klas", pathToTempDirectoryFoxProDb);
        message += klasTable.Status;

        Message("Чтение таблиц FoxPro", message);

        #endregion Производим чтение всех необходимых таблиц

        #region Формируем все необходимые словари

        // Создаем словарь с стандартными изделиями
        Dictionary<string, TableRow> allOkpCodesKlasTable = new Dictionary<string, TableRow>();
        foreach (TableRow row in klasTable.Rows) {
            allOkpCodesKlasTable[row["okp"]] = row;
        }

        #endregion Формируем все необходимые словари

        #region Для выбранных пользователем изделий формируем деревья и наполняем их необходимыми параметрами

        // Передаем в TreeOfTable таблица для последующего формирования деревьев
        TreeOfProduct.SpecTable = specTable;

        // Запрашиваем у пользователя список изделий, для которых необходимо формировать дерево
        string[] listOfSelectedProducts = GetNamesOfProductsFromUser(specTable);

        foreach (string product in listOfSelectedProducts) {

            TreeOfProduct tree = TreeOfProduct.GenerateTree(product);

            // Дополняем ноды дерева информацией о стандартных изделиях
            foreach (TreeNode node in tree.AllNodes) {
                if (allOkpCodesKlasTable.ContainsKey(node["shifr"])) {
                    node["type"] = "Стандартное изделие";
                    node["gost"] = allOkpCodesKlasTable[node["shifr"]]["gost"];
                }
                else {
                    node["type"] = "Не классифицировано";
                    node["gost"] = string.Empty;
                }
            }

            // Генерируем таблицу по стандартным изделиям

            ExcelTableOptions options = new ExcelTableOptions() {
                UseAutoFilter = true,
                NameOfSheet = "Выгрузка стандартных изделий"
            };
            
            ExcelTable table = new ExcelTable(
                    GetDirectory(),
                    string.Format("{0} (стандартные изделия)", product),
                    options
                    );

            // Передаем таблицу в объект состава изделия для ее заполнения
            tree.FillDataForStIzd(table);

            // Заполняем данные о ширине колонок
            table.Columns.Add("Изделие", 20);
            table.Columns.Add("Обозначение", 20);
            table.Columns.Add("Наименование", 20);
            table.Columns.Add("ГОСТ", 20);
            table.Columns.Add("Применяемость", 20);

            table.Generate(new string[] {
                    "Изделие",
                    "Обозначение",
                    "Наименование",
                    "ГОСТ",
                    "Применяемость",
                    });
        }

        Message("Информация", "Формирование всех выгрузок стандартных изделий завершено");

        #endregion Для выбранных пользователем изделий формируем деревья и наполняем их необходимыми параметрами
    }

    #endregion ВыгрузитьСтандартныеИзделия()

    #region ВыгрузитьМатериалы()

    public void ВыгрузитьМатериалы() {

        #region Производим чтение всех необходимых таблиц

        string message = string.Empty;

        Table specTable = new Table("spec", pathToTempDirectoryFoxProDb);
        message += specTable.Status;

        Table klasmTable = new Table("klasm", pathToTempDirectoryFoxProDb);
        message += klasmTable.Status;

        Table katEdizTable = new Table("kat_ediz", pathToTempDirectoryFoxProDb);
        message += katEdizTable.Status;

        Table normTable = new Table("norm", pathToTempDirectoryFoxProDb);
        message += normTable.Status;

        Table marchpTable = new Table("marchp", pathToTempDirectoryFoxProDb);
        message += marchpTable.Status;

        Message("Чтение таблиц FoxPro", message);

        #endregion Производим чтение всех необходимых таблиц

        #region Формируем словари с необходимыми данными

        // Создаем словарь с материалами
        Dictionary<string, TableRow> allOkpCodesKlasmTable = new Dictionary<string, TableRow>();
        foreach (TableRow row in klasmTable.Rows) {
            allOkpCodesKlasmTable[row["okp"]] = row;
        }

        // Создаем словарь с нормами материалов
        Dictionary<string, List<TableRow>> allShifrsNormTable = new Dictionary<string, List<TableRow>>();
        foreach (TableRow row in normTable.Rows) {
            if (allShifrsNormTable.ContainsKey(row["shifr"])) {
                allShifrsNormTable[row["shifr"]].Add(row);
            }
            else {
                allShifrsNormTable[row["shifr"]] = new List<TableRow>();
                allShifrsNormTable[row["shifr"]].Add(row);
            }
        }

        // Создаем словарь с единицами измерения
        Dictionary<int, string> unitsOfMeasurement = new Dictionary<int, string>();
        int kod = 0; // Переменная для хранения кода единицы измерения
        foreach (TableRow row in katEdizTable.Rows) {
            kod = int.Parse(row["kod"]);
            if (!unitsOfMeasurement.ContainsKey(kod)) {
                unitsOfMeasurement[kod] = row["name"];
            }
        }

        #endregion Формируем словари с необходимыми данными

        #region Для выбранных пользователей изделий формируем деревья и наполняем их необходимыми данными

        // Добавляем таблицы, необходимые для заполнения дерева
        TreeOfProduct.SpecTable = specTable;
        TreeOfProduct.MarchpTable = marchpTable;

        string[] listOfSelectedProducts = GetNamesOfProductsFromUser(specTable);

        foreach (string product in listOfSelectedProducts) {
            // Формируем деревья для выбранных изделий
            TreeOfProduct tree = TreeOfProduct.GenerateTree(product);

            // Подключаем к дереву таблицу с маршрутами для определения маршрутов и того, является ли изделие покупным
            tree.GetInfoAboutPurchaseProducts();
            
            // Создаем таблицу и заполняем ее
            ExcelTableOptions options = new ExcelTableOptions() { UseAutoFilter = true, NameOfSheet = "Выгрузка материалов" };
            ExcelTable table = new ExcelTable(
                    GetDirectory(),
                    string.Format("{0} (материалы)", product),
                    options
                    );

            tree.FillDataForMaterial(table, allShifrsNormTable, allOkpCodesKlasmTable, unitsOfMeasurement);

            // Заполняем данные о ширине колонок и генерируем таблицу
            table.Columns.Add("Обозначение", 20);
            table.Columns.Add("Наименование", 20);
            table.Columns.Add("Применяемость", 20);
            table.Columns.Add("Единицы измерения", 20);
            table.Columns.Add("Вид", 20);
            table.Columns.Add("Стандарт", 20);

            table.Generate(new string[] {
                    "Обозначение",
                    "Наименование",
                    "Стандарт",
                    "Применяемость",
                    "Единицы измерения",
                    "Вид"
                    });

        }

        #endregion Для выбранных пользователей изделий формируем деревья и наполняем их необходимыми данными

        Message("Информация", "Формирование выгрузок материалов завершено");
    }

    #endregion ВыгрузитьМатериалы()
    
    #endregion EntryPoints

    #region Методы для копирования файлов базы данных Fox
    
    #region CopyDataBaseFiles

    private static bool CopyDataBaseFiles () {
        // TODO Реализовать новую версию метода по копированию файлов
        // Формируем список объектов, которые хранят данные о названии файла, пути его источика и пути, куда он должен быть сохранен
        List<DataFileInfo> files = arrayOfDbFiles
            .Select(fileName => new DataFileInfo(fileName, pathToSourceDirectoryFoxProDb, pathToTempDirectoryFoxProDb))
            .ToList<DataFileInfo>();

        // Проверяем, все ли файлы есть в папке источнике
        if (!CheckSourceFiles(files))
            return false;

        // Производим копирование отсутствующих файлов
        foreach (DataFileInfo file in files.Where(f => f.NeedCopy)) {
            // Производим копирование файла
            File.Copy(file.PathToSourceFile, file.PathToDestinationFile, true);
        }

        var filesOnUpdate = files
            .Where(f => (!f.NeedCopy) && (f.NeedUpdate))
            .ToList<DataFileInfo>();

        // Спрашиваем, нужно ли обновить файлы, которые уже существовали
        string message = string.Join("\n", filesOnUpdate
                .Select(f => string.Format(
                        "Для файла {0}:\n- в источнике {1}\n- в кэше {2}",
                        f.FileName,
                        f.ModificationSourceFile,
                        f.ModificationDestinationFile
                        ))
                );
        
        // Елси есть файлы, которые нужно обновлять
        if (message != string.Empty) {
            message = string.Format(
                    "Список файлов, кэшированные файлы которых устарели:\n{0}\n\nПроизвести обновление файлов?",
                    message
                    );

            DialogResult result =
                MessageBox.Show(message, "Обновить файлы в кэше?", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

            if (result == DialogResult.Yes) {
                foreach (DataFileInfo file in filesOnUpdate)
                    File.Copy(file.PathToSourceFile, file.PathToDestinationFile, true);
            }
        }

        return true;
    }

    #endregion CopyDataBaseFiles

    #region CheckSourceFiles
    // Метод для проверки наличия всех требуемых файлов в исходной папке
    private static bool CheckSourceFiles(List<DataFileInfo> listOfFiles) {
        string message = string.Empty;
        foreach (DataFileInfo file in listOfFiles) {
            // Проверяем файлы на существование
            if (!file.IsSourceExist)
                message += string.Format("- {0};\n", file.FileName);
        }

        if (message != string.Empty) {
            message = string.Format("Ны были обнаружены следующие файлы, необходимые для работы макроса:\n{0}", message);
            MessageBox.Show(message, "Ошибка");
            return false;
        }

        return true;
    }

    #endregion CheckSourceFiles

    #endregion Методы для копирования файлов базы данных Fox

    #region Service methods
    
    #region Method GetPathToDBFile

    private static string GetPathToDBFile(string nameOfFile)
    {
        nameOfFile = nameOfFile.ToLower();
        if (!nameOfFile.EndsWith(".dbf"))
            nameOfFile += ".dbf";
        return Path.Combine(pathToTempDirectoryFoxProDb, nameOfFile);
    }

    #endregion Method GetPathToDBFile

    #region Method GetNamesOfProduct

    private static string[] GetNamesOfProduct(string[] allProducts) {
        // Пока что возвращаем тестовое обозначение для ускорения тестирования
        SelectProductDialog dialog = new SelectProductDialog(allProducts);
        if (dialog.ShowDialog() == DialogResult.OK) {
            return dialog.SelectedProducts;
        }
        return new string[] {};
    }

    #endregion Method GetNamesOfProduct

    #region Method GetNamesOfProductsFromUser

    private static string[] GetNamesOfProductsFromUser(Table table) {
        // Для начала вычисляем весь список изделий, для которых можно запускать
        // формирование дерева изделия.
        string[] arrayOfVariants = table.GetAllDataFromColumn("shifr+naim").Distinct().ToArray();

        // Выбираем только те изделия, которые есть в первом и втором массиве 
        
        // Формируем данные для вывода в диалоге в формате Обозначение^Наименование
        // для того, чтобы поиск можно было производить и по обозначению, и по наименованию


        // Вызываем диалог для получения результатов выбора пользователя
        SelectProductDialog dialog = new SelectProductDialog(arrayOfVariants);
        if (dialog.ShowDialog() == DialogResult.OK) {
            
            // Возвращаем только обозначения выбранных объектов
            return dialog.SelectedProducts.Select(str => str.Split(new string [] {"  ^  "}, StringSplitOptions.RemoveEmptyEntries)[0]).ToArray();
        }

        return new string[] {};
    }

    #endregion Method GetNamesOfProductsFromUser

    #region Method GetDirectory 

    private static string GetDirectory() {
        string result = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
                "Выгрузка стандартных изделий из Fox"
                );
        if (!Directory.Exists(result))
            Directory.CreateDirectory(result);

        return result;
    }

    #endregion Method GetDirectory 

    #endregion Service methods

    #region Service classes

    #region CopyDataBaseFiles
    private class DataFileInfo {
        public string FileName { get; set; }
        public string PathToSourceFile { get; set; }
        public string PathToDestinationFile { get; set; }
        public DateTime ModificationSourceFile { get; set; }
        public DateTime ModificationDestinationFile { get; set; }
        public bool IsSourceExist { get; set; }
        public bool IsDestinationExist { get; set; }
        public bool NeedCopy { get; set; }
        public bool NeedUpdate { get; set; } = true;

        public DataFileInfo(string fileName, string sourceDirectory, string destinationDirectory) {
            this.FileName = fileName;
            this.PathToSourceFile = Path.Combine(sourceDirectory, fileName);
            this.PathToDestinationFile = Path.Combine(destinationDirectory, fileName);

            // Проверка на существование файлов и проверка даты последней модификации
            this.IsSourceExist = File.Exists(PathToSourceFile);
            this.ModificationSourceFile =
                this.IsSourceExist ? File.GetLastWriteTime(this.PathToSourceFile) : new DateTime(1, 1, 1);
            this.IsDestinationExist = File.Exists(PathToDestinationFile);
            this.ModificationDestinationFile =
                this.IsDestinationExist ? File.GetLastWriteTime(this.PathToDestinationFile) : new DateTime(1, 1, 1);

            this.NeedCopy = IsDestinationExist ? false : true;

            // Вычисляем, необходимо ли обновление для файла 
            if (this.IsDestinationExist && (this.ModificationSourceFile == this.ModificationDestinationFile))
                this.NeedUpdate = false;
        }
    }
    #endregion Class DataFileInfo

    #region SelectProductDialog form
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
            this.Focus();
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

        #region ButtonOk_Click
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
        #endregion ButtonOk_Click

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
            string searchText = inputField.Text.ToUpper();
            if (!string.IsNullOrWhiteSpace(searchText)) {
                // Из запроса формируем массив запросов с разделением по пробелам
                string[] searchQueries = searchText.Split(new char[] {' '}, StringSplitOptions.RemoveEmptyEntries);
                
                List<string> resultOfSearch = this.allProducts.ToList<string>();
                foreach(string query in searchQueries) {
                    resultOfSearch = resultOfSearch.Where(str => str.Contains(query)).ToList<string>();
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
    #endregion SelectProductDialog form

    #region TreeOfProduct and TreeNode classes
    // Классы для формирования дерева изделия

    #region TreeOfProduct class

    private class TreeOfProduct {

        #region Fields and Properties

        public TreeNode HeadOfTree { get; set; }
        public static Table SpecTable { get; set; }
        public static Table MarchpTable { get; set; }
        public List<TreeNode> AllNodes { get; set; } = new List<TreeNode>();

        public bool isRoutesLoad { get; private set; } = false;
        public bool isPurchasedProductLoad { get; private set; } = false;

        #endregion Fields and Properties

        #region Constructors

        private TreeOfProduct() {}

        #endregion Constructors

        #region Generated Tree methods

        public static TreeOfProduct GenerateTree(string shifr) {
            if (SpecTable == null) {
                throw new Exception(
                        string.Format(
                            "При формировании дерева изделия '{0}' возникла ошибка. Отсутствует таблица SPEC",
                            shifr
                            )
                        );
            }

            // Создаем пустой экземпляр дерева для последующего его заполнения
            TreeOfProduct tree = new TreeOfProduct();

            // Находим корневой элемент и для него запускаем рекурсивное формирование состава
            TableRow rootElementRow = SpecTable.Rows.FirstOrDefault(row => row["shifr"] == shifr);
            if (rootElementRow == null) {
                throw new Exception(
                        string.Format(
                            "При формировании дерева изделия '{0}' возникла ошибка. В таблице SPEC отсутствует изделие с заданным шифром",
                            shifr
                            )
                        );
            }

            // Создаем элемент дерева и заполняем его основные параметры
            TreeNode rootNode = new TreeNode(tree);
            rootNode["shifr"] = rootElementRow["shifr"];
            rootNode["name"] = rootElementRow["naim"];
            rootNode.QuantityInParent = 1;
            // Добавляем эту ноду в дерево
            tree.HeadOfTree = rootNode;
            tree.AllNodes.Add(rootNode);

            GenerateTreeRecursively(rootNode, tree);
            
            return tree;
        }

        private static void GenerateTreeRecursively(TreeNode currentNode, TreeOfProduct tree) {

            // Производим поиск с строками всех дочерних элементов в таблице
            List<TableRow> childrenRows
                = SpecTable.Rows.Where(row => row["izd"] == currentNode["shifr"]).ToList<TableRow>();
            // Условие выхода из рекурсии
            if (childrenRows.Count == 0)
                return;

            // Обработка найденных дочерних элементов
            foreach (TableRow children in childrenRows) {
                // Проверяем, нет подключена ли уже такая нода
                TreeNode searchedNode = currentNode.ChildNodes.FirstOrDefault(node => node["shifr"] == children["shifr"]);

                if (searchedNode == null) {
                    // Если в подключенных нодах нет ноды с таким обозначением, создаем новую
                    TreeNode childNode = new TreeNode(tree, currentNode);
                    childNode["shifr"] = children["shifr"];
                    childNode["name"] = children["naim"];
                    childNode.QuantityInParent = int.Parse(children["prim"]);
                    // Подключаем новую ноду
                    currentNode.Add(childNode);
                    tree.AllNodes.Add(childNode);
                    // Вызываем фукнцию рекурсивно
                    GenerateTreeRecursively(childNode, tree);
                }
                else {
                    // В подключенных уже есть нода с таким обозначением, так что мы просто увеличиваем ее количество
                    searchedNode.QuantityInParent += int.Parse(children["prim"]);
                }
            }
        }

        #endregion Generated Tree methods

        #region Methods for generating excel tables

        #region Формирование таблицы с составом изделия

        public void FillDataForTreeOfProduct(ExcelTable exTable) {
            foreach (TreeNode node in this.AllNodes) {
                ExcelRow exRow = new ExcelRow();
                exTable.Add(exRow);

                exRow["Обозначение"] = node["shifr"];
                exRow["Наименование"] = node["name"];
                exRow["Применяемость"] = node.QuantityInTree.ToString();

                // Заполнение параметров, которых может не быть
                if (node.ContainsParameter("gost")) {
                    exRow["Стандарт"] = node["gost"];
                }
                if (node.ContainsParameter("type")) {
                    exRow["Признак покупного"] = node["type"];
                }
                if (node.ContainsParameter("route")) {
                    exRow["Маршрут"] = node["route"];
                }
            }
        }

        #endregion Формирование таблицы с составом изделия

        #region Формирование таблицы по стандартным изделиям

        public void FillDataForStIzd(ExcelTable exTable) {
            Dictionary<string, ExcelRow> alreadyExistDict = new Dictionary<string, ExcelRow>();

            foreach (TreeNode node in this.AllNodes) {
                // Проверяем, является ли данное изделие стандартным
                if (node["type"] == "Стандартное изделие") {
                    // Проверяем, не содержится ли данная нода в уже добавленных в таблицу
                    if (!alreadyExistDict.ContainsKey(node["shifr"])) {
                        ExcelRow exRow = new ExcelRow();
                        // Добавляем эту строку в словарь и в таблицу
                        alreadyExistDict.Add(node["shifr"], exRow);
                        exTable.Add(exRow);

                        // Производим заполнение информации
                        exRow["Изделие"] = this.HeadOfTree["shifr"];
                        exRow["Обозначение"] = node["shifr"];
                        exRow["Наименование"] = node["name"];
                        exRow["ГОСТ"] = node["gost"];
                        exRow["Применяемость"] = node.QuantityInTree.ToString();
                    }
                    // Если для данного изделия уже была создана строка, получаем его и модифицируем данные о применяемости
                    else {
                        int oldPrim = int.Parse(alreadyExistDict[node["shifr"]]["Применяемость"]);
                        alreadyExistDict[node["shifr"]]["Применяемость"] = (oldPrim + node.QuantityInTree).ToString();
                    }
                }
            }
        }

        #endregion Формирование таблицы по стандартным изделиям

        #region Формирование таблицы по материалам

        public void FillDataForMaterial(ExcelTable exTable, Dictionary<string, List<TableRow>> normDict, Dictionary<string, TableRow> klasmTable, Dictionary<int, string> unitsOfMeasurement) {
            //TODO Реализовать метод для формирования таблицы по материалам
            // Проходим по всем объектам дерева

            // Создаем контейнер для строк для того, чтобы суммировать количество по позициям с одинаковым кодом
            Dictionary<string, ExcelRow> containerForRows = new Dictionary<string, ExcelRow>();

            foreach (TreeNode node in this.AllNodes) {
                int quantity = node.QuantityInTree;

                

                if (normDict.ContainsKey(node["shifr"])) {
                    // Итерируемся через строки в таблице норм, относящимся к конкретному изделию
                    foreach (TableRow normDictRow in normDict[node["shifr"]]) {
                        // Получаем количество данной ноды в дереве для получения количества материала
                        // Случай, когда словарь еще не содержит данного материала
                        if (!containerForRows.ContainsKey(normDictRow["okp"])) {
                            // Создаем новую строку Excel таблицы
                            ExcelRow exRow = new ExcelRow();
                            // Добавляем строку в контейнер
                            containerForRows[normDictRow["okp"]] = exRow;

                            // Заполняем параметры строки
                            exRow["Обозначение"] = normDictRow["okp"]; // OKP материала
                            exRow["Наименование"] = klasmTable[normDictRow["okp"]]["name"]; // Наименование материала
                            exRow["Стандарт"] = klasmTable[normDictRow["okp"]]["gost"]; // Стандарт
                            // TODO Пока что непонятно, что делать с единицами измерения, так как они есть в нормах, а так же
                            // в данных матерала. А так же в справочнике с единицами измерения есть переводные коэффициенты,
                            // так что возможно их так же придется использовать. Пока что воспользуюсь единицами измерения, которые указаны в нормах
                            exRow["Единицы измерения"] = unitsOfMeasurement[int.Parse(normDictRow["edizm"])];// Единицы измерения
                            exRow["Применяемость"] = (quantity * decimal.Parse(normDictRow["nmat"])).ToString();// Количество
                            exRow["Вид"] = klasmTable[normDictRow["okp"]]["vid"]; // Вид материала 

                        }
                        // Случай, когда словарь уже содержит данный материал
                        else {
                            // В этом случае мы вычисляем количество данного материала для данного изделия и прибавляем это значение к вычисленному раннее для других изделий
                            containerForRows[normDictRow["okp"]]["Применяемость"] = (decimal.Parse(containerForRows[normDictRow["okp"]]["Применяемость"]) + (quantity * decimal.Parse(normDictRow["nmat"]))).ToString();
                        }
                    }
                }
            }

            // Добавляем сформированные строки в таблицу excel для формирования итоговых значений
            foreach (KeyValuePair<string, ExcelRow> kvp in containerForRows) {
                exTable.Add(kvp.Value);
            }
        }

        #endregion Формирование таблицы по материалам
        
        #endregion Methods for generating excel tables

        public void Print() {
            if (this.HeadOfTree == null) {
                Console.WriteLine("Данное дерево не содержит элементов");
                return;
            }
            
            Console.WriteLine("\nСформированное дерево состава для '{0}'", this.HeadOfTree["shifr"]);
            RecursivePrint(this.HeadOfTree, string.Empty);
            Console.WriteLine();

        }

        private void RecursivePrint(TreeNode currentNode, string indent) {
            Console.WriteLine(string.Format(
                        "{0}{1} ; {2} ; {3} ; {4}",
                        indent,
                        currentNode["shifr"],
                        currentNode["name"],
                        currentNode.QuantityInParent.ToString(),
                        currentNode.QuantityInTree.ToString()
                        ));

            if (currentNode.ChildNodes.Count == 0)
                return;
            
            // Если есть дочерние элементы, производим дальнейшее погружение вглубь
            indent += "-";
            foreach(TreeNode childNode in currentNode.ChildNodes) {
                RecursivePrint(childNode, indent);
            }
        }

        #region Методы для сбора дополнительной информации о составе изделия

        #region GetInfoAboutPurchaseProducts
        // Метод для определения покупных изделий в составе
        public void GetInfoAboutPurchaseProducts() {

            foreach (TreeNode node in this.AllNodes) {
                // Получаем номер первого цеха в изготовлении
                string izg = string.Empty;
                if (this.isRoutesLoad) {
                    izg = node["route"].Split('-')[0];
                }
                else {
                    // Для начала пытаемся получить список цехозаходов для данного изделия
                    List<TableRow> shopCalls = MarchpTable.Rows.Where(row => (row["shifr"] == node["shifr"]) && (row["nper"] == "1")).ToList<TableRow>();
                    if (shopCalls.Count == 0) {
                        node["type"] = "Ошибка";
                        continue;
                    }
                    else if (shopCalls.Count == 1) {
                        izg = shopCalls[0]["izg"];
                    }
                    else {
                        izg = shopCalls.FirstOrDefault(row => row["norm"] == "1")["izg"];
                    }

                    /*
                    try {
                        izg = MarchpTable
                            .Rows
                            .FirstOrDefault(row => 
                                    (row["shifr"] == node["shifr"]) &&
                                    (row["nper"] == "1") &&
                                    (row["norm"] != "0")
                                    )["izg"];
                    }
                    catch {
                        throw new Exception(string.Format("Для изделия '{0}'при определении признака 'Покупное' возникла ошибка", node["shifr"]));
                    }
                    */
                }

                if (!arrayOfUnits.Contains(izg)) {
                    node["type"] = "Покупное";
                }
                else {
                    node["type"] = "Не покупное";
                }
            }

            // Определяем позиции, которые входят в покупные
            foreach (TreeNode node in this.AllNodes.Where(node => node["type"] == "Покупное")) {
                SetPurchaseToChilderRecursively(node);
            }

            this.isPurchasedProductLoad = true;
        }

        private void SetPurchaseToChilderRecursively(TreeNode node) {
            foreach (TreeNode child in node.ChildNodes) {
                if (child["type"] == "Не покупное") {
                    child["type"] = "Входит в покупное";
                }
                SetPurchaseToChilderRecursively(child);
            }
        }

        #endregion GetInfoAboutPurchaseProducts

        #region GetInfoAboutRoutes
        // Метод для получения маршрутов изготовления для элементов состава
        public void GetInfoAboutRoutes() {
            // TODO Переписать метод с учетом того, что маршрута с номером 1 может не быть (тогда основной маршрут будет нулевым)
            foreach (TreeNode node in this.AllNodes) {
                // Получаем сначала все цехозаходы, которые относятся к данному изделию
                var shopCalls = MarchpTable.Rows.Where(row => row["shifr"] == node["shifr"]);
                if (shopCalls.Where(row => row["norm"] == "1").Count() > 0) {
                    node["route"] = string.Join("-", shopCalls
                        .Where(row => row["norm"] == "1")
                        .OrderBy(row => int.Parse(row["nper"]))
                        .Select(row => row["izg"])
                        );
                }
                else {
                    // Проверяем, есть ли маршрут с norm 0
                    if (shopCalls.Where(row => row["norm"] == "0").Count() > 0) {
                        node["route"] = string.Join("-", shopCalls
                                .Where(row => row["norm"] == "0")
                                .OrderBy(row => int.Parse(row["nper"]))
                                .Select(row => row["izg"])
                                );
                    }
                    else {
                        node["route"] = "Ошибка";
                    }
                }
            }
            this.isRoutesLoad = true;
        }

        #endregion GetInfoAboutRoutes

        #endregion Методы для сбора дополнительной информации о составе изделия
    }

    #endregion TreeOfProduct class

    #region TreeNode class

    private class TreeNode {

        #region Fields and Properties

        public TreeOfProduct Tree { get; private set; }
        public TreeNode ParentNode { get; private set; }
        public List<TreeNode> ChildNodes { get; private set; } = new List<TreeNode>();
        private Dictionary<string, string> Parameters { get; set; } = new Dictionary<string, string>();
        public int QuantityInParent { get; set; }
        public int QuantityInTree => GetQuantityInTree();

        #endregion Fields and Properties

        #region Constructors

        public TreeNode(TreeOfProduct tree, TreeNode parent = null) {
            this.ParentNode = parent;
            this.Tree = tree;
        }

        public bool ContainsParameter(string nameOfParameter) {
            return this.Parameters.ContainsKey(nameOfParameter);
        }

        #endregion Constructors

        public void Add(TreeNode node) {
            this.ChildNodes.Add(node);
        }

        public bool Contains(TreeNode node) {
            return this.ChildNodes.Contains(node);
        }

        public void Delete(TreeNode node) {
            // TODO Реализовать функцию удаления ноды из состава
            // В этом случае так же нужно будет удалить ссылку из родительской ноды, а
            // так же пройтись по всем дочерним нодам и так же занулить все ссылки
            // Возможно так же придется реализовать метод Dispose()
        }

        public string GetChainToRootElement(string parameter) {
            string result = string.Empty;
            TreeNode node = this;
            
            try {
                string test = node[parameter];
            }
            catch {
                throw new Exception(string.Format("Параметр '{0}' отсутствует в узлах дерева", parameter));
            }
            
            while (true) {
                node = node.ParentNode;
                if (node != null)
                    result += string.Format("{0} -> {1}", node[parameter], result);
                else
                    break;
            }
            return result;
        }

        // Индексатор для обращения к параметрам
        public string this[string key] {
            get {
                return this.Parameters[key];
            }
            set {
                this.Parameters[key] = value;
            }
        }

        private int  GetQuantityInTree() {
            // Проверяем, если ли у данной ноды родитель
            if (this.ParentNode != null) {
                return (int)(this.QuantityInParent * this.ParentNode.QuantityInTree);
            }
            else
                // Если родителя нет, возвращаем применяемость.
                // Для изделия вернего уровня это количество по умолчанию равно единице
                return this.QuantityInParent;
        }
    }

    #endregion TreeNode class

    #endregion TreeOfProduct and TreeNode classes

    #region Table classes
    // Классы для загрузки информации из таблиц FoxPro в оперативную память

    #region Table class
    // Класс для чтения таблиц из FoxPro

    private class Table {

        #region Fields and Properties

        public string Name { get; private set; } // Название таблицы
        public string PathToTempDirectoryFoxProDb { get; set; }
        public string PathToDBFile { get; private set; }
        public List<TableRow> Rows { get; set; } = new List<TableRow>();
        public List<int> ErrorsId { get; private set; } = new List<int>();
        public int Count => Rows.Count; // Получение количества строк в таблице
        public string Status => string.Format(
                "Прочитана таблица {0} ({1} строк, {2} ошибок)\n",
                this.Name,
                this.Count,
                this.ErrorsId.Count
                );

        #endregion Fields and Properties

        #region Constructor

        public Table(string name, string pathToDirectory) {
            this.Name = name;
            this.PathToTempDirectoryFoxProDb = pathToDirectory;
            this.PathToDBFile = this.GetPathToDbfFile(name);

            this.ReadAllTable();
        }

        #endregion Constructor

        #region Add()

        public void Add(TableRow row) {
            this.Rows.Add(row);
        }

        #endregion Add()

        #region ReadAllTable()
        // Общий метод для чтения dbf таблиц

        public void ReadAllTable() {
            DbfDataReader.DbfDataReaderOptions options = new DbfDataReader.DbfDataReaderOptions() {
                SkipDeletedRecords = true,
                Encoding = System.Text.Encoding.GetEncoding(1251)
            };

            DbfDataReader.DbfDataReader dataReader = new DbfDataReader.DbfDataReader(this.PathToDBFile, options);

            switch (this.Name) {
                case "spec":
                    ReadSpecTable(dataReader);
                    break;
                case "trud":
                    throw new Exception("Работа с таблицей trud еще не реализована");
                    break;
                case "klas":
                    ReadKlasTable(dataReader);
                    break;
                case "klasm":
                    ReadKlasmTable(dataReader);
                    break;
                case "kat_ediz":
                    ReadKatEdizTable(dataReader);
                    break;
                case "norm":
                    ReadNormTable(dataReader);
                    break;
                case "marchp":
                    ReadMarchpTable(dataReader);
                    break;
                default:
                    throw new Exception(
                            string.Format(
                                "Ошибка при чтении таблицы. '{0}' не распознана",
                                this.Name
                                )
                            );
            }
        }

        #region Методы для обработки различных таблиц

        #region ReadSpecTable method
        // Метод для чтения таблицы с составом изделия

        private void ReadSpecTable(DbfDataReader.DbfDataReader dataReader) {
            int counter = 1;
            while (dataReader.Read()) {
                TableRow newRow = new TableRow(counter++);
                try {
                    newRow["shifr"] = dataReader.GetString(1);
                    newRow["naim"] = dataReader.GetString(2);
                    newRow["prim"] = dataReader.GetInt32(3).ToString();
                    newRow["izd"] = dataReader.GetString(4);
                    newRow["shifr+naim"] = string.Format("{0}  ^  {1}", newRow["shifr"], newRow["naim"]);
                }
                catch {
                    this.ErrorsId.Add(counter - 1);
                    continue;
                }
                this.Rows.Add(newRow);
            }
        }

        #endregion ReadSpecTable method

        #region ReadTrudTable method
        // Метод для чтения таблицы с технологическими операциями

        private void ReadTrudTable(DbfDataReader.DbfDataReader dataReader) {
            int counter = 1;
        }

        #endregion ReadTrudTable method

        #region ReadKlasTable method
        // Метод для чтения таблицы с стандартными изделиями

        private void ReadKlasTable(DbfDataReader.DbfDataReader dataReader) {
            int counter = 1;
            while (dataReader.Read()) {
                TableRow newRow = new TableRow(counter++);
                try {
                    newRow["okp"] = dataReader.GetString(0); // Обозначение (код ОКП)
                    newRow["name"] = dataReader.GetString(1); // Наименование
                    newRow["gost"] = dataReader.GetString(2); // Стандарт
                    newRow["vid"] = dataReader.GetString(9); // Вид стандартного изделия
                    try {
                        newRow["edizm"] = dataReader.GetInt32(3).ToString(); // Код единиц измерения
                    }
                    catch {
                        newRow["edizm"] = "0";
                    }
                }
                catch {
                    this.ErrorsId.Add(counter - 1);
                    continue;
                }
                this.Rows.Add(newRow);
            }
        }

        #endregion ReadKlasTable method

        #region ReadKlasmTable method
        // Метод для чтения таблицы с материалами

        private void ReadKlasmTable(DbfDataReader.DbfDataReader dataReader) {
            int counter = 1;
            while (dataReader.Read()) {
                TableRow newRow = new TableRow(counter++);
                try {
                    newRow["okp"] = dataReader.GetString(0); // Обозначение материала (код ОКП)
                    newRow["name"] = dataReader.GetString(1); // Наименование материала
                    newRow["gost"] = dataReader.GetString(2); // Стандарт
                    newRow["vid"] = dataReader.GetString(9); // Вид материала

                    try {
                        newRow["edizm"] = dataReader.GetInt32(3).ToString(); // Код единиц измерения
                    }
                    // Обработка случай, когда колонка не содержит данных о единицах измерения
                    catch {
                        newRow["edizm"] = "0";
                    }
                }
                catch {
                    this.ErrorsId.Add(counter - 1);
                    continue;
                }
                this.Rows.Add(newRow);
            }
        }

        #endregion ReadKlasmTable method

        #region ReadNormTable method
        // Метод для чтения таблицы Norm

        private void ReadNormTable(DbfDataReader.DbfDataReader dataReader) {
            int counter = 1;
            while (dataReader.Read()) {
                TableRow newRow = new TableRow(counter++);
                try {
                    newRow["shifr"] = dataReader.GetString(0);
                    newRow["okp"] = dataReader.GetString(1);
                    newRow["edizm"] = dataReader.GetInt32(2).ToString();
                    newRow["nmat"] = dataReader.GetDecimal(3).ToString();
                }
                catch {
                    this.ErrorsId.Add(counter - 1);
                    continue;
                }
                this.Rows.Add(newRow);
            }
        }

        #endregion ReadNormTable method

        #region ReadKatEdizTable method
        // Метод для чтения таблицы с единицами измерения

        private void ReadKatEdizTable(DbfDataReader.DbfDataReader dataReader) {
            int counter = 1;
            while (dataReader.Read()) {
                TableRow newRow = new TableRow(counter++);
                try {
                    newRow["kod"] = dataReader.GetInt32(0).ToString(); // Код единицы измерения
                    // Так как название единицы измерения здесь в другой кодировке
                    newRow["name"] = System.Text.Encoding.GetEncoding(866).GetString(System.Text.Encoding.GetEncoding(1251).GetBytes(dataReader.GetString(1))); // Наименование единицы измерения
                }
                catch {
                    this.ErrorsId.Add(counter - 1);
                    continue;
                }
                this.Rows.Add(newRow);
            }
        }

        #endregion ReadKatEdizTable method

        #region ReadMarchpTable method
        // Метод для чтения таблицы с маршрутами

        private void ReadMarchpTable(DbfDataReader.DbfDataReader dataReader) {
            //System.Diagnostics.Debugger.Launch();

            int counter = 1;
            while (dataReader.Read()) {
                TableRow newRow = new TableRow(counter++);
                try {
                    newRow["shifr"] = dataReader.GetString(0); // Обозначение изделия
                    try {
                        newRow["nper"] = dataReader.GetInt32(1).ToString(); // Номер перехода
                    }
                    catch {
                        newRow["nper"] = "99"; // Присваеваем такое значение временно
                    }
                    newRow["izg"] = dataReader.GetString(2); // Изготовитель данного перехода
                    newRow["per1"] = dataReader.GetString(3); // Номер перехода для данного цеха изготовителя
                    newRow["potr"] = dataReader.GetString(4); // Потребитель данного цехоперехода
                    newRow["per2"] = dataReader.GetString(5); // Номер перехода для данного цеха потребителя
                    newRow["sdat"] = dataReader.GetString(6); 
                    try {
                        newRow["vrem"] = dataReader.GetDecimal(7).ToString(); // Длительность
                    }
                    catch {
                        newRow["vrem"] = "0";
                    }
                    newRow["norm"] = dataReader.GetString(8); // Параметр, который отвечает за то, действует ли данный цехопереход
                }
                catch {
                    this.ErrorsId.Add(counter - 1);
                    continue;
                }
                this.Rows.Add(newRow);
            }
            //System.Diagnostics.Debugger.Break();
        }

        #endregion ReadMarchpTable method

        #endregion Методы для обработки различных таблиц

        #endregion ReadAllTable()

        #region GetPathToDBFile()
        // Метод для формирования пути к файлу с исходными данными
        private string GetPathToDbfFile(string nameOfFile) {
            // TODO Перенести логину получения пути файла из внешнего метода в данный метод
            nameOfFile = nameOfFile.ToLower();
            if (!nameOfFile.EndsWith(".dbf"))
                nameOfFile += ".dbf";
            return Path.Combine(this.PathToTempDirectoryFoxProDb, nameOfFile);
        }
        #endregion GetPathToDBFile()

        #region PrintAllRowsInConsole()
        // TODO Переработать данный метод для других типов таблиц
        // Данный метод по большей части требуется для отладки, так как он по сути просто отображает в консоли 
        // все те данные, которые были загружены. 
        // Причем этот метод в данный момент корректно будет работать только с таблицей Spec
        public void PrintAllRowsInConsole() {
            switch (this.Name) {
                case "spec" :
                    foreach (TableRow row in this.Rows) {
                        Console.WriteLine(
                                string.Format(
                                    "{0} ; {1} ; {2} ; {3} ; {4}",
                                    row.Id,
                                    row["naim"],
                                    row["shifr"],
                                    row["prim"],
                                    row["izd"]
                                    )
                                );
                    }
                    break;
                default:
                    Console.WriteLine("Данная таблица не поддерживает вывода информации на экран");
                    break;
            }
        }
        #endregion PrintAllRowsInConsole()

        #region FindRow()
        // Метод для поиска записей в таблице.

        public TableRow FindRow(string parameter, string valueOfParameter) {
            return this.Rows.FirstOrDefault(row => row[parameter] == valueOfParameter);
        }
        #endregion FindRow()

        #region GetAllDataFromColumn()
        // Возвращает всю колонку таблицы в виде списка строк.
        // Данный метод удобен для получения сводных данных по обозначению
        public List<string> GetAllDataFromColumn(string nameOfParameter) {
            // Метод для получения данных, содержащихся в одной колонке
            List<string>result = new List<string>();
            foreach (TableRow row in this.Rows)
                result.Add(row[nameOfParameter]);
            return result;
        }
        #endregion GetAllDataFromColumn()

    }
    #endregion Table class

    #region TableRow class

    private class TableRow {

        public int Id { get; private set; }
        private Dictionary<string, string> Parameters { get; set; } = new Dictionary<string, string>();

        public TableRow(int id) {
            this.Id = id;
        }

        // Индексатор
        public string this[string key] {
            get {
                return this.Parameters[key];
            }

            set {
                Parameters[key] = value;
            }
        }
    }

    #endregion TableRow class

    #endregion Table classes

    #region ExcelTableGenerator
    // Классы для генерации Excel файлов

    #region ExcelTable class

    public class ExcelTable {

        #region Fields and Properties

        private string PathToDirectory { get; set; }
        private string NameOfFile { get; set; }
        private string PathToFile { get; set; }
        private List<ExcelRow> Rows { get; set; } = new List<ExcelRow>();
        public Dictionary<string, int> Columns { get; set; } = new Dictionary<string, int>();
        private ExcelTableOptions Options { get; set; }

        #endregion Fields and Properties

        #region Constructors

        public ExcelTable(string pathToDirectory, string nameOfFile, ExcelTableOptions options = null) {
            this.PathToDirectory = pathToDirectory;
            this.NameOfFile = nameOfFile;
            this.PathToFile = Path.Combine(pathToDirectory, string.Format("{0}.xlsx", nameOfFile));

            if (options == null) {
                // Случай, когда опции не заданы
                this.Options = new ExcelTableOptions();
            }
            else
                this.Options = options;
        }

        #endregion Constructors

        public void Add(ExcelRow row) {
            this.Rows.Add(row);
        }

        #region Generate()

        public void Generate(string[] namesOfColumns) {
            // Данный метод генерирует таблицу с заданной последовательностью колонок.
            
            // Производим проверку на правильность входных данных
            foreach (string columnName in namesOfColumns) {
                if (!this.Columns.ContainsKey(columnName))
                    throw new Exception(
                            string.Format(
                                "При генерации таблицы возникла ошибка. ExcelTable не содержит информации о колонке '{0}'",
                                columnName
                                )
                            );
            }

            // Производим подготовительные работы (создаем промежуточные классы, необходимые для работы)
            using (SpreadsheetDocument document =
                    SpreadsheetDocument.Create(this.PathToFile, SpreadsheetDocumentType.Workbook)) {
                WorkbookPart wp = document.AddWorkbookPart();
                wp.Workbook = new Workbook();
                WorksheetPart wsp = wp.AddNewPart<WorksheetPart>();

                // Добавляем в документ новый рабочий лист
                wsp.Worksheet = new Worksheet(new SheetData());

                // Задаем параметры колонок
                Columns lstColumns = new Columns();
                int colCount = namesOfColumns.Length;

                for (int i = 1; i < colCount; i++) {
                    lstColumns.Append(
                            new Column() {
                            Min = (uint)i,
                            Max = (uint)colCount,
                            Width = this.Columns[namesOfColumns[i - 1]],
                            CustomWidth = true
                            }
                            );
                }

                // Добавляем колонки в документ
                wsp.Worksheet.InsertAt(lstColumns, 0);
                
                // Добавляем автофильтр в документ при необходимости
                if (this.Options.UseAutoFilter == true) {
                    // Добавляем фильтр на все колонки, которые будут добавляться в документ
                    AutoFilter af = new AutoFilter() { Reference = string.Format("1:{0}", colCount.ToString())};
                    wsp.Worksheet.Append(af);
                }

                // Создаем лист в книге
                Sheets sheets = wp.Workbook.AppendChild(new Sheets());
                Sheet sheet = new Sheet() { Id = wp.GetIdOfPart(wsp), SheetId = 1, Name = this.Options.NameOfSheet };
                sheets.Append(sheet);
                SheetData sd = wsp.Worksheet.GetFirstChild<SheetData>();

                // Добавляем данные в таблицу

                // Создаем первую строку для того, чтобы разместить в ней шапку документа
                uint counter = 1; // Переменная, которая будет хранить текущей номер строки
                Row row = new Row() { RowIndex = counter++ };
                sd.Append(row);

                // Размещаем шапку документа в первой строке
                InsertHeader(row, namesOfColumns);

                // Размещаем остальные табличные данные
                foreach (ExcelRow exRow in this.Rows) {
                    // Создаем новую строк уи подкючаем ее к документу
                    row = new Row() { RowIndex = counter++ };
                    sd.Append(row);
                    InsertRow(row, namesOfColumns, exRow);
                }

                // Сохраняем документ
                wp.Workbook.Save();
            }
        }

        #region Helper methods

        private void InsertHeader(Row row, string[] headerNames) {
            for (int i = 1; i <= headerNames.Length; i++) {
                InsertCell(row, i, headerNames[i - 1], CellValues.String);
            }
        }

        private void InsertRow(Row row, string[] namesOfColumns, ExcelRow exRow) {
            int counter = 1;
            foreach (string param in namesOfColumns) {
                InsertCell(row, counter++, exRow[param], CellValues.String);
            }
        }

        private void InsertCell(Row row, int index, string value, CellValues type) {
            Cell refCell = null;
            Cell newCell = new Cell() { CellReference = string.Format("{0}:{1}", row.RowIndex.ToString(), index.ToString()) };
            row.InsertBefore(newCell, refCell);

            // Присваиваем значение ячейке
            newCell.CellValue = new CellValue(value);
            newCell.DataType = new EnumValue<CellValues>(type);
        }

        #endregion Helper methods

        #endregion Generate()
    }

    #endregion ExcelTable class

    #region ExcelRow class

    public class ExcelRow {
        private Dictionary<string, string> Data { get; set; } = new Dictionary<string, string>();

        public ExcelRow() {
        }

        public string this[string key] {
            get {
                if (!this.Data.ContainsKey(key))
                    throw new Exception(string.Format("Строка ExcelRow не содержит ключа '{0}'", key));
                return this.Data[key];
            }
            set {
                this.Data[key] = value;
            }
        }

        public void Add(string key, string value) {
            this.Data.Add(key, value);
        }

        public bool ContainsKey(string key) {
            return this.Data.ContainsKey(key);
        }

        public bool ContainsValue(string value) {
            return this.Data.ContainsValue(value);
        }

        public override string ToString() {
            string result = string.Empty;
            foreach (string key in Data.Keys.OrderBy(key => key)) {
                result += string.Format("{0}: {1}; ", key, Data[key]);
            }
            return result;
        }
    }

    #endregion ExcelRow class

    #region ExcelTableOptions class

    public class ExcelTableOptions {
        public bool UseAutoFilter { get; set; } = false;
        public string NameOfSheet { get; set; } = "Выгрузка из Fox";
    }

    #endregion ExcelTableOptions class

    #endregion ExcelTableGenerator

    #endregion Service classes
    
}
