using System;
using System.IO;
using System.Text; // Для работы с кодировкой
using System.Linq;
using System.Collections;
using System.Collections.Generic;
using System.Windows.Forms;

using System.Threading;
using System.Threading.Tasks;

using TFlex.DOCs.Model.Macros;
using TFlex.DOCs.Model.Macros.ObjectModel;
using TFlex.DOCs.Model.Structure;
using TFlex.DOCs.Model.Search;
using TFlex.DOCs.Model.References;

using DbfDataReader;

/*
Так же для работы данного макроса потребуется подключение дополнительных библиотек:

DbfDataReader.dll - для чтения dbf файлов
System.Data.dll - необходим для работы DbfDataReader

*/

/*
TODO: Написать графический интерфейс для запроса у пользователя параметров поиска
TODO: Реализовать параллельный поиск по нескольким таблицам одновременно
TODO: Реализовать поиск сразу по всем колонкам, если колонка не указана пользователем
TODO: Реализовтаь поиск по вхождению
*/

public class Macro : MacroProvider {
    public Macro(MacroContext context)
        : base(context) {
        }

    private const string sourcePath = @"\\fs\FoxProDB\COMDB\PROIZV";
    private readonly string backupPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "Таблицы FoxPro");
    private static readonly string[] metaFilesExtensions = new string[] {
        ".bak",
        ".cdx",
        ".dbc",
        ".dct",
        ".dcx",
        ".fpt",
        ".tbk"
    };

    public override void Run() {

        Backuper backuper = new Backuper(sourcePath, backupPath, this);

        // Получаем словарь с доступными dbf таблицами
        TablesHandler handler = new TablesHandler(backuper.GetPathDbfFiles(), this);

        // Запрашиваем у пользователя значения параметром поиска, производим поиск и выводим результаты поиска в файл
        handler.AskAndSearch();

        //Message("Доступные dbf таблицы", handler.ToString());
        //Message("Доступные колонки", string.Join("\n", handler.AllColumns.Keys()));
        //Message("Все таблицы, в которых есть колонка SHIFR", string.Join("\n", handler.AllColumns["SHIFR"].Name));

        Message("Информация", "Работа макроса завершена");
        
    }

    #region Classes

    private class Backuper {
    // Класс, искапсулирующий всю логику по проведению бэкапа базы данных FoxPro

        private DbRepository SourceRepository { get; set; }
        private DbRepository BackupRepository { get; set; }
        private RepositoryMissedFiles MissedFiles { get; set; }
        private Macro MacroProvider { get; set; }

        public Backuper (string sourcePath, string backupPath, Macro provider) {
            // Инициализируем репозитории
            this.MacroProvider = provider;
            bool showInfo = false;

            // Запрашиваем параметры у пользователя
            GetParametersFromUser(ref sourcePath, ref backupPath, ref showInfo);

            try {
                this.SourceRepository = new DbRepository(sourcePath);
                this.BackupRepository = new DbRepository(backupPath);
            }
            catch (Exception e) {
                throw new Exception(e.Message);
            }

            // Выводим пользователю сообщение об инициализированных репозиториях
            if (showInfo) {
                string message = string.Format("{0}\n\n{1}", this.SourceRepository.ToString(), this.BackupRepository.ToString());
                this.MacroProvider.Message("Информация об инициализированных репозиториях", message);
            }

            // Запускаем сравнение репозиториев
            this.MissedFiles = this.SourceRepository.Compare(this.BackupRepository);

            // Запрашиваем у пользователя подтверждение о начале копирования
            this.MissedFiles.AskAndDownload(this.MacroProvider);
        }

        private void GetParametersFromUser(ref string sourcePath, ref string backupPath, ref bool showInfo) {
            InputDialog dialog = new InputDialog(this.MacroProvider.Context, "Укажите параметры для бэкапа базы данных FoxPro");
            dialog.AddString("Источник", sourcePath);
            dialog.AddString("Бэкап", backupPath);
            dialog.AddFlag("Показывать информацию", showInfo);

            if (dialog.Show()) {

                // Производим проверку пути источника
                if (sourcePath != dialog["Источник"])
                    if (!Directory.Exists(dialog["Источник"]))
                        throw new Exception(string.Format("Директория, указанная для источника не существует:\n{0}", dialog["Источник"]));
                    else
                        sourcePath = dialog["Источник"];

                // Производим проверку пути бэкапа
                if (backupPath != dialog["Бэкап"])
                    backupPath = dialog["Бэкап"];

                if (!Directory.Exists(backupPath))
                    Directory.CreateDirectory(backupPath);

                // Заполняем флаг, показывать ли информацию
                showInfo = dialog["Показывать информацию"];
            }
        }

        public Dictionary<string, string> GetPathDbfFiles() {
            Dictionary<string, string> result = new Dictionary<string, string>(this.BackupRepository.CountDbf);

            foreach (KeyValuePair<string, FileInfo> kvp in this.BackupRepository.DbfFiles)
                result[kvp.Key.Split('.')[0]] = kvp.Value.FullName;

            return result;
        }

        public override string ToString() {
            if (this.MissedFiles.Count == 0)
                return "Локальный репозиторий полностью соответствует удаленному репозиторию";

            string template = "В локальном репозитории найдено {0} несоответствий. Из них:\n{1} - отсутствующие файлы\n{2} - устаревшие файлы";

            return string.Format(template, this.MissedFiles.Count, this.MissedFiles.CountMissed, this.MissedFiles.CountOutdated);
        }

        private class DbRepository {
            public TypeRepository Type { get; private set; } = TypeRepository.None;
            //public StatusRepository Status { get; private set; } = StatusRepository.Empty;
            public string Dir { get; private set; }
            public Dictionary<string, FileInfo> DbfFiles { get; private set; }
            public Dictionary<string, FileInfo> MetaFiles { get; private set; }
            public Dictionary<string, FileInfo> OtherFiles { get; private set; }
            public Dictionary<string, FileInfo> AllFiles { get; private set; }
            public int Count => this.AllFiles.Count;
            public int CountDbf => this.DbfFiles.Count;
            public int CountMeta => this.MetaFiles.Count;
            public int CountOther => this.OtherFiles.Count;

            public DbRepository(string pathToDir) {
                if (Directory.Exists(pathToDir)) {
                    // Присваиваем исходные параметры
                    this.Dir = pathToDir;

                    // Определяем тип репозитория
                    this.Type = Path.GetPathRoot(this.Dir).EndsWith(@"\") ? TypeRepository.Backup : TypeRepository.Source;

                    // Получаем данные по файлам, содержащимся в репозитории
                    UpdateInfo();
                }
                else
                    throw new Exception(string.Format("Директория, переданная в качестве месторасположения репозитория с файлами базы данных\n{0}\n отсутствует в системе", pathToDir));
            }

            public DbRepository(string pathToDir, TypeRepository type) {
                this.Dir = pathToDir;
                this.Type = type;
                UpdateInfo();
            }

            public void UpdateInfo() {

                string[] files = Directory.GetFiles(this.Dir);

                this.DbfFiles = new Dictionary<string, FileInfo>(files.Length);
                this.MetaFiles = new Dictionary<string, FileInfo>(files.Length);
                this.OtherFiles = new Dictionary<string, FileInfo>(files.Length);
                this.AllFiles = new Dictionary<string, FileInfo>(files.Length);

                foreach (string file in files) {
                    FileInfo fileInfo = new FileInfo(file.ToLower());
                    if (fileInfo.Extension == ".dbf") {
                        this.DbfFiles[fileInfo.Name] = fileInfo;
                        this.AllFiles[fileInfo.Name] = fileInfo;
                        continue;
                    }
                    if (metaFilesExtensions.Contains(fileInfo.Extension)) {
                        this.MetaFiles[fileInfo.Name] = fileInfo;
                        this.AllFiles[fileInfo.Name] = fileInfo;
                        continue;
                    }
                    this.OtherFiles[fileInfo.Name] = fileInfo;
                    this.AllFiles[fileInfo.Name] = fileInfo;
                }
            }

            public override string ToString() {
                return string.Format(
                        "Путь директории: {0}\n" +
                        "Тип репозитория: {1}\n" +
                        "Количество DBF файлов: {2}\n" +
                        "Количество Мета файлов: {3}\n" +
                        "Количество Остальных файлов: {4}\n" +
                        "Общее количество файлов: {5}\n",
                        this.Dir,
                        this.Type.ToString(),
                        this.CountDbf.ToString(),
                        this.CountMeta.ToString(),
                        this.CountOther.ToString(),
                        this.Count.ToString()
                        );
            }

            public RepositoryMissedFiles Compare(DbRepository otherRepository) {
                DbRepository source;
                DbRepository backup;

                // Определяем, какой репозиторий с каким будет сравниваться
                if (otherRepository.Type == TypeRepository.Source) {
                    source = otherRepository;
                    backup = this;
                }
                else {
                    source = this;
                    backup = otherRepository;
                }
                
                RepositoryMissedFiles rpf = new RepositoryMissedFiles(source, backup);

                // Проверяем на отсутстующие dbf файлы
                foreach (KeyValuePair<string, FileInfo> kvp in source.DbfFiles) {
                    if (!backup.DbfFiles.ContainsKey(kvp.Key))
                        rpf.AddMissedFile(kvp.Key);
                    else
                        if (kvp.Value.LastWriteTime != backup.DbfFiles[kvp.Key].LastWriteTime)
                            rpf.AddOutdatedFile(kvp.Key, kvp.Value.LastWriteTime.Subtract(backup.DbfFiles[kvp.Key].LastWriteTime));
                }
                
                // Проверяем на отсутствующие meta файлы
                foreach (KeyValuePair<string, FileInfo> kvp in source.MetaFiles) {
                    if (!backup.MetaFiles.ContainsKey(kvp.Key))
                        rpf.AddMissedFile(kvp.Key);
                    else
                        if (kvp.Value.LastWriteTime != backup.MetaFiles[kvp.Key].LastWriteTime)
                            rpf.AddOutdatedFile(kvp.Key, kvp.Value.LastWriteTime.Subtract(backup.MetaFiles[kvp.Key].LastWriteTime));
                }
                
                // Проверяем на отсутствующие остальные файлы
                foreach (KeyValuePair<string, FileInfo> kvp in source.OtherFiles) {
                    if (!backup.OtherFiles.ContainsKey(kvp.Key))
                        rpf.AddMissedFile(kvp.Key);
                    else
                        if (kvp.Value.LastWriteTime != backup.OtherFiles[kvp.Key].LastWriteTime)
                            rpf.AddOutdatedFile(kvp.Key, kvp.Value.LastWriteTime.Subtract(backup.OtherFiles[kvp.Key].LastWriteTime));
                }

                return rpf;
            }

        }

        // Класс, содержащий информацию по отсутстувующим и устаревшим файлав в репозитории бэкапе относительно источника
        private class RepositoryMissedFiles {
            public DbRepository Source { get; private set; }
            public DbRepository Backup { get; private set; }

            public List<string> MissedFiles { get; private set; }
            public Dictionary<string, TimeSpan> OutdatedFiles { get; private set; }

            public int CountMissed => this.MissedFiles.Count;
            public int CountOutdated => this.OutdatedFiles.Count;
            public int Count => this.MissedFiles.Count + this.OutdatedFiles.Count;

            public RepositoryMissedFiles(DbRepository source, DbRepository backup, int capacity = 256) {
                if (source.Type != TypeRepository.Source)
                    throw new Exception(string.Format("Для создания класса {0} требуется файловый репозиторий типа 'Source', а передан {1}", this.GetType().Name, source.Type.ToString()));
                this.Source = source;
                this.Backup = backup;
                this.MissedFiles = new List<string>(capacity);
                this.OutdatedFiles = new Dictionary<string, TimeSpan>(capacity);
            }

            public void AddMissedFile(string name) {
                if (this.MissedFiles.Contains(name))
                    throw new Exception(string.Format("Объект с названием '{0}' уже находился в списке пропущенный файлов", name));
                this.MissedFiles.Add(name);
            }

            public void AddOutdatedFile(string name, TimeSpan span) {
                if (this.OutdatedFiles.ContainsKey(name))
                    throw new Exception(string.Format("Объект с названием '{0}' уже находился в списке устаревших файлов", name));
                this.OutdatedFiles[name] = span;
            }

            public void AskAndDownload(Macro provider) {
                // TODO реализовать метод запроса, какие файлы требуется скичивать в локальный репозиторий
                // Спросить об обновлении отсутствующих файлов
                List<string> tables = new List<string>();

                // Если есть отстутсвующие файлы, спросить, требуется ли их скачивать
                if (this.CountMissed != 0)
                    if (provider.Question(string.Format("Произвести скачивание отсутствующих таблиц? ({0} шт.)", this.MissedFiles.Count)))
                        tables.AddRange(this.MissedFiles.Select(file => file.Split('.')[0]));

                // Если есть устаревшие файлы. спросить, требуется ли их скачивать
                if (this.CountOutdated != 0)
                    if (provider.Question(string.Format("Произвести скачивание устаревших таблиц? ({0} шт.)", this.OutdatedFiles.Count)))
                        tables.AddRange(this.OutdatedFiles.Select(kvp => kvp.Key.Split('.')[0]));

                // Убираем возможные дубликаты
                tables = tables.Distinct().ToList<string>();

                if (tables.Count == 0)
                    return;

                //TODO Добавить проверку на то, что Source и Backup действительно Source и Backup
                DownloadFilesAsync(tables);
            }

            private void DownloadFilesAsync(List<string> tables, int threads = 8) {
                // Получаем список файлов, которые необходимо скачать по названиям таблиц
                List<string> pathsOfTableFiles = new List<string>();

                foreach (string table in tables)
                    pathsOfTableFiles.AddRange(this.GetPathsForTable(table));

                // Производим разделение файлов на потребное количество потоков

                // Количество тредов будет ограничено восемью
                if (threads > 8)
                    threads = 8;
                if (threads < 1) 
                    threads = 1;

                // Формируем список файлов на скачивание
                List<List<string>> pathsForTreads = new List<List<string>>();
                for (int i = 0; i < threads; i++) {
                    pathsForTreads.Add(new List<string>());
                }

                int index = 0;
                foreach (string path in pathsOfTableFiles) {
                    pathsForTreads[index++].Add(path);
                    if (index == threads)
                        index = 0;
                }

                // Запускаем копирование файлов
                List<Task> tasks = new List<Task>();

                foreach (List<string> paths in pathsForTreads) {
                    tasks.Add(Task.Run(() => this.DownloadFiles(paths, this.Backup.Dir)));
                }

                Task.WaitAll(tasks.ToArray<Task>());
            }

            
            private async Task DownloadFiles(List<string> paths, string destinationDirectory) {

                if (!Directory.Exists(destinationDirectory))
                    return;

                foreach (string path in paths)
                    // Для каждого файла производим копирование (флаг true отвечает за перезапись файла, если он
                    // уже существовал)
                    File.Copy(path, Path.Combine(destinationDirectory, Path.GetFileName(path)), true);

                return;
            }
            

            public List<string> GetPathsForTable(string nameOfTable) {
                if (nameOfTable.ToLower().EndsWith(".dbf"))
                    nameOfTable = nameOfTable.Split('.')[0];

                return this.Source.AllFiles
                    .Where(kvp => kvp.Key.Split('.')[0] == nameOfTable.ToLower())
                    .Select(kvp => kvp.Value.FullName)
                    .ToList<string>();
            }

            public override string ToString() {
                string templateString = 
                    "В результате сравнения исходной базы данных и локального бэкапа было обнаружено\n" +
                    "Отсутствующие файлы:\n{0}\nРазница в датах:\n{1}";

                // Получаем строку с сообщением об отсутствующих файлах
                StringBuilder sb = new StringBuilder();
                foreach (string file in this.MissedFiles) {
                    sb.AppendFormat("- {0}\n", file);
                }
                if (sb.Length == 0)
                    sb.AppendLine("В локальной папке нет отсутствующих файлов");

                string missedFilesStr = sb.ToString();
                sb.Clear();

                // Получаем строку с сообщением об устаревших файлах
                foreach (KeyValuePair<string, TimeSpan> kvp in this.OutdatedFiles) {
                    sb.AppendLine(string.Format("Файл {0} устарел на {1:N1} дн.", kvp.Key, kvp.Value.TotalDays));
                }
                if (sb.Length == 0)
                    sb.AppendLine("В локальной папке все файлы актуальные");
                string outdatedFilesString = sb.ToString();

                return string.Format(templateString, missedFilesStr, outdatedFilesString);
            }
        }
    }

    private class BackuperOptions {
        public string SourceDir { get; private set; }
        public string BackupDir { get; private set; }
        public bool ShowInfo { get; private set; }

        public BackuperOptions(string source, string backup, bool showInfo = false) {
            this.SourceDir = source;
            this.BackupDir = backup;
            this.ShowInfo = showInfo;
        }
    }

    private class TablesHandler {
    // Класс, инкапсулирующий всю логику по чтению таблиц FoxPro и поиску в них запрашиваемой информации

        private Macro MacroProvider { get; set; }
        private List<Table> Tables { get; set; }
        private List<string> AllColumns { get; set; }

        public TablesHandler(Dictionary<string, string> tables, Macro provider) {

            this.MacroProvider = provider;
            this.Tables = new List<Table>(tables.Count);
            this.AllColumns = new List<string>();

            foreach(KeyValuePair<string, string> kvp in tables) {
                // Инициируем новую таблицу
                Table newTable = new Table(kvp.Key, kvp.Value);
                this.Tables.Add(newTable);

                // Производим инициализацию списка всех доступных колонок
                foreach (string columnName in newTable.GetColumnNames()) {
                    if (!this.AllColumns.Contains(columnName))
                        this.AllColumns.Add(columnName);
                }
            }

        }

        public void AskAndSearch() {
            // Название полей в диалоге
            string columns = "Колонка";
            string request = "Поисковый запрос";
            string directory = "Директория с результатами поиска";
            string file = "Название файла";
            string caseSensitive = "Регистрозависимый";
            string strictMatch = "Точное совпадение";

            // Конфигурация диалога ввода
            InputDialog dialog = new InputDialog(this.MacroProvider.Context, "Укажите параметры поиска");
            dialog.AddMultiselectFromList(columns, this.AllColumns.OrderBy(col => col).ToList<string>(), true);
            dialog.AddString(request, string.Empty, false, true);
            dialog.AddString(directory, Environment.GetFolderPath(Environment.SpecialFolder.Desktop));
            dialog.AddString(file, "ResultOfSearch.txt");
            dialog.AddFlag(caseSensitive, false);
            dialog.AddFlag(strictMatch, true);

            while (true) {
                dialog.Show();
                string pathToFile = Path.Combine(dialog["Директория с результатами поиска"], dialog["Название файла"]);
                //SearchOptions options = new SearchOptions(searchedColumn, searchedValue);
                //PerformSearch(options);
                //PrintSearchResult(pathToFile);
                if (!this.MacroProvider.Question("Повторить поиск?"))
                    break;
            }
        }

        private void PerformSearch(SearchOptions options) {
            foreach (Table table in this.Tables) {
                if (table.ContainsColumn(options.SearchedColumn, options.CaseSensitive))
                    table.Search(options);
            }
        }

        private void PrintSearchResult(string pathToFile) {
            StringBuilder sb = new StringBuilder();
            foreach (Table table in this.Tables) {
                if (table.HaveResults) {
                    sb.AppendLine(table.PrintResultOfSearch());
                    sb.AppendLine();
                }
                else {
                    continue;
                }
            }

            //TODO: Добавить проверку пути на корректность (есть ли данная директория, или же отсутствует);

            // Производим запись данных в файл
            File.WriteAllText(pathToFile, sb.ToString());
        }

        public override string ToString() {
            string template = "Таблицы, переданные в обработкик:\n{0}";

            return string.Format(template, string.Join("\n", this.Tables.Select(table => table.ToString())));
        }


        #region Inner TableHandler classes

        private class Table {
            public string Name { get; private set; }
            public string Path { get; private set; }
            private DbfTable DbfTable { get; set; }
            private Dictionary<string, DbfColumn> ColumnsDict { get; set; }
            private SearchResult Result { get; set; }
            public bool HaveResults => Result != null;

            public Table(string name, string path) {
                this.Name = name;
                this.Path = path;

                this.DbfTable = new DbfTable(this.Path, Encoding.GetEncoding(1251));
                this.ColumnsDict = new Dictionary<string, DbfColumn>();
                foreach (DbfColumn col in this.DbfTable.Columns)
                    this.ColumnsDict[col.ColumnName] = col;
            }

            public bool ContainsColumn(string searchedColumn, bool caseSensitive = false) {
                // Метод для проверки существования искомой колонки в текущей таблице
                if (caseSensitive == true)
                    return this.ColumnsDict.ContainsKey(searchedColumn);
                else
                    return this.ColumnsDict.Keys.Select(key => key.ToLower()).Contains(searchedColumn.ToLower());
            }

            public List<string> GetColumnNames() {
                return this.ColumnsDict.Select(kvp => kvp.Value.ColumnName).ToList<string>();
            }

            public DbfColumn GetDbfColumn(string searchedColumn, bool caseSensitive = false) {
                // Метод для получения информации о колонке в текущей таблице
                if (caseSensitive == true) {
                    if (this.ColumnsDict.ContainsKey(searchedColumn))
                        return this.ColumnsDict[searchedColumn];
                    else
                        return null;
                }
                if (caseSensitive == false) {
                    List<DbfColumn> resultOfSearch = this.ColumnsDict
                        .Where(kvp => kvp.Key.ToLower() == searchedColumn.ToLower())
                        .Select(kvp => kvp.Value)
                        .ToList<DbfColumn>();
                    if (resultOfSearch.Count == 0)
                        return null;
                    if (resultOfSearch.Count == 1)
                        return resultOfSearch[0];
                    if (resultOfSearch.Count > 1)
                        throw new Exception("При выполнении метода GetDbfColumn класса Table возникла ошибка. При заданных параметрах поиска обнаружено больше одного совпадения");
                }
                return null;
            }

            public void Search(SearchOptions options) {
                // Получаем данные колонки для того, чтобы в дальнейшем получить название колонки а так же тип возвращаемого значения
                DbfColumn column = this.GetDbfColumn(options.SearchedColumn, options.CaseSensitive);
                // Если данная колонка не была обнаружена, завершаем поиск
                if (column == null)
                    return;

                DbfDataReaderOptions dbfOptions = new DbfDataReaderOptions() {
                    SkipDeletedRecords = true,
                    Encoding = Encoding.GetEncoding(1251)
                };
                this.Result = new SearchResult(this.Name, options, this.ColumnsDict.Select(kvp => kvp.Key).ToList<string>());

                using (DbfDataReader.DbfDataReader reader = new DbfDataReader.DbfDataReader(this.DbfTable.Path, dbfOptions)) {
                    while (reader.Read()) {
                        string stringValue = DataObjectToString(reader[column.ColumnName], column.DataType);
                        if (stringValue == options.SearchedValue) {
                            foreach (KeyValuePair<string, DbfColumn> kvp in this.ColumnsDict) {
                                this.Result.Add(kvp.Value.ColumnName, DataObjectToString(reader[kvp.Value.ColumnName], kvp.Value.DataType));
                            }
                        }
                        else
                            continue;

                    }
                }

                if (this.Result.IsEmpty)
                    this.Result = null;
            }

            private static string DataObjectToString(object value, Type typeOfValue) {
                // Проверка на то, что данное поле содержит какую-то информацию
                if (value == null)
                    return string.Empty; // Возвращаем пустую строку в том случае, если запрос из базы данных вернул null

                string result;

                switch (typeOfValue.ToString()) {
                    case "System.Int32":
                        result = ((int)value).ToString();
                        break;
                    case "System.Int64":
                        result = ((long)value).ToString();
                        break;
                    case "System.String":
                        result = (string)value;
                        break;
                    case "System.DateTime":
                        result = ((DateTime)value).ToString("dd.MM.yyyy");
                        break;
                    case "System.Boolean":
                        result = ((bool)value).ToString();
                        break;
                    case "System.Single":
                        result = ((float)value).ToString();
                        break;
                    case "System.Double":
                        result = ((double)value).ToString();
                        break;
                    case "System.Decimal":
                        result = ((decimal)value).ToString();
                        break;
                    default:
                        string template = "При определении типа значения, полученного из таблицы, возникла ошибка.\n Тип {0} не определен";
                        throw new Exception(string.Format(template, typeOfValue.ToString()));
                }
                return result;
            }


            public string PrintResultOfSearch() {
                if (this.Result == null)
                    return string.Empty;
                return this.Result.ToString();
            }


            public override string ToString() {
                return string.Format("{0}  ->  '{1}';", this.Name, this.Path);
            }
        }

        private class SearchResult {

            public string TableName { get; private set; }
            public string SearchedValue { get; private set; }
            public string SearchedColumn { get; private set; }
            private List<string> ColumnNames { get; set; }
            private List<string> ColumnTemplate { get; set; }
            private Dictionary<string, List<string>> Values { get; set; }
            private int RowCount { get; set; }
            private int WidthOfTable { get; set; }
            public bool IsEmpty => this.Values.Select(kvp => kvp.Value.Count).Max() == 0;

            public SearchResult(string tableName, SearchOptions options, List<string> columns) {
                this.TableName = tableName;
                this.SearchedColumn = options.SearchedColumn;
                this.SearchedValue = options.SearchedValue;
                this.ColumnNames = columns;
                this.Values = new Dictionary<string, List<string>>();
                foreach (string columnName in this.ColumnNames)
                    this.Values[columnName] = new List<string>();
            }

            public void Add(string columnName, string cellValue) {

                if (!this.Values.ContainsKey(columnName)) {
                    throw new Exception(string.Format("Таблица '{0}' не содержит колонки '{1}'", this.TableName, columnName));
                }

                this.Values[columnName].Add(cellValue);
            }

            public override string ToString() {
                StringBuilder sb = new StringBuilder();

                // Для начала заполняем шапку таблицы
                sb.AppendFormat("Результаты поиска значения '{0}' колонки '{1}' в таблице '{2}'\n", this.SearchedValue, this.SearchedColumn, this.TableName);

                // Получаем termplate для корректного форматирования каждой из колонок
                this.ColumnTemplate = new List<string>();
                foreach (string column in this.ColumnNames) {
                    this.ColumnTemplate.Add(
                            "{0," +
                            (this.Values[column].Select(val => val.Length).Append(column.Length).Max() + 1).ToString()
                            + "}");
                }

                // Определяем количество строк в результирующей таблице
                this.RowCount = this.Values.Select(kvp => kvp.Value.Count).Max();

                sb.AppendLine();

                int tempSbLength = sb.Length;
                // Для начала выводим название колонок
                for (int colInd = 0; colInd < this.ColumnNames.Count; colInd++) {
                    sb.Append("|");
                    sb.AppendFormat(this.ColumnTemplate[colInd], this.ColumnNames[colInd]);
                }
                sb.Append("|\n");

                // Определяем шинину таблицы
                this.WidthOfTable = sb.Length - tempSbLength;
                // Добавляем разделитель соответствующего разделя
                string splitter = new string('-', this.WidthOfTable - 1);
                sb.AppendLine(splitter);

                // Затем выводим все собранные значения
                for (int rowInd = 0; rowInd < this.RowCount; rowInd++) {
                    for (int colInd = 0; colInd < this.ColumnNames.Count; colInd++) {
                        sb.Append("|");
                        sb.AppendFormat(this.ColumnTemplate[colInd], this.Values[this.ColumnNames[colInd]][rowInd]);
                    }
                    sb.Append("|\n");
                }
                sb.AppendLine(splitter);
                sb.AppendLine();

                return sb.ToString();
            }
            
        }

        private class SearchOptions {
            public string SearchedColumn { get; private set; }
            public string SearchedValue { get; private set; }
            public bool ExactMatch { get; private set; }
            public bool CaseSensitive { get; private set; }

            public SearchOptions(string column, string value, bool exactMatch, bool caseSensitive) {
                this.SearchedColumn = column;
                this.SearchedValue = value;
                this.ExactMatch = exactMatch;
                this.CaseSensitive = caseSensitive;
            }

            public SearchOptions(string column, string value, bool exactMatch) : this(column, value, exactMatch, false) {
            }

            public SearchOptions(string column, string value) : this(column, value, true) {
            }

            public override string ToString() {
                string template = "Поисковый запрос: найти значение '{0}' в колонке '{1}' (Точное совпадение: '{2}'; Регистрозависимость: '{3}')";
                return string.Format(template, this.SearchedValue, this.SearchedColumn, this.ExactMatch.ToString(), this.CaseSensitive.ToString());
            }
        }

        #endregion Inner TableHandler classes

    }

    #endregion Classes

    #region Enums

    public enum TypeRepository {
        None,
        Source,
        Backup
    }

    public enum StatusRepository {
        Empty,
        WithoutDbf,
        WithoutMeta,
        Complete
    }

    #endregion Enums
}


