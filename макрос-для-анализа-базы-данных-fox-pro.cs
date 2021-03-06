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
using Newtonsoft.Json;

/*
Так же для работы данного макроса потребуется подключение дополнительных библиотек:

DbfDataReader.dll - для чтения dbf файлов
System.Data.dll - необходим для работы DbfDataReader

Newtonsoft.Json.dll - необходим для сериализации и сохранения настроек

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

        Settings savedSettings = new Settings("analize-dbf-tables-settings.json");
        savedSettings.Load();

        Backuper backuper = new Backuper(sourcePath, backupPath, this, savedSettings);
        if (backuper.IsExit == true)
            return;

        // Получаем словарь с доступными dbf таблицами
        TablesHandler handler = new TablesHandler(backuper.GetPathDbfFiles(), this, savedSettings);

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
        private Settings Settings { get; set; }

        public bool IsExit { get; set; } = false;

        public Backuper (string sourcePath, string backupPath, Macro provider, Settings settings) {
            // Инициализируем репозитории
            this.MacroProvider = provider;
            this.Settings = settings;
            bool showInfo = false;

            // Запрашиваем параметры у пользователя
            GetParametersFromUser(ref sourcePath, ref backupPath, ref showInfo);

            // Завершаем обработку, если пользлователь нажал отмену в диалоге, который вызывается
            // в методе GetParametersFromUser
            if (this.IsExit == true)
                return;

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

            string sourceField = "Источник";
            string backupField = "Бэкап";
            string showInfoField = "Показывать информацию";

            // Пробуем получить ранее сохраненные параметры пользователя
            sourcePath = this.Settings.RetrieveOrDefault<string>(sourceField, sourcePath);
            backupPath = this.Settings.RetrieveOrDefault<string>(backupField, backupPath);
            showInfo = this.Settings.RetrieveOrDefault<bool>(showInfoField, showInfo);

            dialog.AddString(sourceField, sourcePath);
            dialog.AddString(backupField, backupPath);
            dialog.AddFlag(showInfoField, showInfo);


            if (dialog.Show()) {

                // Производим проверку пути источника
                if (sourcePath != dialog[sourceField])
                    if (!Directory.Exists(dialog[sourceField]))
                        throw new Exception(string.Format("Директория, указанная для источника не существует:\n{0}", dialog[sourceField]));
                    else
                        sourcePath = dialog[sourceField];

                // Производим проверку пути бэкапа
                if (backupPath != dialog[backupField])
                    backupPath = dialog[backupField];

                if (!Directory.Exists(backupPath))
                    Directory.CreateDirectory(backupPath);

                // Заполняем флаг, показывать ли информацию
                showInfo = dialog[showInfoField];

                // Производим обновление и сохранение выбранных пользователем настроек
                if (this.Settings != null) {
                    this.Settings[sourceField] = dialog[sourceField];
                    this.Settings[backupField] = dialog[backupField];
                    this.Settings[showInfoField] = dialog[showInfoField].ToString();

                    this.Settings.Save();
                }
            }
            else
                this.IsExit = true;
        }

        public List<KeyValuePair<string, string>> GetPathDbfFiles() {
            List<KeyValuePair<string,string>> result = new List<KeyValuePair<string, string>>(this.BackupRepository.CountDbf);

            foreach (KeyValuePair<string, FileInfo> kvp in this.BackupRepository.DbfFiles.OrderByDescending(kvp => kvp.Value.Length))
                result.Add(new KeyValuePair<string, string>(kvp.Key.Split('.')[0], kvp.Value.FullName));

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

                if (tasks.Count != 0)
                    Task.WaitAll(tasks.ToArray<Task>());
            }

            
            private void DownloadFiles(List<string> paths, string destinationDirectory) {

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

    private class TablesHandler {
    // Класс, инкапсулирующий всю логику по чтению таблиц FoxPro и поиску в них запрашиваемой информации

        private Macro MacroProvider { get; set; }
        private List<Table> Tables { get; set; }
        private List<string> AllColumns { get; set; }
        private Settings Settings { get; set; }

        public TablesHandler(List<KeyValuePair<string, string>> tables, Macro provider, Settings settings) {

            this.MacroProvider = provider;
            this.Tables = new List<Table>(tables.Count);
            this.AllColumns = new List<string>();
            this.Settings = settings;

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
            string columnsField = "Колонка";
            string requestField = "Поисковый запрос";
            string directoryField = "Директория с результатами поиска";
            string caseSensitiveField = "Регистрозависимый";
            string strictMatchField = "Точное совпадение";

            string request = this.Settings.RetrieveOrDefault<string>(requestField, string.Empty);
            string directory = this.Settings.RetrieveOrDefault<string>(directoryField, Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "Результаты поиска в FoxPro"));
            bool caseSensitive = this.Settings.RetrieveOrDefault<bool>(caseSensitiveField, false);
            bool strictMatch = this.Settings.RetrieveOrDefault<bool>(strictMatchField, false);

            // Конфигурация диалога ввода
            // В том случае, если существуют ранее сохраненные результаты ввода, выдать в диалоге их
            InputDialog dialog = new InputDialog(this.MacroProvider.Context, "Укажите параметры поиска");

            dialog.AddMultiselectFromList(columnsField, this.AllColumns.OrderBy(col => col).ToList<string>(), true);
            dialog.AddString(requestField, request, false, true);
            dialog.AddString(directoryField, directory);
            dialog.AddFlag(caseSensitiveField, caseSensitive);
            dialog.AddFlag(strictMatchField, strictMatch);

            List<string> listOfFileToOpen = new List<string>();
            while (true) {
                if (dialog.Show()) {
                    // Данный участок требуется для преобразования динамически полученного списка List<object> в List<string>
                    List<string> cols = new List<string>();
                    foreach (object obj in dialog[columnsField])
                        cols.Add((string)obj);

                    SearchOptions options = new SearchOptions(cols, dialog[requestField], dialog[strictMatchField], dialog[caseSensitiveField]);
                    PerformSearchAsync(options, 8);
                    string pathToFile = Path.Combine(dialog[directoryField], string.Format("Search '{0}'.txt", dialog[requestField]));
                    listOfFileToOpen.Add(pathToFile);
                    PrintSearchResult(pathToFile);
                    if (!this.MacroProvider.Question("Поиск успешно завершен.\nПовторить поиск с другими параметрами?")) {
                        // Производим сохранение введенных полей пользователем
                        this.Settings[requestField] = dialog[requestField];
                        this.Settings[directoryField] = dialog[directoryField];
                        this.Settings[caseSensitiveField] = dialog[caseSensitiveField].ToString();
                        this.Settings[strictMatchField] = dialog[strictMatchField].ToString();
                        this.Settings.Save();
                        break;
                    }
                }
                else
                    break;
            }

            // Производим открытие файлов в блокноте
            if ((listOfFileToOpen.Count != 0) && (this.MacroProvider.Question("Открыть сгенерированные файлы?"))) {
                foreach (string file in listOfFileToOpen) {
                    System.Diagnostics.Process notepad = new System.Diagnostics.Process();
                    notepad.StartInfo.FileName = "notepad.exe";
                    notepad.StartInfo.Arguments = file;
                    notepad.Start();
                }
            }
        }

        private void PerformSearchAsync(SearchOptions options, int quantityOfThreads = 8) {
            // Производим разделение таблиц на количенство потоков для 
            // произведения асинхроного выполнения поиска
            List<List<Table>> tablesForThreads = new List<List<Table>>();
            for (int i = 0; i < quantityOfThreads; i++) {
                tablesForThreads.Add(new List<Table>());
            }
            
            int index = 0;
            foreach (Table table in this.Tables.Where(t => t.ContainsColumn(options))) {
                tablesForThreads[index++].Add(table);
                if (index == quantityOfThreads)
                    index = 0;
            }

            // Запускаем выполнение поиска
            List<Task> tasks = new List<Task>(quantityOfThreads);

            foreach (List<Table> tables in tablesForThreads) {
                if (tables.Count != 0)
                    tasks.Add(Task.Run(() => this.PerformSearch(options, tables)));
            }

            if (tasks.Count != 0)
                Task.WaitAll(tasks.ToArray<Task>());
        }

        private void PerformSearch(SearchOptions options, List<Table> tables) {
            foreach (Table table in tables)
                table.Search(options);
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

            string directory = Path.GetDirectoryName(pathToFile);
            if (!Directory.Exists(directory))
                Directory.CreateDirectory(directory);

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

            public bool ContainsColumn(SearchOptions options) {
                foreach (string column in options.SearchedColumns) {
                    if (this.ColumnsDict.ContainsKey(column))
                        return true;
                }
                return false;
            }

            public List<string> GetColumnNames() {
                return this.ColumnsDict.Select(kvp => kvp.Value.ColumnName).ToList<string>();
            }

            public List<DbfColumn> GetDbfColumn(SearchOptions options) {
                // Получаем список колонок данной таблицы, по которым нужно вести поиск
                List<DbfColumn> result = this.ColumnsDict
                    .Where(kvp => options.SearchedColumns.Contains(kvp.Key))
                    .Select(kvp => kvp.Value)
                    .ToList<DbfColumn>();

                if (result.Count == 0)
                    return null;

                return result;
            }

            public void Search(SearchOptions options) {
                // Получаем данные колонки для того, чтобы в дальнейшем получить название колонки а так же тип возвращаемого значения
                List<DbfColumn> columns = this.GetDbfColumn(options);
                // Если данная колонка не была обнаружена, завершаем поиск
                if (columns == null)
                    return;

                DbfDataReaderOptions dbfOptions = new DbfDataReaderOptions() {
                    SkipDeletedRecords = true,
                    Encoding = Encoding.GetEncoding(1251)
                };
                this.Result = new SearchResult(this.Name, options, this.ColumnsDict.Select(kvp => kvp.Key).ToList<string>());

                using (DbfDataReader.DbfDataReader reader = new DbfDataReader.DbfDataReader(this.DbfTable.Path, dbfOptions)) {
                    while (reader.Read()) {
                        foreach (DbfColumn column in columns) {
                            string actualValueInString = DataObjectToString(reader[column.ColumnName], column.DataType);

                            if (IsMatch(actualValueInString, options)) {
                                Dictionary<string, string[]> row = new Dictionary<string, string[]>();
                                foreach (KeyValuePair<string, DbfColumn> kvp in this.ColumnsDict) {
                                    row[kvp.Value.ColumnName] = DataObjectToString(reader[kvp.Value.ColumnName], kvp.Value.DataType).Split(new string[] { Environment.NewLine }, StringSplitOptions.RemoveEmptyEntries);
                                }
                                this.Result.AddRow(row);
                                break;
                            }
                        }
                    }
                }

                if (this.Result.IsEmpty)
                    this.Result = null;
            }

            private static bool IsMatch(string actualValue, SearchOptions options) {
                // Готовим исходные данные с учетом регистрозависимости поиска
                string searchedValue;
                if (options.CaseSensitive == false) {
                    actualValue = actualValue.ToLower();
                    searchedValue = options.SearchedValue.ToLower();
                }
                else {
                    searchedValue = options.SearchedValue;
                }

                // Производим сравнение
                if (options.ExactMatch == true)
                    return (searchedValue == actualValue);
                else
                    return actualValue.Contains(searchedValue);
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
            public string SearchedColumns { get; private set; }
            private List<string> ColumnNames { get; set; }
            private List<string> ColumnTemplate { get; set; }
            private Dictionary<string, List<string>> Values { get; set; }
            private int RowCount { get; set; }
            private int WidthOfTable { get; set; }
            public bool IsEmpty => this.Values.Select(kvp => kvp.Value.Count).Max() == 0;

            public SearchResult(string tableName, SearchOptions options, List<string> columns) {
                this.TableName = tableName;
                this.SearchedColumns = string.Join("; ", options.SearchedColumns);
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

            public void AddRow(Dictionary<string, string[]> row) {
                // Так как в ячейке некоторых таблиц может быть текст, разделенный новыми строками,
                // для корректного отображения таких данных их нужно правильно обработать количество и добавить сообветствующее количество новых строк
                // Получаем максимальное количество строк
                int maxCountLinesInRow = row.Select(kvp => kvp.Value.Length).Max();

                foreach (KeyValuePair<string, string[]> kvp in row) {
                    for (int i = 0; i < maxCountLinesInRow; i++) {
                        if (i < kvp.Value.Length)
                            this.Add(kvp.Key, kvp.Value[i]);
                        else
                            this.Add(kvp.Key, string.Empty);
                    }
                }
            }

            public override string ToString() {
                StringBuilder sb = new StringBuilder();

                // Для начала заполняем шапку таблицы
                sb.AppendFormat("Результаты поиска значения '{0}' колонки '{1}' в таблице '{2}'\n", this.SearchedValue, this.SearchedColumns, this.TableName);

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
            public List<string> SearchedColumns { get; private set; }
            public string SearchedValue { get; private set; }
            public bool ExactMatch { get; private set; }
            public bool CaseSensitive { get; private set; }

            public SearchOptions(List<string> columns, string value, bool exactMatch, bool caseSensitive) {
                // TODO: Изменить название на SearchedColumns
                this.SearchedColumns = columns;
                this.SearchedValue = value;
                this.ExactMatch = exactMatch;
                this.CaseSensitive = caseSensitive;
            }

            public SearchOptions(List<string> columns, string value, bool exactMatch) : this(columns, value, exactMatch, false) {
            }

            public SearchOptions(List<string> columns, string value) : this(columns, value, true) {
            }

            public override string ToString() {
                string template = "Поисковый запрос: найти значение '{0}' в колонке '{1}' (Точное совпадение: '{2}'; Регистрозависимость: '{3}')";
                return string.Format(template, this.SearchedValue, this.SearchedColumns, this.ExactMatch.ToString(), this.CaseSensitive.ToString());
            }
        }

        #endregion Inner TableHandler classes

    }

    private class Settings {
        private Dictionary<string, string> Fields { get; set; }
        public string Dir { get; private set; }
        public string Name { get; private set; }
        public string PathToFile => Path.Combine(this.Dir, this.Name);
        public bool Exists => File.Exists(this.PathToFile);
        public int Count => this.Fields.Count;

        #region Constructors

        public Settings(string directory, string name) {
            this.Dir = directory;
            this.Name = name;
            this.Fields = new Dictionary<string, string>();

            if (!Directory.Exists(this.Dir))
                throw new Exception(string.Format("Директория для сохранения настроек отсутствует:\n{0}", this.Dir));
        }

        public Settings(string name) : this(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Temp"), name) {
        }

        public Settings() : this(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Temp"), "Macro-Settings.json") {
        }

        #endregion Constructors

        public bool Load() {
            if (File.Exists(this.PathToFile)) {
                string jsonString = File.ReadAllText(this.PathToFile);
                this.Fields = JsonConvert.DeserializeObject<Dictionary<string, string>>(jsonString);
                return true;
            }
            return false;
        }

        public void Save() {
            string jsonString = JsonConvert.SerializeObject(this.Fields);
            File.WriteAllText(this.PathToFile, jsonString);
        }

        public T RetrieveOrDefault<T>(string key, T defaultValue) {
            if (this.Fields.ContainsKey(key)) {
                switch (defaultValue.GetType().ToString()) {
                    case "System.String":
                        return (T)Convert.ChangeType(this.Fields[key], typeof(T));
                    case "System.Int32":
                        return (T)Convert.ChangeType(this.Fields[key], typeof(T));
                    case "System.Int64":
                        return (T)Convert.ChangeType(this.Fields[key], typeof(T));
                    case "System.DateTime":
                        return (T)Convert.ChangeType(this.Fields[key], typeof(T));
                    case "System.Boolean":
                        return (T)Convert.ChangeType(this.Fields[key], typeof(T));
                    case "System.Single":
                        return (T)Convert.ChangeType(this.Fields[key], typeof(T));
                    case "System.Double":
                        return (T)Convert.ChangeType(this.Fields[key], typeof(T));
                    case "System.Decimal":
                        return (T)Convert.ChangeType(this.Fields[key], typeof(T));
                    default:
                        string template = "При получении значения из словаря сохраненных ранее настроек возникла ошибка:\nТип {0} не поддерживается";
                        throw new Exception(string.Format(template, typeof(T).ToString()));
                }
            }
            else
                return defaultValue;
        }

        public string this[string key] {
            get {
                try {
                    return this.Fields[key];
                }
                catch (Exception e) {
                    throw new Exception(string.Format("При попытке обращения к ключу {0} пользовательский настроек макроса возникла ошибка.\n\n{1}", key, e.Message));
                }
            }
            set {
                this.Fields[key] = value;
            }
        }

        public bool ContainsKey(string key) => this.Fields.ContainsKey(key);

        public bool ContainsValue(string value) => this.Fields.ContainsValue(value);

        public void Clear() {
            this.Fields.Clear();
            this.Save();
        }
        
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


