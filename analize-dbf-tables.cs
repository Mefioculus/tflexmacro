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
TODO: В классе TablesHandler реализовать возможность поиска по таблицам
TODO: Переделать большие ToString через StringBuilder для ускорения работы
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

        //Message("Доступные dbf таблицы", handler.ToString());
        //Message("Доступные колонки", string.Join("\n", handler.GetColumns()));

        Message("Информация", "Работа макроса завершена");
        
    }

    #region Classes

    private class Backuper {
    // Класс, искапсулирующий всю логику по проведению бэкапа базы данных FoxPro

        private DbRepository SourceRepository { get; set; }
        private DbRepository BackupRepository { get; set; }
        private RepositoryMissedFiles MissedFiles { get; set; }
        private Macro MacroProvider { get; set; }

        public Backuper (string sourceDir, string backupDir, Macro provider) {
            // Инициализируем репозитории
            try {
                this.SourceRepository = new DbRepository(sourceDir);
                this.BackupRepository = new DbRepository(backupDir);
            }
            catch (Exception e) {
                throw new Exception(e.Message);
            }
            this.MacroProvider = provider;

            // Выводим пользователю сообщение об инициализированных репозиториях
            this.MacroProvider.Message("Информация об инициализированных репозиториях", string.Format("{0}\n\n{1}", this.SourceRepository.ToString(), this.BackupRepository.ToString()));

            // Запускаем сравнение репозиториев
            this.MissedFiles = this.SourceRepository.Compare(this.BackupRepository);

            // Запрашиваем у пользователя подтверждение о начале копирования
            this.MissedFiles.AskAndDownload(this.MacroProvider);
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

                //TODO: Если директория назначения не существует, не производить копирование
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
                    sb.AppendLine(string.Format("- {0}", file));
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
        private List<String> Columns { get; set; }

        public TablesHandler(Dictionary<string, string> tables, Macro provider) {

            this.MacroProvider = provider;
            this.Tables = new List<Table>(tables.Count);
            this.Columns = new List<string>();

            foreach(KeyValuePair<string, string> kvp in tables) {
                Table newTable = new Table(kvp.Key, kvp.Value);
                this.Tables.Add(newTable);
                this.Columns.AddRange(newTable.Columns);
            }

            this.Columns = this.Columns
                .Distinct()
                .OrderBy()
                .ToList<string>();
        }

        public List<string> GetColumns() {
            return this.Columns;
        }


        public override string ToString() {
            string template = "Таблицы, переданные в обработкик:\n{0}";

            return string.Format(template, string.Join("\n", this.Tables.Select(table => table.ToString())));
        }

        private class Table {
            public string Name { get; private set; }
            public string Path { get; private set; }
            private DbfTable DbfTable { get; set; }
            public List<string> Columns { get; private set; }

            public Table(string name, string path) {
                this.Name = name;
                this.Path = path;

                this.DbfTable = new DbfTable(this.Path, Encoding.GetEncoding(1251));
                this.Columns = new List<string>();

                GetColumns();
            }

            private void GetColumns() {
                foreach (DbfColumn column in this.DbfTable.Columns) {
                    this.Columns.Add(column.ColumnName);
                }
            }

            public string ColumnsToString() {
                return string.Join("; ", this.Columns);
            }

            public override string ToString() {
                return string.Format("{0}  ->  '{1}';", this.Name, this.Path);
            }
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


