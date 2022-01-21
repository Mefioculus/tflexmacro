using System;
using System.IO;
using System.Text;
using System.Linq;
using System.Collections;
using System.Collections.Generic;
using System.Windows.Forms;

using TFlex.DOCs.Model.Macros;
using TFlex.DOCs.Model.Macros.ObjectModel;
using TFlex.DOCs.Model.Structure;
using TFlex.DOCs.Model.Search;
using TFlex.DOCs.Model.References;

/*
Так же для работы данного макроса потребуется подключение дополнительных библиотек:

DbfDataReader.dll - для чтения dbf файлов
System.Data.dll - необходим для работы DbfDataReader

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
        DbRepository sourceRepository = new DbRepository(sourcePath);
        Message("Информация", sourceRepository.ToString());

        DbRepository backupRepository = new DbRepository(backupPath);
        Message("Информация", backupRepository.ToString());

        Message("Сравнение исходного репозитория с локальным", sourceRepository.Compare(backupRepository).ToString());
        Message("Сравнение локального репозитория с исходным", backupRepository.Compare(sourceRepository).ToString());
    }

    private class DbRepository {
        public TypeRepository Type { get; private set; } = TypeRepository.None;
        public StatusRepository Status { get; private set; } = StatusRepository.Empty;
        public string Dir { get; private set; }
        public Dictionary<string, FileInfo> DbfFiles { get; private set; }
        public Dictionary<string, FileInfo> MetaFiles { get; private set; }
        public Dictionary<string, FileInfo> OtherFiles { get; private set; }
        public int Count { get; private set; }
        public int CountDbf => this.DbfFiles.Count;
        public int CountMeta => this.MetaFiles.Count;
        public int CountOther => this.OtherFiles.Count;
        public bool CanWrite { get; private set; } = false;

        public DbRepository(string pathToDir) {
            if (Directory.Exists(pathToDir)) {
                // Присваиваем исходные параметры
                this.Dir = pathToDir;

                // Определяем тип репозитория
                this.Type = Path.GetPathRoot(this.Dir).EndsWith(@"\") ? TypeRepository.Backup : TypeRepository.Source;


                // Начинаем обработку файлов
                string[] files = Directory.GetFiles(this.Dir);
                this.Count = files.Length;

                this.DbfFiles = new Dictionary<string, FileInfo>(this.Count);
                this.MetaFiles = new Dictionary<string, FileInfo>(this.Count);
                this.OtherFiles = new Dictionary<string, FileInfo>(this.Count);

                // Производим добавление информации о файлах
                foreach (string file in files) {
                    FileInfo fileInfo = new FileInfo(file.ToLower());
                    if (fileInfo.Extension == ".dbf") {
                        this.DbfFiles.Add(fileInfo.Name, fileInfo);
                        continue;
                    }
                    if (metaFilesExtensions.Contains(fileInfo.Extension)) {
                        this.MetaFiles.Add(fileInfo.Name, fileInfo);
                        continue;
                    }
                    this.OtherFiles.Add(fileInfo.Name, fileInfo);
                }
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

    private class RepositoryMissedFiles {
        public DbRepository Source { get; private set; }
        public DbRepository Backup { get; private set; }
        public List<string> MissedFiles { get; private set; }
        public Dictionary<string, TimeSpan> OutdatedFiles { get; private set; }
        public int CountMissed => MissedFiles.Count;
        public int CountOutdated => OutdatedFiles.Count;

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

}

