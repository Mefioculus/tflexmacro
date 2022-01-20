using System;
using System.IO;
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
        Message("Информация", Path.GetPathRoot(sourcePath));
        Message("Информация", Path.GetPathRoot(backupPath));


        DbRepository sourceRepository = new DbRepository(TypeRepository.Source, sourcePath);
        Message("Информация", sourceRepository.ToString());

        DbRepository backupRepository = new DbRepository(TypeRepository.Backup, backupPath);
        Message("Информация", backupRepository.ToString());
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

        public DbRepository(TypeRepository type, string pathToDir) {
            if (Directory.Exists(pathToDir)) {
                // Присваиваем исходные параметры
                this.Type = type;
                this.Dir = pathToDir;

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

        public string Compare(DbRepository otherRepository) {
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

            return string.Empty;
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


