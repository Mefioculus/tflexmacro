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
        DbRepository sourceRepository = new DbRepository(TypeRepository.Source, sourcePath);
        Message("Информация", sourceRepository.ToString());
    }

    private class DbRepository {
        public TypeRepository Type { get; private set; } = TypeRepository.None;
        public StatusRepository Status { get; private set; } = StatusRepository.Empty;
        public string Dir { get; private set; }
        public List<FileInfo> DbfFiles { get; private set; }
        public List<FileInfo> MetaFiles { get; private set; }
        public List<FileInfo> OtherFiles { get; private set; }
        public int Count { get; private set; }
        public int CountDbf => this.DbfFiles.Count;
        public int CountMeta => this.MetaFiles.Count;
        public int CountOther => this.OtherFiles.Count;

        public DbRepository(TypeRepository type, string pathToDir) {
            if (Directory.Exists(pathToDir)) {
                this.Type = type;
                this.Dir = pathToDir;
                string[] files = Directory.GetFiles(this.Dir);
                this.Count = files.Length;

                this.DbfFiles = new List<FileInfo>(this.Count);
                this.MetaFiles = new List<FileInfo>(this.Count);
                this.OtherFiles = new List<FileInfo>(this.Count);

                // Производим добавление информации о файлах
                foreach (string file in files) {
                    FileInfo fileInfo = new FileInfo(file.ToLower());
                    if (fileInfo.Extension == ".dbf") {
                        this.DbfFiles.Add(fileInfo);
                        continue;
                    }
                    if (metaFilesExtensions.Contains(fileInfo.Extension)) {
                        this.MetaFiles.Add(fileInfo);
                        continue;
                    }
                    this.OtherFiles.Add(fileInfo);
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


