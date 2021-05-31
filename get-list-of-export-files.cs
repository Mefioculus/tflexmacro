using System;
using System.Collections;
using System.Collections.Generic;
using System.IO.Compression; // Для работы с zip архивами
using System.Web; // Данное пространство имен подключено для того, чтобы произвести декодирование данных с кодировки Url
// Пространства имен DOCs
using TFlex.DOCs.Model.Macros;
using TFlex.DOCs.Model.Macros.ObjectModel; // пространство имен для отображения диалога ожидания 
using TFlex.DOCs.Model.References;
using TFlex.DOCs.Model.References.Files;
using TFlex.DOCs.Model.DataExchange; // пространстпо имен для экспорта данных

// Для работы макроса так же необходимо подключить следующие библиотеки
// System.Web.dll
// System.IO.Compression.dll
// System.IO.Compression.ZipFile.dll
// System.IO.Compression.FileSystem.dll

public class Macro : MacroProvider {
    public Macro(MacroContext context)
        : base(context)
    {
    }

    #region Guids

    private static class Guids {
        public static class References {
            public static Guid Files = new Guid("a0fcd27d-e0f2-4c5a-bba3-8a452508e6b3");
        }
    }

    #endregion Guids

    public override void Run() {
        List<string> listOfFilesOnExport = GetAllExportedFiles(GetExportFile());
        ExportFiles(listOfFilesOnExport);
    }

    #region Get list of files from archive with exported data

    // Метод для получения списка файлов в виде дерева
    private List<string> GetAllExportedFiles(string pathToExportFile) {
        // Создаем список, в котором будем хранить относительные пути ко всем файлам, которые находятся в файле экспорта
        List<string> listOfFiles = new List<string>();

        // Извлекаем список всех файлов, которые находятся в файле экспорта
        using (ZipArchive archive = ZipFile.Open(pathToExportFile, ZipArchiveMode.Read)) {
            foreach (ZipArchiveEntry entry in archive.Entries) {
                string nameOfFile = HttpUtility.UrlDecode(entry.FullName);
                if (nameOfFile.StartsWith("Files")) {
                    if (nameOfFile.Contains("Личные папки")) {
                        listOfFiles.Add(nameOfFile.Remove(0, 6));
                    }
                }
            }
        }

        return listOfFiles;
    }

    #endregion Get list of files from archive with exported data

    #region Export data

    // Функция для экспорта файлов
    private void ExportFiles(List<string>pathToFiles) {
        // Сигнатура метода для проведения экспорта
        //DataExchangeGaneway.Export(ServerConnection, IEnumerable<ReferenceObject>, ExportOptions);
        
        // Создаем новый экземпляр справочника
        FileReference fileReference = new FileReference(Context.Connection);
        List<ReferenceObject> objectToExport = new List<ReferenceObject>();
        string pathNotFound = string.Empty;

        // Производим поиск по справочнику
        int countFindingFiles = 0;
        int countNotFindingFiles = 0;
        int allFiles = pathToFiles.Count;
        
        WaitingDialog.Show("Обработка данных", true);

        foreach (string filePath in pathToFiles) {
            filePath = filePath.Replace("/", "\\");
            ReferenceObject file = fileReference.FindByRelativePath(filePath) as ReferenceObject;

            if (file != null) {
                countFindingFiles++;
                objectToExport.Add(file);
            }
            else {
                countNotFindingFiles++;
                pathNotFound += string.Format("{0}\n", filePath);
            }
            WaitingDialog.NextStep("filePath");
        }
        WaitingDialog.Hide();
        string template = "Всего файлов на экспорт - {0}\nИз них найдено - {1} шт., не найдено - {2} шт.";
        Message("Информация", string.Format(template, allFiles, countFindingFiles, countNotFindingFiles));
        if (pathNotFound != string.Empty)
            Message("Список ненайденных позиций", pathNotFound);

        template = "Начать экспорт {0} найденный позиций из {1} всего?";
        if (!Question("Подтвердите действие", string.Format(template, countFindingFiles, allFiles)))
            return;

        // Написать диалог сохранения файла при помощи WinForms
        string pathToExportFile = GetPathToSaveData();
        if (pathToExportFile == string.Empty) {
            return;
        }

        // Экспорт файлов
        ExportOption option = new ExportOption();

        // Определяем параметры для экспорта
        option.FileName = pathToExportFile;
        // Не экспортировать диалоги
        option.DialogsMode = ExportDialogMode.None;
        // Не экспортировать связанные справочники
        option.LinkedObjectsMode = ExportLinkedObjectMode.None;
        // Экспортировать только выбранные объекты справочника
        option.ObjectMode = ExportObjectMode.Specified;
        // Не экспортировать доступы
        option.IncludeAccesses = false;
        // Не экспортировать связанные макросы
        option.IncludeMacros = false;
        // Не экспортировать прототипы
        option.IncludePrototypes = false;
        // Не экспортировать структуру и типы
        option.IncludeStructure = false;
        // Не экспортировать виды
        option.IncludeViews = false;
        // Упаковать все в один файл
        option.UsePackage = true;

        DataExchangeGateway.Export(Context.Connection, objectToExport, option);
    }

    #endregion Export data

    #region Methods for set path to opening and saving files

    // Метод для выбора файла для его последующего анализа
    private string GetExportFile() {
        string pathToFile = @"C:\Users\gukovry\Desktop\DOCs (Экспорт)\Тестовый экспорт изделий\УЯИС.525455.007 (с связанными объектами связанных справочников).ddx";
        return pathToFile;
    }

    // Функция для получения пути сохранения файла
    private string GetPathToSaveData() {
        string result = string.Empty;
        //TODO Реализовать метод получения пути для сохранения файла экспорта
        return result;
    }

    // Функция для получения пути файла, который требуется открыть
    // TODO

    #endregion Methods for set path to opening and saving files
}
