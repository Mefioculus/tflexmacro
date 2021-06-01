using System;
using System.Collections;
using System.Collections.Generic;
using System.IO.Compression; // Для работы с zip архивами
using System.Web; // Данное пространство имен подключено для того, чтобы произвести декодирование данных с кодировки Url
using Forms = System.Windows.Forms; // Для отображения диалогового окна выбора файла для открытия и для сохранения
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


    #region Entry point

    public override void Run() {
        string pathToExportFile = GetPathToOpenFile();
        if (pathToExportFile == string.Empty)
            return;

        List<string> listOfFilesOnExport = GetAllExportedFiles(pathToExportFile);
        if (listOfFilesOnExport.Count == 0) {
            Message("Информация", "В выбранном экспортируемом файле не были обнаружены объекты файлового справочника для экспорта");
            return;
        }

        ExportFiles(listOfFilesOnExport);
        Message("Информация", "Экспорт завершен");
    }

    #endregion Entry point

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
                    if (nameOfFile.Contains("Личные папки")) { // Remove at job is done
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
            string fPath = filePath.Replace("/", "\\");
            ReferenceObject file = fileReference.FindByRelativePath(fPath) as ReferenceObject;

            if (file != null) {
                countFindingFiles++;
                objectToExport.Add(file);
            }
            else {
                countNotFindingFiles++;
                pathNotFound += string.Format("{0}\n", fPath);
            }
            WaitingDialog.NextStep(filePath);
        }
        WaitingDialog.Hide();
        string template = "Всего файлов на экспорт - {0}\nИз них найдено - {1} шт., не найдено - {2} шт.";
        Message("Информация", string.Format(template, allFiles, countFindingFiles, countNotFindingFiles));
        if (pathNotFound != string.Empty)
            Message("Список ненайденных позиций", pathNotFound);

        template = "Начать экспорт {0} найденный позиций из {1} позиций всего?";
        if (!Question(string.Format(template, countFindingFiles, allFiles)))
            return;

        // Написать диалог сохранения файла при помощи WinForms
        string pathToExportFile = GetPathToSaveData();
        if (pathToExportFile == string.Empty) {
            return;
        }

        // Экспорт файлов
        ExportOptions options = new ExportOptions();

        // Определяем параметры для экспорта
        options.FileName = pathToExportFile;
        // Не экспортировать диалоги
        options.DialogsMode = ExportDialogsMode.None;
        // Не экспортировать связанные справочники
        options.LinkedObjectsMode = ExportLinkedObjectsMode.None;
        // Экспортировать только выбранные объекты справочника
        options.ObjectsMode = ExportObjectsMode.Specified;
        // Не экспортировать доступы
        options.IncludeAccesses = false;
        // Не экспортировать связанные макросы
        options.IncludeMacros = false;
        // Не экспортировать прототипы
        options.IncludePrototypes = false;
        // Не экспортировать структуру и типы
        options.IncludeStructure = false;
        // Не экспортировать виды
        options.IncludeViews = false;
        // Упаковать все в один файл
        options.UsePackage = true;

        DataExchangeGateway.Export(Context.Connection, objectToExport, options);
    }

    #endregion Export data

    #region Methods for set path to opening and saving files

    // Метод для выбора файла для его последующего анализа
    private string GetPathToOpenFile() {
        Forms.OpenFileDialog dialog = new Forms.OpenFileDialog();
        dialog.Filter = "export files (*.ddx)|*.ddx|All files (*.*)|*.*";
        dialog.FilterIndex = 1;
        dialog.RestoreDirectory = true;

        if (dialog.ShowDialog() == Forms.DialogResult.OK)
            return dialog.FileName;
        
        return string.Empty;
    }

    // Функция для получения пути сохранения файла
    private string GetPathToSaveData() {
        Forms.SaveFileDialog dialog = new Forms.SaveFileDialog();
        dialog.Filter = "export files (*.ddx)|*.ddx|All files (*.*)|*.*";
        dialog.FilterIndex = 1;
        dialog.RestoreDirectory = true;
        dialog.AddExtension = true;

        if (dialog.ShowDialog() == Forms.DialogResult.OK)
            return dialog.FileName;
        
        return string.Empty;
    }


    #endregion Methods for set path to opening and saving files
}
