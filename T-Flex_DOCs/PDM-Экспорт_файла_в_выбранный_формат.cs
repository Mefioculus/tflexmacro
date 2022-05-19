using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using TFlex.DOCs.Model.Desktop;
using TFlex.DOCs.Model.FilePreview.CADService;
using TFlex.DOCs.Model.FilePreview.CADService.TFlexCadDocument;
using TFlex.DOCs.Model.Macros;
using TFlex.DOCs.Model.Macros.ObjectModel;
using TFlex.DOCs.Model.References;
using TFlex.DOCs.Model.References.Documents;
using TFlex.DOCs.Model.References.Files;
using TFlex.DOCs.Resources.Strings;

/// <summary>
/// Макрос для экспорта .grb файла в выбранный формат
/// </summary>
public class ExportCadToSelectedFormat : MacroProvider
{
    private static readonly string _tempFolder = Path.Combine(Path.GetTempPath(), "Temp DOCs", "ExportGRB");

    private const string _dialogCaption = "Настройка экспорта документов T-FLEX CAD";
    private const string _dialogTypeDocument = "Тип документа:";
    private const string _dialogResolution = "Разрешение: ";
    private const string _dialogPages = "Страницы: ";
    private const string _dialogExportedFileName = "Имя файла: ";
    private const string _dialogExportToNewFile = "Экспортировать в новый файл";

    /// <summary> Прервать экспорт </summary>
    private bool _breakExport = false;

    /// <summary> Экспортировать в новый файл </summary>
    private bool _exportToNewFile = true;

    /// <summary> Разрешение файла </summary>
    private int _resolution = 300;

    /// <summary> Объект, отвечающий за взаимодействие с системой CAD </summary>
    private CadDocumentProvider _provider;

    /// <summary> Тип экспортируемых файлов </summary>
    private string _extensionDoc = "tif";

    /// <summary> Список поддерживаемых форматов для экспорта из .grb </summary>
    private readonly object[] _supportExtensions =
    {
        "tif", "pdf", "png", "jpg", "bmp"
    };

    /// <summary> Стадия "Аннулировано" </summary>
    public static readonly Guid CanceledStageGuid = new Guid("b04183e6-decb-47b3-8b46-b75a6548d573");

    public ExportCadToSelectedFormat(MacroContext context)
        : base(context)
    {
    }

    public override void Run()
    {
        ExportToFormat(Context.GetSelectedObjects().ToList());
    }

    public ButtonValidator ValidateButton()
    {
        var validator = new ButtonValidator();

        if (DynamicMacro.ExecutionPlace != ExecutionPlace.WpfClient)
        {
            validator.Enable = false;
            validator.Visible = false;
        }

        return validator;
    }

    public void ExportToFormat(List<ReferenceObject> referenceObjects)
    {
        var files = referenceObjects == null ? null : referenceObjects.OfType<FileObject>().Where(fileObject => fileObject.Class.Extension == "grb").ToList();
        if (files == null || !files.Any())
        {
            Message("Сообщение", "Отсутствуют файлы");
            return;
        }

        int filesCount = files.Count;
        if (filesCount > 1)
        {
            if (!Question(String.Format("Выбрано файлов: {0}. Начать экспорт?", filesCount)))
                return;
        }

        if (!Directory.Exists(_tempFolder))
            Directory.CreateDirectory(_tempFolder);

        try
        {
            List<FileObject> uploadedFiles = new List<FileObject>();

            int fileNumber = 1;
            foreach (FileObject grbFile in files)
            {
                if (!LoadGrbFileToLocalPath(grbFile))
                    continue;

                var pagesInfo = GetPagesInfo(grbFile.LocalPath);
                var dialog = CreateExportDialog(grbFile.Name, fileNumber++ < filesCount, pagesInfo);
                if (!dialog.Show())
                {
                    if (_breakExport)
                        break;

                    continue;
                }
                _extensionDoc = dialog.GetValue(_dialogTypeDocument);
                _resolution = dialog[_dialogResolution];
                _exportToNewFile = dialog[_dialogExportToNewFile];

                string exportedFileName = dialog[_dialogExportedFileName];

                string tempExportingFilePath = Path.Combine(_tempFolder, String.Format("{0}.{1}", Guid.NewGuid(), _extensionDoc));
                var exportContext = new ExportContext(tempExportingFilePath);
                exportContext["resolution"] = _resolution;

                var selectedPages = (List<object>)dialog[_dialogPages];
                exportContext.Pages.AddRange(selectedPages.Cast<TFlexPageInfo>().Select(sp => sp.Index));

                ExportGrbToSelectedFormat(exportContext, grbFile.LocalPath);
                var uploadedFile = UploadExportFile(tempExportingFilePath, grbFile.Parent.Path, grbFile, exportedFileName);
                uploadedFiles.Add(uploadedFile);
            }

            if (uploadedFiles.Any())
            {
                RefreshReferenceWindow();
                Message("Информация",
                    String.Format("Экспортированы файлы:{1}{0}", String.Join(Environment.NewLine, uploadedFiles), Environment.NewLine));
            }
        }
        finally
        {
            ClearTemp();
        }
    }

    /// <summary>
    /// Загрузить файл по локальному пути в файловой системе
    /// </summary>
    /// <param name="file"></param>
    private bool LoadGrbFileToLocalPath(FileObject file)
    {
        // Получаем последнюю версию файла
        file.GetHeadRevision();

        if (file.Size == 0)
        {
            Message("Ошибка", "Файл '{0}' не содержит данных", file);
            return false;
        }

        if (File.Exists(file.LocalPath))
        {
            if (file.Reference.IsSlave)
            {
                var contextObject = file.MasterObject;
                if (contextObject is EngineeringDocumentObject document)
                    contextObject = document.GetLinkedNomenclatureObject();

                if (contextObject != null)
                    FileObject.SetOpenDocumentContext(file.LocalPath, contextObject.SystemFields.Guid, contextObject.Reference.Id);
            }

            return true;
        }

        Message("Ошибка", "Ошибка загрузки файла '{0}'", file);
        return false;
    }

    /// <summary>
    /// Показать диалог ввода параметров экспорта
    /// </summary>
    /// <returns></returns>
    private InputDialog CreateExportDialog(string file, bool hasOtherFiles, TFlexPageInfo[] pagesInfo)
    {
        string fileName = Path.GetFileNameWithoutExtension(file);

        var inputDialog = CreateInputDialog(_dialogCaption);

        inputDialog.AddText(String.Format("Исходный файл: '{0}'", file), 1);

        inputDialog.AddSelectFromList(
            _dialogTypeDocument,
            String.IsNullOrEmpty(_extensionDoc) ? "tif" : _extensionDoc,
            true,
            _supportExtensions);

        inputDialog.AddInteger(_dialogResolution, _resolution, true);
        inputDialog.AddMultiselectFromList(_dialogPages, pagesInfo, true);
        inputDialog.AddString(_dialogExportedFileName, fileName, false, true);
        inputDialog.AddFlag(_dialogExportToNewFile, _exportToNewFile, false, true);

        inputDialog.Closing += (bool okButtonClicked, ref bool closeDialog) =>
        {
            if (!okButtonClicked)
            {
                if (hasOtherFiles)
                {
                    _breakExport = Question("Прервать экспорт остальных файлов?");
                    return;
                }
            }

            if (!IsValidFileName(inputDialog[_dialogExportedFileName]))
            {
                Сообщение("Внимание!", "Имя файла содержит некорректные символы");
                closeDialog = false;
            }
        };

        return inputDialog;
    }

    /// <summary>
    /// Получить информацию о страницах в .grb файле
    /// </summary>
    /// <returns></returns>
    private TFlexPageInfo[] GetPagesInfo(string pathToGrbFile)
    {
        TFlexPageInfo[] pagesInfo;

        // Подключаемся к CAD
        _provider = CadDocumentProvider.Connect(Context.Connection, ".grb");

        // Открытие документа grb. Менять документ не будем, поэтому второй аргумент функции (readOnly) равен true.
        using (var document = _provider.OpenDocument(pathToGrbFile, true))
        {
            // Проверяем, был ли открыт документ
            if (document == null)
            {
                var fileName = Path.GetFileName(pathToGrbFile);

                throw new MacroException(String.Format(
                    "Ошибка получения страниц.{0}" +
                    "При операции получения страниц произошла следующая ошибка:{0}" +
                    "Файл '{1}' не может быть открыт", Environment.NewLine, fileName));
            }

            pagesInfo = document.GetTFlexPagesInfo();
        }

        return pagesInfo;
    }

    /// <summary>
    /// Экспорт grb файла в выбранный формат
    /// </summary>
    private void ExportGrbToSelectedFormat(ExportContext exportContext, string pathToGrbFile)
    {
        // Открытие документа grb
        // Менять документ не будем, поэтому второй аргумент функции (readOnly) равен true
        using (var document = _provider.OpenDocument(pathToGrbFile, true))
        {
            var tempGrbFileName = Path.GetFileName(pathToGrbFile);

            // Проверяем был ли открыт документ
            if (document == null)
                throw new MacroException(String.Format("Файл '{0}' не может быть открыт", tempGrbFileName));

            // Экспортируем документ в другой формат на основе контекста настройки
            // Получаем полный путь до экспортированного файла, для дальнейшей проверки экспорта
            var path = document.Export(exportContext);

            // Закрываем grb документ без сохранения
            document.Close(false);

            if (path == null)
            {
                throw new MacroException(String.Format(
                    "Ошибка экспорта.{0}" +
                    "При операции экспорта в '{1}' произошли следующие ошибки:{0}" +
                    "Файл '{2}' не может быть экспортирован",
                    Environment.NewLine, _extensionDoc, tempGrbFileName));
            }
        }
    }

    /// <summary>
    /// Загрузить экспортированный файл на сервер
    /// </summary>
    /// <param name="tempFilePath">Относительный путь к временному файлу</param>
    /// <param name="parentFolderPath">Относительный путь к родительской папке</param>
    /// <param name="grbFileObject">Объект grb файла</param>
    /// <param name="fileName">Имя, с которым файл будет сохранен в справочник</param>
    /// <returns>Созданный файл</returns>
    private FileObject UploadExportFile(string tempFilePath, string parentFolderPath, FileObject grbFileObject, string fileName)
    {
        try
        {
            var fileReference = new FileReference(Context.Connection)
            {
                LoadSettings = { LoadDeleted = true }
            };

            var parentFolder = (TFlex.DOCs.Model.References.Files.FolderObject)fileReference.FindByRelativePath(parentFolderPath);
            if (parentFolder == null)
                throw new MacroException(String.Format("Не найдена родительская папка с именем '{0}'", parentFolderPath));
            parentFolder.Children.Load();

            var uploadingFileName = String.Format("{0}.{1}", fileName, _extensionDoc);

            var exportedFile = parentFolder.Children.AsList
                .FirstOrDefault(child => child.IsFile && child.Name.Value == uploadingFileName) as FileObject;

            if (exportedFile is null)
            {
                FileType fileType = GetFileType(fileReference);
                exportedFile = parentFolder.CreateFile(
                    tempFilePath,
                    String.Empty,
                    uploadingFileName,
                    fileType);
            }
            else
            {
                if (_exportToNewFile)
                {
                    FileType fileType = GetFileType(fileReference);
                    exportedFile = parentFolder.CreateFile(
                        tempFilePath,
                        String.Empty,
                        GetUniqueExportedFileName(uploadingFileName, parentFolder, fileName),
                        fileType);
                }
                else
                {
                    if (!exportedFile.IsCheckedOutByCurrentUser)
                        Desktop.CheckOut(exportedFile, false);

                    File.Copy(tempFilePath, exportedFile.LocalPath, true);
                }
            }

            var masterObject = grbFileObject.MasterObject;
            if (masterObject is EngineeringDocumentObject documentObject)
            {
                if (!exportedFile.Links.ToMany[EngineeringDocumentFields.File]
                    .Objects
                    .Any(document => document.SystemFields.Guid == documentObject.SystemFields.Guid))
                {
                    exportedFile.BeginChanges();
                    exportedFile.AddLinkedObject(EngineeringDocumentFields.File, documentObject);
                    exportedFile.EndChanges();
                }
            }

            Desktop.CheckIn(exportedFile, String.Format(
                "Экспорт файла:{0}'{1}'{0}в формат '{3}':{0}'{2}'",
                Environment.NewLine, grbFileObject.Path, exportedFile.Path, _extensionDoc), false);

            if (masterObject != null && masterObject.Reference.ParameterGroup.Guid == Guids.Изменения)
            {
                if (!masterObject.Links.ToMany[Guids.ИзмененияРабочиеФайлы]
                    .Objects
                    .Any(workingFile => workingFile.SystemFields.Guid == exportedFile.SystemFields.Guid))
                {
                    masterObject.BeginChanges();
                    masterObject.AddLinkedObject(Guids.ИзмененияРабочиеФайлы, exportedFile);
                    masterObject.EndChanges();
                }
            }

            return exportedFile;
        }
        catch (Exception e)
        {
            string exceptionMessage = String.Format(
                "Ошибка загрузки файла на сервер.{0}" +
                "При операции загрузки файла на сервер произошли следующие ошибки:{0}{1}",
                Environment.NewLine, e.Message);
            throw new MacroException(exceptionMessage, e);
        }
    }

    private FileType GetFileType(FileReference fileReference)
    {
        var fileType = fileReference.Classes.GetFileTypeByExtension(_extensionDoc);
        if (fileType is null)
        {
            string typeName = String.Format(Texts.FileNameWithExtension, _extensionDoc.ToUpper());
            fileType = fileReference.Classes.CreateFileType(typeName, String.Empty, _extensionDoc);
        }

        return fileType;
    }

    /// <summary>
    /// Получить уникальное имя экспортируемого файла
    /// </summary>
    /// <param name="exportedFileName">Имя экспортируемого файла</param>
    /// <param name="parentFolder">Родительская папка</param>
    /// <returns>Уникальное имя экспортируемого файла</returns>
    private string GetUniqueExportedFileName(string exportedFileName, TFlex.DOCs.Model.References.Files.FolderObject parentFolder, string fileName)
    {
        var filesName = parentFolder.Children.AsList
            .Where(child => child.IsFile)
            .Select(file => file.Name.Value)
            .ToArray();

        var filesNameSet = new HashSet<string>(filesName);

        var counter = 1;

        while (filesNameSet.Contains(exportedFileName))
        {
            exportedFileName = String.Format("{0}_{1}.{2}", fileName, counter, _extensionDoc);
            counter++;
        }

        return exportedFileName;
    }

    /// <summary>
    /// Удалить временные файлы
    /// </summary>
    private void ClearTemp()
    {
        if (Directory.Exists(_tempFolder))
        {
            foreach (string filePath in Directory.GetFiles(_tempFolder))
                DeleteFile(filePath);
        }
    }

    /// <summary>
    /// Удалить файл по указанному пути
    /// </summary>
    /// <param name="path">Путь к файлу</param>
    private void DeleteFile(string path)
    {
        if (!File.Exists(path))
            return;

        // Получаем атрибуты файла
        var fileAttribute = File.GetAttributes(path);

        if ((fileAttribute & FileAttributes.ReadOnly) == FileAttributes.ReadOnly)
        {
            // Удаляем атрибут 'Только для чтения'
            var removeAttributes = RemoveAttribute(fileAttribute, FileAttributes.ReadOnly);
            File.SetAttributes(path, removeAttributes);
        }

        File.Delete(path);
    }

    // Удалить атрибут
    private FileAttributes RemoveAttribute(FileAttributes attributes, FileAttributes attributesToRemove)
    {
        return attributes & ~attributesToRemove;
    }

    private static bool IsValidFileName(string fileName)
    {
        return !string.IsNullOrWhiteSpace(fileName) &&
               Path.GetInvalidFileNameChars().All(invalidFileNameChar => !fileName.Contains(invalidFileNameChar));
    }

    private static class Guids
    {
        /// <summary> Guid справочника 'Изменения' </summary>
        public static readonly Guid Изменения = new Guid("c9a4bb1b-cacb-4f2d-b61a-265f1bfc7fb9");

        /// <summary> Guid связи 'Рабочие Файлы' справочника 'Файлы' </summary>
        public static readonly Guid ИзмененияРабочиеФайлы = new Guid("6b65a575-3ca4-4fb0-9bfc-4d1655c2d83e");
    }
}

