using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using TFlex.DOCs.Common;
using TFlex.DOCs.Model.Macros;
using TFlex.DOCs.Model.Macros.ObjectModel;
using TFlex.DOCs.Model.References;
using TFlex.DOCs.Model.References.Documents;
using TFlex.DOCs.Model.References.Files;
using TFlex.DOCs.Model.Search;
using TFlex.DOCs.Model.Structure;

namespace PDM_TiffAssemblyMagick
{
    public class Macro : MacroProvider
    {
        private static string _tempFolder = Path.Combine(Path.GetTempPath(), "Temp DOCs", "TiffAssemblyImages");
        private static string _logFile = Path.Combine(_tempFolder, "TFlex.DOCs.TiffAssembly.log");
        private static string _magickEXEFileName = "magick.exe";
        private string _localMagickFilePath;
        private ProcessStartInfo _startInfo;

        public Macro(MacroContext context)
            : base(context)
        {
        }

        public override void Run()
        {
        }

        private string LocalMagickFilePath
        {
            get
            {
                if (String.IsNullOrEmpty(_localMagickFilePath))
                {
                    FileObject magickFile = null;
                    FileReference fileReference = new FileReference(Context.Connection);
                    using (Filter filter = new Filter(fileReference.ParameterGroup.ReferenceInfo))
                    {
                        filter.Terms.AddTerm(fileReference.ParameterGroup[FileReferenceObject.FieldKeys.Name], ComparisonOperator.Equal, _magickEXEFileName);
                        magickFile = fileReference.Find(filter).FirstOrDefault() as FileObject;
                    }
                    if (magickFile != null)
                    {
                        magickFile.GetHeadRevision();
                        _localMagickFilePath = magickFile.LocalPath;
                    }
                }

                return _localMagickFilePath;
            }
        }

        private ProcessStartInfo StartInfo
        {
            get
            {
                if (_startInfo == null)
                {
                    _startInfo = new ProcessStartInfo
                    {
                        FileName = LocalMagickFilePath,
                        WindowStyle = ProcessWindowStyle.Hidden,
                        CreateNoWindow = true,
                        UseShellExecute = false,
                        RedirectStandardOutput = true,
                        RedirectStandardError = true
                    };
                }

                return _startInfo;
            }
        }

        public void СоздатьСводныйФайл()
        {
            List<FileObject> tiffFiles = new List<FileObject>();
            foreach (var file in Context.GetSelectedObjects().OfType<FileObject>())
            {
                if (ТипФайлаПоддерживается(file.Name))
                    tiffFiles.Add(file);
            }

            if (!tiffFiles.Any())
                Error("Выбранные файлы не являются файлами формата TIFF.");

            Directory.CreateDirectory(_tempFolder);

            //Для хранения изображений TIFF
            Dictionary<int, string> tempFiles = new Dictionary<int, string>();

            FolderObject parentFolder = null;
            ParameterGroup linkGroup = null;
            ReferenceObject fileOwner = null;

            string общееНаименование = string.Empty;
            foreach (var file in tiffFiles)
            {
                // д.б. формат: наименованиеФайла_номер
                string наименованиеФайла = Path.GetFileNameWithoutExtension(file.Name);

                int индексРазделителя = наименованиеФайла.LastIndexOf('_');
                if (индексРазделителя > 0)
                {
                    if (индексРазделителя + 1 < наименованиеФайла.Length)
                    {
                        string строкаПослеРазделителя = наименованиеФайла.Substring(индексРазделителя + 1);
                        if (string.IsNullOrEmpty(строкаПослеРазделителя))
                            continue;

                        int номерФайла;
                        if (int.TryParse(строкаПослеРазделителя, out номерФайла))
                        {
                            if (номерФайла >= 0)
                            {
                                if (tempFiles.ContainsKey(номерФайла))
                                    continue;

                                string строкаДоРазделителя = наименованиеФайла.Substring(0, индексРазделителя);

                                if (string.IsNullOrEmpty(общееНаименование))
                                    общееНаименование = строкаДоРазделителя;
                                else
                                {
                                    if (строкаДоРазделителя != общееНаименование)
                                        continue;
                                }

                                string путьКФайлу = Path.Combine(_tempFolder, String.Format("{0}{1}.tiff", общееНаименование, номерФайла));
                                SaveFile(путьКФайлу, file);

                                if (linkGroup is null)
                                    linkGroup = file.Reference.LinkInfo?.LinkGroup;

                                if (fileOwner is null)
                                    fileOwner = file.MasterObject;

                                if (parentFolder == null)
                                {
                                    Объект папка = НайтиОбъект(SystemParameterGroups.Files.ToString(), String.Format("[Дочерние объекты].[ID] = '{0}'", file.SystemFields.Id));
                                    if (папка != null)
                                        parentFolder = (FolderObject)папка;
                                }

                                tempFiles.Add(номерФайла, путьКФайлу);
                            }
                        }
                    }
                }
            }

            if (!tempFiles.Any())
                Error("Выбранные файлы не содержат изображения или наименования не соответствуют формату 'наименованиеФайла_номер'.");

            List<string> sortedFiles = tempFiles.OrderBy(file => file.Key).Select(selector => selector.Value).ToList();
            FileObject originalFile = null;
            try
            {
                // путь к сводному файлу подлинника
                string originalFilePath = Path.Combine(_tempFolder, String.Format("{0}.tiff", общееНаименование));
                // Формирование сводного файла
                AssemblyMultiPageTiff(originalFilePath, sortedFiles);

                if (!File.Exists(originalFilePath))
                    Error("Ошибка создания сводного файла '{0}'.", originalFilePath);

                sortedFiles.Add(originalFilePath);

                var fileReference = new FileReference(Context.Connection);
                originalFile = fileReference.AddFile(originalFilePath, parentFolder);
                if (originalFile is null)
                    Error("Ошибка добавления файла '{0}' в папку '{1}'.", originalFilePath, parentFolder);

                Message("Создание сводного файла *.tiff", "Сформирован сводный файл '{0}'", originalFile.Path);

                LinkFile(fileOwner, linkGroup, originalFile);

                RefreshReferenceWindow();
            }
            catch (Exception e)
            {
                throw e;
            }
            finally
            {
                УдалитьВременныеФайлы(sortedFiles);
            }
        }

        private bool ТипФайлаПоддерживается(string filePath)
        {
            if (String.IsNullOrEmpty(filePath))
                return false;

            string extension = Path.GetExtension(filePath).TrimStart('.').ToLower();
            if (extension == "tif" || extension == "tiff")
                return true;

            return false;
        }

        private void SaveFile(string newFilePath, FileObject file)
        {
            if (!file.IsActualVersionDownloaded())
                file.GetHeadRevision(newFilePath);
            else
                File.Copy(file.LocalPath, newFilePath, true);

            File.SetAttributes(newFilePath, FileAttributes.Normal);
        }

        private void LinkFile(ReferenceObject fileOwner, ParameterGroup linkGroup, FileObject file)
        {
            if (fileOwner is null || linkGroup is null || !linkGroup.IsLinkToMany)
                return;

            if (fileOwner is EngineeringDocumentObject)
            {
                if (!file.Links.ToMany[EngineeringDocumentFields.File]
                    .Objects
                    .Any(document => document.SystemFields.Guid == fileOwner.SystemFields.Guid))
                {
                    file.BeginChanges();
                    file.AddLinkedObject(EngineeringDocumentFields.File, fileOwner);
                    file.EndChanges();
                }
            }
            else
            {
                fileOwner.BeginChanges();
                fileOwner.AddLinkedObject(linkGroup, file);
                fileOwner.EndChanges();
            }
        }

        private void AssemblyMultiPageTiff(string originalFilePath, List<string> filesWithImages)
        {
            StringBuilder errors = new StringBuilder();

            //""A_1.tiff" "A_2.tiff" "A_3.tiff" A.tiff"
            string cmd = String.Format("\"{0}\" \"{1}\"", String.Join("\" \"", filesWithImages), originalFilePath);
            WriteLog(cmd, true);

            StartProcess(cmd, ref errors);

            WriteLog(errors.ToString());
        }

        private string StartProcess(string cmd, ref StringBuilder errors)
        {
            StartInfo.Arguments = cmd;
            Process process = Process.Start(StartInfo);

            string output = process.StandardOutput.ReadToEnd();
            string error = process.StandardError.ReadToEnd();

            process.WaitForExit();

            if (!String.IsNullOrEmpty(error))
                errors.AppendLine(error);

            return output;
        }

        private static void WriteLog(string logText, bool clear = false)
        {
            if (String.IsNullOrEmpty(logText))
                return;

            using (var logWriter = new TFlex.DOCs.Common.LogWriter())
            {
                logWriter.InitializeFileMode(_logFile, clear);
                logWriter.Write(logText);
            }
        }

        private void УдалитьВременныеФайлы(IEnumerable<string> tempFiles)
        {
            if (!Directory.Exists(_tempFolder))
                return;

            List<string> notDeletedFiles = new List<string>();
            foreach (string tempFilePath in tempFiles)
            {
                try
                {
                    DeleteFile(tempFilePath);
                }
                catch
                {
                    notDeletedFiles.Add(tempFilePath);
                }
            }

            if (notDeletedFiles.Any())
                WriteLog(String.Format("Не удалось удалить временные файлы:\r\n{0}", String.Join("\r\n", notDeletedFiles)));
        }

        private void DeleteFile(string filePath)
        {
            FileInfo fileInfo = new FileInfo(filePath);
            if (fileInfo.Exists)
            {
                fileInfo.IsReadOnly = false;
                fileInfo.Delete();
            }
        }
    }
}
