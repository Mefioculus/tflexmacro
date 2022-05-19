using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using TFlex.DOCs.Model.Desktop;
using TFlex.DOCs.Model.FilePreview.CADExchange;
using TFlex.DOCs.Model.FilePreview.CADService;
using TFlex.DOCs.Model.Macros;
using TFlex.DOCs.Model.Processes.Events.Contexts;
using TFlex.DOCs.Model.Processes.Events.Contexts.Data;
using TFlex.DOCs.Model.References;
using TFlex.DOCs.Model.References.ActiveAction;
using TFlex.DOCs.Model.References.ActiveActions;
using TFlex.DOCs.Model.References.Files;
using TFlex.DOCs.Model.References.Processes;
using TFlex.DOCs.Model.Structure;
using TFlex.Model.Technology.References.SetOfDocuments;

namespace TechnologicalPDM_BPCreatePDF
{
    public class Macro : MacroProvider
    {
        #region Guids

        private static class Guids
        {
            public static class Classes
            {
                public static readonly Guid ТехнологическийПроцесс = new Guid("3e93d599-c214-48c8-854f-efe4b475c4d8");
                public static readonly Guid ТехнологическийКомплект = new Guid("dc1cf2a0-6c01-400d-9a42-9642b7496404");
                public static readonly Guid ИзвещениеОбИзменении = new Guid("52ccb35c-67c5-4b82-af4f-e8ceac4e8d02");
            }

            public static class Links
            {
                public static readonly Guid ТПДокументация = new Guid("cc38caed-f747-45ce-9fbf-771566841796");
                public static readonly Guid ТехнологическийКомплектПодлинник = new Guid("148a64ed-3906-4da9-95fc-14bb018669f2");
                public static readonly Guid ИзвещенияИзменения = new Guid("5e46670a-400c-4e36-bb37-d4d651bdf692");
                public static readonly Guid ИзмененияЦелевойВариантТЭ = new Guid("737f68ad-9038-4585-b944-428662256f18");
            }
        }

        #endregion

        private static string[] Extensions = new string[] { ".grb", ".grn", ".grs" };
        private static Guid CadDrawingClassId = new Guid("e440761f-2d0f-465e-a2b7-641427901c9a"); //чертёж системы 2D

        public Macro(MacroContext context)
            : base(context)
        {
        }

        public void СоздатьПодлинникТД()
        {
            // Текущий БП
            ProcessReferenceObject process = null;
            // Текущее действие
            ActiveActionReferenceObject activeAction = null;

            //контекст событий по БП
            EventContext eventContext = Context as EventContext;
            if (eventContext != null)
            {
                var data = eventContext.Data as StateContextData;
                process = data.Process;
                activeAction = data.ActiveAction;
            }
            else
            {
                process = Context.ReferenceObject as ProcessReferenceObject;

                // Guid состояния "Действие на компьютере пользователя" в процедуре
                Guid userMacrosState = Guid.Empty;

                //PDM. Согласование ТД
                if (process.Procedure.SystemFields.Guid == new Guid("71035a47-15d7-43ec-b435-dc1dcd76a94f"))
                    userMacrosState = new Guid("862f1569-f80c-4306-8012-0f4dde7a3c5c");
                //PDM. Согласование технологического ИИ
                else if (process.Procedure.SystemFields.Guid == new Guid("e25dd453-7def-496a-ad00-237cdd2e0bfb"))
                    userMacrosState = new Guid("fd431c1d-b6b2-4564-a946-3be2ee232842");

                activeAction = process.ActiveActions.AsList.FirstOrDefault(action => action.StepGuid == userMacrosState);
            }

            // Данные текущего действия
            ActiveActionData activeActionData = activeAction.GetData<ActiveActionData>();
            // Объекты, подключенные к БП
            List<ReferenceObject> processObjects = activeActionData.GetReferenceObjects().ToList();

            // Сформированные подлинники
            List<ReferenceObject> realFiles = CreateRealFiles(processObjects);
            //if (realFiles.Any())
            //{
            //    // подключаем подлинники к БП
            //    process.BeginChanges();
            //    var existingObjectGuids = activeActionData.ObjectsList.Where(oi => oi.IsEngaged).Select(obj => obj.ObjectGuid).ToArray();
            //    ReferenceObject[] addingObjects = realFiles.Where(obj => (obj != null) && !existingObjectGuids.Contains(obj.SystemFields.Guid)).ToArray();
            //    if (addingObjects.Any())
            //    {
            //        activeActionData.ObjectsList.AddRange(addingObjects.Select(obj => ObjectInfo.Create(obj, true)));
            //        activeAction.SaveData(activeActionData);
            //    }
            //    process.EndChanges();
            //}
        }

        private List<ReferenceObject> CreateRealFiles(List<ReferenceObject> referenceObjects)
        {
            List<ReferenceObject> realFiles = new List<ReferenceObject>();

            foreach (ReferenceObject referenceObject in referenceObjects)
            {
                FileObject realFile = null;

                if (referenceObject.Class.IsInherit(Guids.Classes.ТехнологическийПроцесс))
                {
                    realFile = CreateRealFile(referenceObject);
                    if (realFile != null)
                        realFiles.Add(realFile);
                }
                else if (referenceObject.Class.IsInherit(Guids.Classes.ИзвещениеОбИзменении))
                {
                    var modifications = referenceObject.GetObjects(Guids.Links.ИзвещенияИзменения).ToList();
                    foreach (var modification in modifications)
                    {
                        var tp = modification.GetObject(Guids.Links.ИзмененияЦелевойВариантТЭ);
                        if (tp == null)
                            continue;

                        if (tp.Class.IsInherit(Guids.Classes.ТехнологическийПроцесс))
                        {
                            realFile = CreateRealFile(tp);
                            if (realFile != null)
                                realFiles.Add(realFile);
                        }
                    }
                }
            }

            return realFiles;
        }

        private FileObject CreateRealFile(ReferenceObject tp)
        {
            var rootDocuments = tp.GetObjects(Guids.Links.ТПДокументация);
            if (!rootDocuments.Any())
                return null;

            // комплект документов
            var setOfDocuments = rootDocuments.FirstOrDefault(document => document.Class.IsInherit(Guids.Classes.ТехнологическийКомплект)) as TechnologicalSet;
            if (setOfDocuments == null)
                return null;

            // формируем оригинал для комплекта документов

            FolderObject folder = setOfDocuments.Folder;
            if (folder == null)
                return null;

            List<string> files = new List<string>();
            setOfDocuments.Children.Reload();
            foreach (TechnologicalDocument document in setOfDocuments.Children.OfType<TechnologicalDocument>())
            {
                FileObject documentFile = document.FileObject as FileObject;
                if (documentFile == null)
                    continue;

                try
                {
                    documentFile.GetHeadRevision();
                    files.Add(documentFile.LocalPath);
                }
                catch
                {
                    //%%TODO обработка ошибок
                }
            }

            if (!files.Any())
                return null;

            string outputPath = CreateUniqueFileName(folder, setOfDocuments.ToString());
            bool isEmbedded = true; //%%TODO !setOfDocuments.IsPageCountManually.Value

            CombineFilesProvider provider = new CombineFilesProvider(files, outputPath) { IsEmbedded = isEmbedded };
            provider.Execute(Context.Connection);

            if (!File.Exists(outputPath))
                return null;

            FileObject originalFile = folder.Reference.AddFile(outputPath, folder);
            if (originalFile == null)
                return null;

            Desktop.CheckIn(originalFile, "Автоматическое создание файла комплекта документов", false);

            // Формируем подлинник pdf
            FileObject realFile = ConvertGRBToPDF(originalFile);

            // подключаем к комплекту документации (стадии по БП у комплекта не меняли)
            setOfDocuments.BeginChanges();
            setOfDocuments.File = originalFile;
            if (realFile != null)
                setOfDocuments.SetLinkedObject(Guids.Links.ТехнологическийКомплектПодлинник, realFile);
            setOfDocuments.EndChanges();

            return realFile;
        }

        private static string CreateUniqueFileName(FolderObject folder, string defaultFileName)
        {
            var reference = folder.Reference;
            var localPath = folder.LocalPath;

            var classObject = reference.Classes.TFlexCADFileBase.ChildClasses.Find(CadDrawingClassId) as FileType ??
                reference.Classes.TFlexCADFileBase.ChildClasses.OfType<FileType>()
                .Where(cl => !cl.Hidden && !cl.IsAbstract && cl.Extension.ToLower() == "grb").FirstOrDefault();

            int iteration = 0;
            string fullName;
            Dictionary<Guid, object> dictionary;
            string nowTime = string.Concat(" [", DateTime.Now.ToString("yyyy.MM.dd HH.mm"), "]");
            Guid parentParameterGuid = reference.ParameterGroup.SystemParameters.Find(SystemParameterType.Parent).Guid;
            do
            {
                fullName = string.Concat(
                    localPath,
                    "\\",
                    defaultFileName,
                    iteration == 0 ? string.Empty : nowTime,
                    iteration < 2 ? string.Empty : " (" + (iteration - 1).ToString() + ")",
                    ".grb");
                dictionary = new Dictionary<Guid, object>()
                {
                    { parentParameterGuid, folder },
                    { FileReferenceObject.FieldKeys.Name, Path.GetFileName(fullName) },
                };
                iteration++;
            }
            while (!reference.CanCreateUniqueObject(classObject, dictionary));

            return fullName;
        }

        private FileObject ConvertGRBToPDF(FileObject file)
        {
            file.GetHeadRevision();

            FileObject pdfFile = null;

            string tmpFileFullPath = Path.GetTempFileName();
            string tempPdfFullPath = Path.ChangeExtension(tmpFileFullPath, "pdf");

            try
            {
                string pdfPath = string.Empty;
                CadDocumentProvider provider = CadDocumentProvider.Connect(Context.Connection, ".grb");
                using (var document = provider.OpenDocument(file.LocalPath, true))
                {
                    ExportContext exportContext = new ExportContext(tempPdfFullPath);
                    pdfPath = document.Export(exportContext);
                    document.Close(false);
                }

                if (!string.IsNullOrEmpty(pdfPath))
                {
                    string pdfFileName = string.Format("{0}.pdf", Path.GetFileNameWithoutExtension(file.Name));
                    pdfFile = file.Parent.CreateFile(pdfPath, string.Empty, pdfFileName, file.Reference.Classes.GetFileTypeByExtension("pdf"));
                    if (pdfFile.CanCheckIn)
                    {
                        Desktop.CheckIn(pdfFile, string.Format("Создание подлинника из '{0}'", file.Path), false, null);
                    }
                }
            }
            finally
            {
                DeleteTempFile(tmpFileFullPath);
                DeleteTempFile(tempPdfFullPath);
            }

            return pdfFile;
        }

        private void DeleteTempFile(string path)
        {
            try
            {
                FileInfo fileInfo = new FileInfo(path);
                if (fileInfo.Exists && !fileInfo.IsReadOnly)
                    fileInfo.Delete();
            }
            catch { }
        }
    }
}
