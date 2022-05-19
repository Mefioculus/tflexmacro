using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using TFlex.DOCs.Common;
using TFlex.DOCs.Model.Desktop;
using TFlex.DOCs.Model.Macros;
using TFlex.DOCs.Model.References;
using TFlex.DOCs.Model.References.Documents;
using TFlex.DOCs.Model.References.Files;
using TFlex.DOCs.Model.References.Macros;
using TFlex.DOCs.Model.References.Modifications;
using TFlex.DOCs.Model.References.Nomenclature;

namespace PDM_CheckModificationNotice
{
    public class Macro : MacroProvider
    {
        private static class Guids
        {
            public static readonly Guid ModificationNoticeType = new Guid("52ccb35c-67c5-4b82-af4f-e8ceac4e8d02");
            public static readonly Guid ModificationNoticesSetType = new Guid("322d371b-c672-40ea-9166-316140fc28cb");
            public static readonly Guid ModificationNoticesLink = new Guid("740644db-40a2-44d8-8fe9-841391985d46");

            public static readonly Guid СвязьЦелевойДокумент = new Guid("a0e64cef-bf5b-47b9-ae5d-12155c0db936");
            public static readonly Guid СвязьИсходныйДокумент = new Guid("48b83092-a645-4dbd-83c0-a3ab0a02ee62");
            public static readonly Guid СвязьДокумент = new Guid("d962545d-bccb-40b5-9986-257b57032f6e");
            public static readonly Guid СвязьРабочиеФайлыИзменения = new Guid("6b65a575-3ca4-4fb0-9bfc-4d1655c2d83e");
            public static readonly Guid СвязьЭскиз = new Guid("180e553f-7248-42d8-be38-472386f54ff9");

            public static readonly Guid UsingAreaAddedLink = new Guid("bac1b7b8-bd2e-4198-beb0-3580ee077534");
            public static readonly Guid LinkInitialRevision = new Guid("49028ea0-3bc9-40d5-ae3c-a1e7b12c34e2");

            /// <summary> Стадия "Аннулировано" </summary>
            public static readonly Guid CanceledStageGuid = new Guid("b04183e6-decb-47b3-8b46-b75a6548d573");
        }

        private static readonly string _папкаАрхив = "Архив";
        private static readonly string _папкаПодлинники = Path.Combine(_папкаАрхив, "Подлинники");
        private static readonly string _папкаОригиналы = Path.Combine(_папкаАрхив, "Оригиналы");
        private static readonly string _папкаОригиналыTF = Path.Combine(_папкаОригиналы, "TF");
        private static readonly string _папкаОригиналыMS = Path.Combine(_папкаОригиналы, "MS");
        private static readonly string _папкаОригиналыДругое = Path.Combine(_папкаОригиналы, "Другое");

        public Macro(MacroContext context)
            : base(context)
        {
        }

        public override void Run()
        {
        }

        public List<ReferenceObject> CheckModificationNotice(ReferenceObject modificationNoticeObject, bool useRevisions = false)
        {
            var macro = Context[nameof(MacroKeys.MacroSource)] as CodeMacro;
            if (macro.DebugMode.Value == true)
            {
                System.Diagnostics.Debugger.Launch();
                System.Diagnostics.Debugger.Break();
            }

            var checkingObjects = new List<ReferenceObject>();
            StringBuilder errorsText = new StringBuilder();
            StringBuilder warningsText = new StringBuilder();

            if (modificationNoticeObject is null)
            {
                errorsText.AppendLine("Извещение не найдено.");
                ShowMessage(errorsText.ToString());
                return null;
            }

            var modificationNotices = new List<ReferenceObject>();

            if (modificationNoticeObject.Class.IsInherit(Guids.ModificationNoticesSetType))
            {
                var linkedModificationNotices = modificationNoticeObject.GetObjects(Guids.ModificationNoticesLink);
                if (linkedModificationNotices.IsNullOrEmpty())
                {
                    errorsText.AppendLine($"Комплект извещений '{modificationNoticeObject}' не содержит связанные извещения.");
                    ShowMessage(errorsText.ToString());
                    return null;
                }

                checkingObjects.Add(modificationNoticeObject);
                modificationNotices.AddRange(linkedModificationNotices);
            }
            else
                modificationNotices.Add(modificationNoticeObject);

            checkingObjects.AddRange(modificationNotices);

            DesignContextObject commonDesignContext = null;

            var allModificationReferenceObjects = new List<ModificationReferenceObject>();

            foreach (var modificationNotice in modificationNotices)
            {
                var modifications = modificationNotice
                    .GetObjects(ModificationReferenceObject.RelationKeys.ModificationNotice)
                    .Where(referenceObject => referenceObject.Class.IsInherit(ModificationTypes.Keys.Modification))
                    .OfType<ModificationReferenceObject>().ToList();

                if (modifications.IsNullOrEmpty())
                {
                    warningsText.AppendLine($"Извещение '{modificationNotice}' не содержит изменений.");
                    continue;
                }

                allModificationReferenceObjects.AddRange(modifications);
            }

            if (!allModificationReferenceObjects.IsNullOrEmpty())
            {
                if (useRevisions)
                    commonDesignContext = allModificationReferenceObjects.First().DesignContext;

                checkingObjects.AddRange(allModificationReferenceObjects);

                foreach (var modification in allModificationReferenceObjects)
                {
                    if (useRevisions)
                    {
                        var designContext = modification.DesignContext;
                        if (designContext != commonDesignContext)
                            Error($"У изменений в извещениях указаны разные контексты");

                        if (modification.TryGetObject(ModificationReferenceObject.RelationKeys.PDMObject, out var pdmObject))
                        {
                            if (pdmObject is null)
                            {
                                warningsText.AppendLine($"У изменения '{modification}' отсутствует объект ЭСИ.");
                                continue;
                            }

                            checkingObjects.Add(pdmObject);
                        }
                        else
                            Error($"У изменения '{modification}' отсутствует связь '{ModificationReferenceObject.RelationKeys.PDMObject}'.");

                        var usingAreaObjects = modification.GetObjects(ModificationReferenceObject.RelationKeys.UsingArea);
                        foreach (var usingArea in usingAreaObjects)
                        {
                            var linkMatches = usingArea.GetObjects(ModificationUsingAreaReferenceObject.RelationKeys.HierarchyLinkMatches);
                            foreach (var matchingObject in linkMatches)
                            {
                                var addedLink = matchingObject.GetToOneComplexHierarchyLink(Guids.UsingAreaAddedLink, designContext);
                                if (addedLink is null)
                                {
                                    // %%TODO
                                }
                                else
                                    checkingObjects.Add(addedLink.ParentObject);
                            }
                        }

                        var sourceRevisions = modification.GetObjects(ModificationReferenceObject.RelationKeys.SourceRevisions);
                        foreach (var sourceRevision in sourceRevisions)
                        {
                            var linkedObject = sourceRevision.GetObject(Guids.LinkInitialRevision);
                            if (linkedObject is null)
                                continue;

                            checkingObjects.Add(linkedObject);
                        }


                        var pdmLinkedObject = (pdmObject as NomenclatureObject)?.LinkedObject;
                        if (pdmLinkedObject is null)
                        {
                            warningsText.AppendLine($"У объекта ЭСИ '{pdmObject}' отсутствует связанный объект.");
                        }
                        else
                        {
                            if (pdmLinkedObject is EngineeringDocumentObject targetDocument)
                            {
                                var targetDocumentFiles = targetDocument.GetFiles();
                                if (targetDocumentFiles.IsNullOrEmpty())
                                {
                                    warningsText.AppendLine($"У документа '{targetDocument}' отсутствуют связанные файлы.");
                                }
                                else
                                {
                                    var fileReference = new FileReference(Context.Connection);

                                    fileReference.LoadSettings.LoadDeleted = true;

                                    foreach (var targetFile in targetDocumentFiles)
                                    {
                                        string newFolderPath = String.Empty;

                                        switch (targetFile.Class.Extension.ToLower())
                                        {
                                            case ("tif"):
                                            case ("tiff"):
                                            case ("pdf"):
                                                newFolderPath = _папкаПодлинники;
                                                break;

                                            case ("grb"):
                                                newFolderPath = _папкаОригиналыTF;
                                                break;

                                            case ("xls"):
                                            case ("xlsx"):
                                            case ("doc"):
                                            case ("docx"):
                                                newFolderPath = _папкаОригиналыMS;
                                                break;

                                            default:
                                                newFolderPath = _папкаОригиналыДругое;
                                                break;
                                        }

                                        string newRelativeFilePath = Path.Combine(newFolderPath, targetFile.Name);
                                        var file = fileReference.FindByRelativePath(newRelativeFilePath);
                                        if (file != null)
                                            warningsText.AppendLine(String.Format("Файл '{0}' существует.", newRelativeFilePath));
                                    }

                                    fileReference.LoadSettings.LoadDeleted = false;
                                    checkingObjects.AddRange(targetDocumentFiles);
                                }
                            }
                        }
                    }
                    else
                    {
                        if (modification.TryGetObject(Guids.СвязьЦелевойДокумент, out var targetDocument))
                        {
                            if (targetDocument is null || !(targetDocument is NomenclatureObject nomenclatureObject))
                                errorsText.AppendLine(String.Format("У изменения '{0}' отсутствует целевой вариант.", modification));
                            else
                                checkingObjects.Add(targetDocument);
                        }
                        else
                            Error(String.Format("У изменения '{0}' отсутствует связь '{1}'.", modification, Guids.СвязьЦелевойДокумент.ToString()));

                        if (modification.TryGetObject(Guids.СвязьИсходныйДокумент, out var sourceDocument))
                        {
                            if (sourceDocument != null)
                                checkingObjects.Add(sourceDocument);
                        }
                        else
                            Error(String.Format("У изменения '{0}' отсутствует связь '{1}'.", modification, Guids.СвязьИсходныйДокумент.ToString()));

                        if (modification.TryGetObject(Guids.СвязьЭскиз, out var sketch))
                        {
                            if (sketch != null && sketch is FileObject)
                                checkingObjects.Add(sketch);
                        }
                        else
                            Error(String.Format("У изменения '{0}' отсутствует связь '{1}'.", modification, Guids.СвязьЭскиз.ToString()));

                        IEnumerable<FileObject> modificationFiles = null;
                        if (modification.TryGetObjects(Guids.СвязьРабочиеФайлыИзменения, out var files))
                        {
                            modificationFiles = files.OfType<FileObject>();
                        }
                        else
                            Error(String.Format("У изменения '{0}' отсутствует связь '{1}'.", modification, Guids.СвязьРабочиеФайлыИзменения.ToString()));

                        if (modificationFiles.IsNullOrEmpty())
                            errorsText.AppendLine(String.Format("У изменения '{0}' отсутствуют файлы", modification));
                        else
                            checkingObjects.AddRange(modificationFiles);

                        if (modification.TryGetObject(Guids.СвязьДокумент, out var actualNomenclature))
                        {
                            if (actualNomenclature is null)
                            {
                                errorsText.AppendLine(String.Format("У изменения '{0}' отсутствует актуальный вариант.", modification));
                                ShowMessage(errorsText.ToString());
                                return null;
                            }

                            checkingObjects.Add(actualNomenclature);
                        }
                        else
                            Error(String.Format("У изменения '{0}' отсутствует связь '{1}'.", modification, Guids.СвязьДокумент.ToString()));

                        var actualDocument = (actualNomenclature as NomenclatureObject).LinkedObject as EngineeringDocumentObject;
                        if (actualDocument == null)
                        {
                            errorsText.AppendLine(String.Format("У номенклатурного объекта '{0}' отсутствует связанный объект.", actualNomenclature));
                            ShowMessage(errorsText.ToString());
                            return null;
                        }

                        string actualDocumentName = String.Concat(actualDocument.Denotation, "^", actualDocument.Name);
                        var actualDocumentFiles = actualDocument.GetFiles();
                        if (actualDocumentFiles.IsNullOrEmpty())
                        {
                            errorsText.AppendLine(String.Format("У документа '{0}' отсутствуют связанные файлы", actualDocumentName));
                            ShowMessage(errorsText.ToString());
                            return null;
                        }

                        checkingObjects.AddRange(actualDocumentFiles);

                        var archiveFiles = actualDocumentFiles.Where(file => file.Path.GetString().ToLower().Contains(@"архив\")).ToList();
                        if (archiveFiles.IsNullOrEmpty())
                        {
                            errorsText.AppendLine(String.Format("У документа '{0}' отсутствуют файлы в архиве", actualDocumentName));
                            ShowMessage(errorsText.ToString());
                            return null;
                        }

                        var scriptExtensions = new string[] { "tiff", "tif", "pdf" };
                        var scripts = archiveFiles.Where(docFile => docFile.Path.GetString().ToLower().Contains(@"подлинники")
                                   && scriptExtensions.Contains(docFile.Class.Extension)).ToList();

                        if (scripts.IsNullOrEmpty())
                            errorsText.AppendLine(String.Format("У документа '{0}' отсутствуют подлинники", actualDocumentName));
                        else
                        {
                            foreach (var script in scripts)
                            {
                                var correspondingFileInModification = modificationFiles.FirstOrDefault(mf => mf.Name.GetString() == script.Name.GetString());
                                if (correspondingFileInModification == null)
                                    errorsText.AppendLine(String.Format("У документа '{0}' для файла-подлинника '{1}' не найден соответствующий файл в изменении '{2}'",
                                        actualDocumentName, script, modification));
                            }
                        }

                        var originals = archiveFiles.Where(docFile => docFile.Path.GetString().ToLower().Contains(@"оригиналы") && !scriptExtensions.Contains(docFile.Class.Extension)).ToList();
                        if (originals.IsNullOrEmpty())
                            warningsText.AppendLine(String.Format("У документа '{0}' отсутствуют оригиналы", actualDocumentName));
                        else
                        {
                            foreach (var original in originals)
                            {
                                var correspondingFileInModification = modificationFiles.FirstOrDefault(mf => mf.Name.GetString() == original.Name.GetString());
                                if (correspondingFileInModification == null)
                                    warningsText.AppendLine(String.Format("У документа '{0}' для файла-оригинала '{1}' не найден соответствующий файл в изменении '{2}'",
                                        actualDocumentName, original, modification));
                            }
                        }
                    }
                }
            }

            if (errorsText.Length > 0)
            {
                ShowMessage(errorsText.ToString());
                return null;
            }

            if (warningsText.Length > 0)
            {
                warningsText.AppendLine();
                warningsText.AppendLine("Продолжить?");
                if (!Question(warningsText.ToString()))
                    return null;
            }

            return checkingObjects;
        }

        private void ShowMessage(string text)
        {
            Message("Результат проверки извещения перед согласованием",
                  "{0}{0}{1}{0}",
                  Environment.NewLine,
                  text);
        }
    }

    public static class ReferenceObjectExtensions
    {
        public static void StartUpdate(this ReferenceObject ro)
        {
            if (ro.IsCheckedOut && !ro.IsCheckedOutByCurrentUser)
                throw new InvalidOperationException(String.Format(TFlex.DOCs.Resources.Strings.Messages.CantCheckOutExceptionMessage, ro));

            if (ro.CanCheckOut)
                Desktop.CheckOut(ro, false);

            if (!ro.Changing)
                ro.BeginChanges();
        }

        public static void EndUpdate(this ReferenceObject ro, string comment)
        {
            if (ro.Changing)
                ro.EndChanges();

            if (ro.CanCheckIn)
                Desktop.CheckIn(ro, comment, false);
        }

        public static void CancelUpdate(this ReferenceObject ro)
        {
            if (ro.Changing)
                ro.CancelChanges();

            if (ro.CanUndoCheckOut)
                Desktop.UndoCheckOut(ro);
        }

        public static ComplexHierarchyLink GetToOneComplexHierarchyLink(this ReferenceObject referenceObject,
            Guid linkGuid, DesignContextObject designContext = null, bool applyDesignContext = true, bool applyDate = true)
        {
            var link = referenceObject.Links.ToOneToComplexHierarchy[linkGuid];

            var newConfigurationSettings = new ConfigurationSettings(referenceObject.Reference.Connection)
            {
                DesignContext = designContext,
                ApplyDesignContext = applyDesignContext,
            };
            if (applyDate)
            {
                newConfigurationSettings.Date = DateTime.Now;
                newConfigurationSettings.ApplyDate = true;
            }

            ComplexHierarchyLink linkedComplexLink = null;
            using (link.LinkReference.ChangeAndHoldConfigurationSettings(newConfigurationSettings))
            {
                link.Reload();
                linkedComplexLink = link.LinkedComplexLink;
            }
            return linkedComplexLink;
        }
    }
}
