using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using TFlex.DOCs.Common;
using TFlex.DOCs.Model.Desktop;
using TFlex.DOCs.Model.Macros;
using TFlex.DOCs.Model.Macros.ObjectModel;
using TFlex.DOCs.Model.Macros.Processes;
using TFlex.DOCs.Model.Processes.Events.Contexts;
using TFlex.DOCs.Model.References;
using TFlex.DOCs.Model.References.Documents;
using TFlex.DOCs.Model.References.Files;
using TFlex.DOCs.Model.References.Macros;
using TFlex.DOCs.Model.References.Modifications;
using TFlex.DOCs.Model.References.Nomenclature;
using TFlex.DOCs.Model.Search;
using TFlex.DOCs.Model.Stages;
using TFlex.DOCs.Resources.Strings;

namespace PDM
{
    public class PDM___Извещения : ProcessActionMacroProvider
    {
        /// <summary> Стадия по умолчанию - "Хранение" </summary>
        private Guid _defaultStageGuid = new Guid("9826e84a-a0a7-404b-bf0e-d61b902e346a");
        /// <summary> Стадия для извещения по умолчанию - "Хранение" </summary>
        private Guid _defaultModificationNoticeStageGuid = new Guid("9826e84a-a0a7-404b-bf0e-d61b902e346a");
        /// <summary> Стадия "Исправление" </summary>
        private Guid _changingStageGuid = new Guid("006dfe49-78c6-453d-a65d-c02b24fc4f98");

        public static readonly string InventoryCardMacroGuid = "a184d20e-8f23-49fa-964f-c61b81b266a6";

        /// <summary> Флаг удаления объектов из корзины </summary>
        private static readonly bool _deleteFromRecycleBin = false;

        #region Папки

        private static readonly string _папкаАрхив = "Архив";
        private static readonly string _папкаПодлинники = Path.Combine(_папкаАрхив, "Подлинники");
        private static readonly string _папкаОригиналы = Path.Combine(_папкаАрхив, "Оригиналы");
        private static readonly string _папкаОригиналыTF = Path.Combine(_папкаОригиналы, "TF");
        private static readonly string _папкаОригиналыMS = Path.Combine(_папкаОригиналы, "MS");
        private static readonly string _папкаОригиналыДругое = Path.Combine(_папкаОригиналы, "Другое");

        #endregion

        private static class Guids
        {
            public static readonly Guid ModificationNoticeType = new Guid("52ccb35c-67c5-4b82-af4f-e8ceac4e8d02");
            public static readonly Guid ModificationNoticesSetType = new Guid("322d371b-c672-40ea-9166-316140fc28cb");

            public static readonly Guid ModificationNoticeToModificationsLink = new Guid("5e46670a-400c-4e36-bb37-d4d651bdf692");

            public static readonly Guid СвязьИсходнаяРевизия = new Guid("49028ea0-3bc9-40d5-ae3c-a1e7b12c34e2");

            public static readonly Guid UsingAreaSourceLink = new Guid("b0a49e16-0a7e-49a3-a897-b7fdbf3fccee");
            public static readonly Guid UsingAreaAddedLink = new Guid("bac1b7b8-bd2e-4198-beb0-3580ee077534");
            public static readonly Guid UsingAreaDeletedLink = new Guid("37210965-e2fc-4a44-b0ce-87dde109c458");

            public static readonly Guid ПараметрЦелеваяСтадияРевизии = new Guid("895675ef-e25f-449e-a589-2d6a3cf1026e");
            public static readonly Guid ПараметрНомерИзменения = new Guid("91486563-d044-4045-814b-3432b67812f1");

            public static readonly Guid СправочникИзменения = new Guid("c9a4bb1b-cacb-4f2d-b61a-265f1bfc7fb9");

            public static readonly Guid МакросПроверкиИИ = new Guid("26c5d6ef-a79d-434a-9a9e-fff341463b81");
        }

        private FileReference _fileReference;
        private readonly List<TFlex.DOCs.Model.References.Files.FolderObject> _folders = new List<TFlex.DOCs.Model.References.Files.FolderObject>();

        public PDM___Извещения(EventContext context)
            : base(context)
        {
        }

        private FileReference FileReferenceInstance => _fileReference ?? (_fileReference = new FileReference(Context.Connection));

        public override void Run()
        {
        }

        public bool ПроверитьИИ(ReferenceObject modificationNotice)
        {
            var macro = Context[nameof(MacroKeys.MacroSource)] as CodeMacro;
            if (macro.DebugMode.Value == true)
            {
                Debugger.Launch();
                Debugger.Break();
            }

            RefObjList includedObjects = RunMacro(Guids.МакросПроверкиИИ.ToString(), "CheckModificationNotice", modificationNotice, true);
            if (includedObjects.IsNullOrEmpty())
                return false;

            var includedReferenceObjects = includedObjects.To<ReferenceObject>();

            var lockedObjects = includedReferenceObjects
                .Where(referenceObject => !referenceObject.Reference.ParameterGroup.SupportsDesktop && referenceObject.IsCheckedOut)
                .ToList();
            if (lockedObjects.Any())
            {
                var text = new StringBuilder();

                foreach (var lockedObject in lockedObjects)
                {
                    string str = String.Format("'{0}' [тип '{1}'], редактируется пользователем '{2}'",
                        lockedObject.ToString(),
                        lockedObject.Class,
                        lockedObject.SystemFields.ClientView.Name
                       );

                    text.AppendLine(str);
                }

                Message("Сообщение",
                    String.Format("Объекты заблокированы, процесс применения извещения не будет запущен:{0}{0}{1}",
                    Environment.NewLine,
                    text.ToString()));
                return false;
            }

            var checkingObjects = includedReferenceObjects
                .Where(referenceObject => referenceObject.Reference.ParameterGroup.SupportsDesktop && referenceObject.IsCheckedOut)
                .ToList();
            if (checkingObjects.Any())
            {
                var text = new StringBuilder();

                foreach (var editingObject in checkingObjects)
                {
                    string str = String.Format("'{0}' [тип '{1}']",
                        editingObject.ToString(),
                        editingObject.Class);

                    text.AppendLine(str);
                }

                if (!Question(String.Format(
                    "Перед запуском процесса применения извещения необходимо применить изменения к следующим объектам:{0}{0}{1}{0}Продолжить?",
                    Environment.NewLine,
                    text.ToString())))
                    return false;

                var checkedInObjects = Desktop.CheckIn(checkingObjects, String.Empty, true);
                if (checkedInObjects.IsNullOrEmpty())
                {
                    Message("Сообщение", "К объектам не были применены изменения: процесс применения извещения не будет запущен.");
                    return false;
                }
            }

            return true;
        }

        public void ПроверитьИзвещениеПередПрименением()
        {
            var macro = Context[nameof(MacroKeys.MacroSource)] as CodeMacro;
            if (macro.DebugMode.Value == true)
            {
                Debugger.Launch();
                Debugger.Break();
            }

            var modificationNotices = GetModificationNotices();
            if (modificationNotices.IsNullOrEmpty())
            {
                Переменные["Результат проверки возможности применения изменений"] = "Не найдено извещение";
                return;
            }

            var errors = new StringBuilder();

            var allModifications = new List<ReferenceObject>();
            foreach (var modificationNotice in modificationNotices)
            {
                var modifications = modificationNotice.GetObjects(Guids.ModificationNoticeToModificationsLink).ToList();
                if (modifications.IsNullOrEmpty())
                    errors.AppendLine($"Извещение '{modificationNotice}' не содержит изменений.");
                else
                    allModifications.AddRange(modifications);
            }

            if (allModifications.IsNullOrEmpty())
            {
                Переменные["Результат проверки возможности применения изменений"] = errors.ToString();
                return;
            }

            foreach (var modification in allModifications.OfType<ModificationReferenceObject>())
            {
                var pdmObject = modification.GetObject(ModificationReferenceObject.RelationKeys.PDMObject);
                if (pdmObject is null)
                {
                    errors.AppendLine($"У изменения '{modification}' отсутствует объект ЭСИ.");
                    continue;
                }

                var designContext = modification.DesignContext;
                string designContextName = designContext is null ? Texts.MainContextText : designContext.ToString();

                var modificationNotice = modification.GetObject(Guids.ModificationNoticeToModificationsLink);

                var configurationSettingsInContext = new ConfigurationSettings(Context.Connection)
                {
                    DesignContext = designContext,
                    ApplyDesignContext = true,
                    Date = Texts.TodayText,
                    ApplyDate = true,
                };

                var configurationSettingsInMainContext = new ConfigurationSettings(Context.Connection)
                {
                    DesignContext = null,
                    ApplyDesignContext = true,
                    Date = Texts.TodayText,
                    ApplyDate = true,
                };

                var usingAreaObjects = modification.GetObjects(ModificationReferenceObject.RelationKeys.UsingArea);
                foreach (var usingArea in usingAreaObjects)
                {
                    var hierarchyLinkMatches = usingArea.GetObjects(ModificationUsingAreaReferenceObject.RelationKeys.HierarchyLinkMatches);
                    foreach (var matches in hierarchyLinkMatches)
                    {
                        bool hasSourceLinkInMainContext = false;

                        bool hasSourceLinkInContext = false;
                        bool hasAddedLinkInContext = false;
                        bool hasDeletedLinkInContext = false;

                        ReferenceObject parent = null;

                        var sourceLink = matches.Links.ToOneToComplexHierarchy[Guids.UsingAreaSourceLink];
                        using (sourceLink.LinkReference.ChangeAndHoldConfigurationSettings(configurationSettingsInMainContext))
                        {
                            sourceLink.Reload();

                            if (sourceLink.LinkedComplexLink != null)
                            {
                                hasSourceLinkInMainContext = true;
                                parent = sourceLink.LinkedComplexLink.ParentObject;
                            }
                        }

                        if (!hasSourceLinkInMainContext)
                        {
                            // По связи в исходном может быть подключение, добавленное в выбранном контексте!
                            // при этом не должно быть удаленного (его не надо проверять)
                            using (sourceLink.LinkReference.ChangeAndHoldConfigurationSettings(configurationSettingsInContext))
                            {
                                sourceLink.Reload();

                                if (sourceLink.LinkedComplexLink != null)
                                {
                                    hasSourceLinkInContext = true;
                                    parent = sourceLink.LinkedComplexLink.ParentObject;
                                }
                            }
                        }

                        var addedLink = matches.Links.ToOneToComplexHierarchy[Guids.UsingAreaAddedLink];
                        using (addedLink.LinkReference.ChangeAndHoldConfigurationSettings(configurationSettingsInContext))
                        {
                            addedLink.Reload();

                            if (addedLink.LinkedComplexLink != null)
                            {
                                hasAddedLinkInContext = true;
                                if (parent is null)
                                    parent = addedLink.LinkedComplexLink.ParentObject;
                            }
                        }

                        // если было исходное подключение в основном контексте, должно быть удаленное в контексте
                        if (hasSourceLinkInMainContext || !hasSourceLinkInContext)
                        {
                            var deletedLink = matches.Links.ToOneToComplexHierarchy[Guids.UsingAreaDeletedLink];
                            using (deletedLink.LinkReference.ChangeAndHoldConfigurationSettings(configurationSettingsInContext))
                            {
                                deletedLink.Reload();

                                if (deletedLink.LinkedComplexLink != null)
                                {
                                    hasDeletedLinkInContext = true;
                                    if (parent is null)
                                        parent = deletedLink.LinkedComplexLink.ParentObject;
                                }
                            }
                        }

                        if (!hasAddedLinkInContext)
                        {
                            errors.AppendLine($"Извещение '{modificationNotice}':" +
                                $" в изменении документа '{pdmObject}' отсутствует подключение к '{parent}'," +
                                $" добавленное в контексте '{designContextName}'. Проверьте область применения.");
                        }

                        if (!hasSourceLinkInContext && !hasDeletedLinkInContext)
                        {
                            errors.AppendLine($"Извещение '{modificationNotice}':" +
                                $" в изменении документа '{pdmObject}' отсутствует подключение к '{parent}'," +
                                $" удалённое в контексте '{designContextName}'. Проверьте область применения.");
                        }
                    }
                }

                var actions = modification.GetObjects(ModificationReferenceObject.RelationKeys.ModificationActions).OfType<ModificationActionReferenceObject>();
                foreach (var action in actions)
                {
                    // исключаем из проверки действие "Изменение объекта"
                    if (action.Class.IsEditObjectAction)
                        continue;

                    var actionObjectsLink = action.Links.ToManyToComplexHierarchy
                        .Find(ModificationActionReferenceObject.RelationKeys.ModificationActionObjectsHierarchyLink);

                    using (actionObjectsLink.LinkReference.ChangeAndHoldConfigurationSettings(configurationSettingsInContext))
                    {
                        actionObjectsLink.LinkReference.Refresh();

                        var actionObjects = actionObjectsLink.GetLinkedComplexLinks().OfType<NomenclatureHierarchyLink>().ToList();
                        if (actionObjects.IsNullOrEmpty())
                        {
                            errors.AppendLine($"Извещение '{modificationNotice}': в изменении документа '{pdmObject}' у действия '{action.Class}'" +
                                $" отсутствуют связанные подключения в контексте '{designContextName}'.");
                        }
                    }
                }
            }

            string errorsText = errors.ToString();
            if (String.IsNullOrEmpty(errorsText))
                Переменные["Результат проверки возможности применения изменений"] = "Готово к применению";
            else
                Переменные["Результат проверки возможности применения изменений"] = errorsText;
        }

        public void ПрименитьИИ()
        {
            var macro = Context[nameof(MacroKeys.MacroSource)] as CodeMacro;
            if (macro.DebugMode.Value == true)
            {
                Debugger.Launch();
                Debugger.Break();
            }

            var modificationNotices = GetModificationNotices();

            foreach (var modificationNotice in modificationNotices)
                ProcessModificationNotice(modificationNotice);

            var modificationNoticeSets = GetModificationNoticeSets();
            ChangeStage(_defaultModificationNoticeStageGuid, modificationNoticeSets);
        }

        private List<ReferenceObject> GetModificationNotices()
            => Objects
            .Select(obj => (ReferenceObject)obj)
            .Where(obj => obj.Class.IsInherit(Guids.ModificationNoticeType)).ToList();

        private List<ReferenceObject> GetModificationNoticeSets()
            => Objects
            .Select(obj => (ReferenceObject)obj)
            .Where(obj => obj.Class.IsInherit(Guids.ModificationNoticesSetType)).ToList();

        private void ProcessModificationNotice(ReferenceObject modificationNotice)
        {
            var setStages = new Dictionary<Guid, List<ReferenceObject>>
            {
                { _defaultStageGuid, new List<ReferenceObject>() }
            };

            var deletingFiles = new List<ReferenceObject>();

            var modificationReference = Context.Connection.CreateReference(Guids.СправочникИзменения);

            var modifications = modificationNotice.GetObjects(Guids.ModificationNoticeToModificationsLink).ToList();
            foreach (var modification in modifications)
            {
                ApplyHierarchyLinks(modification);

                setStages[_defaultStageGuid].Add(modification);

                var targetRevision = modification.GetObject(ModificationReferenceObject.RelationKeys.PDMObject);
                if (targetRevision.Class.Scheme != null)
                    setStages[_defaultStageGuid].Add(targetRevision);

                var document = (targetRevision as NomenclatureObject)?.LinkedObject as EngineeringDocumentObject;
                if (document != null)
                {
                    var allTargetRevisionFiles = document.GetFiles();
                    if (allTargetRevisionFiles.Any())
                    {
                        var storageFiles = ProcessTargetRevisionFiles(document, allTargetRevisionFiles, out var nonDuplicatedFiles);
                        var setDefaultStageFiles = storageFiles.Where(file => file.SystemFields.Stage?.Stage?.Guid != _defaultStageGuid).ToList<ReferenceObject>();
                        if (setDefaultStageFiles.Any())
                            setStages[_defaultStageGuid].AddRange(setDefaultStageFiles);

                        deletingFiles.AddRange(allTargetRevisionFiles.Except(storageFiles).Except(nonDuplicatedFiles));

                        var inventoryCard = (ReferenceObject)RunMacro(InventoryCardMacroGuid, "FindOrCreateDocumentInventoryCard", document, true);

                        foreach (var file in storageFiles)
                        {
                            RunMacro(InventoryCardMacroGuid, "ProcessInventoryCard",
                                inventoryCard,
                                document,
                                file,
                                false,
                                true);
                        }
                    }
                }

                var initialRevisionsList = modification.GetObjects(ModificationReferenceObject.RelationKeys.SourceRevisions);
                foreach (var revisionListObject in initialRevisionsList)
                {
                    Guid stageGuid = _defaultStageGuid;

                    if (revisionListObject.ParameterValues.Contains(Guids.ПараметрЦелеваяСтадияРевизии))
                    {
                        var settingStage = (Guid)revisionListObject[Guids.ПараметрЦелеваяСтадияРевизии].Value;
                        if (settingStage != null && settingStage != Guid.Empty)
                            stageGuid = settingStage;
                    }

                    var initialRevision = revisionListObject.GetObject(Guids.СвязьИсходнаяРевизия) as NomenclatureObject;
                    var initialRevisionFiles = initialRevision
                        .GetAllLinkedFiles()
                        .Where(file => file.SystemFields.Stage?.Stage?.Guid != stageGuid).ToList<ReferenceObject>();

                    if (initialRevisionFiles.Any())
                        AddToSetStage(setStages, initialRevisionFiles, stageGuid);

                    if (initialRevision.Class.Scheme != null)
                    {
                        if (initialRevision.SystemFields.Stage?.Stage?.Guid != stageGuid)
                            AddToSetStage(setStages, new List<ReferenceObject>() { initialRevision }, stageGuid);
                    }
                }

                modificationReference.Objects.Load();

                var existingModifications = modificationReference.Objects
                     .Where(mod =>
                     mod.TryGetObject(ModificationReferenceObject.RelationKeys.PDMObject, out var revision)
                     && revision != null
                     && revision.SystemFields.LogicalObjectGuid == targetRevision.SystemFields.LogicalObjectGuid
                     );

                int maxValue = 0;
                if (existingModifications.Any())
                {
                    maxValue = existingModifications.Max(mod =>
                    {
                        string modValue = mod[Guids.ПараметрНомерИзменения].Value.ToString();
                        if (String.IsNullOrEmpty(modValue))
                            return 0;

                        string[] parts = modValue.Split('.');
                        if (parts.Length == 2 && parts[0].ToLower() == "ии" && Int32.TryParse(parts[1], out int number))
                            return number;

                        return 0;
                    });
                }

                modification.BeginChanges();
                modification[Guids.ПараметрНомерИзменения].Value = String.Format("ИИ.{0}", ++maxValue);
                modification.EndChanges();
            }

            // Изменение стадий объектов
            foreach (KeyValuePair<Guid, List<ReferenceObject>> stagePair in setStages)
            {
                ChangeStage(stagePair.Key, stagePair.Value);
            }

            // По умолчанию удаляем все файлы
            string objectErrors = DeleteReferenceObjects(deletingFiles, _deleteFromRecycleBin);
            if (!String.IsNullOrEmpty(objectErrors))
            {
                //%%TODO
            }

            ChangeStage(_defaultModificationNoticeStageGuid, new List<ReferenceObject>() { modificationNotice });
        }

        private static void AddToSetStage(Dictionary<Guid, List<ReferenceObject>> setStages, List<ReferenceObject> objects, Guid stageGuid)
        {
            if (setStages.TryGetValue(stageGuid, out List<ReferenceObject> referenceObjects))
                referenceObjects.AddRange(objects);
            else
                setStages.Add(stageGuid, objects);
        }

        private void ApplyHierarchyLinks(ReferenceObject modification)
        {
            var designContext = (DesignContextObject)modification.GetObject(ModificationReferenceObject.RelationKeys.DesignContext);

            var newConfigurationSettings = new ConfigurationSettings(Context.Connection)
            {
                DesignContext = designContext,
                ApplyDesignContext = true,
                Date = Texts.TodayText,
                ApplyDate = true,
            };

            // Обработка подключений применяемости
            var usingAreaObjects = modification.GetObjects(ModificationReferenceObject.RelationKeys.UsingArea);
            foreach (var usingArea in usingAreaObjects)
            {
                var hierarchyLinkMatches = usingArea.GetObjects(ModificationUsingAreaReferenceObject.RelationKeys.HierarchyLinkMatches);
                foreach (var matches in hierarchyLinkMatches)
                {
                    var addedLink = matches.Links.ToOneToComplexHierarchy[Guids.UsingAreaAddedLink];
                    using (addedLink.LinkReference.ChangeAndHoldConfigurationSettings(newConfigurationSettings))
                    {
                        addedLink.Reload();
                        ProcessHierarchyLink(addedLink.LinkedComplexLink as NomenclatureHierarchyLink);
                    }

                    var deletedLink = matches.Links.ToOneToComplexHierarchy[Guids.UsingAreaDeletedLink];
                    using (deletedLink.LinkReference.ChangeAndHoldConfigurationSettings(newConfigurationSettings))
                    {
                        deletedLink.Reload();
                        ProcessHierarchyLink(deletedLink.LinkedComplexLink as NomenclatureHierarchyLink);
                    }
                }
            }

            // Обработка подключений в действиях изменения
            var actions = modification.GetObjects(ModificationReferenceObject.RelationKeys.ModificationActions).OfType<ModificationActionReferenceObject>().ToList();
            foreach (var action in actions)
            {
                if (action.Class.IsEditObjectAction)
                    continue;

                var actionObjectsLink = action.Links.ToManyToComplexHierarchy
                    .Find(ModificationActionReferenceObject.RelationKeys.ModificationActionObjectsHierarchyLink);

                using (actionObjectsLink.LinkReference.ChangeAndHoldConfigurationSettings(newConfigurationSettings))
                {
                    actionObjectsLink.LinkReference.Refresh();

                    var actionObjects = actionObjectsLink.GetLinkedComplexLinks().OfType<NomenclatureHierarchyLink>().ToList();
                    foreach (var linkedHierarchyLink in actionObjects.Where(
                        lnk => lnk.DesignContextId.Value == (designContext?.Id ?? 0)))
                    {
                        ProcessHierarchyLink(linkedHierarchyLink);
                    }
                }
            }
        }

        private void ProcessHierarchyLink(NomenclatureHierarchyLink link)
        {
            if (link is null || link.DesignContextId == 0)
                return;

            link.ApplyDesignContextChangesToMainContext();
        }

        private List<ReferenceObject> ChangeStage(Guid stageGuid, List<ReferenceObject> objects)
        {
            if (!objects.Any())
                return new List<ReferenceObject>(0);

            var stage = Stage.Find(Context.Connection, stageGuid);
            if (stage is null)
                return new List<ReferenceObject>(0);

            return stage.Set(objects);
        }

        private List<FileObject> ProcessTargetRevisionFiles(EngineeringDocumentObject targetRevisionDocument, List<FileObject> files, out List<FileObject> nonDuplicatedFiles)
        {
            nonDuplicatedFiles = new List<FileObject>();
            var storageFiles = new List<FileObject>();

            if (!files.Any())
                return storageFiles;

            var fileReference = new FileReference(Context.Connection);
            var parameterGroup = fileReference.ParameterGroup;
            var filter = new Filter(parameterGroup);

            var relativePathParameter = parameterGroup.Parameters.Find(FileReferenceObject.FieldKeys.Path);

            //Перенос утверждённых файлов в архив
            foreach (var revisionFile in files)
            {
                if (revisionFile.Path.GetString().ToLower().Contains(@"архив\"))
                {
                    storageFiles.Add(revisionFile);
                    continue;
                }

                // выбор папки для хранения утвержденных файлов
                TFlex.DOCs.Model.References.Files.FolderObject newFolder;
                switch (revisionFile.Class.Extension.ToLower())
                {
                    case ("tif"):
                    case ("tiff"):
                    case ("pdf"):
                        newFolder = GetFolder(_папкаПодлинники);
                        break;

                    case ("grb"):
                        newFolder = GetFolder(_папкаОригиналыTF);
                        break;

                    case ("xls"):
                    case ("xlsx"):
                    case ("doc"):
                    case ("docx"):
                        newFolder = GetFolder(_папкаОригиналыMS);
                        break;

                    default:
                        newFolder = GetFolder(_папкаОригиналыДругое);
                        break;
                }

                try
                {
                    filter.Terms.Clear();

                    string newRelativeFilePath = Path.Combine(newFolder.Path, revisionFile.Name);

                    filter.Terms.AddTerm(
                        relativePathParameter,
                        ComparisonOperator.Equal,
                        newRelativeFilePath);

                    var existingFiles = fileReference.Find(filter);
                    if (existingFiles.IsNullOrEmpty())
                    {
                        // делаем дубликат файла в соответствующей папке
                        var duplicatedFile = revisionFile.CreateFileDuplicate(newFolder);
                        // добавляем в список объектов для смены стадии на "Хранение"
                        storageFiles.Add(duplicatedFile);
                    }
                    else
                    {
                        //%%TODO 
                        revisionFile.GetHeadRevision();
                        if (!File.Exists(revisionFile.LocalPath))
                            continue;

                        var existingFile = existingFiles.FirstOrDefault();

                        var fileInChangingStage = ChangeStage(_changingStageGuid, new List<ReferenceObject>(1) { existingFile })
                            .OfType<FileObject>()
                            .FirstOrDefault();

                        if (fileInChangingStage is null)
                            continue; //%%TODO

                        fileInChangingStage.StartUpdate();

                        File.Copy(revisionFile.LocalPath, fileInChangingStage.LocalPath, true);
                        new FileInfo(fileInChangingStage.LocalPath).IsReadOnly = false;

                        fileInChangingStage.AddLinkedObject(EngineeringDocumentFields.File, targetRevisionDocument);

                        fileInChangingStage.EndUpdate(String.Format("Копирование файла из '{0}' в '{1}'", revisionFile.Path, newRelativeFilePath));

                        // добавляем в список объектов для смены стадии на "Хранение"
                        storageFiles.Add(fileInChangingStage);
                    }
                }
                catch (Exception)
                {
                    nonDuplicatedFiles.Add(revisionFile);
                }
            }

            return storageFiles;
        }

        private string DeleteReferenceObjects(List<ReferenceObject> referenceObjects, bool deleteFromRecycleBin = true)
        {
            if (referenceObjects.IsNullOrEmpty())
                return String.Empty;

            var deletingObjects = referenceObjects.Where(referenceObject => referenceObject != null);

            var referenceObjectsWithSupportDesktop = deletingObjects.Where(ro => ro.Reference.ParameterGroup.SupportsDesktop).ToList();

            var referenceObjectsWithoutSupportDesktop = deletingObjects.Except(referenceObjectsWithSupportDesktop).Where(ro => ro.CanDelete).ToArray();
            Array.ForEach(referenceObjectsWithoutSupportDesktop, ro => ro.Delete());

            var addedObjectsWithSupportDesktop = referenceObjectsWithSupportDesktop.Where(ro => ro.LockState == ReferenceObjectLockState.CheckedOutForAdd).ToArray();
            Desktop.UndoCheckOut(addedObjectsWithSupportDesktop);

            var savedObjectsWithSupportDesktop = referenceObjectsWithSupportDesktop.Except(addedObjectsWithSupportDesktop).ToList();
            var checkedOutObjectsWithSupportDesktop = savedObjectsWithSupportDesktop.Where(ro => ro.IsCheckedOutByCurrentUser).ToList();
            Desktop.UndoCheckOut(checkedOutObjectsWithSupportDesktop);

            var canCheckOutObjects = savedObjectsWithSupportDesktop.Where(ro => ro.CanCheckOut).ToList();
            Desktop.CheckOut(canCheckOutObjects, true);

            var objectNames = new StringBuilder();
            foreach (var referenceObject in canCheckOutObjects)
            {
                if (referenceObject is FileObject)
                    objectNames.AppendLine(((FileObject)referenceObject).Path);
                else
                    objectNames.AppendLine(referenceObject.ToString());
            }

            var deletedObjects = Desktop.CheckIn(
                canCheckOutObjects,
                String.Format("Автоматическое удаление объектов:{0}{1}", Environment.NewLine, objectNames),
                false);

            string clearRecycleBinResult = String.Empty;

            if (deleteFromRecycleBin)
            {
                var objectsInRecycleBin = deletingObjects.Union(deletedObjects).Where(desktopObject => ((ReferenceObject)desktopObject).IsInRecycleBin);
                try
                {
                    Desktop.ClearRecycleBin(objectsInRecycleBin);
                }
                catch (EmptyRecycleBinError exception)
                {
                    clearRecycleBinResult = exception.Message;
                }
            }

            return clearRecycleBinResult;
        }

        private TFlex.DOCs.Model.References.Files.FolderObject GetFolder(string folderRelativePath)
        {
            if (FileReferenceInstance is null)
                return null;

            if (String.IsNullOrEmpty(folderRelativePath))
                return null;

            var addedFolder = _folders.FirstOrDefault(folderObject => folderObject.Path == folderRelativePath);
            if (addedFolder != null)
                return addedFolder;

            if (!(FileReferenceInstance.FindByRelativePath(folderRelativePath) is TFlex.DOCs.Model.References.Files.FolderObject folder))
            {
                folder = FileReferenceInstance.CreatePath(folderRelativePath, null);
                if (folder is null)
                    Message("Создание папки", TFlex.DOCs.Resources.Strings.Messages.UnableToCreateFolderMessage, folderRelativePath);
                else
                    Desktop.CheckIn(folder, TFlex.DOCs.Resources.Strings.Texts.AutoCreateText, false);
            }

            if (folder != null)
                _folders.Add(folder);

            return folder;
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
    }
}

