using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using TFlex.DOCs.Client;
using TFlex.DOCs.Client.ViewModels;
using TFlex.DOCs.Client.ViewModels.Base;
using TFlex.DOCs.Client.ViewModels.Commands.Base;
using TFlex.DOCs.Client.ViewModels.Commands.References;
using TFlex.DOCs.Client.ViewModels.Commands.References.DesktopOperationsCommand;
using TFlex.DOCs.Client.ViewModels.Commands.Tree;
using TFlex.DOCs.Client.ViewModels.References;
using TFlex.DOCs.Client.ViewModels.References.Modifications;
using TFlex.DOCs.Client.ViewModels.References.Nomenclature;
using TFlex.DOCs.Client.ViewModels.References.SelectionDialogs;
using TFlex.DOCs.Common;
using TFlex.DOCs.Model.Macros;
using TFlex.DOCs.Model.Macros.ObjectModel;
using TFlex.DOCs.Model.References;
using TFlex.DOCs.Model.References.Documents;
using TFlex.DOCs.Model.References.Files;
using TFlex.DOCs.Model.References.Macros;
using TFlex.DOCs.Model.References.Modifications;
using TFlex.DOCs.Model.References.Nomenclature;
using TFlex.DOCs.Model.References.Users;
using TFlex.DOCs.Resources.Strings;

namespace PDM
{
    public class PDM___Изменения : MacroProvider
    {
        private static class Guids
        {
            public static readonly Guid ModificationNumberParameter = new Guid("91486563-d044-4045-814b-3432b67812f1");

            public static readonly Guid LinkInitialRevision = new Guid("49028ea0-3bc9-40d5-ae3c-a1e7b12c34e2");
            public static readonly Guid RevisionStageParameter = new Guid("895675ef-e25f-449e-a589-2d6a3cf1026e");

            public static readonly Guid UsingAreaInitialRevision = new Guid("00f158e5-eab7-4006-b672-14b6b6f4c92f");
            public static readonly Guid UsingAreaAddedLink = new Guid("bac1b7b8-bd2e-4198-beb0-3580ee077534");
            public static readonly Guid UsingAreaDeletedLink = new Guid("37210965-e2fc-4a44-b0ce-87dde109c458");
            public static readonly Guid UsingAreaSourceLink = new Guid("b0a49e16-0a7e-49a3-a897-b7fdbf3fccee");

            public static readonly Guid ИИВыпущеноНа = new Guid("7ee36a71-a877-4327-87cc-c1d34c92d9e4");
            public static readonly Guid ИзвещениеОбИзменении = new Guid("52ccb35c-67c5-4b82-af4f-e8ceac4e8d02");

            /// <summary> Стадия "Хранение" </summary>
            public static readonly Guid StorageStageGuid = new Guid("9826e84a-a0a7-404b-bf0e-d61b902e346a");
        }

        private static class Resources
        {

            public const string DialogTargetRevisionFileString = "Файл целевой ревизии: ";
            public const string DialogInitialRevisionString = "Исходная ревизия: ";
        }
        private static readonly string _шаблонНаименованияВарианта = "КИ.{0}-{1}";
        private static readonly string _шаблонДаты = "dd.MM.yy"; //"dd.MM.yy.HH.mm.ss";
        private static readonly string _tempFolder = Path.Combine(Path.GetTempPath(), "Temp DOCs", "CompareFiles");

        public PDM___Изменения(MacroContext context) : base(context)
        {
        }

        public override void Run()
        {
        }

        public ButtonValidator ValidateCreateModificationNoticeButton()
        {
            var validator = new ButtonValidator();

            if (ApplicationManager.IsLaunchedPdm || !(Context is UIMacroContext))
            {
                validator.Enable = false;
                validator.Visible = false;
            }

            return validator;
        }

        public void СоздатьИзвещение()
        {
            var macro = Context[nameof(MacroKeys.MacroSource)] as CodeMacro;
            if (macro.DebugMode.Value == true)
            {
                Debugger.Launch();
                Debugger.Break();
            }

            if (!(Context is UIMacroContext uiContext))
                return;

            List<NomenclatureObjectViewModel> selectedObjects = null;

            var currentObjects = (uiContext.OwnerViewModel as ISupportSelection)?.SelectedObjects.OfType<NomenclatureObjectViewModel>().ToList();

            var newConfigurationSettings = new ConfigurationSettings(Context.Reference.ConfigurationSettings)
            {
                ShowDeletedInDesignContextLinks = false,
                Date = Texts.TodayText,
                ApplyDate = true,
                ApplyDesignContext = true
            };

            using (Context.Reference.ChangeAndHoldConfigurationSettings(newConfigurationSettings))
            {
                using (var selectViewModel = new SelectPDMObjectViewModel(currentObjects, true, true, Context.Reference))
                {
                    if (ApplicationManager.OpenDialog(selectViewModel, default(IDialogOwnerWindowSource)))
                        selectedObjects = selectViewModel.GetSelectedObjects().OfType<NomenclatureObjectViewModel>().ToList();
                }
            }

            if (selectedObjects.IsNullOrEmpty())
                return;

            var modificationObjects = CreateModifications(selectedObjects);
            CreateModificationNotice(modificationObjects);
        }

        private List<ReferenceObject> CreateModifications(List<NomenclatureObjectViewModel> selectedObjects, ReferenceObject modificationNotice = null)
        {
            var modificationObjects = new List<ReferenceObject>();

            var modificationReference = Context.Connection.ReferenceCatalog.Find(ModificationReference.ReferenceId)?.CreateReference();

            string currentDate = DateTime.Now.ToString(_шаблонДаты);
            string userName = CurrentUser[User.Fields.LastName.ToString()];

            var groupedObjectsByParent = selectedObjects.GroupBy(obj => obj.AllHierarchyLinks?.FirstOrDefault()?.ParentObject);

            foreach (var groupElement in groupedObjectsByParent)
            {
                var parent = groupElement.Key;

                foreach (var element in groupElement)
                {
                    var pdmObject = element.ReferenceObject as NomenclatureReferenceObject;

                    var addedLinks = new List<NomenclatureHierarchyLink>();
                    var deletedLinks = new List<NomenclatureHierarchyLink>();
                    var sourceLinks = new List<NomenclatureHierarchyLink>();

                    var hierarchyLinks = element.AllHierarchyLinks.OfType<NomenclatureHierarchyLink>().ToList();
                    if (!hierarchyLinks.IsNullOrEmpty())
                    {
                        var linksNotDeletedInContext = hierarchyLinks.Where(lnk => lnk.DesignContextId > 0 && !lnk.DeletedInDesignContext);

                        //Добавленные подключения
                        addedLinks = linksNotDeletedInContext.Where(lnk => lnk.SubstitutedLinkId == 0).ToList();
                        if (addedLinks.Any())
                        {
                            if (parent != null)
                            {
                                deletedLinks = parent.Children.GetHierarchyLinks().OfType<NomenclatureHierarchyLink>()
                                      .Where(lnk =>
                                      !addedLinks.Select(l => l.Id).Contains(lnk.Id) &&
                                      lnk.DeletedInDesignContext &&
                                      lnk.ChildObject.SystemFields.LogicalObjectGuid == pdmObject.SystemFields.LogicalObjectGuid).ToList();

                                if (deletedLinks.Any())
                                {
                                    var reference = new NomenclatureReference(Context.Connection);
                                    using (reference.ClearAndHoldUseConfigurationSettings())
                                    {
                                        var parentObject = reference.Find(parent.Id);
                                        sourceLinks = parentObject.Children.GetHierarchyLinks().OfType<NomenclatureHierarchyLink>()
                                             .Where(lnk => deletedLinks.Select(l => l.SubstitutedLinkId.Value).Contains(lnk.Id)).ToList();
                                    }
                                }
                            }
                        }

                        //Измененные подключения
                        var editedLinks = linksNotDeletedInContext.Where(lnk => lnk.SubstitutedLinkId > 0).ToList();
                        if (editedLinks.Any())
                        {
                            // Нужно в изменении для родительского объекта создать действие "Изменение"

                            var modificationForParent = modificationObjects.FirstOrDefault(
                                     mod => mod.GetObject(ModificationReferenceObject.RelationKeys.PDMObject) == parent);

                            if (modificationForParent != null)
                            {
                                modificationForParent.BeginChanges();

                                //Создать действие "Изменение"
                                var action = modificationForParent.CreateListObject(
                                    ModificationReferenceObject.RelationKeys.ModificationActions, ModificationActionTypes.Keys.EditAction)
                                    as ModificationActionReferenceObject;

                                foreach (var editedLink in editedLinks)
                                {
                                    action.Links.ToManyToComplexHierarchy[ModificationActionReferenceObject.RelationKeys.ModificationActionObjectsHierarchyLink]
                                        .AddLinkedComplexLink(editedLink);
                                }

                                //Сформировать текст
                                var macroContext = new UIMacroContext(new ReferenceObjectViewModel(action));
                                macroContext.RunMacro("cd2247eb-47ee-4944-9369-e73ee2fb5e5e",
                                     "CreateActionText",
                                     editedLinks);

                                action.EndChanges();

                                modificationForParent.EndChanges();
                            }
                        }
                    }

                    var existingModification = modificationObjects.FirstOrDefault(
                        mod => mod.GetObject(ModificationReferenceObject.RelationKeys.PDMObject) == pdmObject);

                    if (existingModification != null)
                        continue;

                    //создаем изменение
                    var modificationClass = modificationReference.Classes.Find(ModificationTypes.Keys.Modification);
                    var modification = modificationReference.CreateReferenceObject(modificationClass);
                    modification[Guids.ModificationNumberParameter].Value = String.Format(_шаблонНаименованияВарианта, userName, currentDate);
                    modification.SetLinkedObject(ModificationReferenceObject.RelationKeys.DesignContext, Context.ConfigurationSettings.DesignContext);
                    modification.SetLinkedObject(ModificationReferenceObject.RelationKeys.PDMObject, pdmObject);

                    //Создаём действие "Изменение объекта" для объекта ЭСИ
                    var editAction = modification.CreateListObject(
                        ModificationReferenceObject.RelationKeys.ModificationActions, ModificationActionTypes.Keys.EditObjectAction)
                        as ModificationActionReferenceObject;

                    var context = new UIMacroContext(new ReferenceObjectViewModel(editAction));
                    context.RunMacro("cd2247eb-47ee-4944-9369-e73ee2fb5e5e",
                         "CreateActionTextForPdmObject",
                         pdmObject,
                         sourceLinks,
                         addedLinks
                         );

                    editAction.EndChanges();

                    //Создать область применения
                    if (!addedLinks.IsNullOrEmpty() && !deletedLinks.IsNullOrEmpty() && !sourceLinks.IsNullOrEmpty())
                    {
                        var usingArea = modification.CreateListObject(ModificationReferenceObject.RelationKeys.UsingArea, ModificationUsingAreaTypes.Keys.UsingArea)
                            as ModificationUsingAreaReferenceObject;

                        foreach (var addedLink in addedLinks)
                        {
                            int currentIndex = addedLinks.IndexOf(addedLink);
                            var deletedLink = deletedLinks.ElementAtOrDefault(currentIndex);
                            var sourceLink = sourceLinks.ElementAtOrDefault(currentIndex);

                            if (deletedLink is null || sourceLink is null)
                                break;

                            CreateMatchingLinksObject(addedLink, deletedLink, sourceLink, usingArea);
                        }

                        //Сформировать текст области применения
                        var macroContext = new UIMacroContext(new ReferenceObjectViewModel(usingArea));
                        macroContext.RunMacro("cd2247eb-47ee-4944-9369-e73ee2fb5e5e",
                             "CreateUsingAreaText",
                             sourceLinks);

                        usingArea.EndChanges();

                        //Создаем объект списка Исходные ревизии типа Исходная ревизия
                        var childObject = sourceLinks.First().ChildObject;
                        var sourceRevisionListObject = modification.CreateListObject(ModificationReferenceObject.RelationKeys.SourceRevisions);
                        sourceRevisionListObject.SetLinkedObject(Guids.LinkInitialRevision, childObject);

                        if (childObject.Class.Scheme != null && childObject.SystemFields.Stage != null)
                            sourceRevisionListObject[Guids.RevisionStageParameter].Value = childObject.SystemFields.Stage.Stage.Guid;

                        sourceRevisionListObject.EndChanges();
                    }

                    if (modificationNotice != null)
                        modification.SetLinkedObject(ModificationReferenceObject.RelationKeys.ModificationNotice, modificationNotice);

                    modification.EndChanges();

                    modificationObjects.Add(modification);
                }
            }
            return modificationObjects;
        }

        static void CreateMatchingLinksObject(NomenclatureHierarchyLink addedLink, NomenclatureHierarchyLink deletedLink, NomenclatureHierarchyLink sourceLink, ModificationUsingAreaReferenceObject usingArea)
        {
            var matchingLinksObject = usingArea.CreateListObject(ModificationUsingAreaReferenceObject.RelationKeys.HierarchyLinkMatches);
            matchingLinksObject.Links.ToOneToComplexHierarchy[Guids.UsingAreaSourceLink].SetLinkedComplexLink(sourceLink);
            matchingLinksObject.Links.ToOneToComplexHierarchy[Guids.UsingAreaAddedLink].SetLinkedComplexLink(addedLink);
            matchingLinksObject.Links.ToOneToComplexHierarchy[Guids.UsingAreaDeletedLink].SetLinkedComplexLink(deletedLink);
            matchingLinksObject.SetLinkedObject(Guids.UsingAreaInitialRevision, sourceLink.ChildObject);
            matchingLinksObject.EndChanges();
        }

        private void CreateModificationNotice(List<ReferenceObject> modificationObjects)
        {
            if (modificationObjects.IsNullOrEmpty())
                return;

            //Справочник "Извещения об изменениях", тип "Извещение об изменении"
            var modificationNoticeReferenceInfo = Context.Connection.ReferenceCatalog.Find(SystemParameterGroups.ModificationNotices);
            var modificationNoticeReference = modificationNoticeReferenceInfo.CreateReference();
            var modificationNoticeClass = modificationNoticeReference.Classes.Find(Guids.ИзвещениеОбИзменении);

            var modificationNotice = modificationNoticeReference.CreateReferenceObject(modificationNoticeClass);

            //Заполняем параметр "Выпущено на", если объект один
            if (modificationObjects.Count == 1)
            {
                var linkedObject = modificationObjects[0].GetObject(ModificationReferenceObject.RelationKeys.PDMObject);
                if (linkedObject != null)
                    modificationNotice[Guids.ИИВыпущеноНа].Value = linkedObject[NomenclatureReferenceObject.FieldKeys.Object].Value.ToString();
            }

            //Связываем все выбранные изменения с созданным извещением
            foreach (var modification in modificationObjects)
            {
                modificationNotice.AddLinkedObject(ModificationReferenceObject.RelationKeys.ModificationNotice, modification);
            }

            modificationNotice.EndChanges();

            ShowPropertyDialog(RefObj.CreateInstance(modificationNotice, Context), true);
        }

        public ButtonValidator ValidateAddDocumentButton()
        {
            var validator = new ButtonValidator();

            if (Context.Reference.IsSlave)
            {
                var masterObject = Context.Reference.LinkInfo.MasterObject;
                if (!masterObject.Class.IsInherit(Guids.ИзвещениеОбИзменении) || masterObject.SystemFields.Stage?.Stage?.Guid == Guids.StorageStageGuid)
                {
                    validator.Enable = false;
                    validator.Visible = false;
                }
            }
            else
            {
                validator.Enable = false;
                validator.Visible = false;
            }

            return validator;
        }

        public void ДобавитьДокументВИзвещение()
        {
            AddDocumentInModificationNotice();
        }

        private void AddDocumentInModificationNotice()
        {
            var macro = Context[nameof(MacroKeys.MacroSource)] as CodeMacro;
            if (macro.DebugMode.Value == true)
            {
                Debugger.Launch();
                Debugger.Break();
            }

            if (!Context.Reference.IsSlave)
                return;

            List<NomenclatureObjectViewModel> selectedObjects = null;

            var selectedModification = Context.ReferenceObject as ModificationReferenceObject;
            var designContext = selectedModification is null
                ? Context.ConfigurationSettings.DesignContext
                : selectedModification.DesignContext;

            var reference = new NomenclatureReference(Context.Connection);

            var newConfigurationSettings = new ConfigurationSettings(reference.ConfigurationSettings)
            {
                ShowDeletedInDesignContextLinks = false,
                Date = Texts.TodayText,
                ApplyDate = true,
                DesignContext = designContext,
                ApplyDesignContext = true,
            };

            using (reference.ChangeAndHoldConfigurationSettings(newConfigurationSettings))
            {
                using (var selectViewModel = new SelectPDMObjectViewModel(null, false, false, reference))
                {
                    if (ApplicationManager.OpenDialog(selectViewModel, default(IDialogOwnerWindowSource)))
                        selectedObjects = selectViewModel.GetSelectedObjects().OfType<NomenclatureObjectViewModel>().ToList();
                }
            }

            if (selectedObjects.IsNullOrEmpty())
                return;

            var modificationNotice = Context.Reference.LinkInfo.MasterObject;
            CreateModifications(selectedObjects, modificationNotice);

            RefreshControls("LinkedModifications");
        }

        public void СравнитьПодлинники()
        {
            CompareRevisionsFiles();
        }

        /// <summary> Сравнение подлинников </summary>
        private void CompareRevisionsFiles()
        {
            var modification = Context.ReferenceObject;

            var pdmObject = modification.GetObject(ModificationReferenceObject.RelationKeys.PDMObject) as NomenclatureObject;
            if (pdmObject is null)
                Error("Не задан объект ЭСИ.");

            var pdmObjectFiles = new List<FileObject>();
            var document = pdmObject.LinkedObject as EngineeringDocumentObject;
            if (document != null)
            {
                var allPdmObjectFiles = document.GetFiles();

                pdmObjectFiles.AddRange(allPdmObjectFiles.Where(file => IsTiffFile(file.Name.GetString())));
            }

            if (!pdmObjectFiles.Any())
                Error($"Не найден подлинник объекта '{pdmObject}'.");


            var initialRevisionListObjects = modification.GetObjects(ModificationReferenceObject.RelationKeys.SourceRevisions);
            var revisions = initialRevisionListObjects
                .Select(listObject => listObject.GetObject(Guids.LinkInitialRevision) as NomenclatureObject)
                .Where(revision =>
                    revision != null
                    && revision.SystemFields.LogicalObjectGuid == pdmObject.SystemFields.LogicalObjectGuid)
                .ToList();

            var targetFile = pdmObjectFiles.FirstOrDefault(file => file.Path.GetString().ToLower().Contains(@"архив\подлинники"));
            var files = targetFile is null ? pdmObjectFiles : new List<FileObject>(1) { targetFile };

            var dialog = CreateDialog(files, revisions);
            if (!dialog.Show())
                return;

            targetFile = dialog[Resources.DialogTargetRevisionFileString];
            NomenclatureObject initialRevision = dialog[Resources.DialogInitialRevisionString];

            if (targetFile is null)
                Error("Не найден подлинник целевой ревизии.");

            if (initialRevision is null)
                Error("Не выбрана ревизия.");


            FileObject initialFile = null;

            var initialDocument = initialRevision.LinkedObject as EngineeringDocumentObject;
            if (initialDocument != null)
            {
                var allInitialRevisionFiles = initialDocument.GetFiles();

                initialFile = allInitialRevisionFiles.FirstOrDefault(
                    file => file.Path.GetString().ToLower().Contains(@"архив\подлинники") && IsTiffFile(file.Name.GetString()));
            }

            if (initialFile is null)
                Error("Не найден подлинник исходной ревизии.");


            // папка для сравнения файлов

            if (!Directory.Exists(_tempFolder))
                Directory.CreateDirectory(_tempFolder);

            string targetRevisionFilePath = Path.Combine(_tempFolder, "ПодлинникЦелевойРевизиии.tiff");
            string initialRevisionFilePath = Path.Combine(_tempFolder, "ПодлинникИсходнойРевизии.tiff");

            if (targetFile.IsAdded || targetFile.IsCheckedOutByCurrentUser)
            {
                File.Copy(targetFile.LocalPath, targetRevisionFilePath, true);
            }
            else
            {
                // загружаем с сервера версию файла
                targetFile.GetFileVersion(targetRevisionFilePath, targetFile.SystemFields.Version);
            }

            new FileInfo(targetRevisionFilePath).IsReadOnly = false;


            if (initialFile.IsAdded || initialFile.IsCheckedOutByCurrentUser)
            {
                // берем локальный файл
                File.Copy(initialFile.LocalPath, initialRevisionFilePath, true);
            }
            else
            {
                // загружаем с сервера версию файла изменения
                initialFile.GetFileVersion(initialRevisionFilePath, initialFile.SystemFields.Version);
            }

            new FileInfo(initialRevisionFilePath).IsReadOnly = false;

            //CompareTiff.exe
            string directoryPath = Path.GetDirectoryName(typeof(MacroContext).Assembly.Location);
            string tiffComparatorSourcePath = Path.Combine(directoryPath, "CompareTiff.exe");
            if (!File.Exists(tiffComparatorSourcePath))
                Error("Файл '{0}' не найден.", tiffComparatorSourcePath);

            var compareTiffProcess = Process.Start(tiffComparatorSourcePath, String.Format("-multi \"{0}\" \"{1}\"", initialRevisionFilePath, targetRevisionFilePath));
            compareTiffProcess.WaitForExit();
        }

        private InputDialog CreateDialog(List<FileObject> files, List<NomenclatureObject> revisions)
        {
            var inputDialog = CreateInputDialog("");

            inputDialog.AddSelectFromList(Resources.DialogTargetRevisionFileString, files.FirstOrDefault(), true, files);
            inputDialog.AddSelectFromList(Resources.DialogInitialRevisionString, revisions.FirstOrDefault(), true, revisions);

            inputDialog.Closing += (bool okButtonClicked, ref bool closeDialog) =>
            {
                if (okButtonClicked)
                {
                    if (!HasFile(inputDialog[Resources.DialogInitialRevisionString]))
                    {
                        Сообщение("Внимание!", "Не найден подлинник выбранной ревизии");
                        closeDialog = false;
                    }
                }
            };

            return inputDialog;
        }

        private bool HasFile(NomenclatureObject initialRevision)
        {
            FileObject initialFile = null;

            var initialDocument = initialRevision.LinkedObject as EngineeringDocumentObject;
            if (initialDocument != null)
            {
                var allInitialRevisionFiles = initialDocument.GetFiles();

                initialFile = allInitialRevisionFiles.FirstOrDefault(
                    file => file.Path.GetString().ToLower().Contains(@"архив\подлинники") && IsTiffFile(file.Name.GetString()));
            }

            return initialFile != null;
        }

        private bool IsTiffFile(string filePath)
        {
            if (String.IsNullOrEmpty(filePath))
                return false;

            string extension = Path.GetExtension(filePath).TrimStart('.').ToLower();
            if (extension == "tif" || extension == "tiff")
                return true;

            return false;
        }


        public void СозданиеДействияИзменение()
        {
            CreatingEditObjectAction();
        }

        public void CreatingEditObjectAction()
        {
            if (!CurrentObject.Class.IsInherit(ModificationActionTypes.Keys.EditObjectAction))
                return;

            UncheckIsAutoText();
        }

        private void UncheckIsAutoText()
        {
            if (Context.ReferenceObject is ModificationActionReferenceObject action)
                action.IsAutoText.Value = false;
        }

        public void ЗавершениеИзмененияПараметра() => OnParameterChanged();

        public void OnParameterChanged()
        {
            var parameterGuid = Context.ChangedParameter.ParameterInfo.Guid;
            if (parameterGuid == ModificationActionReferenceObject.FieldKeys.IsAutoText && Context.ChangedParameter.GetBoolean() == true)
            {
                if (Context.ReferenceObject is ModificationActionReferenceObject action)
                {
                    if (action.Class.IsEditObjectAction)
                        return;

                    var designContext = (DesignContextObject)action.MasterObject.GetObject(ModificationReferenceObject.RelationKeys.DesignContext);

                    var actionObjectsLink = action.Links.ToManyToComplexHierarchy
                    .Find(ModificationActionReferenceObject.RelationKeys.ModificationActionObjectsHierarchyLink);

                    var newConfigurationSettings = new ConfigurationSettings(Context.Connection)
                    {
                        DesignContext = designContext,
                        ApplyDesignContext = true,
                    };

                    using (actionObjectsLink.LinkReference.ChangeAndHoldConfigurationSettings(newConfigurationSettings))
                    {
                        actionObjectsLink.LinkReference.Refresh();

                        var actionLinks = actionObjectsLink.GetLinkedComplexLinks().OfType<NomenclatureHierarchyLink>().ToList();

                        var macroContext = new UIMacroContext(new ReferenceObjectViewModel(Context.ReferenceObject));
                        macroContext.RunMacro("cd2247eb-47ee-4944-9369-e73ee2fb5e5e",
                             "CreateActionText",
                             actionLinks);
                    }
                }
            }
            else if (parameterGuid == ModificationUsingAreaReferenceObject.FieldKeys.IsAutoText && Context.ChangedParameter.GetBoolean() == true)
            {
                var usingArea = Context.ReferenceObject;

                var sourceLinks = new List<NomenclatureHierarchyLink>();
                var hierarchyLinkMatches = usingArea.GetObjects(ModificationUsingAreaReferenceObject.RelationKeys.HierarchyLinkMatches);
                foreach (var matches in hierarchyLinkMatches)
                {
                    var sourceLink = ModificationHelper.GetToOneComplexHierarchyLink(matches, Guids.UsingAreaSourceLink) as NomenclatureHierarchyLink;
                    sourceLinks.Add(sourceLink);
                }

                var macroContext = new UIMacroContext(new ReferenceObjectViewModel(Context.ReferenceObject));
                macroContext.RunMacro("cd2247eb-47ee-4944-9369-e73ee2fb5e5e",
                     "CreateUsingAreaText",
                     sourceLinks);
            }
        }

        public void ЗавершениеРедактированияСвойствОбъекта() => OnPropertyEndChanges();

        public void OnPropertyEndChanges()
        {
            var macro = Context[nameof(MacroKeys.MacroSource)] as CodeMacro;
            if (macro.DebugMode.Value == true)
            {
                Debugger.Launch();
                Debugger.Break();
            }

            if (Context is UIMacroContext uiContext)
            {
                var usingAreaVM = uiContext.ObjectViewModel as UsingAreaObjectViewModel;
                if (usingAreaVM is null)
                    return;

                var modification = usingAreaVM.Modification;

                var currentRevisionListObjects = modification.GetObjects(ModificationReferenceObject.RelationKeys.SourceRevisions);

                if (uiContext.CancellingChanges)
                {
                    // При отмене изменений восстанавливаем список исходных ревизий

                    if (usingAreaVM.ExistingSourceRevisions.IsNullOrEmpty())
                    {
                        foreach (var listObject in currentRevisionListObjects)
                        {
                            var linkedRevision = listObject.GetObject(ModificationHelper.Guids.LinkInitialRevision);
                            if (linkedRevision is null)
                                continue; // %%TODO

                            ModificationHelper.DeleteInitialRevisionListObject(usingAreaVM.ReferenceObject, linkedRevision);
                        }
                    }
                    else
                    {
                        foreach (var listObject in currentRevisionListObjects)
                        {
                            var linkedRevision = listObject.GetObject(ModificationHelper.Guids.LinkInitialRevision);
                            if (linkedRevision is null)
                                continue; // %%TODO

                            if (usingAreaVM.ExistingSourceRevisions.Contains(linkedRevision))
                                continue;

                            ModificationHelper.DeleteInitialRevisionListObject(usingAreaVM.ReferenceObject, linkedRevision);
                        }

                        foreach (var existingRevision in usingAreaVM.ExistingSourceRevisions)
                        {
                            var listObject = currentRevisionListObjects.FirstOrDefault(
                                obj => obj.GetObject(ModificationHelper.Guids.LinkInitialRevision)?.Id == existingRevision.Id);

                            if (listObject is null)
                                ModificationHelper.CreateInitialRevisionListObject(modification, existingRevision);
                        }
                    }
                }
                else
                {
                    foreach (var checkedRootObject in usingAreaVM.CheckedRootObjectsWithoutChildren)
                        ModificationHelper.CreateInitialRevisionListObject(modification, checkedRootObject.ReferenceObject);

                    foreach (var uncheckedObject in usingAreaVM.UncheckedRootObjectsWithoutChildren)
                        ModificationHelper.DeleteInitialRevisionListObject(usingAreaVM.ReferenceObject, uncheckedObject.ReferenceObject);

                    var designContext = (DesignContextObject)modification.GetObject(ModificationReferenceObject.RelationKeys.DesignContext);
                    var pdmObject = ModificationHelper.GetLinkedObject(modification, ModificationReferenceObject.RelationKeys.PDMObject, designContext);

                    foreach (var checkedObject in usingAreaVM.CheckedChildObjects)
                    {
                        var parent = checkedObject.Parent as InitialRevisionObjectViewModel;
                        ModificationHelper.CreateInitialRevisionListObject(modification, parent.ReferenceObject);

                        if (pdmObject is not null)
                        {
                            foreach (var initialLink in checkedObject.AllHierarchyLinks)
                                ModificationHelper.ReplaceRevisionLink(usingAreaVM.ReferenceObject, pdmObject, initialLink);
                        }
                    }

                    foreach (var uncheckedObject in usingAreaVM.UncheckedChildObjects)
                    {
                        foreach (var link in uncheckedObject.AllHierarchyLinks)
                            ModificationHelper.RemoveMatchingLinksObject(usingAreaVM.ReferenceObject, link);
                    }

                    // Удаляем объекты списка с пустыми ревизиями по связи
                    foreach (var listObject in currentRevisionListObjects)
                    {
                        var linkedRevision = listObject.GetObject(ModificationHelper.Guids.LinkInitialRevision);
                        if (linkedRevision is null)
                            listObject.Delete();
                    }
                }
            }
        }

        /// <summary>
        /// View model для выбора объектов ЭСИ
        /// </summary>
        public class SelectPDMObjectViewModel : SelectReferenceObjectFromTreeWithCheckboxes
        {
            private List<NomenclatureObjectViewModel> _rootObjects;
            bool _isExpandRootNodes;
            bool _isAutoCheck;

            public SelectPDMObjectViewModel(List<NomenclatureObjectViewModel> rootObjects, bool isExpandRootNodes, bool isAutoCheck, Reference reference)
                : base(reference, true)
            {
                _rootObjects = rootObjects;
                _isExpandRootNodes = isExpandRootNodes;
                _isAutoCheck = isAutoCheck;

                Caption = "Выбор документов для включения в извещение";
            }

            protected override LayoutViewModel CreateContentViewModel(CancellationToken cancellationToken)
                => new PDMObjectsTreeViewModel(_rootObjects, _isExpandRootNodes, _isAutoCheck,
                    this, new ReferenceDataProvider(Reference));

            protected override IEnumerable<IReferenceObjectViewModel> GetContentSelectedObjects()
            {
                var tree = ContentViewModel as ReferenceTreeViewModel;
                var checkedObjects = new List<ReferenceObjectViewModel>();
                GetAllCheckedObjects(tree.DataSource, checkedObjects);
                return checkedObjects;
            }

            private void GetAllCheckedObjects(ReferenceObjectUICollection dataSource, List<ReferenceObjectViewModel> checkedObjects)
            {
                if (dataSource is null)
                    return;

                foreach (var referenceObject in dataSource.OfType<ReferenceObjectViewModel>())
                {
                    if (referenceObject.IsChecked == true)
                        checkedObjects.Add(referenceObject);

                    if (referenceObject.HasChildren)
                        GetAllCheckedObjects(referenceObject.Children, checkedObjects);
                }
            }
        }

        public class PDMObjectsTreeViewModel : NomenclatureTreeViewModel
        {
            private List<NomenclatureObjectViewModel> _rootObjects;
            private readonly bool _isAutoCheck;

            public PDMObjectsTreeViewModel(List<NomenclatureObjectViewModel> rootObjects, bool isExpandRootNodes, bool isAutoCheck,
                LayoutViewModel owner, ReferenceDataProvider dataProvider)
                : base(owner, dataProvider)
            {
                _rootObjects = rootObjects;
                RootObject = null;
                ShowRootNode = false;
                ShowCheckboxes = true;
                MultipleSelection = true;
                PropertyPanelLocation = TFlex.DOCs.Client.ViewModels.ObjectProperties.Location.None;

                IsExpandRootNodes = isExpandRootNodes;
                _isAutoCheck = isAutoCheck;

                Appearance.AllowEditMultipeRowsCommand = false;
            }

            protected override bool ShowStandardTreeCommands => false;

            protected override ReferenceEntryViewModel CreateRootNode() => null;

            protected override void FillChildren(ReferenceEntryViewModel parentObject, ReferenceObjectUICollection collection, CancellationToken cancellationToken)
            {
                if (parentObject is null)
                {
                    if (_rootObjects.IsNullOrEmpty())
                    {
                        base.FillChildren(parentObject, collection, cancellationToken);
                    }
                    else
                    {
                        var viewModels = new List<NomenclatureObjectViewModel>();

                        foreach (var rootObject in _rootObjects)
                        {
                            var hierarchyLinks = rootObject.AllHierarchyLinks.ToList();
                            if (hierarchyLinks.IsNullOrEmpty())
                                viewModels.Add(new NomenclatureObjectViewModel(rootObject.NomenclatureObject) { IsChecked = true });
                            else
                            {
                                var notDeletedLinks = hierarchyLinks.Where(link => (link as NomenclatureHierarchyLink).DeletedInDesignContext == false).ToList();
                                if (notDeletedLinks.IsNullOrEmpty())
                                    return;

                                if (notDeletedLinks.CountMoreThan(1))
                                    viewModels.Add(new NomenclatureObjectViewModel(rootObject.NomenclatureObject, notDeletedLinks) { IsChecked = true });
                                else
                                    viewModels.Add(new NomenclatureObjectViewModel(rootObject.HierarchyLink, false) { IsChecked = true });
                            }
                        }

                        if (viewModels.Any())
                            collection.AddRange(viewModels);
                    }
                }
                else
                {
                    base.FillChildren(parentObject, collection, cancellationToken);

                    if (_isAutoCheck && !collection.IsNullOrEmpty())
                    {
                        int parentId = (parentObject as ReferenceObjectViewModel)?.ReferenceObject.Id ?? -1;
                        bool isRootParent = _rootObjects.FirstOrDefault(root => root.ReferenceObject.Id == parentId) != null;

                        if (isRootParent)
                        {
                            foreach (var child in collection.OfType<NomenclatureObjectViewModel>())
                            {
                                // Отмечаем только подключения, добавленные в контексте
                                var changedLink = child.AllHierarchyLinks?.OfType<NomenclatureHierarchyLink>()
                                      .FirstOrDefault(link => !link.DeletedInDesignContext && link.DesignContextId > 0 && link.SubstitutedLinkId == 0);
                                if (changedLink != null)
                                    child.IsChecked = true;
                            }
                        }
                    }
                }
            }

            protected override void OnGetCommands(CommandsList commands, CommandCollection parentCommand,
                TFlex.DOCs.Client.ViewModels.Commands.Toolbars.BarType barType, CancellationToken cancellationToken)
            {
                base.OnGetCommands(commands, parentCommand, barType, cancellationToken);

                commands.RemoveCommand<CheckOutOperationCommand>();
                commands.RemoveCommand<CheckInOperationCommand>();
                commands.RemoveCommand<UndoCheckOutOperationCommand>();
                commands.RemoveCommand<CommunicationCommandCollection>();

                commands.RemoveCommand<ReportCommandCollection>();
                commands.RemoveCommand<UserEventCollectionCommand>();
                commands.RemoveCommand<UserEventCommand>();
                commands.RemoveCommand<OperationCommandCollection>();

                if (parentCommand is null)
                {
                    commands.InsertFirst(GetCommand<CollapseChildrenCommand<ReferenceEntryViewModel, ReferenceObjectUICollection>>());
                    commands.InsertFirst(GetCommand<CollapseAllCommand<ReferenceEntryViewModel, ReferenceObjectUICollection>>());
                    commands.InsertFirst(GetCommand<ExpandChildrenCommand<ReferenceEntryViewModel, ReferenceObjectUICollection>>());
                    commands.InsertFirst(GetCommand<ExpandAllCommand<ReferenceEntryViewModel, ReferenceObjectUICollection>>());
                    commands.InsertFirst(GetCommand<UnCheckChildrenCommand<ReferenceEntryViewModel, ReferenceObjectUICollection>>());
                    commands.InsertFirst(GetCommand<UnCheckAllCommand<ReferenceEntryViewModel, ReferenceObjectUICollection>>());
                    commands.InsertFirst(GetCommand<CheckChildrenCommand<ReferenceEntryViewModel, ReferenceObjectUICollection>>());
                    commands.InsertFirst(GetCommand<CheckAllCommand<ReferenceEntryViewModel, ReferenceObjectUICollection>>());
                }
            }

            protected override void FillCreateObjectCommands(CommandsList commands, CancellationToken cancellationToken)
            {
                return;
            }
        }
    }
}

