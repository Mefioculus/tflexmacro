using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using TFlex.DOCs.Client;
using TFlex.DOCs.Client.ViewModels;
using TFlex.DOCs.Client.ViewModels.Base;
using TFlex.DOCs.Client.ViewModels.ObjectProperties.References;
using TFlex.DOCs.Common;
using TFlex.DOCs.Model.Desktop;
using TFlex.DOCs.Model.Macros;
using TFlex.DOCs.Model.Macros.ObjectModel;
using TFlex.DOCs.Model.References;
using TFlex.DOCs.Model.References.Macros;
using TFlex.DOCs.Model.References.Modifications;
using TFlex.DOCs.Model.References.Nomenclature;
using TFlex.DOCs.Model.References.Nomenclature.ModificationNotices;
using TFlex.DOCs.Model.References.Procedures;
using TFlex.DOCs.Resources.Strings;

namespace PDM_RunProcessModificationNotice
{
    public class Macro : MacroProvider
    {
        private static class Guids
        {
            public static readonly Guid ModificationNoticeType = new Guid("52ccb35c-67c5-4b82-af4f-e8ceac4e8d02");
            public static readonly Guid ModificationNoticesSetType = new Guid("322d371b-c672-40ea-9166-316140fc28cb");
            public static readonly Guid ModificationNoticesLink = new Guid("740644db-40a2-44d8-8fe9-841391985d46");

            public static readonly Guid МакросПроверкиИИ = new Guid("26c5d6ef-a79d-434a-9a9e-fff341463b81");
            public static readonly Guid МакросПроверкиТехнологическогоИИ = new Guid("73bc4b49-3761-4f06-9cc6-4eaba15be565");

            public static readonly Guid ПроцедураСогласованиеИИ = new Guid("67206bf8-c5e2-4b30-9493-8449cb30928b");
            public static readonly Guid ПроцедураСогласованиеИзвещения = new Guid("8a12b741-41e0-4d61-8a96-4c8f0156f78e");
            public static readonly Guid ПроцедураСогласованиеИзвещенияВОсновномКонтексте = new Guid("3dc63b17-0e8a-4dfb-aa53-a07e0bf1f8ea");
            public static readonly Guid ПроцедураСогласованиеТехнологическогоИИ = new Guid("e25dd453-7def-496a-ad00-237cdd2e0bfb");

            public static readonly Guid TechnologicalModificationType = new Guid("a8c61089-dc59-43c8-9296-f2629fec2c4e");

            public static readonly Guid СвязьДокумент = new Guid("d962545d-bccb-40b5-9986-257b57032f6e");

            /// <summary> Стадия "Разработка" </summary>
            public static readonly Guid DevelopmentStage = new Guid("527f5234-4c94-43d1-a38d-d3d7fd5d15af");
        }

        public Macro(MacroContext context)
            : base(context)
        {
        }

        public override void Run()
        {
        }

        public ButtonValidator ValidateButton()
        {
            var button = new ButtonValidator();

            var selectedObjects = Context.GetSelectedObjects();
            if (selectedObjects.Length == 1)
            {
                var modificationNoticeObject = selectedObjects
                    .FirstOrDefault(obj => obj.Reference.ParameterGroup.Guid == ModificationNoticesReference.ModificationNoticesReferenceGuid);

                if (modificationNoticeObject != null)
                {
                    if (modificationNoticeObject.Class.IsInherit(Guids.ModificationNoticeType) ||
                        modificationNoticeObject.Class.IsInherit(Guids.ModificationNoticesSetType))
                    {
                        var modificationStage = modificationNoticeObject.SystemFields.Stage?.Stage;
                        if (modificationStage != null && modificationStage.Guid == Guids.DevelopmentStage)
                            return button;
                    }
                }
            }

            button.Enable = false;
            button.Visible = false;
            return button;
        }

        public void СогласоватьИИ()
        {
            var macro = Context[nameof(MacroKeys.MacroSource)] as CodeMacro;
            if (macro.DebugMode.Value == true)
            {
                System.Diagnostics.Debugger.Launch();
                System.Diagnostics.Debugger.Break();
            }

            ReferenceObject modificationNoticeObject = null;

            var referenceObject = Context.ReferenceObject;
            if (referenceObject.Reference.ParameterGroup.Guid == ModificationNoticesReference.ModificationNoticesReferenceGuid)
                modificationNoticeObject = referenceObject;
            else if (referenceObject.Reference.ParameterGroup.Guid == ModificationReference.ReferenceId)
                modificationNoticeObject = (referenceObject as ModificationReferenceObject).ModificationNotice;

            RunModificationNoticeEndorsement(modificationNoticeObject);
        }

        private void RunModificationNoticeEndorsement(ReferenceObject modificationNoticeObject)
        {
            if (modificationNoticeObject is null)
                return;

            bool isPdmModification = false;
            bool useRevisions = false;

            var modificationNotices = new List<ReferenceObject>();
            var modificationObjects = new List<ReferenceObject>();

            modificationNotices.Add(modificationNoticeObject);

            if (modificationNoticeObject.Class.IsInherit(Guids.ModificationNoticesSetType))
            {
                var linkedModificationNotices = modificationNoticeObject.GetObjects(Guids.ModificationNoticesLink);
                if (!linkedModificationNotices.IsNullOrEmpty())
                {
                    modificationNotices.AddRange(linkedModificationNotices);

                    foreach (var linkedModificationNotice in linkedModificationNotices)
                        modificationObjects.AddRange(linkedModificationNotice.GetObjects(ModificationReferenceObject.RelationKeys.ModificationNotice));
                }
            }
            else
                modificationObjects.AddRange(modificationNoticeObject.GetObjects(ModificationReferenceObject.RelationKeys.ModificationNotice));

            bool isMainContext = false;

            RefObjList checkingObjects = null;
            if (modificationObjects.All(referenceObject => referenceObject.Class.IsInherit(ModificationTypes.Keys.Modification)))
            {
                isPdmModification = true;

                useRevisions = modificationObjects.All(referenceObject =>
                  {
                      referenceObject.TryGetObject(Guids.СвязьДокумент, out var linkedDocument);
                      return linkedDocument is null;
                  });

                isMainContext = modificationObjects.OfType<ModificationReferenceObject>().All(obj => obj.DesignContext is null);

                checkingObjects = RunMacro(Guids.МакросПроверкиИИ.ToString(), "CheckModificationNotice", modificationNoticeObject, useRevisions);
            }
            else if (modificationObjects.All(referenceObject => referenceObject.Class.IsInherit(Guids.TechnologicalModificationType)))
            {
                checkingObjects = RunMacro(Guids.МакросПроверкиТехнологическогоИИ.ToString(), "CheckModificationNotice", modificationNoticeObject);
            }

            if (checkingObjects.IsNullOrEmpty())
                return;

            var checkingReferenceObjects = checkingObjects.To<ReferenceObject>();

            var lockedObjects = checkingReferenceObjects
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
                    String.Format("Объекты заблокированы, бизнес-процесс не будет запущен:{0}{0}{1}",
                    Environment.NewLine,
                    text.ToString()));
                return;
            }

            var checkedOutObjects = checkingReferenceObjects
                .Where(referenceObject => referenceObject.Reference.ParameterGroup.SupportsDesktop && referenceObject.IsCheckedOut)
                .ToList();
            if (checkedOutObjects.Any())
            {
                var text = new StringBuilder();

                foreach (var editingObject in checkedOutObjects)
                {
                    string str = String.Format("'{0}' [тип '{1}']",
                        editingObject.ToString(),
                        editingObject.Class);

                    text.AppendLine(str);
                }

                if (!Question(String.Format(
                    "Перед запуском бизнес-процесса согласования извещения необходимо применить изменения к следующим объектам:{0}{0}{1}{0}Продолжить?",
                    Environment.NewLine,
                    text.ToString())))
                    return;

                if (CanCheckInEditableObjects(checkedOutObjects))
                    ExecuteCheckInOperation(checkedOutObjects);
            }

            RunProcess(modificationNotices, isPdmModification, useRevisions, isMainContext);
        }

        private void ExecuteCheckInOperation(List<ReferenceObject> checkedOutObjects)
        {
            var checkedInObjects = Desktop.CheckIn(checkedOutObjects, String.Empty, true).OfType<ReferenceObject>().ToArray();
            if (checkedInObjects.IsNullOrEmpty())
                Error("К объектам не были применены изменения: бизнес-процесс не будет запущен.");
        }

        private bool CanCheckInEditableObjects(List<ReferenceObject> checkedOutObjects)
        {
            if (!(ApplicationManager.MainViewModel is MainViewModel mainViewModel))
                return true;

            var openedObjects = mainViewModel.OpenedMdiItems.OfType<ReferenceObjectMDIViewModel>().Select(o => o.Object).ToList();
            if (openedObjects.IsNullOrEmpty())
                return true;

            var found = openedObjects.Where(o =>
            {
                if (checkedOutObjects.Contains(o))
                    return true;

                if (o is NomenclatureObject nomenclatureObject &&
                    checkedOutObjects.Contains(nomenclatureObject.LinkedObject))
                    return true;

                return false;
            }).ToArray();

            if (found.Any())
            {
                var sb = new StringBuilder();
                foreach (var referenceObject in found)
                {
                    if (sb.Length > 0)
                        sb.Append(", ");

                    sb.Append(String.Concat("\"", referenceObject.ToString(), "\""));
                }

                Error(Messages.WrongDoOpearationOnEditingObjectWarning, sb);
            }

            return true;
        }

        private void RunProcess(List<ReferenceObject> objects, bool isPdmModification, bool useRevisions, bool isMainContext)
        {
            var procedureGuid =
                isPdmModification
                ? (useRevisions ? (isMainContext ? Guids.ПроцедураСогласованиеИзвещенияВОсновномКонтексте : Guids.ПроцедураСогласованиеИзвещения) : Guids.ПроцедураСогласованиеИИ)
                : Guids.ПроцедураСогласованиеТехнологическогоИИ;

            var procedures = Context.Connection.ReferenceCatalog.Find(ProceduresReference.ReferenceGuid).CreateReference();
            var procedure = procedures.Find(procedureGuid) as BusinessProcedureReferenceObject;
            if (procedure is null)
                throw new NullReferenceException("procedure");

            if (Context is UIMacroContext uiContext)
            {
                var context = new TFlex.DOCs.Client.ViewModels.Processes.References.Procedures.UICreateWorkflowContext(procedure, objects)
                {
                    AlwaysShowDialog = procedure.IsShowWizard,
                };

                (uiContext.OwnerViewModel as ISupportProgressIndicator).Hide(() => procedure.CreateProcess(context));
            }
        }
    }
}
