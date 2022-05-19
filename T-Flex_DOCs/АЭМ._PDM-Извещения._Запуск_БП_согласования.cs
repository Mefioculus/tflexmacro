using System;
using System.Linq;
using System.Text;
using TFlex.DOCs.Client.ViewModels.Base;
using TFlex.DOCs.Common;
using TFlex.DOCs.Model.Desktop;
using TFlex.DOCs.Model.Macros;
using TFlex.DOCs.Model.Macros.ObjectModel;
using TFlex.DOCs.Model.Processes.Events.Contexts;
using TFlex.DOCs.Model.References;
using TFlex.DOCs.Model.References.Modifications;
using TFlex.DOCs.Model.References.Nomenclature.ModificationNotices;
using TFlex.DOCs.Model.References.Procedures;

namespace PDM_RunProcessModificationNotice
{
    public class Macro : MacroProvider
    {
        private static class Guids
        {
            public static readonly Guid МакросПроверкиИИ = new Guid("26c5d6ef-a79d-434a-9a9e-fff341463b81");
            public static readonly Guid МакросПроверкиТехнологическогоИИ = new Guid("73bc4b49-3761-4f06-9cc6-4eaba15be565");

            public static readonly Guid ПроцедураСогласованиеИИ = new Guid("67206bf8-c5e2-4b30-9493-8449cb30928b");
            public static readonly Guid ПроцедураСогласованиеИзвещения = new Guid("7a0816a0-47a6-4563-8627-e1dfff696fb3");//8a12b741-41e0-4d61-8a96-4c8f0156f78e
            public static readonly Guid ПроцедураСогласованиеТехнологическогоИИ = new Guid("e25dd453-7def-496a-ad00-237cdd2e0bfb");

            public static readonly Guid ТипИИ = new Guid("52ccb35c-67c5-4b82-af4f-e8ceac4e8d02");
            public static readonly Guid ТипИзменение = new Guid("f40ea698-bfaa-4143-9534-6276ddec0955");
            public static readonly Guid ТипТехнологическоеИзменение = new Guid("a8c61089-dc59-43c8-9296-f2629fec2c4e");

            public static readonly Guid СвязьИзвещенияИзменения = new Guid("5e46670a-400c-4e36-bb37-d4d651bdf692");

            public static readonly Guid СвязьДокумент = new Guid("d962545d-bccb-40b5-9986-257b57032f6e");
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
            var modificationNotice = Context
                .GetSelectedObjects()
                .FirstOrDefault(obj => obj.Reference.ParameterGroup.Guid == ModificationNoticesReference.ModificationNoticesReferenceGuid);

            if (modificationNotice.Class.IsInherit(Guids.ТипИИ))
            {
                var modificationStage = modificationNotice.SystemFields.Stage == null ? null : modificationNotice.SystemFields.Stage.Stage;
                if (modificationStage != null && modificationStage.Name == "Разработка")
                    return button;
            }

            button.Enable = false;
            button.Visible = false;
            return button;
        }

        public void СогласоватьИИ()
        {
            ReferenceObject modificationNotice = null;

            var referenceObject = Context.ReferenceObject;
            if (referenceObject.Reference.ParameterGroup.Guid == ModificationNoticesReference.ModificationNoticesReferenceGuid)
                modificationNotice = referenceObject;
            else if (referenceObject.Reference.ParameterGroup.Guid == ModificationReference.ReferenceId)
                modificationNotice = (referenceObject as ModificationReferenceObject).ModificationNotice;

            RunModificationNoticeEndorsement(modificationNotice);
        }

        private void RunModificationNoticeEndorsement(ReferenceObject modificationNotice)
        {
            if (modificationNotice is null)
                return;

            bool isPdmModification = false;
            bool useRevisions = false;

            RefObjList includedObjects = null;

            var modificationObjects = modificationNotice.GetObjects(Guids.СвязьИзвещенияИзменения);
            if (modificationObjects.All(referenceObject => referenceObject.Class.IsInherit(Guids.ТипИзменение)))
            {
                isPdmModification = true;

                useRevisions = modificationObjects.All(referenceObject =>
                  {
                      referenceObject.TryGetObject(Guids.СвязьДокумент, out var linkedDocument);
                      return linkedDocument is null;
                  });

                includedObjects = RunMacro(Guids.МакросПроверкиИИ.ToString(), "CheckModificationNotice", modificationNotice, useRevisions);
            }
            else if (modificationObjects.All(referenceObject => referenceObject.Class.IsInherit(Guids.ТипТехнологическоеИзменение)))
            {
                includedObjects = RunMacro(Guids.МакросПроверкиТехнологическогоИИ.ToString(), "CheckModificationNotice", modificationNotice);
            }

            if (includedObjects.IsNullOrEmpty())
                return;

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
                    String.Format("Объекты заблокированы, бизнес-процесс не будет запущен:{0}{0}{1}",
                    Environment.NewLine,
                    text.ToString()));
                return;
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
                    "Перед запуском бизнес-процесса согласования извещения необходимо применить изменения к следующим объектам:{0}{0}{1}{0}Продолжить?",
                    Environment.NewLine,
                    text.ToString())))
                    return;

                var checkedInObjects = Desktop.CheckIn(checkingObjects, String.Empty, true);
                if (checkedInObjects.IsNullOrEmpty())
                {
                    Message("Сообщение", "К объектам не были применены изменения: бизнес-процесс не будет запущен.");
                    return;
                }
            }

            RunProcess(new ReferenceObject[] { modificationNotice },
                isPdmModification
                ? (useRevisions ? Guids.ПроцедураСогласованиеИзвещения : Guids.ПроцедураСогласованиеИИ)
                : Guids.ПроцедураСогласованиеТехнологическогоИИ);
        }

        private void RunProcess(ReferenceObject[] objects, Guid procedureGuid)
        {
            var procedures = Context.Connection.ReferenceCatalog.Find(ProceduresReference.ReferenceGuid).CreateReference();
            var procedure = procedures.Find(procedureGuid) as BusinessProcedureReferenceObject;
            if (procedure is null)
                throw new NullReferenceException("procedure");

            CreateWorkflowContext context;

            if (Context is TFlex.DOCs.UI.Objects.Managers.UIMacroContext)
            {
                context = new TFlex.DOCs.UI.Objects.References.Procedures.UICreateWorkflowContext(procedure, objects)
                {
                    AlwaysShowDialog = procedure.IsShowWizard,
                };

                procedure.CreateProcess(context);
            }
            else if (Context is TFlex.DOCs.Client.ViewModels.UIMacroContext)
            {
                TFlex.DOCs.Client.ViewModels.UIMacroContext uiContext = (TFlex.DOCs.Client.ViewModels.UIMacroContext)Context;

                context = new TFlex.DOCs.Client.ViewModels.Processes.References.Procedures.UICreateWorkflowContext(procedure, objects)
                {
                    AlwaysShowDialog = procedure.IsShowWizard,
                };

                (uiContext.OwnerViewModel as ISupportProgressIndicator).Hide(() => procedure.CreateProcess(context));
            }
        }
    }
}
