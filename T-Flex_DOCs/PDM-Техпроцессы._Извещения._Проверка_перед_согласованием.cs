using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using TFlex.DOCs.Common;
using TFlex.DOCs.Model.Macros;
using TFlex.DOCs.Model.References;
using TFlex.Model.Technology.References.SetOfDocuments;
using TFlex.Technology;

namespace TechnologicalPDM_CheckModificationNotice
{
    public class Macro : MacroProvider
    {
        private static class Guids
        {
            public static readonly Guid ТипТехнологическоеИзменение = new Guid("a8c61089-dc59-43c8-9296-f2629fec2c4e");

            public static readonly Guid СвязьИзвещенияИзменения = new Guid("5e46670a-400c-4e36-bb37-d4d651bdf692");
            public static readonly Guid СвязьИзмененияАктуальныйВариантТЭ = new Guid("254c4753-4b42-454e-84cc-f5abc82b2448");
            public static readonly Guid СвязьИзмененияИсходныйВариантТЭ = new Guid("87dab3c7-c8f5-40a4-91dd-da6734ee1f3b");
            public static readonly Guid СвязьИзмененияЦелевойВариантТЭ = new Guid("737f68ad-9038-4585-b944-428662256f18");

            public static readonly Guid ТПДокументация = new Guid("cc38caed-f747-45ce-9fbf-771566841796");
        }

        public Macro(MacroContext context)
            : base(context)
        {
        }

        public override void Run()
        {
        }

        public List<ReferenceObject> CheckModificationNotice(ReferenceObject modificationNotice)
        {
            var checkingObjects = new List<ReferenceObject>();
            StringBuilder errorsText = new StringBuilder();

            if (modificationNotice == null)
            {
                errorsText.AppendLine("Извещение не найдено.");
                ShowErrors(errorsText.ToString());
                return null;
            }

            checkingObjects.Add(modificationNotice);

            var modifications = modificationNotice
                .GetObjects(Guids.СвязьИзвещенияИзменения)
                .Where(referenceObject => referenceObject.Class.IsInherit(Guids.ТипТехнологическоеИзменение));
            if (modifications.IsNullOrEmpty())
            {
                errorsText.AppendLine(String.Format("Извещение '{0}' не содержит изменений.", modificationNotice));
                ShowErrors(errorsText.ToString());
                return null;
            }

            checkingObjects.AddRange(modifications);

            foreach (var modification in modifications)
            {
                var actualTP = modification.GetObject(Guids.СвязьИзмененияАктуальныйВариантТЭ);
                if (actualTP == null || !actualTP.Class.IsInherit(Technology2012Classes.StructuredTechnologicalProcessType))
                    errorsText.AppendLine(String.Format("У изменения '{0}' отсутствует актуальный вариант технологического процесса.", modification));
                else
                    checkingObjects.Add(actualTP);

                var targetTP = modification.GetObject(Guids.СвязьИзмененияЦелевойВариантТЭ);
                if (targetTP == null || !targetTP.Class.IsInherit(Technology2012Classes.StructuredTechnologicalProcessType))
                {
                    errorsText.AppendLine(String.Format("У изменения '{0}' отсутствует целевой вариант технологического процесса.", modification));
                    ShowErrors(errorsText.ToString());
                    return null;
                }

                checkingObjects.Add(targetTP);

                ReferenceObject setOfDocuments = null;
                var targetDocuments = targetTP.GetObjects(Guids.ТПДокументация);
                if (targetDocuments != null)
                    setOfDocuments = targetDocuments.FirstOrDefault(document => document.Class.IsInherit(SetOfDocumentsTypes.Keys.SetOfDocuments));
                if (setOfDocuments == null)
                {
                    errorsText.AppendLine(String.Format("Целевой вариант технологического процесса '{0}' не содержит комплект документов.", targetTP));
                    ShowErrors(errorsText.ToString());
                    return null;
                }

                checkingObjects.Add(setOfDocuments);
            }

            if (String.IsNullOrEmpty(errorsText.ToString()))
                return checkingObjects;

            ShowErrors(errorsText.ToString());
            return null;
        }

        private void ShowErrors(string errorsText)
        {
            Message("Результат проверки извещения перед согласованием",
                  "{0}{0}{1}{0}",
                  Environment.NewLine,
                  errorsText);
        }
    }
}
