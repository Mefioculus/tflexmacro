using System.Collections.Generic;
using TFlex.DOCs.Model.Macros;
using TFlex.DOCs.Model.References.Nomenclature;

namespace PDM
{
    public class PDM___Тексты_изменения : MacroProvider
    {
        // Guid макроса "PDM - Формирование текстов изменений"
        private static readonly string macroGuid = "0a7413f8-f208-4978-acf8-89ff3de93801";

        public PDM___Тексты_изменения(MacroContext context) : base(context)
        {
        }

        public void CreateActionText(List<NomenclatureHierarchyLink> checkedLinks)
        {
            if (DynamicMacro.ExecutionPlace == ExecutionPlace.WpfClient)
                RunMacro(macroGuid, "CreateActionTextCore", checkedLinks);
        }

        public void CreateActionTextForPdmObject(NomenclatureReferenceObject referenceObject,
            List<NomenclatureHierarchyLink> sourceLinks, List<NomenclatureHierarchyLink> newLinks)
        {
            if (DynamicMacro.ExecutionPlace == ExecutionPlace.WpfClient)
                RunMacro(macroGuid, "CreateActionTextForPdmObjectCore", referenceObject, sourceLinks, newLinks);
        }

        public void CreateUsingAreaText(List<NomenclatureHierarchyLink> checkedLinks)
        {
            if (DynamicMacro.ExecutionPlace == ExecutionPlace.WpfClient)
                RunMacro(macroGuid, "CreateUsingAreaTextCore", checkedLinks);
        }

        public void CreateModificationTexts()
        {
            if (DynamicMacro.ExecutionPlace == ExecutionPlace.WpfClient)
                RunMacro(macroGuid, "CreateModificationTextsCore");
        }
    }
}

