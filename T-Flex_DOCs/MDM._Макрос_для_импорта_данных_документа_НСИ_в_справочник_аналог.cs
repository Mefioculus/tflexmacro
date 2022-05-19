using System;
using System.Linq;
using TFlex.DOCs.Client.ViewModels;
using TFlex.DOCs.Client.ViewModels.Layout;
using TFlex.DOCs.Client.ViewModels.References;
using TFlex.DOCs.Client.ViewModels.WorkingPages;
using TFlex.DOCs.Macro.NormativeReferenceInfo.Handlers.DataCleanupSettings;
using TFlex.DOCs.Model.Macros;
using TFlex.DOCs.Model.References;


namespace Macroses
{
    /// <summary>
    /// (MDM и НСИ) Макрос для импорта данных документа НСИ в справочник аналог
    /// </summary>
    public class MdmNsiImportDocNsiFromAnalogReferenceMacro : MacroProvider
    {
        public MdmNsiImportDocNsiFromAnalogReferenceMacro(MacroContext context)
            : base(context)
        {
        }

        public override void Run()
        {
            //System.Diagnostics.Debugger.Launch();
            //System.Diagnostics.Debugger.Break();

            var uiContext = Context as UIMacroContext;

            if (uiContext == null || !(uiContext.OwnerViewModel is WorkingPageViewModel))
                return;

            var wp = (WorkingPageViewModel) uiContext.OwnerViewModel;
            var referenceWindowVm = wp.Content.FindFirst("item13") as ReferenceWindowLayoutItemViewModel;

            var selecteReferenceObjects = GetSelecteReferenceObjects(referenceWindowVm);

            if (selecteReferenceObjects.Length == 0)
                throw new MacroException("Не выбраны объекты в справочнике");

            new ImportDocumentNsiDataFromAnalogReference(selecteReferenceObjects, this).RunImport();
        }

        private static ReferenceObject[] GetSelecteReferenceObjects(
            ReferenceWindowLayoutItemViewModel referenceWindowVm)
        {
            if (referenceWindowVm == null)
                throw new MacroException("Элемент управления item13 не является элементом управления справочника");

            if (referenceWindowVm.InnerViewModel == null)
                throw new MacroException("Справочник не выбран");

            var explorer = referenceWindowVm.InnerViewModel as ReferenceExplorerViewModel;

            if (explorer != null)
            {
                if (explorer.GridViewModel != null)
                {
                    return explorer.GridViewModel.SelectedObjects.OfType<ReferenceObjectViewModel>()
                        .ToArray()
                        .Select(o => o.ReferenceObject)
                        .ToArray();
                }

                if (explorer.TreeViewModel != null)
                {
                    return explorer.TreeViewModel.SelectedObjects.OfType<ReferenceObjectViewModel>()
                        .ToArray()
                        .Select(o => o.ReferenceObject)
                        .ToArray();
                }
            }
            else if (referenceWindowVm.InnerViewModel is ReferenceGridViewModel)
            {
                var grid = (ReferenceGridViewModel) referenceWindowVm.InnerViewModel;

                if (grid.SelectedObjects != null)
                {
                    return grid.SelectedObjects.OfType<ReferenceObjectViewModel>()
                        .ToArray()
                        .Select(o => o.ReferenceObject)
                        .ToArray();
                }
            }

            return Array.Empty<ReferenceObject>();
        }
    }
}

