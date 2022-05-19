/*
TFlex.DOCs.ui.Client.dll
DevExpress.Xpf.Grid.v19.2.dll
*/

using System;
using System.Linq;
using TFlex.DOCs.Client.ViewModels.Base;
using TFlex.DOCs.Client.ViewModels.Layout;
using TFlex.DOCs.Client.ViewModels.References;
using TFlex.DOCs.Model.Macros;
using TFlex.DOCs.Model.Macros.ObjectModel;


public class MdmTakeObjectsMacro : MacroProvider
{
    public MdmTakeObjectsMacro(MacroContext context)
        : base(context)
    {
    }

    public override void Run()
    {
    }

    /// <summary>
    /// Получает значение ячейки из окна справочника, структуры
    /// </summary>
    /// <param name="itemName">Имя элемента управления (окна справочника, структуры)</param>
    /// <returns>Значение ячейки</returns>
    public string ПолучитьЯчейку(string itemName)
    {
        var supportSelectionVM = GetSupportSelection(itemName);
        return supportSelectionVM?.CurrentColumn.GetValue(supportSelectionVM.FocusedObject)?.ToString();
    }

    /// <summary>
    /// Получает наименование колонки из окна справочника, структуры
    /// </summary>
    /// <param name="itemName">Имя элемента управления (окна справочника, структуры)</param>
    /// <returns>Наименование колонки</returns>
    public string ПолучитьКолонку(string itemName)
    {
        var supportSelectionVM = GetSupportSelection(itemName);
        return supportSelectionVM?.CurrentColumn.Caption;
    }

    /// <summary>
    /// Получает выбранные объекты из окна справочника, структуры
    /// </summary>
    /// <param name="itemName">Имя элемента управления (окна справочника, структуры)</param>
    /// <returns>Выбранные объекты</returns>
    public Объекты ПолучитьВыбраныеОбъекты(string itemName)
    {
        var supportSelectionVM = GetSupportSelection(itemName);
        if (supportSelectionVM == null)
            return new Объекты();

        var selectedObjets = supportSelectionVM.SelectedObjects;
        var objs = selectedObjets.OfType<ReferenceObjectViewModel>().Select(roVM => roVM.ReferenceObject).ToList();
        return new Объекты(objs, Context);
    }

    /// <summary>
    /// Получает выбранный объект из окна справочника, структуры
    /// </summary>
    /// <param name="itemName">Имя элемента управления (окна справочника, структуры)</param>
    /// <returns>Выбранный объект</returns>
    public Объект ПолучитьВыбраныйОбъект(string itemName)
    {
        var supportSelectionVM = GetSupportSelection(itemName);
        if (supportSelectionVM == null)
            return null;

        return supportSelectionVM.FocusedObject is ReferenceObjectViewModel obj
            ? Объект.CreateInstance(obj.ReferenceObject, Context)
            : null;
    }

    /// <summary>
    /// Получает Guid справочника из окна справочника, структуры
    /// </summary>
    /// <param name="itemName">Имя элемента управления (окна справочника, структуры)</param>
    /// <returns>Guid справочника</returns>
    public Guid ПолучитьGuidСправочника(string itemName)
    {
        var currentWindow = Context.GetCurrentWindow();
        var foundItem = currentWindow?.FindItem(itemName);

        if (!(foundItem is ReferenceWindowLayoutItemViewModel referenceWindowLayoutItemVM))
            return Guid.Empty;

        switch (referenceWindowLayoutItemVM.InnerViewModel)
        {
            case ReferenceExplorerViewModel explorerViewModel:
                return explorerViewModel.Reference.ParameterGroup.Guid;
            case ReferenceGridViewModel gridViewModel:
                return ((ISupportReference) gridViewModel).Reference.ParameterGroup.Guid;
            default:
                return Guid.Empty;
        }
    }

    private ISupportSelection GetSupportSelection(string itemName)
    {
        var currentWindow = Context.GetCurrentWindow();
        var foundItem = currentWindow?.FindItem(itemName);

        if (foundItem == null)
            return null;

        switch (foundItem)
        {
            // Окно справочника
            case ReferenceWindowLayoutItemViewModel referenceWindowLayoutItemVM:
                return referenceWindowLayoutItemVM.InnerViewModel as ISupportSelection;
            // Окно структуры объекта
            case ObjectStructureLayoutItemViewModel objectStructureLayoutItemViewModel:
                return objectStructureLayoutItemViewModel.InnerViewModel;
            // Диалог ввода, связанные объекты
            case LinkToManyLayoutItemViewModel linkToManyLayoutItemViewModel:
                return linkToManyLayoutItemViewModel.LinkContent as ISupportSelection;
            default:
                return null;
        }
    }
}

