/*
TFlex.DOCs.UI.Types.dll
TFlex.DOCs.UI.Client.dll

*/
using System.Linq;
using TFlex.DOCs.Client.ViewModels.Base;
using TFlex.DOCs.Client.ViewModels.Layout;
using TFlex.DOCs.Model.Macros;
using TFlex.DOCs.Model.Macros.ObjectModel;
using TFlex.DOCs.Model.Macros.ObjectModel.Layout;

public class Macro : MacroProvider
{
    public Macro(MacroContext context)
        : base(context)
    {
        /*
        if (Вопрос("Хотите запустить в режиме отладки?"))
        {
            System.Diagnostics.Debugger.Launch();
            System.Diagnostics.Debugger.Break();
        }
        */
    }

    public override void Run()
    {
    }

    public Объект GetObjNewClient(string itemName) //Возвращает один выделенный объект, где itemName - имя ЭУ с выдеденным объектом.
    {
        var supportSelectionVM = GetSupportSelection(itemName);
        if (supportSelectionVM == null)
            return null;
        var focusedObject = supportSelectionVM.FocusedObject;
        var obj = focusedObject as TFlex.DOCs.Client.ViewModels.References.ReferenceObjectViewModel;
        if (obj == null)
            return null;
        return Объект.CreateInstance(obj.ReferenceObject, Context);
    }

    public Объекты GetObjsNewClient(string itemName) // Возвращает выделенные объекты, где itemName - имя ЭУ с выделенными объектами
    {
        var supportSelectionVM = GetSupportSelection(itemName);
        if (supportSelectionVM == null)
            return null;
        var focusedObjects = supportSelectionVM.SelectedObjects;

        var objs = focusedObjects.Select(ob => (ob as TFlex.DOCs.Client.ViewModels.References.ReferenceObjectViewModel));
        if (objs == null)
            return null;
        return Объекты.CreateInstance(objs.Select(ob => ob.ReferenceObject), Context);
    }

    private ISupportSelection GetSupportSelection(string itemName)
    {
        IWindow currentWindow = Context.GetCurrentWindow();
        ILayoutItem foundItem = currentWindow.FindItem(itemName);

        if (foundItem == null)
            return null;

        //окно справочника
        ReferenceWindowLayoutItemViewModel referenceWindowLayoutItemVM = foundItem as ReferenceWindowLayoutItemViewModel;
        if (referenceWindowLayoutItemVM != null)
        {
            return referenceWindowLayoutItemVM.InnerViewModel as ISupportSelection;
        }
        return null;
    }
}
