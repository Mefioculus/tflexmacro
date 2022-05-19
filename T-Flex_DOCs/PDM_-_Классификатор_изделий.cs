using System;
using System.Collections.Generic;
using System.Linq;
using TFlex.DOCs.Common;
using TFlex.DOCs.Model.Macros;
using TFlex.DOCs.Model.References;
using TFlex.DOCs.Model.References.Nomenclature;

public class Macro : MacroProvider
{
    public Macro(MacroContext context)
        : base(context)
    {
    }

    public override void Run()
    {
    }

    public void ИзменениеСвязиОпций()
    {
        var isUIContext = Context is TFlex.DOCs.Client.ViewModels.IUIMacroContext || Context is TFlex.DOCs.UI.Objects.Managers.IUIMacroContext;
        if (isUIContext)
        {
            var linkChangedArgs = Context.ModelChangedArgs as ObjectLinkChangedEventArgsBase;

            if (linkChangedArgs.Link.LinkGroup.Guid == ProductsClassifierOptionsReferenceObject.RelationKeys.OptionsValuesLink
            && linkChangedArgs.RemovedObject != null)
            {
                var currentOption = Context.ReferenceObject as ProductsClassifierOptionsReferenceObject;
                var removingValue = linkChangedArgs.RemovedObject as ProductOptionValuesReferenceObject;
                var currentProduct = currentOption.MasterObject as ProductsClassifierReferenceObject;

                var changingChildren = new List<ProductsClassifierReferenceObject>();
                GetChangingChildrenProducts(currentOption, removingValue, currentProduct, changingChildren);
                if (changingChildren.Any())
                {
                    string message = String.Format($"Изменение затронет структуру следующих изделий: {String.Join(", ", changingChildren)}. Продолжить?");
                    if (!Context.ShowQuestion(message))
                        Cancel();
                }
            }
        }
    }

    private void GetChangingChildrenProducts(ProductsClassifierOptionsReferenceObject changingOption,
        ProductOptionValuesReferenceObject removingValue,
        ProductsClassifierReferenceObject parentProduct,
        List<ProductsClassifierReferenceObject> changingChildren)
    {
        foreach (var childProduct in parentProduct.Children.OfType<ProductsClassifierReferenceObject>())
        {
            var childOption = childProduct.Options.FirstOrDefault(opt => opt.Option == changingOption.Option);
            if (childOption != null)
            {
                var existingOptionValues = childOption.OptionValues.AsList.ToList();
                if (existingOptionValues.Any(currentValue => currentValue == removingValue))
                {
                    changingChildren.Add(childProduct);
                    GetChangingChildrenProducts(changingOption, removingValue, childProduct, changingChildren);
                }
            }
        }
    }

    public void ВыполнитьДействияОпций()
    {
        var currentObject = (ReferenceObject)CurrentObject;
        if (currentObject is ProductsClassifierOptionsReferenceObject currentOption)
        {
            var currentProduct = (ProductsClassifierReferenceObject)currentOption.MasterObject;
            if (!currentProduct.SpecificStructure)
                return;

            var allProductActions = GetAllProductOptionActions(currentProduct);
            var existingProductOptions = currentProduct.Options.ToList();
            DoOptionActions(currentOption, currentProduct, allProductActions, existingProductOptions);
        }
        else if (currentObject is ProductsClassifierReferenceObject currentProduct)
        {
            if (Context.ChangedParameter.ParameterInfo.Guid == ProductsClassifierReferenceObject.FieldKeys.SpecificStructure
                && currentProduct.SpecificStructure)
            {
                var allProductActions = GetAllProductOptionActions(currentProduct);
                var existingProductOptions = currentProduct.Options.ToList();
                foreach (var option in existingProductOptions)
                    DoOptionActions(option, currentProduct, allProductActions, existingProductOptions);
            }
        }
    }

    private void DoOptionActions(
        ProductsClassifierOptionsReferenceObject currentOption,
        ProductsClassifierReferenceObject currentProduct,
        List<ProductsClassifierActionReferenceObject> allProductActions,
        List<ProductsClassifierOptionsReferenceObject> existingProductOptions)
    {
        var currentLinkedOption = currentOption.Option;
        var currentOptionValues = currentOption.OptionValues.AsList.ToList();
        if (currentLinkedOption is null || currentOptionValues.IsNullOrEmpty() || currentOptionValues.CountMoreThan(1))
            return;

        var currentOptionValue = currentOptionValues[0];
        var currentOptionActions = allProductActions.Where(action
             =>
             action.SelectedOption != null &&
             action.SelectedOption.Id == currentLinkedOption.Id &&
             action.SelectedOptionValue != null &&
             action.SelectedOptionValue.Id == currentOptionValue.Id).ToList();

        if (!currentOptionActions.Any())
            return;

        var changedOptions = new Dictionary<int, ProductsClassifierOptionsReferenceObject>();
        foreach (var action in currentOptionActions)
        {
            var actionChangingOption = action.ChangingOption;
            if (actionChangingOption is null)
                continue;

            int actionChangingOptionId = actionChangingOption.Id;
            var changingProductOption = existingProductOptions.FirstOrDefault(op => op.Option != null && op.Option.Id == actionChangingOptionId);
            if (changingProductOption is null)
                continue;

            if (action.OptionActionType == OptionActionType.SetValue)
            {
                var newOptionValues = action.ChangingOptionValues.AsList.ToList();
                if (newOptionValues.IsNullOrEmpty() || newOptionValues.CountMoreThan(1))
                    continue;

                var newValue = newOptionValues[0];
                SetOptionValueAction(changingProductOption, newValue);
            }
            else if (action.OptionActionType == OptionActionType.AllowEdit || action.OptionActionType == OptionActionType.DisableEdit)
            {
                SetEditableModeOptionAction(changingProductOption, action.OptionActionType == OptionActionType.DisableEdit);
            }
            else if (action.OptionActionType == OptionActionType.Hide || action.OptionActionType == OptionActionType.Show)
            {
                var hidingOptionValues = action.ChangingOptionValues.AsList.ToList();
                HideOptionAction(changingProductOption, hidingOptionValues, action.AllValues, action.OptionActionType == OptionActionType.Hide);
            }

            if (changingProductOption.Changing)
            {
                if (!changedOptions.ContainsKey(actionChangingOptionId))
                    changedOptions.Add(actionChangingOptionId, changingProductOption);
            }
        }

        foreach (var changedProductOption in changedOptions)
        {
            try
            {
                changedProductOption.Value.EndChanges();
            }
            catch
            {
                if (changedProductOption.Value.Changing)
                    changedProductOption.Value.CancelChanges();
            }
        }
    }

    private void SetOptionValueAction(ProductsClassifierOptionsReferenceObject changingProductOption, ProductOptionValuesReferenceObject newValue)
    {
        var existingOptionValues = changingProductOption.OptionValues.AsList.ToList();

        foreach (var existingValue in existingOptionValues)
        {
            if (existingValue.Id == newValue.Id)
                continue;

            if (!changingProductOption.Changing)
                changingProductOption.BeginChanges();

            changingProductOption.RemoveOptionValue(existingValue);
        }

        if (changingProductOption.OptionValues.AsList.IsNullOrEmpty())
        {
            if (!changingProductOption.Changing)
                changingProductOption.BeginChanges();

            changingProductOption.AddOptionValue(newValue);
        }
    }

    private void SetEditableModeOptionAction(ProductsClassifierOptionsReferenceObject changingProductOption, bool settingValue)
    {
        bool currentValue = changingProductOption.IsNotEditable;
        if (currentValue == settingValue)
            return;

        if (!changingProductOption.Changing)
            changingProductOption.BeginChanges();
        changingProductOption.IsNotEditable.Value = settingValue;
    }

    private void HideOptionAction(ProductsClassifierOptionsReferenceObject changingProductOption, List<ProductOptionValuesReferenceObject> hidingOptionValues, bool hideOption, bool hide)
    {
        if (hideOption || hidingOptionValues.IsNullOrEmpty())
        {
            bool currentValue = changingProductOption.IsHidden;
            if (currentValue == hide)
                return;

            if (!changingProductOption.Changing)
                changingProductOption.BeginChanges();
            changingProductOption.IsHidden.Value = hide;
        }
        else
        {
            if (!hide)
                return;

            var currentOptionValue = changingProductOption.OptionValues.AsList.FirstOrDefault();
            if (currentOptionValue is null)
                return;

            var currentOptionValueId = currentOptionValue.Id;

            if (!hidingOptionValues.Select(newValue => newValue.Id).Contains(currentOptionValueId))
                return;

            var availableValueIds = GetVisibleOptionValueIdList(changingProductOption);
            int newValueId = availableValueIds.FirstOrDefault(id => id != currentOptionValueId);

            var optionValuesReference = new ProductOptionValuesReference(Context.Connection);
            var newOptionValue = optionValuesReference.Find(newValueId) as ProductOptionValuesReferenceObject;
            if (newOptionValue is null)
                return;

            if (!changingProductOption.Changing)
                changingProductOption.BeginChanges();
            changingProductOption.RemoveOptionValue(currentOptionValue);
            changingProductOption.AddOptionValue(newOptionValue);
        }
    }

    public List<int> ПолучитьСписокИдентификаторовСкрытыхОпций()
    {
        return GetHiddenOptionIdList();
    }

    private List<int> GetHiddenOptionIdList()
    {
        var existingOptionIdList = new List<int>(0);

        var product = (ProductsClassifierReferenceObject)CurrentObject.MasterObject;
        var existingOptions = product.Options;

        foreach (var existingOption in existingOptions)
        {
            var existingLinkedOption = existingOption.Option;
            if (existingLinkedOption is null)
                continue;

            int existingLinkedOptionId = existingLinkedOption.SystemFields.Id;
            existingOptionIdList.Add(existingLinkedOptionId);
        }

        return existingOptionIdList;
    }

    public List<int> ПолучитьСписокИдентификаторовВидимыхЗначенийОпций()
    {
        return GetVisibleOptionValueIdList((ProductsClassifierOptionsReferenceObject)CurrentObject);
    }

    private List<int> GetVisibleOptionValueIdList(ProductsClassifierOptionsReferenceObject currentOption)
    {
        var product = (ProductsClassifierReferenceObject)currentOption.MasterObject;
        if (product is null)
            return new List<int>(0);

        var currentLinkedOption = currentOption.Option;
        if (currentLinkedOption is null)
            return new List<int>(0);

        ProductsClassifierOptionsReferenceObject parentOption = null;

        if (product.Parent != null)
        {
            var parentOptionsList = product.Parent.Options;
            parentOption = parentOptionsList.FirstOrDefault(option => option.Option != null && option.Option.Id == currentLinkedOption.Id);
        }

        List<int> allPossibleOptionValueIds =
            parentOption is null
            ? currentLinkedOption.PossibleOptionValues.Select(optionValue => optionValue.Id).ToList<int>()
            : parentOption.OptionValues.AsList.Select(optionValue => optionValue.Id).ToList<int>();

        if (!product.SpecificStructure)
            return allPossibleOptionValueIds;

        var visibleByActionValueIdList = new List<int>(0);
        var hiddenByActionValueIdList = new List<int>(0);

        var existingProductOptions = product.Options;

        var allActions = GetAllProductOptionActions(product);

        foreach (var action in allActions.Where(action
            => (action.OptionActionType == OptionActionType.Hide || action.OptionActionType == OptionActionType.Show) &&
            action.SelectedOption != null &&
            !action.ChangingOptionValues.IsNullOrEmpty()))
        {
            var matchingProductOption = existingProductOptions.FirstOrDefault(option => option.Option != null && option.Option.Id == action.SelectedOption.Id);
            if (matchingProductOption is null)
                continue;

            var changingOptionValueIds = action.ChangingOptionValues.AsList.Select(optionValue => optionValue.Id);

            if (action.SelectedOptionValue is null)
            {
                if (action.OptionActionType == OptionActionType.Show)
                    visibleByActionValueIdList.AddRange(changingOptionValueIds);
                else
                    hiddenByActionValueIdList.AddRange(changingOptionValueIds);
            }
            else
            {
                if (matchingProductOption.OptionValues.AsList.Select(optionValue => optionValue.Id).Contains(action.SelectedOptionValue.Id))
                {
                    if (action.OptionActionType == OptionActionType.Show)
                        visibleByActionValueIdList.AddRange(changingOptionValueIds);
                    else
                        hiddenByActionValueIdList.AddRange(changingOptionValueIds);
                }
            }
        }

        foreach (int visibleId in visibleByActionValueIdList)
        {
            if (hiddenByActionValueIdList.Contains(visibleId))
                hiddenByActionValueIdList.Remove(visibleId);
        }

        foreach (int hiddenId in hiddenByActionValueIdList.Distinct())
        {
            if (allPossibleOptionValueIds.Contains(hiddenId))
                allPossibleOptionValueIds.Remove(hiddenId);
        }

        return allPossibleOptionValueIds;
    }

    public object КолонкаПолучитьРезультатДействия()
    {
        return CurrentObject?.ToString();
    }

    public void ЗавершениеИзмененияПараметра()
    {
        Guid parameterGuid = Context.ChangedParameter.ParameterInfo.Guid;
        if (parameterGuid == ProductsClassifierActionReferenceObject.FieldKeys.ActionType)
        {
            var action = (ProductsClassifierActionReferenceObject)CurrentObject;
            switch (action.OptionActionType)
            {
                case OptionActionType.AllowEdit:
                case OptionActionType.DisableEdit:
                    action.AllValues.Value = true;
                    break;
                case OptionActionType.Hide:
                case OptionActionType.Show:
                    break;
                case OptionActionType.SetValue:
                    action.AllValues.Value = false;
                    break;
            }
        }
        else if (parameterGuid == ProductsClassifierActionReferenceObject.FieldKeys.AllValues)
        {
            var action = (ProductsClassifierActionReferenceObject)CurrentObject;
            var optionValues = action.ChangingOptionValues.AsList.ToList();
            foreach (var optionValue in optionValues)
                action.RemoveChangingOptionValue(optionValue);
        }
    }

    public List<int> ПолучитьСписокДоступныхУправляющихОпцийДляДействия()
    {
        var action = (ProductsClassifierActionReferenceObject)CurrentObject;
        return GetOptionsForActionExcept(action, action.ChangingOption);
    }

    public List<int> ПолучитьСписокДоступныхИзменяемыхОпцийДляДействия()
    {
        var action = (ProductsClassifierActionReferenceObject)CurrentObject;
        return GetOptionsForActionExcept(action, action.SelectedOption);
    }

    private List<int> GetOptionsForActionExcept(ProductsClassifierActionReferenceObject action, ProductOptionsReferenceObject exceptingOption)
    {
        var product = (ProductsClassifierReferenceObject)action.MasterObject;
        var existingOptions = product.Options;
        return existingOptions
            .Select(option => option.Option?.Id ?? -1)
            .Except(new List<int>() { exceptingOption?.Id ?? -1 }).ToList();
    }

    public List<int> ПолучитьСписокДоступныхЗначенийДляУправляющейОпции()
    {
        var action = (ProductsClassifierActionReferenceObject)CurrentObject;
        return GetValuesForActionOption(action, action.SelectedOption);
    }

    public List<int> ПолучитьСписокДоступныхЗначенийДляИзменяемойОпции()
    {
        var action = (ProductsClassifierActionReferenceObject)CurrentObject;
        return GetValuesForActionOption(action, action.ChangingOption);
    }

    private List<int> GetValuesForActionOption(ProductsClassifierActionReferenceObject action, ProductOptionsReferenceObject actionOption)
    {
        if (actionOption is null)
            return new List<int>() { -1 };

        var product = (ProductsClassifierReferenceObject)action.MasterObject;
        var productOption = product.Options.FirstOrDefault(option => option.Option != null && option.Option.Id == actionOption.Id);
        return productOption?.OptionValues.AsList.Select(optionValue => optionValue.Id).ToList()
            ?? new List<int>() { -1 };
    }

    private List<ProductsClassifierActionReferenceObject> GetAllProductOptionActions(ProductsClassifierReferenceObject product)
    {
        if (!product.SpecificStructure)
            return new List<ProductsClassifierActionReferenceObject>(0);

        var optionActions = product.OptionActions.ToList();

        var parentProduct = product.Parent as ProductsClassifierReferenceObject;
        while (parentProduct != null)
        {
            optionActions.AddRange(parentProduct.OptionActions);
            parentProduct = parentProduct.Parent;
        }

        return optionActions;
    }
}
