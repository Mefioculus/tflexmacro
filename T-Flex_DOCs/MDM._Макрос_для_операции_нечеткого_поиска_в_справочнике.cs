/* Дополнительные ссылки
PresentationFramework.dll
TFlex.DOCs.UI.Client.dll
TFlex.DOCs.UI.Utils.dll
DevExpress.Mvvm.v20.1.dll
WindowsBase.dll
*/

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Windows;
using TFlex.DOCs.Client;
using TFlex.DOCs.Client.ViewModels;
using TFlex.DOCs.Client.ViewModels.Columns;
using TFlex.DOCs.Client.ViewModels.References;
using TFlex.DOCs.Client.ViewModels.References.SelectionDialogs;
using TFlex.DOCs.Macro.NormativeReferenceInfo.Handlers.FuzzySearch;
using TFlex.DOCs.Model;
using TFlex.DOCs.Model.Macros;
using TFlex.DOCs.Model.Macros.ObjectModel;
using TFlex.DOCs.Model.References;
using TFlex.DOCs.Model.Search;
using TFlex.DOCs.Model.Structure;
using TFlex.DOCs.UI.Utils.Helpers;


/// <summary>
/// Макрос для операции нечеткого поиска в справочнике 'Эталоны СтИ' (MDM и НСИ)
/// </summary>
public class MdmNsiFuzzySearchMacro : MacroProvider
{
    private const string SearchStringName = "Строка поиска:";
    private const string SearchParametersName = "Параметр справочника:";
    private const string SearchAlgorithName = "Алгоритм поиска:";
    private const string CoefficientSimilarity = "Коэффициент соответствия(%):";
    private const string ResultCountName = "Количество результатов:";
    private const string ProcessedStage = "6e3af1a3-e9b7-4ab6-b045-ecaa6e763001";

    // Список алгоритмов для выбора в интерфейсе
    private static readonly Dictionary<string, FuzzySearchType> FuzzySearchTypes =
        new Dictionary<string, FuzzySearchType>
        {
            {"Расстояние Левенштейна", FuzzySearchType.Levenstein},
            {"Сходство Джаро — Винклера", FuzzySearchType.JaroWinkler},
            {"N - грамм", FuzzySearchType.QGramsDistance},
            {"Алгоритм Смита — Ватермана — Гото", FuzzySearchType.SmithWatermanGotoh}
        };

    public MdmNsiFuzzySearchMacro(MacroContext context)
        : base(context)
    {
    }

    public void ЗапуститьПоискИзАРМ()
    {
        string значениеЯчейки = ВыполнитьМакрос("ea0d8d3c-395b-48d5-9d42-dab98371522c", "ПолучитьЯчейку", "Записи");
        Объекты выбранныеОбъекты = ВыполнитьМакрос("ea0d8d3c-395b-48d5-9d42-dab98371522c", "ПолучитьВыбраныеОбъекты", "Записи"); //АРМ.Взять объект ячейку
        if (выбранныеОбъекты.Count == 0)
            return;

        var гуидВыбранногоСправочника = выбранныеОбъекты[0].Справочник.Guid;

        var гуидСправочникаЭталона = ПолучитьСправочникЭталонДляСправочника(гуидВыбранногоСправочника, out var документНСИ);
        if (гуидСправочникаЭталона == Guid.Empty)
            Ошибка("У параметра документа НСИ не указан справочник эталона");

        var referenceInfo = Context.Connection.ReferenceCatalog.Find(гуидСправочникаЭталона);
        if (referenceInfo == null)
            throw new MacroException("Справочник эталона не найден в каталоге справочников");

        var etalonReference = referenceInfo.CreateReference();

        var inputDialog = СоздатьДиалогВвода("Поиск");
        if (!ShowDialog(inputDialog, etalonReference, значениеЯчейки))
            return;

        string searchString = inputDialog[SearchStringName];
        string searchTypeName = inputDialog[SearchAlgorithName];
        int resultCount = inputDialog[ResultCountName];
        string searchParameterName = inputDialog[SearchParametersName];
        int сoefficientSimilarity = inputDialog[CoefficientSimilarity];

        var result = FuzzySearchReferenceObjects.Find(
            searchString,
            new[] { searchParameterName },
            resultCount,
            etalonReference,
            null,
            FuzzySearchTypes[searchTypeName],
            (double)сoefficientSimilarity / 100);

        var filter = new Filter(etalonReference.ParameterGroup);

        var objectId = etalonReference.ParameterGroup.Parameters.Find(SystemParameterType.ObjectId);
        int[] objectIds = result.Select(s => s.Object.SystemFields.Id).ToArray();

        filter.Terms.AddTerm(objectId, ComparisonOperator.IsOneOf, objectIds);

        var etalonObject = GetEtalon(result, referenceInfo, filter) ?? throw new MacroException("Эталон не найден");

        var эталон = Объект.CreateInstance(etalonObject, Context);
        var документНСИЭталона = эталон.СвязанныйОбъект["Документ НСИ"];
        if (документНСИЭталона == null)
            Ошибка("У эталона не указан документ НСИ");

        if (документНСИ["ID"] != документНСИЭталона["ID"])
            Ошибка("Внимание! Операция не может быть завершена: текущий Документ НСИ отличается от документа эталона.");

        var errors = new StringBuilder();
        foreach (var объект in выбранныеОбъекты)
        {
            try
            {
                объект.Изменить();
                объект.СвязанныйОбъект["Эталон"] = эталон;
                объект.Сохранить();
                if (!объект.ИзменитьСтадию(ProcessedStage))
                    errors.AppendLine($"Объект: {объект} невозможно изменить стадию");
            }
            catch (Exception e)
            {
                errors.AppendLine($"Объект: {объект} ошибка при изменении{Environment.NewLine}{e}");
            }
        }

        if (errors.Length == 0)
            Сообщение("Выполнено", "Эталоны были подключены");
        else
            Сообщение("Предупреждение", $"Во время изменения объектов возникли ошибки{Environment.NewLine}{errors}");
    }

    private Guid ПолучитьСправочникЭталонДляСправочника(Guid гуидВыбранногоСправочника, out Объект документНСИ)
    {
        документНСИ = НайтиОбъект("Документы НСИ", "79d316f0-0f13-4ee9-9316-e41f27389333", гуидВыбранногоСправочника.ToString());
        if (документНСИ == null)
            throw new MacroException("Не найден документ НСИ: на текущий справочник");

        var параметрДокументаНСИ = документНСИ.СвязанныйОбъект["9677778a-6a67-42ac-9608-19df7af56946"];
        if (параметрДокументаНСИ == null)
            throw new MacroException("У документа НСИ не указан параметр");

        return параметрДокументаНСИ["ae1c005a-92b0-42f0-8cb5-67aa9ce83856"];
    }

    public override void Run()
    {
        if (Context.Reference == null)
            throw new MacroException("Не возможно начать поиск без справочника");

        if (!(Context is UIMacroContext uiContext))
            return;

        var inputDialog = СоздатьДиалогВвода("Поиск");

        if (!ShowDialog(inputDialog, Context.Reference))
            return;

        string searchString = inputDialog[SearchStringName];
        string searchTypeName = inputDialog[SearchAlgorithName];
        int resultCount = inputDialog[ResultCountName];
        string searchParameterName = inputDialog[SearchParametersName];
        int сoefficientSimilarity = inputDialog[CoefficientSimilarity];

        var result = FuzzySearchReferenceObjects.Find(
            searchString,
            new[] { searchParameterName },
            resultCount,
            Context.Reference,
            null,
            FuzzySearchTypes[searchTypeName],
            (double)сoefficientSimilarity / 100);

        FillReference(uiContext, result);
    }

    public void DropFilter()
    {
        if (!(Context is UIMacroContext uiContext))
            return;

        var gridViewModel = (ReferenceGridViewModel)uiContext.OwnerViewModel;
        gridViewModel.Filter = null;
        ОбновитьОкноСправочника();
    }

    private static bool ShowDialog(ДиалогВвода inputDialog, Reference reference, string baseValue = "")
    {
        string defaultParameterName = reference.ParameterGroup.DefaultVisibleParameter.Name;
        string[] searchParametersNames = reference.ParameterGroup.Parameters
            .Where(p => !p.IsSystem && !p.IsExtended)
            .Select(p => p.Name)
            .ToArray();

        inputDialog.ДобавитьСтроковое(SearchStringName, baseValue, обязательное: true, использоватьВсюШирину: true);
        inputDialog.ДобавитьВыборИзСписка(SearchParametersName, defaultParameterName, true,
            searchParametersNames.Select(s => (object)s).ToArray());
        inputDialog.ДобавитьВыборИзСписка(SearchAlgorithName, "Расстояние Левенштейна", true,
            FuzzySearchTypes.Keys.Select(s => (object)s).ToArray());
        inputDialog.ДобавитьЦелое(CoefficientSimilarity, 85, true);
        inputDialog.ДобавитьЦелое(ResultCountName, 10, true);

        return inputDialog.Показать();
    }

    private void FillReference(UIMacroContext uiContext, List<FuzzySearchItem> referenceObjects)
    {
        var gridViewModel = (ReferenceGridViewModel)uiContext.OwnerViewModel;

        var filter = new Filter(Context.Reference.ParameterGroup);

        var objectId = Context.Reference.ParameterGroup.Parameters.Find(SystemParameterType.ObjectId);
        int[] objectIds = referenceObjects.Select(s => s.Object.SystemFields.Id).ToArray();

        filter.Terms.AddTerm(objectId, ComparisonOperator.IsOneOf, objectIds);
        gridViewModel.Filter = filter;
        gridViewModel.CustomColumnsGenerator = (collection, token) =>
        {
            var column = collection.FirstOrDefault(c => c.Name == "CoefficientSimilarity1");

            if (column != null)
                collection.Remove(column);

            var columnViewModel = new CustomColumn<ReferenceObjectViewModel, int>(
                "Коэффициент сходства",
                "CoefficientSimilarity1", o =>
                {
                    var item = referenceObjects.FirstOrDefault(i => i.Object.Guid == o.Guid);
                    return item != null ? (int)(item.CoefficientSimilarity * 100) : 0;
                }, null)
            {
                Visible = true,
                VisibleIndex = 100
            };

            collection.Add(columnViewModel);
        };

        DispatcherHelper.CheckBeginInvokeOnUI(() => gridViewModel.ReloadData(true));

        //ОбновитьОкноСправочника();
    }

    private ReferenceObject GetEtalon(List<FuzzySearchItem> items, ReferenceInfo etalonReference, Filter filter)
    {
        var uiContext = (UIMacroContext)Context;

        var dialog = new MySelectReferenceObjectViewModel(items, etalonReference, false, uiContext.OwnerViewModel)
        {
            ContextFilter = filter,
            DataLoadingMacroContextObject = Context?.ReferenceObject
        };

        using (dialog)
        {
            if (ApplicationManager.OpenDialog(dialog, uiContext.OwnerViewModel))
                return dialog.GetSelectedReferenceObjects().FirstOrDefault();
        }

        return null;
    }

    private class CustomColumn<TViewModel, TParameter> : DelegateColumnViewModel<TViewModel, TParameter>
    {
        public CustomColumn(string caption,
            string name,
            Func<TViewModel, TParameter> getter,
            Action<TViewModel, TParameter> setter)
            : base(caption, name, getter, setter)
        {
            ShowInColumnChooser = false;
            HorizontalContentAlignment = HorizontalAlignment.Left;
        }

        protected override bool CanSaveToSettings => false;
    }

    private class MySelectReferenceObjectViewModel : SelectReferenceObjectViewModel
    {
        private readonly Action<ColumnsCollection, CancellationToken> _customColumnsGenerator;

        public MySelectReferenceObjectViewModel(List<FuzzySearchItem> referenceObjects,
            ReferenceInfo referenceInfo,
            bool multipleSelection = false,
            LayoutViewModel owner = null)
            : base(referenceInfo, multipleSelection, owner)
        {
            _customColumnsGenerator = (collection, token) =>
            {
                var column = collection.FirstOrDefault(c => c.Name == "CoefficientSimilarity1");

                if (column != null)
                    collection.Remove(column);

                var columnViewModel = new CustomColumn<ReferenceObjectViewModel, int>(
                    "Коэффициент сходства",
                    "CoefficientSimilarity1", o =>
                    {
                        var item = referenceObjects.FirstOrDefault(i => i.Object.Guid == o.Guid);
                        return item != null ? (int)(item.CoefficientSimilarity * 100) : 0;
                    }, null)
                {
                    Visible = true,
                    VisibleIndex = 100
                };

                collection.Add(columnViewModel);
            };
        }

        protected override LayoutViewModel CreateContentViewModel(CancellationToken cancellationToken)
        {
            var viewModell = base.CreateContentViewModel(cancellationToken);

            switch (viewModell)
            {
                case ReferenceGridViewModel gridViewModel:
                    gridViewModel.CustomColumnsGenerator = _customColumnsGenerator;
                    break;
                case ReferenceTreeViewModel treeViewModel:
                    treeViewModel.CustomColumnsGenerator = _customColumnsGenerator;
                    break;
                case ReferenceExplorerViewModel explorerViewModel:
                    explorerViewModel.PropertyChanged += (sender, args) =>
                    {
                        switch (args.PropertyName)
                        {
                            case nameof(ReferenceExplorerViewModel.GridViewModel):
                                if (explorerViewModel.GridViewModel is ReferenceGridViewModel grid)
                                    grid.CustomColumnsGenerator = _customColumnsGenerator;
                                break;
                            case nameof(ReferenceExplorerViewModel.TreeViewModel):
                                if (explorerViewModel.TreeViewModel is ReferenceTreeViewModel tree)
                                    tree.CustomColumnsGenerator = _customColumnsGenerator;
                                break;
                        }
                    };
                    break;
            }

            return viewModell;
        }
    }
}
