using System;
using System.Collections.Generic;
using TFlex.DOCs.Model.Macros;
using TFlex.DOCs.Model.References;
using TFlex.DOCs.Model.Search;
using System.Linq;
using System.Text;
using TFlex.DOCs.Common;

public class Macro : MacroProvider
{
    private List<ErrorWrapper> _errorsList;
    public Macro(MacroContext context)
        : base(context)
    {
        _errorsList = new List<ErrorWrapper>();
    }

    public override void Run()
    {
        if (Context.ReferenceObject == null)
            Error("Нельзя вызывать нумерацию узлов на корневом объекте справочника");

        RefreshReferenceWindow(); // требуется, если пользователь отсортировал данные в справочнике вручную, тогда данные модели и данные вью модели будут синхронизированы
        string numerationSettingsComboBoxName = "Укажите настройки нумерации";
        string integerFieldName = "Начать с";
        string rootPrefix = "Префикс для узла первого уровня";
        // прочитать настройки нумерации
        var numerationSettingsCollection = NumerationSettings.Load(Context).ToArray();

        int startValue = 1;
        var inputDialog = CreateInputDialog("Выбор настроек нумерации");
        if (numerationSettingsCollection.Length > 0)
        {
            inputDialog.AddSelectFromList(numerationSettingsComboBoxName, numerationSettingsCollection.First(), true,
                numerationSettingsCollection);
        }
        else
        {
            inputDialog.AddSelectFromList(numerationSettingsComboBoxName, String.Empty, true,
                numerationSettingsCollection);
        }

        inputDialog.AddInteger(integerFieldName, startValue, true);
        inputDialog.AddString(rootPrefix, String.Empty);
        if (!inputDialog.Show())
            return;

        if (!(inputDialog.GetValue(numerationSettingsComboBoxName) is NumerationSettings selectedNumerationSettings))
            return;

        startValue = inputDialog.GetValue(integerFieldName);
        string rootPrefixValue = inputDialog.GetValue(rootPrefix);
        SetNumbersToTree(Context.ReferenceObject, startValue, selectedNumerationSettings, rootPrefixValue);
        HandleErrors();
        RefreshReferenceWindow();
    }

    public ButtonValidator GetButtonValidator()
    {
        bool visible = Context.ReferenceObject != null;
        return new ButtonValidator()
        {
            Enable = true,
            Visible = visible
        };
    }

    public void ClearNumbers()
    {
        // прочитать настройки нумерации
        var numerationSettingsCollection = NumerationSettings.Load(Context).ToArray();

        if (numerationSettingsCollection.Length == 0)
            Error("Внимание!", "Нет настроек нумерации, чтобы очистить номера узлов");

        var selectedNumerationSettings = numerationSettingsCollection.First();
        var contextReferenceObject = Context.ReferenceObject;
        if (contextReferenceObject == null)
            Error("Отсутствует объект для очистки нумерации!");

        using var referenceObjectSaveSet = new ReferenceObjectSaveSet();
        var allChildren = contextReferenceObject.Children.RecursiveLoad();
        allChildren.Add(contextReferenceObject);
        foreach (var child in allChildren)
        {
            try
            {
                child.BeginChanges();
                child[selectedNumerationSettings.ParameterNameGuid].Value = String.Empty;
                referenceObjectSaveSet.Add(child);
            }
            catch (Exception e)
            {
                string message = $"Не удалось очистить номер у объекта {child} из-за внутренней ошибки";
                RegisterError(message, e, child);
            }
        }

        referenceObjectSaveSet.EndChanges();
        HandleErrors();
        RefreshReferenceWindow();
    }

    /// <summary>
    /// Устанавливает узлам дерева нумерацию в соответствии с настройками
    /// </summary>
    /// <param name="referenceObject">Начальный объект для простановки номеров</param>
    /// <param name="currentNumber">номер, с которого начинаем</param>
    /// <param name="numberingSettings">выбранные настройки нумерации</param>
    /// <param name="rootPrefixValue">префикс для узла</param>
    private void SetNumbersToTree(ReferenceObject referenceObject, int currentNumber,
        NumerationSettings numberingSettings, string rootPrefixValue)
    {
        bool startNumberAlreadySet = false;
        string startNumber = String.Empty;

        // если родитель используется в нумерации
        var firstParentReferenceObject = referenceObject.Parent;
        if (firstParentReferenceObject != null && CheckClassIsSameOrChildType(numberingSettings, firstParentReferenceObject))
        {
            string parentNumber = firstParentReferenceObject[numberingSettings.ParameterNameGuid].GetString();
            if (String.IsNullOrEmpty(parentNumber))
                throw new InvalidOperationException(
                    "Для правильной нумерации Вы должны присвоить номер родительскому объекту!");

            startNumber = parentNumber + "." + currentNumber;
        }

        // если указан префикс, то он приоритетней, чем родительский префикс
        if (!String.IsNullOrEmpty(rootPrefixValue))
            startNumber = rootPrefixValue.TrimEnd('.') + "." + currentNumber;

        //если номер до сих пор пуст, но мы должны его проставить, потому что запрашиваемый объект участвует в нумерации
        if (CheckClassIsSameOrChildType(numberingSettings, referenceObject))
        {
            if (String.IsNullOrEmpty(startNumber))
                startNumber = currentNumber.ToString();

            // если рассматриваемый элемент входит в список тех, которые требуют нумерации, то проставим ему
            SetNumberToObject(referenceObject, startNumber, numberingSettings);
            startNumberAlreadySet = true;
        }

        var stack = new Stack<StackParameters>();
        stack.Push(new StackParameters(startNumber, referenceObject.Children));
        bool firstChildren = true;
        while (stack.Count > 0)
        {
            StackParameters numberContainer = stack.Pop();
            int count = 1;
            if (firstChildren && !startNumberAlreadySet)
            {
                count = currentNumber;
                startNumberAlreadySet = true;
                firstChildren = false;
            }

            foreach (var childReferenceObject in numberContainer.ObjectCollection)
            {
                if (!CheckClassIsSameOrChildType(numberingSettings, childReferenceObject))
                    continue;

                string number = String.IsNullOrEmpty(numberContainer.Number)
                    ? count.ToString()
                    : $"{numberContainer.Number}.{count}";

                SetNumberToObject(childReferenceObject, number, numberingSettings);
                stack.Push(new StackParameters(number, childReferenceObject.Children));
                count++;
            }
        }
    }

    /// <summary>
    /// Задаем номер узлу в базе данных
    /// </summary>
    /// <param name="referenceObject">Узел</param>
    /// <param name="number">номер</param>
    /// <param name="numberingSettings">настройки нумерации</param>
    private void SetNumberToObject(ReferenceObject referenceObject, string number, NumerationSettings numberingSettings)
    {
        if (referenceObject is null)
            throw new ArgumentNullException(nameof(referenceObject));

        if (number is null)
            throw new ArgumentNullException(nameof(number));

        if (numberingSettings is null)
            throw new ArgumentNullException(nameof(numberingSettings));

        string numberToSet = numberingSettings.TypeSeparatorInTheEnd && !String.IsNullOrEmpty(number)
            ? number + "."
            : number;

        try
        {
            referenceObject.Modify(
                refObject => { refObject[numberingSettings.ParameterNameGuid].Value = numberToSet; });
        }
        catch (Exception e)
        {
            string message = $"Не удалось присвоить номер {numberToSet} объекту с содержимым \"{referenceObject}\"";
            RegisterError(message, e, referenceObject);
        }
    }

    private void RegisterError(string message, Exception ex, ReferenceObject referenceObject)
    {
        _errorsList.Add(new ErrorWrapper() {ReferenceObject = referenceObject, Exception = ex, Message = message});
    }

    private void HandleErrors()
    {
        if (_errorsList.Count == 0)
            return;

        var stringBuilder = new StringBuilder();
        foreach (var error in _errorsList)
        {
            stringBuilder.AppendLine(error.Message);
        }

        stringBuilder.AppendLine("Подробную информацию по ошибкам смотрите в логе приложения");
        string message = stringBuilder.ToString();
        var aggregationException = new AggregateException(_errorsList.Select(error => error.Exception));
        LogWriterLight.Instance.WriteToLogFile(aggregationException, message);
        Error(message);
    }

    private bool CheckClassIsSameOrChildType(NumerationSettings numberingSettings, ReferenceObject referenceObject)
    {
        return numberingSettings.ClassObjectCollectionGuid.Any(guid => referenceObject.Class.IsInherit(guid));
    }

    /// <summary>
    /// Вложенный класс для хранения параметров стека, а именно номера и дочерних элементов для прохода
    /// </summary>
    private class StackParameters
    {
        public StackParameters(string number, ReferenceObjectCollection collection)
        {
            Number = number ?? throw new ArgumentNullException(nameof(number));
            ObjectCollection = collection ?? throw new ArgumentNullException(nameof(collection));
        }

        /// <summary>
        /// Номер в иерархии
        /// </summary>
        public string Number { get; }

        /// <summary>
        /// Набор дочерних объектов
        /// </summary>
        public ReferenceObjectCollection ObjectCollection { get; }
    }
}

/// <summary>
/// Класс определяет настройки для нумераций
/// </summary>
public class NumerationSettings
{
    /// <summary>
    /// GUID справочника Типы нумерации
    /// </summary>
    private static readonly Guid NumerationSettingsCatalogGuid = new Guid("16175b87-d5d6-4a22-a67f-710ea07c49a6");

    /// <summary>
    /// GUID списка типов внутри справочника Типы нумерации
    /// </summary>
    private static readonly Guid ObjectListRelationGuid = new Guid("f2873c41-235e-45c6-8a92-4732e74953ad");

    /// <summary>
    /// GUID параметра Наименование
    /// </summary>
    private static readonly Guid NameParameterGuid = new Guid("825240b7-e8e1-4513-b17f-8b0bae4bd86e");

    /// <summary>
    /// GUID параметра Проставлять разделитель в конце номера
    /// </summary>
    private static readonly Guid PlaceSeparatorAfterNumberParameterGuid =
        new Guid("ab786788-8036-48cd-be2c-7c358257c53e");

    /// <summary>
    /// GUID параметра "Справочник".
    /// </summary>
    private static readonly Guid CatalogGuidParameter = new Guid("ab3fb803-2325-4e98-a12d-d21d678b57c0");

    /// <summary>
    /// GUID параметра, куда будет проставляться номер.
    /// </summary>
    private static readonly Guid NumberParameterGuid = new Guid("c9d39af0-915b-4e58-b69b-8fdb1c08a25b");

    /// <summary>
    /// GUID параметра Тип объекта у дочерних объектов нумерации
    /// </summary>
    private static readonly Guid ClassGuidParameter = new Guid("b4ac523a-8482-4c3b-8c81-b251b6c88bfd");

    private NumerationSettings(string name, Guid catalogGuid, Guid parameterNameGuid, bool typeEndSeparator)
    {
        if (String.IsNullOrEmpty(name))
            throw new ArgumentNullException(nameof(name), "Не задано наименование для настройки нумерации!");

        if (catalogGuid == Guid.Empty)
            throw new ArgumentException(nameof(catalogGuid));

        if (parameterNameGuid == Guid.Empty)
            throw new ArgumentException(nameof(parameterNameGuid));

        Name = name;
        CatalogGuid = catalogGuid;
        TypeSeparatorInTheEnd = typeEndSeparator;
        ParameterNameGuid = parameterNameGuid;
        ClassObjectCollectionGuid = new HashSet<Guid>();
    }

    /// <summary>
    /// Имя настроек нумерации
    /// </summary>
    public string Name { get; private set; }

    /// <summary>
    /// Строковое представление GUID параметра "Параметр для нумерации"
    /// </summary>
    public Guid ParameterNameGuid { get; set; }

    /// <summary>
    /// Справочник
    /// </summary>
    public Guid CatalogGuid { get; set; }

    /// <summary>
    /// Набор типов, которые будут использоваться именно в этой нумерации
    /// </summary>
    public HashSet<Guid> ClassObjectCollectionGuid { get; private set; }

    /// <summary>
    /// Ставить в конце разделитель
    /// </summary>
    public bool TypeSeparatorInTheEnd { get; private set; }

    /// <summary>
    /// Загрузить настройки нумерации из базы данных
    /// </summary> 
    public static ICollection<NumerationSettings> Load(MacroContext context)
    {
        var settingsCollection = new List<NumerationSettings>();
        var referenceInfo = context.Connection.ReferenceCatalog.Find(NumerationSettingsCatalogGuid);
        if (referenceInfo == null)
            throw new InvalidOperationException(
                $"Невозможно получить справочник Настройки нумерации по данному GUID {NumerationSettingsCatalogGuid}");

        var reference = referenceInfo.CreateReference();
        var filter = new Filter(referenceInfo);
        filter.Terms.AddTerm(reference.ParameterGroup[CatalogGuidParameter],
            ComparisonOperator.Equal, context.Reference.ParameterGroup.Guid);

        List<ReferenceObject> objectList = reference.Find(filter);

        // Обработка загруженных объектов      
        foreach (ReferenceObject referenceObject in objectList)
        {
            // читаем основные данные о нумерации
            string name = referenceObject[NameParameterGuid].GetString();
            bool typeEndSeparator = referenceObject[PlaceSeparatorAfterNumberParameterGuid].GetBoolean();
            Guid catalogGuid = referenceObject[CatalogGuidParameter].GetGuid();
            string parameterName = referenceObject[NumberParameterGuid].GetString();
            Guid parameterNameGuid = Guid.Empty;
            if (!String.IsNullOrEmpty(parameterName))
            {
                string[] lines = parameterName.Split(new char[] {'.'});
                parameterNameGuid = new Guid(lines[lines.Length - 1].Trim(new char[] {'[', ']'}));
            }

            var numerationSettings = new NumerationSettings(name, catalogGuid, parameterNameGuid, typeEndSeparator);

            // список типов внутри нумерации
            List<ReferenceObject> linkedObjects = referenceObject.GetObjects(ObjectListRelationGuid);
            foreach (var child in linkedObjects)
            {
                Guid typeGuid = child[ClassGuidParameter].GetGuid();
                if (typeGuid == Guid.Empty)
                {
                    string message = $"В дочернем объекте нумерации {name} не задан тип!";
                    throw new InvalidOperationException(message);
                }

                numerationSettings.ClassObjectCollectionGuid.Add(typeGuid);
            }

            settingsCollection.Add(numerationSettings);
        }

        return settingsCollection;
    }

    public override string ToString()
    {
        return Name;
    }
}

/// <summary>
/// Класс-обертка для ошибки, содержащий ссылку на объект справочника
/// </summary>
public class ErrorWrapper
{
    /// <summary>
    /// Объект справочника, где произошла ошибка
    /// </summary>
    public ReferenceObject ReferenceObject { get; set; }

    /// <summary>
    /// Ссылка на исключение
    /// </summary>
    public Exception Exception { get; set; }

    /// <summary>
    /// Сообщение для отчетности
    /// </summary>
    public string Message { get; set; }
}
