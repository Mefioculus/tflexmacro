using System;
using System.Collections.Generic;
using System.Linq;
using TFlex.DOCs.Model.Macros;
using TFlex.DOCs.Model.References;
using TFlex.DOCs.Model.Search;
using TFlex.DOCs.Model.References.Codifier;
using TFlex.DOCs.Model.Search.Path;
using System.Diagnostics;

public class MacroGenerateNewNumber : MacroProvider
{
    public MacroGenerateNewNumber(MacroContext context)
        : base(context)
    {
    }

    public override void Run()
    {
        var contextReferenceObject = Context.ReferenceObject;

        var settings = CodifierSettings.Load(Context);
        if (settings.Count == 0)
        {
            Сообщение("Предупреждение!", "Для текущего справочника и типа нет настроек в справочнике Кодификатор");
            return;
        }

        var currentSettings = settings.First();
        var autoNumeratorReferenceObject = currentSettings.ReferenceObject as CodifierReferenceObject;
        if (autoNumeratorReferenceObject is null)
        {
            Сообщение("Предупреждение!",
                "Не удалось привести найденный объект к типу объекта справочника Кодификатор");
            return;
        }

        string number = String.Empty;
        int delayTime = 1000;
        bool success = false;
        int attemptsCount = 3;

        IGetNextNumber nextNumberGetter = contextReferenceObject.Reference.ParameterGroup.SupportsRevisions
            ? new ReferenceWithRevisionsSupportNextNumberGetter(currentSettings, Context)
            : new CommonReferenceNextNumberGetter();

        for (int i = 0; i < attemptsCount; i++)
        {
            try
            {
                number = nextNumberGetter.GetNextNumber(autoNumeratorReferenceObject, contextReferenceObject);
                success = true;
                break;
            }
            catch
            {
                // ловим в течение некоторого времени все подряд
            }

            System.Threading.Thread.Sleep(delayTime);
            delayTime *= 2; // увеличиваем время ожидания в 2 раза
        }

        // последняя попытка. Если что не так, это распространение исключения
        if (!success)
        {
            number = autoNumeratorReferenceObject.GetNextNumber(contextReferenceObject);
            success = !String.IsNullOrEmpty(number);
        }

        if (!success)
            return;

        contextReferenceObject.Modify(
            refObject => { refObject[currentSettings.ParameterNameGuid].Value = number; });

        if (!contextReferenceObject.IsNew)
            RefreshReferenceWindow();
    }
}

/// <summary>
/// Класс определяет настройки для получения нового номера из справочника Кодификатор
/// </summary>
public class CodifierSettings
{
    /// <summary>
    /// Создает экземпляр настроек автонумерации
    /// </summary>
    private CodifierSettings(Guid parameterNameGuid, ReferenceObject referenceObject)
    {
        if (parameterNameGuid == Guid.Empty)
            throw new ArgumentException(nameof(parameterNameGuid));

        if (referenceObject == null)
            throw new ArgumentNullException(nameof(referenceObject));

        ParameterNameGuid = parameterNameGuid;
        ReferenceObject = referenceObject;
    }

    /// <summary>
    /// GUID параметра "Параметр для нумерации"
    /// </summary>
    public Guid ParameterNameGuid { get; set; }

    /// <summary>
    /// Объект справочника, который умеет получать номер для указанного справочника и типа
    /// </summary>
    public ReferenceObject ReferenceObject { get; set; }

    /// <summary>
    /// Подобрать объект кодификатора из базы данных
    /// </summary> 
    public static ICollection<CodifierSettings> Load(MacroContext context)
    {
        var settingsCollection = new List<CodifierSettings>();
        var referenceInfo = context.Connection.ReferenceCatalog.Find(CodifierReference.ReferenceId);
        if (referenceInfo is null)
            throw new InvalidOperationException(
                $"Невозможно получить справочник Кодификатор по данному GUID {CodifierReference.ReferenceId}");

        var reference = referenceInfo.CreateReference();
        string str = $"[Справочник] = '{context.Reference.ParameterGroup.Guid}' И [Типы объектов справочника].[Идентификатор типа объекта] = '{context.ReferenceObject.Class.Guid}'";
        var filter = Filter.Parse(str, reference.ParameterGroup);
        var referenceObjectList = reference.Find(filter);

        if (referenceObjectList.Count == 0)
            return settingsCollection;
        
        // Найдены наборы правил для формирования кода. Теперь требуется найти один подходящий набор по контексту текущего объекта
        // Контекстом называется любой предок текущего объекта - это может быть папка или любой другой родительский объект
        ReferenceObject parent = context.ReferenceObject.Parent;
        ReferenceObject codifierReferenceObjectWithContext = null;
        ReferenceObject codifierReferenceObjectWithoutContext = null;
        ReferenceObject codifierReferenceObjectToUse = null;
        // есть коллекция возможных кодификаторов, требуется найти подходящий
        // рассматриваются все родители объекта, и если находится такой, что он является контекстом, то присваивается контекстный кодификатор,
        // иначе пользуемся по умолчанию
        while (parent is not null)
        {
            foreach (ReferenceObject codifierObjectFromList in referenceObjectList)
            {
                string path = codifierObjectFromList[CodifierReferenceObject.FieldKeys.Context].GetString();
                if (String.IsNullOrEmpty(path))
                {
                    codifierReferenceObjectWithoutContext = codifierObjectFromList;
                    continue;
                }
                
                Guid referenceObjectGuid = new(path.Split(new char[]{ '^' }, StringSplitOptions.RemoveEmptyEntries)[1]);
                if (referenceObjectGuid != parent.Guid)
                    continue;

                codifierReferenceObjectWithContext = codifierObjectFromList;
                break;
            }

            if (codifierReferenceObjectWithContext is not null)
                break;

            parent = parent.Parent;
        }

        if (codifierReferenceObjectWithoutContext is null)
        {
            foreach (ReferenceObject codifierObjectFromList in referenceObjectList)
            {
                string path = codifierObjectFromList[CodifierReferenceObject.FieldKeys.Context].GetString();
                if (String.IsNullOrEmpty(path))
                {
                    codifierReferenceObjectWithoutContext = codifierObjectFromList;
                    break;
                }
            }
        }

        codifierReferenceObjectToUse = codifierReferenceObjectWithContext ?? codifierReferenceObjectWithoutContext;
        if (codifierReferenceObjectToUse is null)
            return settingsCollection;

        string parameterName = codifierReferenceObjectToUse[CodifierReferenceObject.FieldKeys.NumberParameter].GetString();
        if (String.IsNullOrEmpty(parameterName)) 
            return settingsCollection;

        string[] lines = parameterName.Split(new char[] {'.'});
        var parameterNameGuid = new Guid(lines[lines.Length - 1].Trim(new char[] {'[', ']'}));
        var numerationSettings = new CodifierSettings(parameterNameGuid, codifierReferenceObjectToUse);
        settingsCollection.Add(numerationSettings);

        return settingsCollection;
    }
}

/// <summary>
/// Интерфейс с методом, который умеет получать код
/// </summary>
public interface IGetNextNumber
{
    /// <summary>
    /// Получить следующий идентификатор кода
    /// </summary>
    /// <param name="autoNumeratorReferenceObject"></param>
    /// <param name="contextReferenceObject"></param>
    /// <returns></returns>
    string GetNextNumber(CodifierReferenceObject autoNumeratorReferenceObject,
        ReferenceObject contextReferenceObject);
}

/// <summary>
/// Класс для получения кода справочниками, самыми обычными, без каких-либо поддержек ревизий, стадий и т.д.
/// </summary>
public class CommonReferenceNextNumberGetter : IGetNextNumber
{
    public virtual string GetNextNumber(CodifierReferenceObject codifierReferenceObject,
        ReferenceObject contextReferenceObject)
    {
        return codifierReferenceObject.GetNextNumber(contextReferenceObject);
    }
}

/// <summary>
/// Класс для получения кода для объекта справочника с поддержкой ревизий
/// </summary>
public class ReferenceWithRevisionsSupportNextNumberGetter : CommonReferenceNextNumberGetter
{
    private readonly CodifierSettings _settings;
    private readonly MacroContext _context;

    public ReferenceWithRevisionsSupportNextNumberGetter(CodifierSettings settings, MacroContext context)
    {
        _settings = settings;
        _context = context;
    }

    public override string GetNextNumber(CodifierReferenceObject codifierReferenceObject,
        ReferenceObject contextReferenceObject)
    {
        bool forRevision = (bool)_context["ForRevision"];

        return forRevision 
            ? contextReferenceObject.Prototype[_settings.ParameterNameGuid].GetString() 
            : base.GetNextNumber(codifierReferenceObject, contextReferenceObject);
    }
}
