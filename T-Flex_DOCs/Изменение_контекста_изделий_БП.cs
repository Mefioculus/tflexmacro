/*
Ссылки
TFlex.DOCs.Model.Processes.dll
Tflex.DOCs.Common.dll
*/

using System;
using System.Linq;
using System.Collections.Generic;
using TFlex.DOCs.Model.Macros;
using TFlex.DOCs.Model.Macros.Processes;
using TFlex.DOCs.Model.Processes.Events.Contexts;
using TFlex.DOCs.Model.References.Nomenclature;

using TFlex.DOCs.Common.Extensions;
using TFlex.DOCs.Model.References;

public class MacroChangeMainContextBP : ProcessActionMacroProvider
{
    /// <summary>
    /// Разделитель идентификаторов подключений
    /// </summary>
    private static readonly char _separatorLinkList = ';';

    /// <summary>
    /// Разделитель для записи идентификатора объект - подключение
    /// </summary>
    private static readonly char _separatorLinkId = '-';

    /// <summary>
    /// Режим для тестирования
    /// Если стоит true то при любой ошибке БП, будет генерировать исключения
    /// </summary>
    private const bool _testMode = true;

    private NomenclatureReference _nomenclatureReference;

    public MacroChangeMainContextBP(EventContext context)
        : base(context)
    {
    }

    /// <summary>
    /// режимы отладки, после тестирования можно удалить
    /// 
    /// </summary>
    private void TryRunDebug()
    {
        if (Context.Connection.ClientView.HostName == "MOSIN" && Вопрос("Хотите запустить в режиме отладки?"))
            RunDebug();
    }

    private void RunDebug()
    {
        System.Diagnostics.Debugger.Launch();
        System.Diagnostics.Debugger.Break();
    }

    public override void Run()
    {

    }

    public void ИзменитьКонтексПроектированияБП()
    {
        TryRunDebug();

        //RunDebug();

        //Получаем строку с идентификаторами подключений из переменной БП
        var hierarhyLinksStringValue = Переменные["Список подключений"];

        //Если строка пуста, то генерируем ошибку
        if (String.IsNullOrEmpty(hierarhyLinksStringValue))
            throw new MacroException("Переменная 'Список подключений' не содержит данных или пустая");

        //Получаем контекст проектирования, проверяем, что он действительно является контекстом проектирования
        var контекстПроектирования = Переменные["Контекст проектирования"];
        if (контекстПроектирования is null)
            throw new MacroException("Переменная 'Контекст проектирования' пуста");

        var structContext = (ReferenceObject)контекстПроектирования;

        if (structContext is not DesignContextObject designContextObject)
            throw new MacroException("Переданый объект не является контекстом проектирования");

        //Создаем справочник
        _nomenclatureReference = new NomenclatureReference(Context.Connection);

        //Очищаем все настройки для поиска всех объектов
        using (_nomenclatureReference.ClearAndHoldUseConfigurationSettings())
        {
            //Назначем контекст проектирования
            _nomenclatureReference.ConfigurationSettings.DesignContext = designContextObject;

            //Находим все подключения
            var hierarhyLinks = ПолучитьПодключенияИзСтроки(hierarhyLinksStringValue);
            if (hierarhyLinks.Count == 0)
                throw new MacroException("Не удалось найти подключения");

            foreach (var hierarchyLink in hierarhyLinks)
            {
                try
                {
                    //Изменяем контекст проектирования на основной
                    hierarchyLink.ApplyDesignContextChangesToMainContext();
                }
                catch { }
            }
        }
    }

    /// <summary>
    /// Метод из строки, которая была записана в переменной, распарсит значение и найдет нужные подключения ЭСИ
    /// </summary>
    /// <param name="value"></param>
    /// <returns></returns>
    private List<NomenclatureHierarchyLink> ПолучитьПодключенияИзСтроки(string value)
    {
        var result = new List<NomenclatureHierarchyLink>();
        if (String.IsNullOrEmpty(value))
            throw new MacroException("value is null");

        //Разбиваем строку на коллекцию, в которой будут содержаться данные о подключениях (objectId - linkId)
        var linkDataArray = value.Split(_separatorLinkList);
        foreach (var linkDataString in linkDataArray)
        {
            //Разбиваем строку на objectId и linkId
            var linkIdData = linkDataString.Split(_separatorLinkId);

            //Если после разбития у нас получилось не 2 элемента, как должно быть, то переходим к следуюущему значению в списке
            if (linkIdData.Length != 2)
            {
                if (_testMode)
                    throw new MacroException("linkIdData.Length != 2");
                continue;
            }
            //continue;

            //Преобразуем строковое значение в число, для поиска
            if (!Int32.TryParse(linkIdData[0], out var objectId))
            {
                if (_testMode)
                    throw new MacroException("Int32.TryParse(linkIdData[0], out var objectId)");
                continue;
            }

            if (!Int32.TryParse(linkIdData[1], out var linkId))
            {
                if (_testMode)
                    throw new MacroException("Int32.TryParse(linkIdData[1], out var linkId)");
                continue;
            }
            //continue;

            //Находим объект, который был создан в неосновном контексте
            var childObject = _nomenclatureReference.Find(objectId);
            if (childObject is null)
            {
                if (_testMode)
                    throw new MacroException($"childObject is null");
                continue;
            }
            //continue;

            //Если по какой-то причине у объекта нет родителей, то выходим
            if (childObject.Parents is null || !childObject.Parents.Any())
            {
                if (_testMode)
                    throw new MacroException("childObject.Parents is null || childObject.Parents.Count() == 0");
                continue;
            }
            //continue;

            foreach (var parentObject in childObject.Parents)
            {
                //Находим все подключения дочернего к родителю
                var links = childObject.GetParentLinks(parentObject).OfType<NomenclatureHierarchyLink>();

                //Находим первое попавщееся подключение с заданным ID
                var foundedLink = links.FirstOrDefault(l => l.Id == linkId);
                if (foundedLink is not null)
                    result.Add(foundedLink);
            }
        }

        //Возвращаем результаты всех найденных подключений
        return result;
    }
}
