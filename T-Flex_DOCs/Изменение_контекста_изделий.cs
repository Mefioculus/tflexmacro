/*
Ссылки
TFlex.DOCs.Ui.Client.Dll
Tflex.DOCs.Common.dll
*/

using System;
using System.Linq;
using System.Collections.Generic;
using TFlex.DOCs.Client.ViewModels;
using TFlex.DOCs.Client.ViewModels.Base;
using TFlex.DOCs.Client.ViewModels.References;
using TFlex.DOCs.Model.Macros;
using TFlex.DOCs.Model.References;
using TFlex.DOCs.Model.References.Nomenclature;
using System.Text;
using TFlex.DOCs.Common.Extensions;

public class MacroBPMainContext : MacroProvider
{
    private static readonly string _наименованиеБП = "Изменение контекста проектирования";

    /// <summary>
    /// Разделитель идентификаторов подключений
    /// </summary>
    private static readonly char _separatorLinkList = ';';

    /// <summary>
    /// Разделитель для записи идентификатора объект - подключение
    /// </summary>
    private static readonly char _separatorLinkId = '-';

    /// <summary>
    /// Гуид контекста проектирования, на котором будет работать кнопка
    /// </summary>
    private static readonly Guid _techContexGuid = new Guid("371f983e-64cc-46fb-8fdf-c8c990a3c8bd");
    
    /// <summary>
    /// ID контекста проектирования, на котором будет работать кнопка
    /// </summary>
    private static readonly int _techContexId = 4;

    /// <summary>
    /// Гуид технологической структуры
    /// </summary>
    private static readonly Guid _tethStrucTypeGuid = new Guid("633f08c5-4aef-44f8-924b-81c3e7339aea");
    
    public MacroBPMainContext(MacroContext context)
        : base(context)
    {
    }

    private void TryRunDebug()
    {
        if (Context.Connection.ClientView.HostName == "MOSIN" && Вопрос("Хотите запустить в режиме отладки?"))
        {
            System.Diagnostics.Debugger.Launch();
            System.Diagnostics.Debugger.Break();
        }
    }

    public ButtonValidator ValidateButton()
    {
    	
    	//TryRunDebug();
        var validator = new ButtonValidator();
        validator.Visible = false;

        //Если текущий справочник не открыт в нужном контексте проектирования, то не отображаем кнопку
        
        //Определение контекста по Guid
        if (Context.Reference.ConfigurationSettings.DesignContext?.Guid != _techContexGuid)
        
        //Определение контекста по наименованию
        //if (!Context.Reference.ConfigurationSettings.DesignContext.Name.Equals(_techContexName))
                return validator;

        //Получаем выбранные подключения
        var selectedHierarhyLinks = GetSelectedHierarhyLinks();
        if (selectedHierarhyLinks.Count == 0)
            return validator;

        //Если хотя бы одно из подключений принадлежит технологическому контексту, то кнопку отображать
        if (selectedHierarhyLinks.Any(l => l.DesignContextId == _techContexId))
                validator.Visible = true;

        return validator;
    }

    public void ЗапуститьБП()
    {
        TryRunDebug();

        //Получаем все выбранные подключения
        var hierarhyLinks = GetSelectedHierarhyLinks();
        if (hierarhyLinks.Count == 0)
            Error("Не выбраны подключения");

        //Из выбранных подключений находим все подключения, которые принадлежат к Технологической структуре
        //var techStructHierarhyLinks = hierarhyLinks.FindAll(l => l.StructureTypes.Any(s => s.Guid == _tethStrucTypeGuid));
        var techStructHierarhyLinks = hierarhyLinks.FindAll(l => l.StructureTypes.Any(s => s.Guid == _tethStrucTypeGuid)
                                    || (l[new Guid("e56d2f86-11e1-4312-ad75-ca3f15896717")].Value != null 
                                        && l.StructureTypes.Any(s => s.Guid == _tethStrucTypeGuid))); // Заменяемое подключение);

        if (techStructHierarhyLinks.Count == 0)
            throw new MacroException("У выбранных подключений не найдены подключения с технологической структурой");

        //Получаем строку, в которой будут указаны все данные
        string value = СкомпоноватьПодключенияВСтроку(hierarhyLinks);

        //Создаем словарь с переменными, для передачи в БП
        var переменные = new Dictionary<string, object>
        {
            { "Список подключений", value },
            { "Контекст проектирования", Context.ConfigurationSettings.DesignContext }
        };

        //Запускаем БП, в случае, если БП не удалось запустить, генерируем ошибку
        if (БизнесПроцессы.Запустить(_наименованиеБП, переменные, показыватьДиалог: false))
            Сообщение("Информация", "Процесс сохранения технологической структуры успешно запущен");
        else
            throw new MacroException("Не удалось запустить процесс сохранения технологической структуры");
    }

    private static string СкомпоноватьПодключенияВСтроку(List<NomenclatureHierarchyLink> hierarhyLinks)
    {
        //Компонует все подключения в виде IdОбъекта-IdПодключения
        var result = hierarhyLinks.Select(l => $"{l.ChildObject.Id}{_separatorLinkId}{l.Id}");

        return String.Join(_separatorLinkList.ToString(), result);
    }

    private List<NomenclatureHierarchyLink> GetSelectedHierarhyLinks()
    {
        var result = new List<NomenclatureHierarchyLink>();

        //Приводим контекст к UIMacroContext, т.е. мы работаем с интерефейсом
        if (Context is not UIMacroContext uIMacroContext)
            return result;

        //Проверяем, что родительское окно поддерживает выборку записей
        if (uIMacroContext.OwnerViewModel is not ISupportSelection selection)
            return result;

        //Получаем выбранные объекты
        var selectedObjects = selection.SelectedObjects.OfType<IReferenceObjectViewModel>();

        //У выбранных объектов получаем подключения ЭСИ
        result = selectedObjects.Select(c => c.HierarchyLink).OfType<NomenclatureHierarchyLink>().ToList();

        return result;
    }
}
