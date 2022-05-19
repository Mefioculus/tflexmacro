/*
TFlex.DOCs.UI.Client.dll
*/

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading;
using TFlex.DOCs.Client;
using TFlex.DOCs.Client.Utils;
using TFlex.DOCs.Client.ViewModels;
using TFlex.DOCs.Client.ViewModels.Base;
using TFlex.DOCs.Client.ViewModels.Layout;
using TFlex.DOCs.Client.ViewModels.References;
using TFlex.DOCs.Client.ViewModels.References.Stages;
using TFlex.DOCs.Common;
using TFlex.DOCs.Model.Macros;
using TFlex.DOCs.Model.Macros.ObjectModel;
using TFlex.DOCs.Model.References;
using TFlex.DOCs.Model.Stages;
using TFlex.DOCs.Model.Structure;

public class MDMCreateSettings : MacroProvider
{
    public MDMCreateSettings(MacroContext context)
        : base(context)
    {
        if (Context.Connection.ClientView.HostName == "MOSINS" && Вопрос("Хотите запустить в режиме отладки?"))
        {
            Debugger.Launch();
            Debugger.Break();
        }
    }

    private static readonly string _стадияНормализация = "02b8bbcd-e24d-4aac-8853-ba93cbbec5f0";

    private readonly List<string> списокПараметровДляПравилаОбмена = new List<string>()
    {
        "Сводное наименование" ,
        "Комментарий" ,
        "Guid записи аналога" ,
        "Guid справочника аналога"
    };

    public override void Run()
    {
    }

    private Объекты ПолучитьВыбранныеОбъектыСРабочейСтраницы(string itemName) => ВыполнитьМакрос("ea0d8d3c-395b-48d5-9d42-dab98371522c", "ПолучитьВыбраныеОбъекты", itemName);

    /// <summary>
    /// АРМ Специалиста НСИ - Подготовка - Подключить Тип НСИ и Документ НСИ
    /// </summary>
    public void ПодключитьТипНСИ()
    {
        var выбранныеОбъекты = ПолучитьВыбранныеОбъектыСРабочейСтраницы("item13");
        if (выбранныеОбъекты.Count == 0)
            return;

        var группы = выбранныеОбъекты.GroupBy(ob => ob["Тип НСИ"]);
        if (группы.Count() > 1)
            Ошибка("Невозможно выполнить операцию для объектов с разными Типами НСИ. Выберите группу объектов с одинаковым Типом НСИ");

        var выбранныйОбъект = выбранныеОбъекты.FirstOrDefault();

        Guid гуидТипаНСИ = выбранныйОбъект["Тип НСИ"];
        if (гуидТипаНСИ == Guid.Empty)
            гуидТипаНСИ = ПолучитьТипНСИ(выбранныйОбъект.Справочник.Guid);

        if (гуидТипаНСИ == Guid.Empty)
            return;

        string фильтрДокументовНСИ = String.Format("[Параметры документа НСИ]->[Тип НСИ]->[Guid] = '{0}'", гуидТипаНСИ);
        //Этого не может быть,но на всякий случай проверка
        var ПоискДокументовНСИ = НайтиОбъекты("Документы НСИ", фильтрДокументовНСИ);
        if (ПоискДокументовНСИ.Count == 0)
            Ошибка("Не найден документ НСИ для выбранного типа НСИ");

        var гуидДокументаНСИ = ПолучитьДокументНСИ(фильтрДокументовНСИ);
        if (гуидДокументаНСИ == Guid.Empty)
            return;

        if (!Вопрос("Параметры: Документ НСИ, Тип НСИ будут перезаписаны." + Environment.NewLine + "Продолжить выполнение?"))
            return;

        var errors = new StringBuilder();
        foreach (var объект in выбранныеОбъекты)
        {
            try
            {
                объект.Изменить();
                объект["Тип НСИ"] = гуидТипаНСИ;
                объект["Документ НСИ"] = гуидДокументаНСИ;
                объект.Сохранить();
            }
            catch (Exception e)
            {
                errors.Append(String.Format("Ошибка при обработке объекта: {0}{1}{2}{1}",
                    объект, Environment.NewLine, e.Message));
            }
        }

        if (errors.Length > 0)
            Сообщение("Предупреждение", errors.ToString());
        else
            ПокзатьСообщениеВыпонено();
    }

    public void ПроверитьЗаписи()
    {
        var выбранныеОбъекты = ПолучитьВыбранныеОбъектыСРабочейСтраницы("Записи");
        if (выбранныеОбъекты.Count == 0)
            return;

        var errors = new StringBuilder();
        foreach (var объект in выбранныеОбъекты)
        {
            try
            {
                string checkObjectResult = ВыполнитьМакрос("75b95647-4432-457f-aaf9-44afae6575c3", "ПроверитьЗаписьМакрос", объект);
                if (!String.IsNullOrEmpty(checkObjectResult))
                    errors.AppendLine(checkObjectResult);
            }
            catch (Exception e)
            {
                errors.Append(String.Format("Ошибка при обработке объекта: {0}{1}{2}{1}",
                    объект, Environment.NewLine, e.Message));
            }
        }

        if (errors.Length > 0)
            Сообщение("Предупреждение", errors.ToString());
        else
            ПокзатьСообщениеВыпонено();
    }

    private Guid ПолучитьДокументНСИ(string фильтрДокументаНСИ)
    {
        var диалог = СоздатьДиалогВыбораОбъектов("Документы НСИ");
        диалог.Фильтр = фильтрДокументаНСИ;
        диалог.ПоказатьПанельКнопок = false;
        диалог.МножественныйВыбор = false;
        if (!диалог.Показать())
            return Guid.Empty;

        return диалог.ФокусированныйОбъект["Guid"];
    }

    /// <summary>
    /// Находит тип НСИ для указанно справочника показывает диалог с выбором типа НСИ
    /// </summary>
    /// <param name="referenceGuid"></param>
    /// <returns>Возвращает гуид типа НСИ+</returns>
    private Guid ПолучитьТипНСИ(Guid referenceGuid)
    {
        string filterString = String.Format("[Использование в пользовательских системах]->[Справочник-аналог] = '{0}'", referenceGuid);

        var найденыеТипы = НайтиОбъекты("00bf7ef0-6080-4edd-a548-95b44df465c4", filterString);
        if (найденыеТипы.Count == 0)
            Сообщение("Предупреждение", String.Format("Нет связанных типов НСИ, выбор будет осуществляться из всех доступных"));

        var диалог = СоздатьДиалогВыбораОбъектов("00bf7ef0-6080-4edd-a548-95b44df465c4");//Типы НСИ
        //Если нет подходящих типов, то отображаем все
        if (найденыеТипы.Count != 0)
            диалог.Фильтр = filterString;
        диалог.ПоказатьПанельКнопок = false;
        диалог.МножественныйВыбор = false;
        if (!диалог.Показать())
            return Guid.Empty;

        return диалог.ФокусированныйОбъект["Guid"];
    }

    /// <summary>
    /// АРМ Специалиста НСИ - Подготовка - Создать или подключить настройки соответствия
    /// </summary>
    public void СформироватьНастройкиСоответствия()
    {
        var выбранныеОбъекты = ПолучитьВыбранныеОбъектыСРабочейСтраницы("item13");
        if (выбранныеОбъекты.Count == 0)
            throw new MacroException("Нет выбранных объектов");

        var objectsGroup = выбранныеОбъекты.GroupBy(ob => ob["Документ НСИ"].ToString());
        if (objectsGroup.Count() > 1)
            throw new MacroException("Невозможно создать настройки, выбранные объекты связаны с разными Документами НСИ");

        var выбранныйОбъект = выбранныеОбъекты[0];
        var справочникВременный = ПолучитьСправочникВременный(выбранныеОбъекты[0]);
        var справочникАналог = ((ReferenceObject)выбранныйОбъект).Reference;
        Guid гуидДокументаНСИ = выбранныйОбъект["Документ НСИ"];

        var гуидВременногоСправочника = выбранныйОбъект.Справочник.Guid;
        var errors = new StringBuilder();

        Объект правилоКонвертации = null;

        var правилаКонвертации = НайтиПравилоОбменаДаннымиДляСправочников(справочникВременный.ParameterGroup.Guid.ToString(), гуидВременногоСправочника.ToString());
        if (правилаКонвертации.IsNullOrEmpty())
        {
            правилоКонвертации = СоздатьПравилоКонвертацииДанных(выбранныйОбъект, справочникВременный, справочникАналог);
        }
        else if (правилаКонвертации.Count == 1)
        {
            if (Вопрос("Правило конвертации для указанных справочников уже существует." + Environment.NewLine + "Подключить?"))
            {
                правилоКонвертации = правилаКонвертации[0];
            }
            else
            {
                if (!Вопрос("Продолжить создание нового правила?"))
                    return;

                правилоКонвертации = СоздатьПравилоКонвертацииДанных(выбранныйОбъект, справочникВременный, справочникАналог);
            }
        }
        else
        {
            var диалог = СоздатьДиалогВыбораПравил(правилаКонвертации);
            if (диалог.Показать())
            {
                правилоКонвертации = диалог.ФокусированныйОбъект;
            }
            else
            {
                if (!Вопрос("Продолжить создание нового правила?"))
                    return;

                правилоКонвертации = СоздатьПравилоКонвертацииДанных(выбранныйОбъект, справочникВременный, справочникАналог);
            }
        }

        Guid гуидПравила = правилоКонвертации["Guid"];
        ЗаполнитьПравилоВОбъекты(выбранныеОбъекты, гуидДокументаНСИ, гуидПравила, errors);

        if (errors.Length > 0)
            Сообщение("Предупреждение", errors.ToString());

        ПоказатьДиалогСвойств(правилоКонвертации);
    }

    private ДиалогВыбораОбъектов СоздатьДиалогВыбораПравил(Объекты правилаКонвертации)
    {
        var диалог = СоздатьДиалогВыбораОбъектов(Настройки.ПравилоОбменаДанными);
        диалог.Заголовок = "Найдено несколько правил конвертации, выберите правило для подключения";
        диалог.Фильтр = String.Format("[ID] Входит в список '{0}'", String.Join(", ", правилаКонвертации.Select(ob => ob["ID"])));
        диалог.МножественныйВыбор = false;
        диалог.ПоказатьПанельКнопок = false;
        return диалог;
    }

    /// <summary>
    /// Возвращает справочник аналог от документа НСИ
    /// </summary>
    /// <param name="объект">Объект временного справочника</param>
    /// <returns></returns>
    private Reference ПолучитьСправочникВременный(Объект объект)
    {
        Guid гуидДокументНСИ = объект["Документ НСИ"]; //Документ НСИ
        if (гуидДокументНСИ == Guid.Empty)
            throw new MacroException("Не указан документ НСИ");

        var документНСИ = НайтиОбъект("a169916f-fa02-417b-b52f-63de54b06a59", "Guid", гуидДокументНСИ);//Документ НСИ
        if (документНСИ == null)
            throw new MacroException("Не найден подключенный объект 'Документ НСИ'");

        string строкаГуидаВременногоСправочника = документНСИ["79d316f0-0f13-4ee9-9316-e41f27389333"]; //Guid сгенерированного справочника
        if (строкаГуидаВременногоСправочника == String.Empty)
            throw new MacroException("В документе НСИ не указан сгенерированный справочник");

        if (!Guid.TryParse(строкаГуидаВременногоСправочника, out var гуидАналога))
            throw new MacroException("В документе НСИ не указан тип параметра Guid");

        var справочникВременный = FindReference(гуидАналога);
        if (справочникВременный == null)
            throw new MacroException("В системе не найден справочник аналог, указанный в документе НСИ");

        return справочникВременный;
    }

    private Объект СоздатьПравилоКонвертацииДанных(Объект выбранныйОбъект, Reference временныйСправочник, Reference аналогСправочник)
    {
        var правилоКонвертации = СоздатьОбъект(Настройки.ПравилоОбменаДанными,//Правила обмена данными
            "2a343a98-791f-4f79-8cac-fd55f46cf81a");//Правило конвертации данных

        //Компонует строку в виде "Справочник аналог - Временный справочник"
        string наименованиеПравилаОбмена = String.Format("{0} - {1}", выбранныйОбъект.Справочник.Имя, временныйСправочник.Name);

        if (наименованиеПравилаОбмена.Length > 254)
            наименованиеПравилаОбмена = наименованиеПравилаОбмена.Substring(0, 254);

        правилоКонвертации["85f67608-ebdd-4232-84c9-41c8ac0ec241"] = наименованиеПравилаОбмена;// Наименование

        var правилоСинхронизацииСправочника = правилоКонвертации.СоздатьОбъектСписка("5397a778-a6d6-4161-8e95-dfb151f652e5", "eb35f865-37a0-47ef-ac80-5dbc61d1d4fa");
        правилоСинхронизацииСправочника["1ffbdf28-5dff-4d14-affb-6c39116d0cf2"] = наименованиеПравилаОбмена;//Описание
        правилоСинхронизацииСправочника["79dee7ec-1e0c-460c-8fc1-9b051c038799"] = временныйСправочник.ParameterGroup.Guid;//Справочник
        правилоСинхронизацииСправочника["c04fb9b1-9369-4041-b877-92ece79b1dae"] = аналогСправочник.ParameterGroup.Guid.ToString();//Внешний справочник

        СоздатьОписанияПараметров(правилоСинхронизацииСправочника, временныйСправочник, аналогСправочник);

        правилоСинхронизацииСправочника.Сохранить();

        правилоКонвертации.Сохранить();
        return правилоКонвертации;
    }

    private void СоздатьОписанияПараметров(Объект правилоСинхронизацииСправочника, Reference аналогСправочник, Reference временныйСправочник)
    {
        foreach (var параметрАналог in списокПараметровДляПравилаОбмена)
        {
            string tempParameterDescription;
            string analogParameterDescription;
            switch (параметрАналог)
            {
                case "Сводное наименование":
                    tempParameterDescription = ПолучитьОписаниеПараметраДляПравилаОбмена(временныйСправочник, "Наименование");
                    analogParameterDescription = ПолучитьОписаниеПараметраДляПравилаОбмена(аналогСправочник, параметрАналог);
                    СоздатьОписаниеПараметра(правилоСинхронизацииСправочника, analogParameterDescription, tempParameterDescription);
                    break;
                case "Комментарий":
                    tempParameterDescription = ПолучитьОписаниеПараметраДляПравилаОбмена(временныйСправочник, "Комментарий");
                    analogParameterDescription = ПолучитьОписаниеПараметраДляПравилаОбмена(аналогСправочник, параметрАналог);
                    СоздатьОписаниеПараметра(правилоСинхронизацииСправочника, analogParameterDescription, tempParameterDescription);
                    break;
                case "Guid записи аналога":
                    tempParameterDescription = ПолучитьОписаниеПараметраДляПравилаОбмена(временныйСправочник, "Guid");
                    analogParameterDescription = ПолучитьОписаниеПараметраДляПравилаОбмена(аналогСправочник, параметрАналог);
                    СоздатьОписаниеПараметра(правилоСинхронизацииСправочника, analogParameterDescription, tempParameterDescription, type: "Правило соответствия идентификатора");
                    break;
                case "Guid справочника аналога":
                    analogParameterDescription = ПолучитьОписаниеПараметраДляПравилаОбмена(аналогСправочник, параметрАналог);
                    СоздатьОписаниеПараметра(правилоСинхронизацииСправочника, analogParameterDescription, formula: "return Context.DataParameter.DataReference.SourceKey.ToString();");
                    break;
                default:
                    break;
            }
        }

        var analogParameters = аналогСправочник.ParameterGroup.OneToOneParameters.Where(p => !p.IsSystem && !списокПараметровДляПравилаОбмена.Contains(p.Name));
        var tempParamters = временныйСправочник.ParameterGroup.OneToOneParameters.Where(p => !p.IsSystem);
        foreach(var analogParameter in analogParameters)
        {
            var foundedTempParameter = tempParamters.FirstOrDefault(p => p.Type == analogParameter.Type && p.Name == analogParameter.Name);
            if (foundedTempParameter is null)
                continue;

            var tempParameterDescription = ПолучитьОписаниеПараметраДляПравилаОбмена(foundedTempParameter);
            var analogParameterDescription = ПолучитьОписаниеПараметраДляПравилаОбмена(analogParameter);
            СоздатьОписаниеПараметра(правилоСинхронизацииСправочника, analogParameterDescription, tempParameterDescription);
        }
    }

    private Объект СоздатьОписаниеПараметра(Объект правилоСинхронизации, string analogParameter, string tempParameter = "",
        string type = "Правило соответствия параметра", string formula = "")
    {
        var правилоПараметра = правилоСинхронизации.СоздатьОбъектСписка("Правила соответствия параметров", type);
        правилоПараметра["Путь"] = analogParameter;
        правилоПараметра["Внешний путь"] = tempParameter;
        правилоПараметра["Формула получения значения"] = formula;
        правилоПараметра.Сохранить();
        return правилоПараметра;
    }

    private string ПолучитьОписаниеПараметраДляПравилаОбмена(Reference reference, string parameterName)
    {
        var tempParameterInfo = reference.ParameterGroup.Parameters.FindByName(parameterName);
        if (tempParameterInfo == null)
            return String.Empty;

        return String.Format("[{0}].[{1}]", reference.ParameterGroup.Guid, tempParameterInfo.Guid);
    }

    private string ПолучитьОписаниеПараметраДляПравилаОбмена(ParameterInfo parameter) 
        => String.Format("[{0}].[{1}]", parameter.Group.ReferenceGroup.Guid, parameter.Guid);

    public void ПодключитьНастройкуОчисткиДанных() => throw new NotImplementedException();

    public void ВыгрузитьВДокументНСИ()
    {
        var errors = new StringBuilder();
        var errorsChangeStageObjects = new List<string>();
        var выбранныеОбъекты = ПолучитьВыбранныеОбъектыСРабочейСтраницы("item13");
        var группировка = выбранныеОбъекты.GroupBy(ob => (Guid)ob["Настройка соответствия"]);
        foreach (var element in группировка)
        {
            var наименованиеПравилаНастройки = element.Key;
            if (наименованиеПравилаНастройки == Guid.Empty)
            {
                errors.Append(String.Format("У объектов {0} не заданы Настройки соответствия{1}.", String.Join(", ", element.ToList()), Environment.NewLine));
                continue;
            }

            var result = ОбменДанными.ИмпортироватьОбъекты(наименованиеПравилаНастройки.ToString(), element, false);
            if (result.ИмеютсяОшибки)
            {
                errors.Append(String.Format("Во время обработки правила обмена для объектов {0} возникли ошибки.{2}{1}{2}", String.Join(", ", element.ToList()), result.ТекстОшибки, Environment.NewLine));
            }
            else
            {
                foreach (var объект in element)
                {
                    if (!объект.ИзменитьСтадию(_стадияНормализация))
                        errorsChangeStageObjects.Add(объект.ToString());
                }
            }
        }

        if (errorsChangeStageObjects.Count > 0)
            errors.Append(String.Format("У объектов {0} стадия не была изменена.{1}", String.Join(", ", errorsChangeStageObjects), Environment.NewLine));

        if (errors.Length > 0)
        {
            throw new MacroException("В время обработки возникли ошибки." + Environment.NewLine + errors.ToString());
        }
        else
        {
            Сообщение("Выполнено", "Операция выполнена успешно");
        }
    }

    public void ОткрытьОкноСменыСтадий()
    {
        var viewModel = GetSupportSelection("Записи");

        if (viewModel == null)
            throw new MacroException("Внимание! Не найден элемент управления 'Записи'");

        var selectedObjects = viewModel.GetSelectedReferenceObjects().ToArray();

        if (selectedObjects.Length == 0)
            throw new MacroException("Внимание! Не выбраны объекты для смены стадии");

        if (selectedObjects.Any(o => !o.Reference.ParameterGroup.SupportsStages))
            throw new MacroException("Внимание! Не все выбранные объекты поддерживают смену стадий");

        if (!selectedObjects[0].Reference.Connection.IsAdministrator &&
            selectedObjects.Any(o => o.SystemFields.Stage.Transitions.Count == 0 ||
                                     o.SystemFields.Stage.Transitions.Any(t => !t.IsManual)))
        {
            throw new MacroException(
                "Внимание! У выбранных объектов нет доступных переходов для ручного переключения стадии");
        }

        OpenChangeStageDialog(selectedObjects, viewModel);
    }

    private void ЗаполнитьПравилоВОбъекты(Объекты объекты, Guid гуидДокументНСИ, Guid гуидПравила, StringBuilder errors)
    {
        foreach (var объект in объекты)
        {
            try
            {
                Guid гуидОбъекта = объект["Документ НСИ"];

                if (гуидОбъекта != гуидДокументНСИ) //Документ НСИ)
                    continue;

                объект.Изменить();
                объект["Настройка соответствия"] = гуидПравила; //Настройка соответствия
                объект.Сохранить();
            }
            catch
            {
                errors.Append(String.Format("Ошибка при изменении объекта: {0} возможно объект заблокирован другим пользователем{1}", объект, Environment.NewLine));
            }
        }
    }

    private Объекты НайтиПравилоОбменаДаннымиДляСправочников(string справочник, string внешнийСправочник)
    {
        string filterString = ПолучитьСтрокуПоисаПравилаОбмена(справочник, внешнийСправочник);
        return НайтиОбъекты(Настройки.ПравилоОбменаДанными, filterString);
    }

    private string ПолучитьСтрокуПоисаПравилаОбмена(string справочник, string внешнийСправочник)
        => String.Format("[Правила синхронизации справочников].[Справочник] = '{0}' И " +
                             "[Правила синхронизации справочников].[Внешний справочник] = '{1}'",
                              справочник, внешнийСправочник);

    private Reference FindReference(Guid referenceGuid)
    {
        var referenceInfo = Context.Connection.ReferenceCatalog.Find(referenceGuid);
        if (referenceInfo == null)
            return null;
        return referenceInfo.CreateReference();
    }

    private ISupportSelection GetSupportSelection(string itemName)
    {
        var currentWindow = Context.GetCurrentWindow();
        var foundItem = currentWindow?.FindItem(itemName);

        if (foundItem == null)
            return null;

        return foundItem switch
        {
            // Окно справочника
            ReferenceWindowLayoutItemViewModel referenceWindowLayoutItemVM => referenceWindowLayoutItemVM.InnerViewModel as ISupportSelection,
            // Окно структуры объекта
            ObjectStructureLayoutItemViewModel objectStructureLayoutItemViewModel => objectStructureLayoutItemViewModel.InnerViewModel,
            // Диалог ввода, связанные объекты
            LinkToManyLayoutItemViewModel linkToManyLayoutItemViewModel => linkToManyLayoutItemViewModel.LinkContent as ISupportSelection,
            _ => null,
        };
    }

    private static async void OpenChangeStageDialog(ReferenceObject[] selectedObjects,
        ISupportSelection viewModel)
    {
        using var dialog = new ChangeStageDialogViewModel(viewModel as LayoutViewModel, selectedObjects[0].Reference, selectedObjects);
        if (!await ApplicationManager.OpenDialogAsync(dialog, viewModel))
            return;

        if (dialog.Objects.Count == 0)
            return;

        await DOCsTaskFactory.Run(() =>
        {
            if (dialog.NewStage == null)
                Stage.Clear(dialog.Connection, dialog.Objects);
            else if (dialog.IgnoreScheme)
                dialog.NewStage.Set(dialog.Objects);
            else
                dialog.NewStage.Change(dialog.Objects);
        }, CancellationToken.None);

        ReloadView(viewModel);
    }

    private static void ReloadView(ISupportSelection viewModel)
    {
        switch (viewModel)
        {
            case ReferenceExplorerViewModel explorerViewModel:
                explorerViewModel.GridViewModel?.ReloadData();
                explorerViewModel.TreeViewModel?.ReloadData();
                break;
            case IDataSourceViewModel dataSourceViewModel:
                dataSourceViewModel.ReloadData();
                break;
        }
    }

    private void ПокзатьСообщениеВыпонено() => Сообщение("Выполнено", "Обработка завершена, ошибок не обнаружено");

    private class Настройки
    {
        public static string ПравилоОбменаДанными = "212a5ec8-3f36-4501-bb46-082d200ba05f";
    }
}
