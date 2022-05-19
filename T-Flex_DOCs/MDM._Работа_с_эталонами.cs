/* Дополнительные ссылки
TFlex.DOCs.UI.Client.dll
TFlex.DOCs.Common.dll
 */

/* Дополнительные файлы
TFlex.DOCs.Model.NormativeReferenceInfo.dll
*/

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using TFlex.DOCs.Client.ViewModels;
using TFlex.DOCs.Client.ViewModels.Base;
using TFlex.DOCs.Client.ViewModels.Layout;
using TFlex.DOCs.Client.ViewModels.References;
using TFlex.DOCs.Client.ViewModels.WorkingPages;
using TFlex.DOCs.Model.Desktop;
using TFlex.DOCs.Model.Macros;
using TFlex.DOCs.Model.Macros.ObjectModel;
using TFlex.DOCs.Model.Macros.ObjectModel.Layout;
using TFlex.DOCs.Model.Parameters;
using TFlex.DOCs.Model.References;
using TFlex.DOCs.Model.References.NormativeReferenceInfo.RequestsNsi;
using TFlex.DOCs.Model.Search;
using TFlex.DOCs.Model.Stages;
using TFlex.DOCs.Model.Structure;
using TFlex.DOCs.Model.Structure.Builders;

public class WorkWithEtalonsMacro : MacroProvider
{
    private List<ParameterInfo> _extendedParameters;

    private List<ParameterType> _allParameterTypeList;

    private readonly StringBuilder _errors = new StringBuilder();

    /// <summary> Guid макроса - 'MDM. АРМ. Взять объект ячейку' </summary>
    private const string TakeCellObjectMacroGuid = "ea0d8d3c-395b-48d5-9d42-dab98371522c";



    private static string startSeparatorStr = "[";
    private static string endSeparatorStr = "]";

    public WorkWithEtalonsMacro(MacroContext context)
        : base(context)
    {
        if (Context.Connection.ClientView.HostName == "MOSINS")
            if (Вопрос("Хотите запустить в режиме отладки?"))
            {
                System.Diagnostics.Debugger.Launch();
                System.Diagnostics.Debugger.Break();
            }
    }

    private List<ParameterType> AllParameterTypeList => _allParameterTypeList ??= ParameterType.GetTypeList(true);

    private List<ParameterInfo> ExtendedParameters => _extendedParameters ??= Context.Connection.ExtendedParameters.GetExtendedParameters().ToList();

    private new Объекты ВыбранныеОбъекты => ВыполнитьМакрос(TakeCellObjectMacroGuid, "ПолучитьВыбраныеОбъекты", "Записи");

    public override void Run()
    {
    }

    /// <summary>
    /// Справочник эталонов, диалог свойств
    /// </summary>
    public void ПодключитьДопПараметрыКСправочникуЭталонов()
    {
        if (ТекущийОбъект == null)
            Ошибка("Нет выбранного объекта");

        var checkObjectHelper = new CheckObjectHelper(ТекущийОбъект, CheckObjectHelper.ObjectType.Etalon);
        List<Объект> параметрыДокумента = ПолучитьПараметрыДокументаНСИУОбъекта(checkObjectHelper, out var etalonReference);

        var errors = new StringBuilder();
        foreach (var параметрДокумента in параметрыДокумента)
        {
            bool дополнительный = параметрДокумента["Дополнительный"];
            if (!дополнительный)
                continue;

            var имяДопПараметра = параметрДокумента["Наименование"];
            int типПараметра = параметрДокумента["Тип параметра"];
            var глоссарий = параметрДокумента.СвязанныйОбъект["Параметр глоссария"];
            if (глоссарий == null)
            {
                errors.Append(String.Format("Ошибка в описании дополнительного параметра: {0}{1}Не указан глоссарий{1}", имяДопПараметра, Environment.NewLine));
                continue;
            }
            try
            {
                var extendedParameterInfoAlias = СоздатьДополнительныйПараметрИПсевдоним(etalonReference, GetParameterType(типПараметра), глоссарий.ToString(), имяДопПараметра);
                if (extendedParameterInfoAlias == null)
                {
                    errors.Append(String.Format("Ошибка при создании дополнительного параметра: {0}{1}", имяДопПараметра, Environment.NewLine));
                    continue;
                }
            }
            catch (Exception e)
            {
                errors.Append(String.Format("Ошибка в создании дополнительного параметра: {0}{1}{2}{1}", имяДопПараметра, Environment.NewLine, e.Message));
                continue;
            }
        }

        etalonReference.ParameterGroup.ReferenceInfo.RefreshDescription();
        etalonReference.Refresh(false);

        if (errors.Length > 0)
            Ошибка($"Во время проверки возникли ошибки{Environment.NewLine}{errors}");
        else
            Сообщение("Выполнено", "Проверка выполнена, все дополнительные параметры были подключены");
        //Context.Reference.Refresh(false);
    }

    private List<Объект> ПолучитьПараметрыДокументаНСИУОбъекта(CheckObjectHelper checkObjectHelper, out Reference etalonReference)
    {
        var документНСИ = checkObjectHelper.CheckObjectType switch
        {
            CheckObjectHelper.ObjectType.Etalon
            => checkObjectHelper.Объект.СвязанныйОбъект["Документ НСИ"] ?? throw new MacroException("У объекта не указан документ НСИ"),
            CheckObjectHelper.ObjectType.TempObject
            => НайтиОбъект("Документы НСИ", "Guid сгенерированного справочника", checkObjectHelper.ReferenceObject.Reference.ParameterGroup.Guid.ToString()) ?? throw new MacroException("Не найден документ НСИ на выбранный справочник"),
            _ => throw new NotSupportedException(nameof(checkObjectHelper.CheckObjectType)),
        };
        etalonReference = GetEtalonReference(документНСИ);
        return ПолучитьПараметрыДокументаНСИ(документНСИ);
    }

    public void ЗаполнитьЗначенияДопПараметров()
    {
        if (ТекущийОбъект == null)
            Ошибка("Нет выбранного объекта");

        var документНСИ = ТекущийОбъект.СвязанныйОбъект["Документ НСИ"] ?? throw new MacroException("У объекта не указан документ НСИ");

        var параметрыДокумента = ПолучитьПараметрыДокументаНСИ(документНСИ).FindAll(p => p["Дополнительный"] == true);
        if (параметрыДокумента.Count == 0)
            return;

        var currentObject = Context.ReferenceObject;
        var parameters = currentObject.ParameterValues.Select(s => s.ParameterInfo);
        foreach (var параметр in параметрыДокумента)
        {
            var foundParameter = parameters.FirstOrDefault(p => p.Name == параметр["Наименование"]);
            if (foundParameter == null)
                continue;

            currentObject[foundParameter].Value ??= foundParameter.Type.DefaultValue;
        }
    }

    public void ДобавитьДопПараметрКЭталону()
    {
        var документНСИ = ТекущийОбъект.СвязанныйОбъект["Документ НСИ"];
        if (документНСИ == null)
            Ошибка("У объекта не указан документ НСИ");

        var referenceExtendedParameter = Context.ReferenceObject.Reference.ParameterGroup.Parameters.Where(p => p.IsExtended).ToList();

        var foundedExtendedParameters = new List<string>();

        var параметрыДокумента = ПолучитьПараметрыДокументаНСИ(документНСИ);
        foreach (var параметрДокумента in параметрыДокумента)
        {
            bool дополнительный = параметрДокумента["Дополнительный"];
            if (!дополнительный)
                continue;

            var имяДопПараметра = параметрДокумента["Наименование"];
            if (referenceExtendedParameter.Exists(p => p.Name == имяДопПараметра))
                foundedExtendedParameters.Add(имяДопПараметра);
        }

        if (foundedExtendedParameters.Count == 0)
            Ошибка("Не найдены подходящие параметры для указанного документа НСИ");

        var диалог = СоздатьДиалогВвода("Выберите дополнительный параметр");
        диалог.ДобавитьВыборИзСписка("Параметр", String.Empty, true, foundedExtendedParameters.ToArray());
        if (!диалог.Показать())
            return;

        var foundParameter = referenceExtendedParameter.Find(p => p.Name == диалог["Параметр"]);
        if (foundParameter == null)
            Ошибка("Ошибка поиска параметра в текущем справочнике");

        Context.ReferenceObject[foundParameter].Value = foundParameter.Type.DefaultValue;
    }

    /// <summary>
    /// АРМ Эксперта - Создание эталонов - Работа с документами НСИ
    /// </summary>
    public void ПодключитьЭталон()
    {
        var выбранныеОбъекты = ВыбранныеОбъекты;
        if (выбранныеОбъекты.Count == 0)
            Ошибка("Не найдены выбранные объекты");

        ПроверитьСтадииВыбранныхОбъектов(выбранныеОбъекты);

        var временныйСправочник = ((ReferenceObject)выбранныеОбъекты[0]).Reference;
        var referenceGuid = временныйСправочник.ParameterGroup.Guid;

        var документНСИ = НайтиОбъект("Документы НСИ", "79d316f0-0f13-4ee9-9316-e41f27389333", referenceGuid.ToString()) ?? throw new MacroException("Не найден документ НСИ: на текущий справочник");
        var справочникЭталона = GetEtalonReference(документНСИ);

        var диалогВыбораЭталона = СоздатьДиалогВыбораОбъектов(справочникЭталона.Name);
        диалогВыбораЭталона.МножественныйВыбор = false;
        if (!диалогВыбораЭталона.Показать())
            return;

        var эталон = диалогВыбораЭталона.ФокусированныйОбъект ?? throw new MacroException("Не выбран объект в диалоге");

        var документНСИЭталона = эталон.СвязанныйОбъект["Документ НСИ"] ?? throw new MacroException("У выбранного эталона не указан документ НСИ");
        if (документНСИЭталона["ID"] != документНСИ["ID"])
            throw new MacroException("У эталона отличается документ НСИ");

        foreach (var объект in выбранныеОбъекты)
        {
            try
            {
                объект.Изменить();
                объект.СвязанныйОбъект["Эталон"] = эталон;
                объект.Сохранить();
                ИзменитьСтадиюЗаписи(объект, _errors);
            }
            catch (Exception e)
            {
                _errors.Append(String.Format("Ошибка при обработке объекта: {0}{1}{2}{1}", объект, Environment.NewLine, e.Message));
            }
        }

        if (_errors.Length > 0)
            Сообщение("Предупреждение", _errors.ToString());
        else
            Сообщение("Сообщение", "Эталон подключен");
    }

    private void ПроверитьСтадииВыбранныхОбъектов(Объекты выбранныеОбъекты)
    {
        if (выбранныеОбъекты.Any(объект => ((ReferenceObject)объект).SystemFields.Stage.Guid != StageGuids.PreparedStageGuid))
            Ошибка("Один из выбранных объектов не находится на стадии подготовлено");
    }

    /// <summary>
    /// АРМ Эксперта - Создание эталонов - Работа с документами НСИ
    /// </summary>
    public void СоздатьЭталон()
    {
        var выбранныеОбъекты = ВыбранныеОбъекты.ToList();
        if (выбранныеОбъекты.Count == 0)
            Ошибка("Не найдены выбранные объекты");

        ПроверитьСтадииВыбранныхОбъектов(выбранныеОбъекты);

        var временныйСправочник = ((ReferenceObject)выбранныеОбъекты[0]).Reference;
        var referenceGuid = временныйСправочник.ParameterGroup.Guid;
        var документНСИ = НайтиОбъект("Документы НСИ", "79d316f0-0f13-4ee9-9316-e41f27389333", referenceGuid.ToString())
             ?? throw new MacroException("Не найден документ НСИ: на текущий справочник");

        var параметрыДокумента = ПолучитьПараметрыДокументаНСИ(документНСИ);
        var etalonReference = GetEtalonReference(документНСИ);
        var списокПараметров = ПолучитьСписокПараметров(параметрыДокумента, _errors, etalonReference, временныйСправочник);
        if (_errors.Length > 0)
            Ошибка(_errors.ToString());

        //Блок для поиска эталона с заданными параметрами
        var соответствиеЭталонов = ПолучитьСоответствияЭталонов(выбранныеОбъекты, etalonReference, списокПараметров); //ПолучитьСоответствияЭталонов
        if (!ВопросОбОбновленииЭталонов(etalonReference, соответствиеЭталонов))
            return;

        string типЭталона = ПолучитьБазовыйТипСправочника(etalonReference);
        var checkInList = new List<ReferenceObject>();
        etalonReference.Refresh(false);
        foreach (var etalonManager in соответствиеЭталонов)
        {
            try
            {
                if (etalonManager.EtalonObject == null)
                {
                    string справочник = etalonReference.ParameterGroup.Guid.ToString();

                    var эталон = СоздатьОбъект(справочник, типЭталона);
                    ЗаполнитьПараметрыЭталона(списокПараметров, etalonManager, эталон);
                    эталон.СвязанныйОбъект["Документ НСИ"] = документНСИ;
                    эталон.Сохранить();
                    checkInList.Add((ReferenceObject)эталон);

                    etalonManager.Объект.СвязанныйОбъект["Эталон"] = эталон;
                    etalonManager.Объект.Сохранить();
                    ИзменитьСтадиюЗаписи(etalonManager.Объект, _errors);
                }
                else if (etalonManager.LinkEtalon)
                {
                    var эталон = Объект.CreateInstance(etalonManager.EtalonObject, Context);

                    if (etalonManager.UpdateEtalon)
                    {
                        ЗаполнитьПараметрыЭталона(списокПараметров, etalonManager, эталон);
                        эталон.СвязанныйОбъект["Документ НСИ"] = документНСИ;
                        эталон.Сохранить();
                        checkInList.Add((ReferenceObject)эталон);
                    }

                    etalonManager.Объект.Изменить();
                    etalonManager.Объект.СвязанныйОбъект["Эталон"] = эталон;
                    etalonManager.Объект.Сохранить();
                    ИзменитьСтадиюЗаписи(etalonManager.Объект, _errors);
                }
            }
            catch (Exception e)
            {
                _errors.Append(String.Format("Ошибка при создании объекта эталона: {0}{1}{2}{1}",
                    etalonManager, Environment.NewLine, e.Message));
            }
        }

        if (checkInList.Count > 0)
            Desktop.CheckIn(checkInList, "Создание эталона", true);

        if (_errors.Length > 0)
            Сообщение("Предупреждение", _errors.ToString());
        else
            Сообщение("Сообщение", "Операция завершена успешно");
    }

    private void ЗаполнитьПараметрыЭталона(List<ParameterData> списокПараметров, CreateEtalonManager etalonManager, Объект эталон)
    {
        foreach (var параметр in списокПараметров)
        {
            try
            {
                ОбработатьПараметр(etalonManager.Объект, эталон, параметр);
            }
            catch (Exception e)
            {
                _errors.Append(String.Format("Ошибка при записи параметр объекта эталона: {0}{1}{2}{1}",
                    параметр.TempParameter, Environment.NewLine, e.Message));
            }
        }
    }

    /// <summary>
    /// Показывает справочник с выбором эталонов которые нужно подключить
    /// Показывает вопрос, какие эталоны стоит обновить
    /// </summary>
    /// <param name="etalonReference"></param>
    /// <param name="соответствиеЭталонов"></param>
    /// <returns>Возвращает true если необходимо продолжить обработку</returns>
    private bool ВопросОбОбновленииЭталонов(Reference etalonReference, List<CreateEtalonManager> соответствиеЭталонов)
    {
        var найденыеЭталоны = соответствиеЭталонов.FindAll(key => key.EtalonObject != null);
        if (найденыеЭталоны.Count == 0)
            return true;

        var диалог = СоздатьДиалогВыбораОбъектов(etalonReference.Name);
        диалог.Заголовок = "Выберите эталоны к которым хотите подключить к записям";
        диалог.МножественныйВыбор = true;
        диалог.Фильтр = CreateEtalonManager.GetIdFilterString(найденыеЭталоны);
        if (диалог.Показать() && диалог.ВыбранныеОбъекты.Count != 0)
        {
            foreach (var объект in диалог.ВыбранныеОбъекты)
            {
                var эталоны = найденыеЭталоны.FindAll(e => e.IdEtalon == объект["ID"]);
                if (эталоны.Count == 0)
                    continue;

                эталоны.ForEach(e => e.LinkEtalon = true);
            }

            if (соответствиеЭталонов.Count == 1 && Вопрос("Обновить данные у существующих эталонов?"))
                найденыеЭталоны.FindAll(e => e.LinkEtalon).ForEach(e => e.UpdateEtalon = true);

            return true;
        }
        else
        {
            return Вопрос("Продолжить выполнение операции?");
        }
    }

    private List<CreateEtalonManager> ПолучитьСоответствияЭталонов(List<Объект> выбранныеОбъекты, Reference etalonReference, List<ParameterData> списокПараметров)
    {
        var списокКлючевыхПараметров = списокПараметров.FindAll(par => par.IsKey);
        var соответствиеЭталонов = new List<CreateEtalonManager>();
        foreach (var выбранныйОбъект in выбранныеОбъекты)
        {
            var createEtalonManajer = new CreateEtalonManager()
            {
                Объект = выбранныйОбъект,
            };

            соответствиеЭталонов.Add(createEtalonManajer);

            if (списокКлючевыхПараметров.Count == 0)
                continue;
            //Найти эталон
            var filter = new Filter(etalonReference.ParameterGroup);
            foreach (var ключевойПараметр in списокКлючевыхПараметров)
            {
                filter.Terms.AddTerm(ключевойПараметр.EtalonParameterInfo, ComparisonOperator.Equal, выбранныйОбъект[ключевойПараметр.TempParameter]);
            }

            var foundedEtalons = etalonReference.Find(filter);
            if (foundedEtalons.Count > 1)
                _errors.Append(String.Format("Для объекта: {0} Найдено несколько соответствующих эталонов, по ключевым полям{1}", выбранныйОбъект, Environment.NewLine));

            createEtalonManajer.EtalonObject = foundedEtalons.FirstOrDefault();
        }

        return соответствиеЭталонов;
    }

    /// <summary>
    /// Создает параметры, подключает связи.
    /// </summary>
    /// <param name="errors"></param>
    /// <param name="записьВременногоСправочника"></param>
    /// <param name="эталон"></param>
    /// <param name="parameterData"></param>
    private void ОбработатьПараметр(Объект записьВременногоСправочника, Объект эталон, ParameterData parameterData)
    {
        if (parameterData.IsLink && parameterData.LinkData != null)
        {
            var связанныйОбъект =
                НайтиОбъект(parameterData.LinkData.ReferenceGuid.ToString(),
                parameterData.LinkData.LinkedParameter.ToString(),
                записьВременногоСправочника[parameterData.TempParameter]);
            if (связанныйОбъект == null)
                return;

            эталон.СвязанныйОбъект[parameterData.EtalonParameter.ToString()] = связанныйОбъект;
        }
        else
        {
            эталон[parameterData.EtalonParameter.ToString()] = записьВременногоСправочника[parameterData.TempParameterInfo.Guid.ToString()];
        }
    }

    private void ИзменитьСтадиюЗаписи(Объект записьВременногоСправочника, StringBuilder errors)
    {
        if (записьВременногоСправочника.ИзменитьСтадию(StageGuids.ProcessedStageGuid))
        {
            if ((Guid)записьВременногоСправочника["Гуид справочника аналога"] == Guid.Empty || (Guid)записьВременногоСправочника["Гуид записи аналога"] == Guid.Empty)
            {
                //errors.Append(String.Format("У объекта: {0}, не указанна запись аналога{1}", записьВременногоСправочника, Environment.NewLine));
                return;
            }
            try
            {
                var объектАналог = НайтиОбъект(записьВременногоСправочника["Гуид справочника аналога"], "Guid", записьВременногоСправочника["Гуид записи аналога"]);
                if (объектАналог == null)
                {
                    errors.Append(String.Format("Объект аналог для записи: {0}, не найден{0}", записьВременногоСправочника, Environment.NewLine));
                    return;
                }
                if (!объектАналог.ИзменитьСтадию(StageGuids.ProcessedStageGuid))
                {
                    errors.Append(String.Format("Ошибка при изменении стадии аналог объекта: {0}{1}", объектАналог, Environment.NewLine));
                }
            }
            catch
            {
                errors.Append(String.Format("Ошибка при поиске объекта аналога для записи: {0}{1}Возможно указан несуществующий справочник{1}", записьВременногоСправочника, Environment.NewLine));
            }
        }
        else
        {
            errors.Append(String.Format("Ошибка при изменении стадии объекта: {0}{1}", записьВременногоСправочника, Environment.NewLine));
        }
    }

    /// <summary>
    /// АРМ Эксперта - Создание эталонов - Работа с Эталонами
    /// </summary>
    public void СоздатьЭталонЭксперта()
    {
        string справочник = ВыполнитьМакрос("9b941dc9-2930-4e25-b493-5371de8fb6ae", "ПолучитьСправочникЭксперта");//АРМ НСИ. Отображение справочника
        if (String.IsNullOrEmpty(справочник))
            throw new MacroException("Не указан справочник");

        var типНСИ = НайтиОбъект("Типы НСИ", $"[Справочник эталона] = '{справочник}'") ?? throw new MacroException("Не найден тип НСИ для выбранного справочника");

        var диалог = СоздатьДиалогВыбораОбъектов("Документы НСИ");
        диалог.МножественныйВыбор = false;
        диалог.Фильтр = $"[Тип] = '{типНСИ["Тип документа НСИ"]}'";
        if (!диалог.Показать())
            return;

        var документНСИ = диалог.ФокусированныйОбъект ?? throw new MacroException("Не выбран документ НСИ");

        var etalonReference = FindReference(справочник);
        var параметрыДокумента = ПолучитьПараметрыДокументаНСИ(документНСИ);
        var errors = new StringBuilder();
        List<ParameterData> parameterDataList = ПолучитьВсеСоответствияПараметров(параметрыДокумента, errors, etalonReference, false);
        etalonReference.Refresh(false);

        var etalonObject = etalonReference.CreateReferenceObject();
        foreach (var parameterData in parameterDataList.FindAll(p => p.IsExtendedParameter))
        {
            if (parameterData.EtalonParameterInfo == null)
            {
                errors.AppendLine("Внутренняя ошибка, не указано описание дополнительного параметра");
                continue;
            }
            etalonObject[parameterData.EtalonParameterInfo].Value = parameterData.EtalonParameterInfo.Type.DefaultValue;
        }
        var эталон = Объект.CreateInstance(etalonObject, Context);
        эталон.СвязанныйОбъект["Документ НСИ"] = документНСИ;
        ПоказатьДиалогСвойств(Объект.CreateInstance(etalonObject, Context));
    }

    public void ПроверитьЭталоны()
    {
        var выбранныеОбъекты = ВыбранныеОбъекты;
        if (выбранныеОбъекты.Count == 0)
            Ошибка("Не найдены выбранные объекты");

        var errors = new StringBuilder();

        var гуидТекущегоСправочника = выбранныеОбъекты[0].Справочник.Guid;
        var документНСИ = НайтиОбъект("Документы НСИ", "Guid генерированного справочника", гуидТекущегоСправочника);
        if (документНСИ == null)
            throw new MacroException("Не найден документ НСИ на выбранный справочник");

        var параметрыДокумента = ПолучитьПараметрыДокументаНСИ(документНСИ);

        foreach (var объектВременногоСправочника in выбранныеОбъекты)
        {
            try
            {
                var эталон = объектВременногоСправочника.СвязанныйОбъект["Эталон"];
                var checkObjectHelper = new CheckObjectHelper(эталон, CheckObjectHelper.ObjectType.Etalon);
                var result = ValidateCheckObject(checkObjectHelper, параметрыДокумента);
                if (result.HasErrors)
                    errors.AppendLine($"Во время проверки эталона возникли ошибки:{Environment.NewLine}{result.GetErrorMessage()}");
            }
            catch (Exception e)
            {
                errors.Append(String.Format("Ошибка при обработке объекта: {0}{1}{2}{1}",
                    объектВременногоСправочника, Environment.NewLine, e.Message));
            }
        }

        if (errors.Length > 0)
            Сообщение("Предупреждение", errors.ToString());
        else
            Сообщение("Сообщение", "Эталон проверен, ошибок не найдено");
    }

    /// <summary>
    /// Справочник эталонов - диалог свойств - кнопка "Проверить значения"
    /// </summary>
    public void ПроверитьЭталонДиалогСвойств() => ValidateCheckObject(new CheckObjectHelper(ТекущийОбъект, CheckObjectHelper.ObjectType.Etalon), recordResult: true);

    /// <summary>
    /// Возвращает строковое представление результата проверки эталона
    /// </summary>
    /// <returns></returns>
    public string ПроверитьЭталонМакрос(Объект эталон) => ValidateCheckObject(new CheckObjectHelper(эталон, CheckObjectHelper.ObjectType.Etalon)).GetErrorMessage().ToString();

    public string ПроверитьЗаписьМакрос(Объект временнаяЗапись)
        => ValidateCheckObject(new CheckObjectHelper(временнаяЗапись, CheckObjectHelper.ObjectType.TempObject), recordResult: true).GetErrorMessage().ToString();

    /// <summary>
    /// Записывает в объект результат проверки
    /// </summary>
    /// <param name="result"></param>
    private void WriteResultToObject(ChecObjectResult result)
    {
        if (!TryBeginChanges(result.SourseObject))
            return;

        var resultParameterInfo = FindParameterInfo(result.SourseObject, "Проверка пройдена");
        var resultTextParameterInfo = FindParameterInfo(result.SourseObject, "Комментарий");
        if (resultParameterInfo == null || resultTextParameterInfo == null)
            return;

        result.SourseObject[resultParameterInfo].Value = result.HasErrors ? 1 : 2;
        result.SourseObject[resultTextParameterInfo].Value = result.GetErrorMessage().ToString();
        result.SourseObject.EndChanges();
    }

    public void УтвердитьЭталоны()
    {

    }

    /// <summary>
    /// Метод производит проверку эталона с заданными параметрами документа НСИ
    /// Если параметры документа НСИ не указанно, то они будут получены от эталона
    /// </summary>
    /// <param name="объект">Объект справочника "Эталоны"</param>
    /// <param name="параметрыДокумента">Список объектов параметры документа НСИ</param>
    /// <returns>Возвращает результат проверки с текстом всех ошибок</returns>
    private ChecObjectResult ValidateCheckObject(CheckObjectHelper objectHelper, List<Объект> параметрыДокумента = null, bool recordResult = false)
    {
        var result = new ChecObjectResult(objectHelper.ReferenceObject);
        if (параметрыДокумента == null || параметрыДокумента.Count == 0)
        {
            параметрыДокумента = ПолучитьПараметрыДокументаНСИУОбъекта(objectHelper, out _);
            if (параметрыДокумента.Count == 0)
            {
                result.AddError("У объекта не заданны параметры документа НСИ");
                return result;
            }
        }

        objectHelper.ReferenceObject.Reload();
        objectHelper.Parameters = параметрыДокумента.Select(p => p.ToString()).ToList();
        foreach (var параметрОбъект in параметрыДокумента)
        {
            try
            {
                ПроверитьПараметрУОбъекта(objectHelper, (ReferenceObject)параметрОбъект, result);
            }
            catch (Exception e)
            {
                result.AddError($"Необработанное исключение при обработке параметра: {параметрОбъект}{Environment.NewLine}{e.Message}");
            }
        }

        if (recordResult)
            WriteResultToObject(result);

        return result;
    }

    /// <summary>
    /// Диалог свойств справочника эталона
    /// </summary>
    public void СформироватьСводноеНаименование()
    {
        string parameterName = "Сводное наименование";
        var objectHelper = new CheckObjectHelper(ТекущийОбъект, CheckObjectHelper.ObjectType.Etalon);
        ValidateGenerateName(objectHelper, parameterName);
    }

    public void СформироватьПараметрАРМВременнаяЗапись()
    {
        var выбранныеОбъекты = ВыбранныеОбъекты;
        string parameterName = ВыполнитьМакрос("ea0d8d3c-395b-48d5-9d42-dab98371522c", "ПолучитьКолонку", "Записи");
        if (String.IsNullOrEmpty(parameterName))
            Ошибка("Не удалось получить наименование текущей колонки");

        //string parameterName = "Сводное наименование";
        foreach (var объект in выбранныеОбъекты)
        {
            var objectHelper = new CheckObjectHelper(объект, CheckObjectHelper.ObjectType.TempObject);
            ValidateGenerateName(objectHelper, parameterName);
            объект.Сохранить();
        }
    }

    private void ValidateGenerateName(CheckObjectHelper objectHelper, string parameterName)
    {
        objectHelper.ReferenceObject.Reload();
        var параметрыДокумента = ПолучитьПараметрыДокументаНСИУОбъекта(objectHelper, out _);
        objectHelper.Parameters = параметрыДокумента.Select(p => p.ToString()).ToList();
        if (objectHelper.Parameters.Count == 0)
            Ошибка("У объекта не заданны параметры документа НСИ");

        var параметрСводноеНаименование = параметрыДокумента.Find(p => p.ToString() == parameterName);
        if (параметрСводноеНаименование == null)
            Ошибка("В параметрах не задан параметр: " + parameterName);

        string formulaResult = параметрСводноеНаименование[DescriptionParametersNSI.Formula.ToString()];
        if (String.IsNullOrWhiteSpace(formulaResult))
            Ошибка("У параметра документа НСИ не задана формула");

        if (!РасчитатьФормулуПоПараметрам(objectHelper, ref formulaResult))
            Ошибка("Ошибка при разборе формулы");

        formulaResult = ParseNumFormula((ReferenceObject)параметрСводноеНаименование, formulaResult);
        objectHelper.Объект[parameterName] = formulaResult;
    }

    /// <summary>
    /// АРМ Эксперта - Создание эталонов - Работа с Эталонами
    /// </summary>
    public void ДобавитьВвыбраннуюЗаявку_РабочаяСтраница()
        => AddToSelectedRequest(GetEtalonsFromReferenceControl());

    public void ДобавитьВвыбраннуюЗаявку_Справочник()
        => AddToSelectedRequest(Context.GetSelectedObjects());

    private void AddToSelectedRequest(ReferenceObject[] selectedEtalons)
    {
        if (selectedEtalons == null || selectedEtalons.Length == 0)
            return;

        if (selectedEtalons.Any(s => s.SystemFields.Stage.Stage.Guid != StageGuids.ApprovedStageGuid))
        {
            Сообщение("Внимание!", "Добавляемые объекты должны находиться в стадии 'Утверждено'");
            return;
        }

        string filter =
            $"[{RequestNsiObject.RelationKeys.Executor}] Текущий пользователь " +
            $"И ([Тип] = '{RequestNsiReference.TypesKeys.CreationEtalonObject}' " +
            $"ИЛИ [Тип] = '{RequestNsiReference.TypesKeys.AdjustmentEtalonObject}' " +
            $"ИЛИ [Тип] = '{RequestNsiReference.TypesKeys.СanceledEtalonObject}') " +
            $"И [Стадия] = '{StageGuids.ProcessingStageGuid}'";

        var selectObjectDialog = СоздатьДиалогВыбораОбъектов(RequestNsiReference.ReferenceId.ToString());
        selectObjectDialog.Фильтр = filter;
        selectObjectDialog.МножественныйВыбор = false;

        if (!selectObjectDialog.Показать())
            return;

        RequestNsiObject requestNsiObject = (ReferenceObject)selectObjectDialog.ФокусированныйОбъект;

        requestNsiObject.Main.Modify(obj =>
        {
            foreach (var etalon in selectedEtalons)
            {
                requestNsiObject.AddEtalonRecord(etalon);
            }
        }, null, true, false, true);

        Сообщение("Информация", "Операция завершена успешно");
    }

    /// <summary>
    /// АРМ Эксперта - Создание эталонов - Работа с Эталонами
    /// </summary>
    public void УстановитьСтадиюАннулировать_РабочаяСтраница()
        => SetStageCanceled(GetEtalonsFromReferenceControl());

    public void УстановитьСтадиюАннулировать_Справочник()
        => SetStageCanceled(Context.GetSelectedObjects());

    private void SetStageCanceled(ReferenceObject[] selectedEtalons)
    {
        if (selectedEtalons == null || selectedEtalons.Length == 0)
            return;

        if (selectedEtalons.Any(s => !s.IsCheckedOutByCurrentUser && !s.CanCheckOut))
        {
            Сообщение("Внимание!",
                "Невозможно осуществить операцию. Нет доступа на редактирование объектов");
            return;
        }

        if (!Вопрос("Вы уверены, что хотите аннулировать выбранные объекты ?"))
            return;

        if (selectedEtalons.All(s =>
        {
            if (s.SystemFields.Stage.Stage.Guid != StageGuids.ApprovedStageGuid)
                return false;

            var parameter = s.ParameterValues.FirstOrDefault(p =>
                p.ParameterInfo.Name == "Состояние" && p.ParameterInfo.Type == ParameterType.Int);

            return parameter != null && parameter.GetInt32() == 0; // Утвержден
        }))
        {
            foreach (var selectedEtalon in selectedEtalons)
            {
                Desktop.CheckOut(selectedEtalon, false);

                selectedEtalon.Modify(o =>
                {
                    var parameter = o.ParameterValues.First(p =>
                        p.ParameterInfo.Name == "Состояние" && p.ParameterInfo.Type == ParameterType.Int);

                    parameter.Value = 2; // Аннулирован
                }, true, false, true);

                // Desktop.CheckIn(selectedEtalon, "Смена состояния - Аннулирован", false);
            }

            Сообщение("Информация", "Операция завершена успешно");
        }
        else
        {
            Сообщение("Внимание!",
                "Невозможно осуществить операцию. Проверьте стадию и состояние объектов");
        }
    }

    /// <summary>
    /// АРМ Эксперта - Создание эталонов - Работа с Эталонами
    /// </summary>
    public void КорректировкаЗаписиЭталона_РабочаяСтраница()
        => ReferenceRecordAdjustment(GetEtalonsFromReferenceControl());

    public void КорректировкаЗаписиЭталона_Справочник()
        => ReferenceRecordAdjustment(Context.GetSelectedObjects());

    private void ReferenceRecordAdjustment(ReferenceObject[] selectedEtalons)
    {
        if (selectedEtalons == null || selectedEtalons.Length == 0)
            return;

        if (!Вопрос("Выбранные объекты будут взяты на корректировку?"))
            return;

        if (selectedEtalons.Any(s => !s.IsCheckedOutByCurrentUser && !s.CanCheckOut))
        {
            Сообщение("Внимание!",
                "Невозможно осуществить операцию. Нет доступа на редактирование объектов");
            return;
        }

        if (selectedEtalons.All(s => s.SystemFields.Stage.Stage.Guid == StageGuids.ApprovedStageGuid))
        {
            var stage = Stage.Find(Context.Connection, StageGuids.AdjustmentStageGuid);
            if (stage == null)
                throw new MacroException("Внимание! Не найдена стадия 'Корректировка'");

            var adjustmentObjects = stage.AutomaticChange(selectedEtalons);

            if (adjustmentObjects == null || adjustmentObjects.Count == 0)
                throw new MacroException("Внимание! Невозможно перейти на стадию 'Корректировка'");

            foreach (var selectedEtalon in adjustmentObjects)
            {
                Desktop.CheckOut(selectedEtalon, false);
            }

            if (selectedEtalons.Length == 1)
            {
                ПоказатьДиалогСвойств(Объект.CreateInstance(selectedEtalons[0], Context));
            }
            else
            {
                Сообщение("Информация", "Операция завершена успешно");
            }
        }
        else
        {
            Сообщение("Внимание!",
                "Невозможно осуществить операцию. Проверьте стадию объектов");
        }
    }

    //public void СформироватьДополнительныеПараметры()
    //{
    //    const string additionalParameters = "Дополнительные параметры";

    //    if (ТекущийОбъект == null)
    //        return;

    //    var descParametersNs = GetDescParametersNsi();
    //    if (descParametersNs.Count == 0)
    //        return;

    //    List<DescParameterNsiObject> newParameters;

    //    if (ТекущийОбъект.СвязанныеОбъекты[additionalParameters].Count == 0)
    //    {
    //        newParameters = descParametersNs;
    //    }
    //    else
    //    {
    //        newParameters = new List<DescParameterNsiObject>();

    //        var existingParameters = ТекущийОбъект.СвязанныеОбъекты[additionalParameters]
    //            .Select(p => (AdditionalParameterObject)(ReferenceObject)p).ToArray();

    //        newParameters.AddRange(descParametersNs.Where(descParameterNsi =>
    //            existingParameters.All(p => descParameterNsi.Guid != p.DocumentParameterNsiGuid)));
    //    }

    //    if (newParameters.Count == 0)
    //    {
    //        Сообщение("Внимание!", "Все дополнительные параметры уже подключены к текущему эталону");
    //        return;
    //    }

    //    var reference = AdditionalParametersReference.CreateReference(Context.Connection);

    //    foreach (var descParameterNsi in newParameters)
    //    {
    //        var newParameter = reference.CreateAdditionalParameterObject();
    //        newParameter.DocumentParameterNsiGuid.Value = descParameterNsi.Guid;
    //        newParameter.Name.Value = descParameterNsi.Name.Value;
    //        newParameter.DataType.Value = descParameterNsi.ParameterType.Value;

    //        var glossaryParameter = descParameterNsi.GlossaryParameter;

    //        if (glossaryParameter != null)
    //        {
    //            newParameter.GlossaryObject = glossaryParameter;
    //            newParameter.Units.Value = glossaryParameter.UnitMeasureType.Value;
    //        }

    //        newParameter.EndChanges();

    //        ТекущийОбъект.Подключить(additionalParameters, Объект.CreateInstance(newParameter, Context));
    //    }

    //    Сообщение("Информация", "Операция формирования дополнительных параметров завершена");
    //}

    //private List<DescParameterNsiObject> GetDescParametersNsi()
    //{
    //    var docObj = ТекущийОбъект.СвязанныйОбъект["Документ НСИ"];
    //    if (docObj == null)
    //        throw new MacroException("Не найден объект по связи 'Документ НСИ'");

    //    if (ТекущийОбъект.СвязанныеОбъекты["Дополнительные параметры"] == null)
    //        throw new MacroException($"У эталона '{ТекущийОбъект}' отсутствует связь 'Дополнительные параметры'");

    //    var documentNsi = (DocumentNsiObject)(ReferenceObject)docObj;
    //    var documentParameterNsi = documentNsi.DocumentParameterNsi;

    //    if (documentParameterNsi == null)
    //    {
    //        throw new MacroException(
    //            $"У документа НСИ '{documentNsi.Name}' отсутствует объект по связи 'Параметр документа НСИ'");
    //    }

    //    var parametersNsi = documentParameterNsi.ParametersNsi;

    //    var additionalParameters = parametersNsi
    //        .Where(p => p.Class.Guid == DescParametersNsiReference.TypesKeys.DescriptionParameterClass
    //                    && p[DocumentParameterNsiObject.FieldKeys.Additional].GetBoolean()).ToArray();

    //    if (additionalParameters.Length == 0)
    //    {
    //        Сообщение("Внимание!",
    //            $"У объекта '{documentParameterNsi.Name}' справочника 'Параметры документов НСИ' нет дополнительных параметров");
    //        return new List<DescParameterNsiObject>();
    //    }

    //    return additionalParameters.Select(p => (DescParameterNsiObject)p).ToList();
    //}

    private ReferenceObject[] GetEtalonsFromReferenceControl()
    {
        const string itemName = "EtalonsRecordsControl1";

        var supportSelectionVM = GetSupportSelection(itemName);
        if (supportSelectionVM.SelectedObjects == null || supportSelectionVM.SelectedObjects.Count == 0)
            return Array.Empty<ReferenceObject>();

        var refObject = new List<ReferenceObject>();

        foreach (var selectedObject in supportSelectionVM.SelectedObjects)
        {
            if (selectedObject is ReferenceObjectViewModel vm)
            {
                refObject.Add(vm.ReferenceObject);
            }
        }

        return refObject.ToArray();
    }

    private ISupportSelection GetSupportSelection(string itemName)
    {
        var uiMacroContext = Context as UIMacroContext;
        if (uiMacroContext?.OwnerViewModel == null)
            return null;

        var workingPage = GetWorkingPage(uiMacroContext.OwnerViewModel);
        if (workingPage == null)
            return null;

        IWindow window = workingPage;
        var foundItem = window.FindItem(itemName);

        var referenceWindowLayoutItemVM = foundItem as ReferenceWindowLayoutItemViewModel;

        return referenceWindowLayoutItemVM?.InnerViewModel as ISupportSelection;
    }

    private static WorkingPageViewModel GetWorkingPage(LayoutViewModel vm)
    {
        if (vm is WorkingPageViewModel workingPage)
            return workingPage;

        var owner = vm.Owner;

        while (owner != null)
        {
            if (owner is WorkingPageViewModel page)
                return page;

            owner = owner.Owner;
        }

        return null;
    }

    /// <summary>
    /// Производит проверки над объектом эталона, с указанным параметром документа НСИ
    /// Выдает исключение если были ошибки во время проверок
    /// </summary>
    /// <param name="referenceObject"></param>
    /// <param name="parameterObject"></param>
    private void ПроверитьПараметрУОбъекта(CheckObjectHelper checkObjectHelper, ReferenceObject parameterObject, ChecObjectResult result)
    {
        if (!parameterObject[DescriptionParametersNSI.Control].GetBoolean())
            return;

        Parameter objectParameter = null;
        switch (checkObjectHelper.CheckObjectType)
        {
            case CheckObjectHelper.ObjectType.Etalon:
                if (!TryGetEtalonParameter(checkObjectHelper.ReferenceObject, parameterObject, result, out objectParameter))
                    return;
                break;
            case CheckObjectHelper.ObjectType.TempObject:
                objectParameter = checkObjectHelper.ReferenceObject.ParameterValues.FirstOrDefault(p => p.ParameterInfo.Name == parameterObject.ToString());
                break;
            default:
                return;
        }

        if (objectParameter is null)
        {
            result.AddError($"Не найден параметр: {parameterObject}");
            return;
        }

        bool isMandatory = parameterObject[DescriptionParametersNSI.Mandatory].GetBoolean();
        if (isMandatory && (objectParameter.Value == null || objectParameter.Value == objectParameter.ParameterInfo.Type.DefaultValue))
        {
            result.AddError($"Параметр {parameterObject} не заполнен");
            return;
        }

        if (objectParameter.Value == null)
        {
            result.AddError($"Параметр {parameterObject} имеет значение null");
            return;
        }

        var objectParameterValueString = objectParameter.Value.ToString();
        ПроверитьПараметрНаСписокЗначений(objectParameterValueString, parameterObject, result);
        ПроверитьДиапазонЗначений(objectParameterValueString, parameterObject, result);
        ПроверитьЗначениеПоФормуле(checkObjectHelper, objectParameterValueString, parameterObject, result);
        ПроверитьРегулярноеВыражение(objectParameterValueString, parameterObject, result);
    }

    /// <summary>
    /// Получает параметр объекта эталона
    /// </summary>
    /// <param name="etalonReferenceObject">Объект эталона</param>
    /// <param name="parameterObject"></param>
    /// <param name="result">Для записи ошибок</param>
    /// <param name="objectParameter">Выходной параметр</param>
    /// <returns>возвращает true если удалось получить параметр, в противном случае false</returns>
    private bool TryGetEtalonParameter(ReferenceObject etalonReferenceObject, ReferenceObject parameterObject, ChecObjectResult result, out Parameter objectParameter)
    {
        //Если параметр дополнительный у него не указана связь с гуидом
        if (parameterObject[DescriptionParametersNSI.Additional].GetBoolean())
        {
            objectParameter = etalonReferenceObject.ParameterValues.FirstOrDefault(p => p.ParameterInfo.Name == parameterObject.ToString());
            if (objectParameter == null)
            {
                result.AddError($"Не найден дополнительный параметр: {parameterObject}");
                return false;
            }
        }
        else
        {
            if (!TryParsePapameter(parameterObject[DescriptionParametersNSI.EtalonParameter].GetString(), out var etalonParameterGuid))
            {
                objectParameter = null;
                result.AddError("Некорректно указан параметр");
                return false;
            }

            objectParameter = etalonReferenceObject[etalonParameterGuid];
        }

        return true;
    }

    private void ПроверитьРегулярноеВыражение(string etalonParameterValueString, ReferenceObject parameterObject, ChecObjectResult result)
    {
        var regexTemplate = parameterObject[DescriptionParametersNSI.RegexTemplate].GetString();
        if (String.IsNullOrWhiteSpace(regexTemplate))
            return;

        try
        {
            var regex = new Regex(regexTemplate);
            if (!regex.IsMatch(etalonParameterValueString))
                result.AddError($"Параметр '{parameterObject}' не проходит проверку по шаблону строки");
        }
        catch (Exception e)
        {
            result.AddError($"Параметр '{parameterObject}' ошибка при разборе шаблона строки:{Environment.NewLine}{e.Message}");
        }
    }

    private void ПроверитьПараметрНаСписокЗначений(string etalonParameterValueString, ReferenceObject parameterObject, ChecObjectResult result)
    {
        var objectValueList = parameterObject.GetObjects(DescriptionParametersNSI.ListValues);
        if (objectValueList.Count == 0)
            return;

        if (!objectValueList.Exists(ob => ob[DescriptionParametersNSI.ListValueParameters.Value].Value.ToString() == etalonParameterValueString))
            result.AddError($"Параметр '{parameterObject}' не попадает в список значений");
    }

    private void ПроверитьДиапазонЗначений(string etalonParameterValue, ReferenceObject parameterObject, ChecObjectResult result)
    {
        bool useMaxValue = parameterObject[DescriptionParametersNSI.UseMaxRange].GetBoolean();
        bool useMinValue = parameterObject[DescriptionParametersNSI.UseMinRange].GetBoolean();
        if (!useMaxValue && !useMinValue)
            return;

        if (!Double.TryParse(etalonParameterValue, out double value))
            result.AddError($"Неверно указан тип параметра: {parameterObject}");

        if (useMaxValue)
        {
            var maxValue = parameterObject[DescriptionParametersNSI.RangeMax].GetDouble();
            if (maxValue < value)
            {
                result.AddError($"Значение параметра: {parameterObject} не попадает в максимальный диапазон");
                return;
            }
        }

        if (useMinValue)
        {
            var minValue = parameterObject[DescriptionParametersNSI.RangeMin].GetDouble();
            if (minValue > value)
            {
                result.AddError($"Значение параметра: {parameterObject} не попадает в минимальный диапазон");
            }
        }
    }

    private void ПроверитьЗначениеПоФормуле(CheckObjectHelper checkObjectHelper, string etalonParameterValueString, ReferenceObject parameterObject, ChecObjectResult objectResult)
    {
        var formulaResult = parameterObject[DescriptionParametersNSI.Formula].GetString();
        if (String.IsNullOrWhiteSpace(formulaResult))
            return;

        if (!РасчитатьФормулуПоПараметрам(checkObjectHelper, ref formulaResult))
        {
            objectResult.AddError($"В параметре: '{parameterObject}' некорректно указанна формула");
            return;
        }
        formulaResult = ParseNumFormula(parameterObject, formulaResult);

        if (formulaResult != etalonParameterValueString)
            objectResult.AddError($"Значение параметра: '{parameterObject}' не совпадает со значением формулы");
    }

    private static string ParseNumFormula(ReferenceObject parameterObject, string formulaResult)
    {
        int parameterType = parameterObject[DescriptionParametersNSI.ParameterType].GetInt32();
        if (parameterType == 4 || parameterType == 3) //Целое число действительное число
        {
            formulaResult = formulaResult.Replace(" ", String.Empty);
            return MathParser.MathParser.Вычислить(formulaResult).ToString();
        }

        return formulaResult;
    }

    private bool РасчитатьФормулуПоПараметрам(CheckObjectHelper checkObjectHelper, ref string sourceFormula)
    {
        /// "Винт " + Диаметр резьбы + "-" + Длина винта + "-" - Стандарт
        foreach (var parameter in checkObjectHelper.Parameters)
        {
            string foundedStringParameter = String.Format("{0}{1}{2}", "{", parameter, "}");
            if (!sourceFormula.Contains(foundedStringParameter))
                continue;

            var objectParameter = checkObjectHelper.ReferenceObject.ParameterValues.FirstOrDefault(p => p.ParameterInfo.Name == parameter);
            if (parameter == null)
                continue;

            var parameterValue = objectParameter.Value;
            if (parameterValue == null)
                sourceFormula = sourceFormula.Replace(foundedStringParameter, "ParameterIsNull");
            else
                sourceFormula = sourceFormula.Replace(foundedStringParameter, parameterValue.ToString());
        }

        return CheckFormulaForSeparator(ref sourceFormula);
    }

    /// <summary>
    /// Проверяет формулу на разделители
    /// </summary>
    /// <param name="sourceFormula"></param>
    /// <returns></returns>
    private bool CheckFormulaForSeparator(ref string sourceFormula)
    {
        var startIndexs = AllIndexOf(sourceFormula, startSeparatorStr);
        var endIndexs = AllIndexOf(sourceFormula, endSeparatorStr);
        if (startIndexs.Count == 0 && endIndexs.Count == 0)
            return true;

        // Если количество разное, то ошибка в разборе формулы
        if (startIndexs.Count != endIndexs.Count)
            return false;

        bool process = true;
        while (process)
        {
            // Смотрим есть ли в формуле стартовый символ
            if (!sourceFormula.Contains(startSeparatorStr))
                return true;

            // Получаем строку с которой будем производить поиск
            var startIndex = sourceFormula.LastIndexOf('[');
            var str = sourceFormula.Substring(startIndex);
            // Если в обрабатываемой строке нет символа окончания
            if (!str.Contains(endSeparatorStr))
                return false;

            var lenght = str.IndexOf(']') + 1;
            // Получаем значение без 
            str = str.Substring(0, lenght); //Получаем {par2} - wef
            if (str.Contains("ParameterIsNull"))
            {
                sourceFormula = sourceFormula.Remove(startIndex, lenght);
            }
            else
            {
                sourceFormula = sourceFormula.Remove(startIndex + lenght - 1, 1);
                sourceFormula = sourceFormula.Remove(startIndex, 1);
            }
        }
        return true;
    }


    /// <summary>
    /// Находит все вхождения символов в строке
    /// </summary>
    /// <param name="source">Исходная строка</param> 
    /// <param name="substring">Искомая строка</param>
    /// <returns>Список индексов вхождения строки</returns>
    private List<int> AllIndexOf(string source, string substring)
    {
        var indices = new List<int>();

        int index = source.IndexOf(substring, 0);
        while (index > -1)
        {
            indices.Add(index);
            index = source.IndexOf(substring, index + substring.Length);
        }

        return indices;
    }

    private object CalculateFormula(ReferenceObject referenceObject, string formulaString) => new FormulaMacro(formulaString).Calculate(new MacroContext(referenceObject));

    private string ПолучитьБазовыйТипСправочника(Reference reference) => reference.Classes.BaseClasses[0].Name;

    /// <summary>
    /// Получает словарь параметров из параметров документа
    /// </summary>
    /// <param name="параметрыДокумента"></param>
    /// <param name="errors"></param>
    /// <param name="useKey">Только ключевые параметры</param>
    /// <returns>ключ, параметр временного справочника, значение параметр справочника аналога</returns>
    private List<ParameterData> ПолучитьСписокПараметров(List<Объект> параметрыДокумента, StringBuilder errors, Reference etalonReference, Reference tempReference)
    {
        List<ParameterData> parameterDataList = ПолучитьВсеСоответствияПараметров(параметрыДокумента, errors, etalonReference);

        etalonReference.Refresh();

        foreach (var parameterData in parameterDataList)
        {
            //Если связь то пропускаем, будем искать параметры
            if (parameterData.IsLink)
                continue;

            ParameterInfo tempParameterInfo = tempReference.ParameterGroup.Parameters.FindByName(parameterData.TempParameter);
            if (tempParameterInfo == null)
            {
                errors.Append(String.Format("Ошибка в описании параметра: {0}{1}Параметр не найден во временном справочнике{1}", parameterData.Name, Environment.NewLine));
                continue;
            }

            parameterData.TempParameterInfo = tempParameterInfo;
            parameterData.EtalonReference = etalonReference;
            if (parameterData.EtalonParameterInfo != null)
            {
                parameterData.EtalonParameter = parameterData.EtalonParameterInfo.Guid;
                continue;
            }

            var etalonParameterInfo = etalonReference.ParameterGroup.Parameters.Find(parameterData.EtalonParameter);
            if (etalonParameterInfo == null)
            {
                errors.Append(String.Format("Ошибка в описании параметра: {0}{1}Параметр не найден в справочнике эталона{1}", parameterData.Name, Environment.NewLine));
                continue;
            }

            parameterData.EtalonParameterInfo = etalonParameterInfo;
        }

        return parameterDataList;
    }

    private List<ParameterData> ПолучитьВсеСоответствияПараметров(List<Объект> параметрыДокумента, StringBuilder errors, Reference etalonReference, bool useKey = false)
    {
        var parameterDataList = new List<ParameterData>();

        foreach (var параметрДокумента in параметрыДокумента)
        {
            bool isKey = параметрДокумента["Ключ"];
            if (useKey && !isKey)
                continue;

            bool isLink = параметрДокумента.Тип == "Связь";
            string параметрВременногоСправочника = параметрДокумента["Наименование"];
            if (String.IsNullOrEmpty(параметрВременногоСправочника))
                continue;

            string параметрЭталона = параметрДокумента["Параметр эталона"];
            var parameterEtalonGuid = Guid.Empty;
            int типПараметра = параметрДокумента["Тип параметра"];

            var parameterData = new ParameterData();

            bool дополнительныйПараметр = параметрДокумента["Дополнительный"];
            if (дополнительныйПараметр)
            {
                var глоссарий = параметрДокумента.СвязанныйОбъект["Параметр глоссария"];
                if (глоссарий == null)
                {
                    errors.Append(String.Format("Ошибка в описании дополнительного параметра: {0}{1}Не указан глоссарий{1}", параметрВременногоСправочника, Environment.NewLine));
                    continue;
                }
                try
                {
                    var extendedParameterInfoAlias = СоздатьДополнительныйПараметрИПсевдоним(etalonReference, GetParameterType(типПараметра), глоссарий.ToString(), параметрВременногоСправочника);
                    if (extendedParameterInfoAlias == null)
                    {
                        errors.Append(String.Format("Ошибка в описании дополнительного параметра: {0}{1}", параметрВременногоСправочника, Environment.NewLine));
                        continue;
                    }

                    parameterData.EtalonParameterInfo = extendedParameterInfoAlias;
                }
                catch (Exception e)
                {
                    errors.Append(String.Format("Ошибка в описании дополнительного параметра: {0}{1}{2}{1}", параметрВременногоСправочника, Environment.NewLine, e.Message));
                    continue;
                }

            }
            else if (!TryParsePapameter(параметрЭталона, out parameterEtalonGuid))
            {
                if (isKey)
                    errors.Append(String.Format("Ошибка в описании ключевого параметра: {0}{1}", параметрВременногоСправочника, Environment.NewLine));

                continue;
            }

            parameterData.IsKey = isKey;
            parameterData.ParameterType = типПараметра;
            parameterData.TempParameter = параметрВременногоСправочника;
            parameterData.EtalonParameter = parameterEtalonGuid;
            parameterData.IsExtendedParameter = дополнительныйПараметр;

            if (isLink)
            {
                Guid referenceGuid = параметрДокумента["Справочник по связи"];//Справочник для связи
                if (referenceGuid == Guid.Empty)
                    continue;

                string параметрПоСвязи = параметрДокумента["Параметр поиска"];
                if (!TryParsePapameter(параметрПоСвязи, out var parameterToLinkGuid))
                    continue;

                var linkData = new LinkData()
                {
                    ReferenceGuid = referenceGuid,
                    LinkedParameter = parameterToLinkGuid
                };

                parameterData.IsLink = true;
                parameterData.LinkData = linkData;
            }

            parameterDataList.Add(parameterData);
        }
        etalonReference.ParameterGroup.ReferenceInfo.RefreshDescription();
        etalonReference.Refresh(false);

        return parameterDataList;
    }

    private ParameterInfo СоздатьДополнительныйПараметрИПсевдоним(Reference etalonReference, ParameterType parameterType, string extendedParameterName, string aliasParameterName)
    {
        if (!etalonReference.ParameterGroup.SupportsExtendedParameters)
            return null;

        var baseClass = etalonReference.Classes.BaseClasses[0];
        ExtendedParameterInfoBuilder extendedParameterInfoBuilder;
        var extendedParameter = ExtendedParameters.Find(p => p.Name == extendedParameterName && p.Type == parameterType);
        if (extendedParameter == null)
        {
            extendedParameterInfoBuilder = new ExtendedParameterInfoBuilder(Context.Connection, etalonReference.ParameterGroup, baseClass)
            {
                Name = extendedParameterName,
                Type = parameterType
            };
            extendedParameterInfoBuilder.Save();
            extendedParameter = extendedParameterInfoBuilder.ParameterInfo;
            ExtendedParameters.Add(extendedParameter);
        }
        else
        {
            extendedParameterInfoBuilder = new ExtendedParameterInfoBuilder(Context.Connection, etalonReference.ParameterGroup, extendedParameter);
        }

        if (etalonReference.ParameterGroup.Parameters.Find(extendedParameter.Id) == null)
        {
            extendedParameter = Context.Connection.ExtendedParameters.AttachExtendedParameterToParameterGroup(extendedParameter, etalonReference.ParameterGroup, baseClass);
            extendedParameterInfoBuilder = new ExtendedParameterInfoBuilder(Context.Connection, extendedParameter);
        }

        if (extendedParameterName == aliasParameterName)
            return extendedParameter;

        if (!extendedParameterInfoBuilder.CanEditAliases)
            return null;

        var alias = extendedParameterInfoBuilder.Aliases.FirstOrDefault(al => al.Alias == aliasParameterName && al.GroupId == etalonReference.ParameterGroup.Id);
        if (alias == null)
        {
            alias = extendedParameterInfoBuilder.AddAlias(extendedParameter, etalonReference.Id, baseClass.Id, aliasParameterName);
            extendedParameterInfoBuilder.Save();
        }

        return ExtendedParameterInfoBuilder.BuildAliasedParameterInfo(extendedParameter, alias);
    }

    /// <summary>
    /// Преобразует строковое представление параметра из ЭУ "Выбор параметра" в гуид
    /// </summary>
    /// <param name="fullParameterPath"></param>
    /// <param name="parameter"></param>
    /// <returns></returns>
    private static bool TryParsePapameter(string fullParameterPath, out Guid parameter)
    {
        parameter = Guid.Empty;
        if (fullParameterPath.Length == 77)
            if (Guid.TryParse(fullParameterPath.Substring(40, 36), out parameter))
                return true;

        return false;
    }

    private static string GetParameterName(int v) => GetParameterType(v).Name;

    private static ParameterType GetParameterType(int v) => ParameterType.GetType(v);

    private Reference FindReference(Guid referenceGuid)
        => Context.Connection.ReferenceCatalog.Find(referenceGuid)?.CreateReference() ?? throw new MacroException($"Не найден справочник с идентификатором {referenceGuid}");
    private Reference FindReference(string referenceName)
        => Context.Connection.ReferenceCatalog.Find(referenceName)?.CreateReference() ?? throw new MacroException($"Не найден справочник с наименованием {referenceName}");

    private List<Объект> ПолучитьПараметрыДокументаНСИ(Объект документНСИ)
    {
        var параметрДокументаНСИ = документНСИ.СвязанныйОбъект["9677778a-6a67-42ac-9608-19df7af56946"] ?? throw new MacroException("У документа НСИ не указан параметр");
        var параметрыДокумента = параметрДокументаНСИ.СвязанныеОбъекты["Описание параметров НСИ"].ToList();
        if (параметрыДокумента.Count == 0)
            throw new MacroException("У параметра документа НСИ не указанны параметры");

        return параметрыДокумента;
    }

    private Reference GetEtalonReference(Объект документНСИ)
    {
        var параметрДокументаНСИ = документНСИ.СвязанныйОбъект["9677778a-6a67-42ac-9608-19df7af56946"] ?? throw new MacroException("У документа НСИ не указан параметр");
        Guid гуидСправочникаЭталона = параметрДокументаНСИ["ae1c005a-92b0-42f0-8cb5-67aa9ce83856"];
        if (гуидСправочникаЭталона == Guid.Empty)
            throw new MacroException("У параметра документа НСИ не указан справочник эталона");

        var etalonReference = FindReference(гуидСправочникаЭталона);
        return etalonReference ?? throw new MacroException("Указан несуществующий справочник эталона");
    }

    /// <summary>
    /// Получает
    /// </summary>
    /// <param name="объект"></param>
    /// <param name="parameterDataList"></param>
    /// <returns></returns>
    private static string ПолучитьФильтрДляПоискаЭталона(Объект объект, List<ParameterData> parameterDataList)
    {
        var keyParameters = parameterDataList.Where(p => p.IsKey).ToList();
        if (keyParameters.Count == 0)
            return "[ID] > 0";

        var etalonReference = keyParameters[0].EtalonReference;
        var filter = new Filter(etalonReference.ParameterGroup);
        foreach (var parameterData in keyParameters)
        {
            filter.Terms.AddTerm(parameterData.EtalonParameterInfo, ComparisonOperator.Equal, объект[parameterData.TempParameter]);
        }

        return filter.ToString();
    }

    private ParameterInfo FindParameterInfo(ReferenceObject referenceObject, string v) => referenceObject.Reference.ParameterGroup.OneToOneParameters.FindByName(v);

    private static bool TryBeginChangesReferenceObject(ReferenceObject referenceObject)
    {
        if (referenceObject.CanEdit)
        {
            referenceObject.BeginChanges();
            return true;
        }

        return referenceObject.Changing;
    }

    /// <summary>
    /// Берет объект с историей изменений на редактирование
    /// </summary>
    /// <param name="referenceObject"></param>
    /// <returns></returns>
    private static bool TryBeginChanges(ReferenceObject referenceObject)
    {
        if (referenceObject.IsCheckedOut)
        {
            if (referenceObject.IsCheckedOutByCurrentUser)
            {
                return TryBeginChangesReferenceObject(referenceObject);
            }
        }
        else if (referenceObject.CanCheckOut)
        {
            return TryBeginChangesReferenceObject(referenceObject);
        }

        return TryBeginChangesReferenceObject(referenceObject);
    }

    #region classes

    public class CheckObjectHelper
    {
        public CheckObjectHelper(Объект объект, ObjectType checkObjectType)
        {
            Объект = объект ?? throw new ArgumentNullException(nameof(объект));
            CheckObjectType = checkObjectType;
        }

        public enum ObjectType
        {
            Etalon,
            TempObject
        }

        public ReferenceObject ReferenceObject => (ReferenceObject)Объект;

        public Объект Объект { get; }

        public ObjectType CheckObjectType { get; set; }

        public List<string> Parameters { get; set; }
    }
    public class CreateEtalonManager
    {
        /// <summary>
        /// Компонует 
        /// </summary>
        /// <param name="createEtalonManagers"></param>
        /// <returns></returns>
        public static string GetIdFilterString(List<CreateEtalonManager> createEtalonManagers)
        {
            var idList = createEtalonManagers.Select(c => c.EtalonObject.SystemFields.Id.ToString());
            return String.Format("[ID] входит в список '{0}'", String.Join(", ", idList));
        }

        public int IdObject => Объект["ID"];
        public Объект Объект { get; set; }

        public int IdEtalon => EtalonObject.SystemFields.Id;
        public ReferenceObject EtalonObject { get; set; }

        /// <summary>
        /// Признак подключать эталон
        /// </summary>
        public bool LinkEtalon { get; set; }
        /// <summary>
        /// Признак, обновлять эталон
        /// </summary>
        public bool UpdateEtalon { get; set; }
    }

    public class ParameterData
    {
        /// <summary>
        /// Наименование обрабатываемого параметра
        /// </summary>
        public string Name => TempParameter;
        /// <summary>
        /// Признак ключевой параметр
        /// </summary>
        public bool IsKey { get; set; }
        /// <summary>
        /// Наименование параметра во временном справочнике
        /// </summary>
        public string TempParameter { get; set; }
        /// <summary>
        /// Описание параметра во временном справочнике
        /// </summary>
        public ParameterInfo TempParameterInfo { get; set; }
        /// <summary>
        /// Строковое представление гуида параметра в эталонном справочнике
        /// </summary>
        public Guid EtalonParameter { get; set; }
        /// <summary>
        /// Описание параметра в эталонном справочнике
        /// </summary>
        public ParameterInfo EtalonParameterInfo { get; set; }
        public Reference EtalonReference { get; set; }
        /// <summary>
        /// Признак является дополнительным
        /// </summary>
        public bool IsExtendedParameter { get; set; }
        /// <summary>
        /// Признак является связью
        /// </summary>
        public bool IsLink { get; set; }
        /// <summary>
        /// Указывает тип значения параметра: строка, целое итд.
        /// </summary>
        public int ParameterType { get; set; }
        /// <summary>
        /// Параметры для связи
        /// </summary>
        public LinkData LinkData { get; set; }
    }
    public class LinkData
    {
        /// <summary>
        /// Справочник по связи
        /// </summary>
        public Guid ReferenceGuid { get; set; }
        /// <summary>
        /// Параметр по которому надо искать объект
        /// </summary>
        public Guid LinkedParameter { get; set; }
    }

    private class ChecObjectResult
    {
        private readonly StringBuilder _fullErrorMessage = new StringBuilder();

        public ChecObjectResult(ReferenceObject referenceObject) => SourseObject = referenceObject ?? throw new ArgumentNullException("");

        public ReferenceObject SourseObject { get; }

        public bool HasErrors { get; set; }

        public void AddError(string errorString)
        {
            if (!HasErrors)
                HasErrors = true;

            _fullErrorMessage.AppendLine(errorString);
        }

        public StringBuilder GetErrorMessage() => _fullErrorMessage;
    }

    private class StageGuids
    {
        /// <summary> Guid стадии "Утверждено" </summary>
        public static readonly Guid ApprovedStageGuid = new Guid("6246d014-e8da-434f-be92-75dddebff0a6");

        /// <summary> Guid стадии "Обработка" </summary>
        public static readonly Guid ProcessingStageGuid = new Guid("e294a1d1-bbc8-4230-b740-32d7b3e0a566");

        /// <summary> Guid стадии "Обработано" </summary>
        public static readonly string ProcessedStageGuid = "6e3af1a3-e9b7-4ab6-b045-ecaa6e763001";

        /// <summary> Guid стадии "Корректировка" </summary>
        public static readonly Guid AdjustmentStageGuid = new Guid("18df455a-0dc8-43a9-b256-c0fd6898df1b");

        /// <summary> Guid стадии "Подготовлено" </summary>
        public static readonly Guid PreparedStageGuid = new Guid("cd731353-9912-43a2-b569-e56c9be84ea9");
    }

    private class DescriptionParametersNSI
    {
        /// <summary>
        /// Тип параметра
        /// </summary>
        public static Guid ParameterType = new Guid("5f095520-bc54-4eb6-a6b8-b64c024a433a");

        /// <summary>
        /// Параметр эталон (строковые)
        /// </summary>
        public static Guid EtalonParameter = new Guid("ad6b0ccd-55c9-4bf6-babe-3c1a3f0cfaf8");

        /// <summary>
        /// Обязательный (Boolean)
        /// </summary>
        public static Guid Mandatory = new Guid("67833e7d-f188-4e3e-b658-eeacaaff7e77");

        /// <summary>
        /// Дополнительный
        /// </summary>
        public static Guid Additional = new Guid("b5b76838-15c2-4c07-929e-010576bf90cb");

        /// <summary>
        /// Контрольный (Boolean)
        /// </summary>
        public static Guid Control = new Guid("12be7c62-da05-4fb8-a16e-bd45e4561ec1");

        /// <summary>
        /// Использовать диапазон (Min) (Boolean)
        /// </summary>
        public static Guid UseMinRange = new Guid("b3b76d8a-e4bd-40cf-b0e1-cd02d455f0e2");

        /// <summary>
        /// Использовать диапазон (Max) (Boolean)
        /// </summary>
        public static Guid UseMaxRange = new Guid("6554e18f-7dde-40ab-8f4a-13c5f32b66b4");

        /// <summary>
        /// Диапазон значений (Max)
        /// </summary>
        public static Guid RangeMax = new Guid("bd75703a-47e0-4b06-9bae-730cf0c6ce5d");

        /// <summary>
        /// Диапазон значений (Min)
        /// </summary>
        public static Guid RangeMin = new Guid("5eccbd7d-9aa6-4538-b2ae-6deb672abe6f");

        /// <summary>
        /// Формула
        /// </summary>
        public static Guid Formula = new Guid("961e03ed-6fce-4b3d-91ed-94855037e89b");

        /// <summary>
        /// Формула
        /// </summary>
        public static Guid RegexTemplate = new Guid("25ca5612-e0b3-4b6b-9a28-08de3e19fa0c");
        //25ca5612-e0b3-4b6b-9a28-08de3e19fa0c

        /// <summary>
        /// Список значений (Список объектов)
        /// </summary>
        public static Guid ListValues = new Guid("6b463a86-bed1-4544-b054-a6f7a11cf1c9");

        public class ListValueParameters
        {
            //Основные параметры
            public static Guid Ikonka = new Guid("344277d8-f2d0-4d26-9dd6-0d4526c34635");
            public static Guid Name = new Guid("a46662b8-98c9-4234-bf8c-ecf7412b9631");
            public static Guid Value = new Guid("cecab80f-604e-4f86-bd3b-ac74bc2d93b6");
            public static Guid Description = new Guid("2de5e3a0-856f-4a2f-bb99-316dbf910a11");

            //string
            public static Guid String_Value = new Guid("49f82936-dde7-4144-b73e-e34e4150e18d");
            public static Guid MultilineString_value = new Guid("0c76a22c-3db7-4c32-af17-bf3ad51b7450");
            public static Guid HTML_Value = new Guid("da14c5b2-06ee-4f89-961c-b40f899bf11a");
            //number
            public static Guid Int_Value = new Guid("e45a6406-5e8d-4d5e-a24c-dd52142bbdb4");
            public static Guid Double_Value = new Guid("97fcdc92-ec0a-4da4-822f-0cdaa5609b32");

            public static Guid Guid_Value = new Guid("935ee344-a885-4b76-a2fe-91fae1b3781c");
            public static Guid Boolean_Value = new Guid("0b66e5f3-a9de-4ab4-88d7-aa720e0ea605");
            public static Guid DateTime_Value = new Guid("3577a2da-2fc3-43ec-bf8b-0753b3245f56");
        }
    }
    #endregion
}
