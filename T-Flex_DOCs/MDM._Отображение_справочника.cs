/*
TFlex.DOCs.UI.Client.dll
*/

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using TFlex.DOCs.Client.ViewModels.Layout;
using TFlex.DOCs.Client.ViewModels.References;
using TFlex.DOCs.Model.Macros;
using TFlex.DOCs.Model.Macros.ObjectModel;
using TFlex.DOCs.Model.References;


public class AsuNsiARMStatic : MacroProvider
{
    /// <summary> Guid макроса - 'MDM. АРМ. Взять объект ячейку' </summary>
    private const string TakeCellObjectMacroGuid = "ea0d8d3c-395b-48d5-9d42-dab98371522c";

    /// <summary> Guid стадии "Утверждено" </summary>
    private static readonly Guid ApprovedStageGuid = new Guid("6246d014-e8da-434f-be92-75dddebff0a6");

    /// <summary>
    /// Документ НСИ - Связь Разделы классификатора НСИ
    /// </summary>
    private static readonly string ClassifierSections = "68af68e1-93ab-4c23-bf25-fd7e4677db34";

    private static Guid _справочникАналог = Guid.Empty;

    private static string _меткаАналог = String.Empty;
    private static string _справочникЭксперта = String.Empty;
    private static string _объектДокументНСИ = String.Empty;
    private static string _справочникВременный = String.Empty;
    private static string _меткаВременный = String.Empty;
    private static string _меткаСправочникаЭксперта = String.Empty;

    public AsuNsiARMStatic(MacroContext context)
        : base(context)
    {
    }

    public override void Run()
    {
    }

    //ВыполнитьМакрос("d0bf469b-6ba3-45cc-84b5-d201ad619d4a","ПолучитьСправочникАналог");
    public string ПолучитьСправочникАналог()
        => _справочникАналог.ToString();

    //ВыполнитьМакрос("d0bf469b-6ba3-45cc-84b5-d201ad619d4a","ПолучитьМеткаАналог");
    public string ПолучитьМеткаАналог()
        => _меткаАналог;

    //ВыполнитьМакрос("d0bf469b-6ba3-45cc-84b5-d201ad619d4a","ПолучитьСправочникВременный");
    public string ПолучитьСправочникВременный()
        => _справочникВременный;

    //ВыполнитьМакрос("d0bf469b-6ba3-45cc-84b5-d201ad619d4a","ПолучитьМеткаВременный");
    public string ПолучитьМеткаВременный()
        => _меткаВременный;

    //ВыполнитьМакрос("d0bf469b-6ba3-45cc-84b5-d201ad619d4a","ПолучитьСвойстваДокументНСИ");
    public string ПолучитьСвойстваДокументНСИ()
        => _объектДокументНСИ;

    public string ПолучитьСправочникЭксперта()
        => _справочникЭксперта;

    public string ПолучитьМеткуСправочникаЭксперта()
        => _меткаСправочникаЭксперта;

    public void ЗадатьЭксперта()
    {
        var диалог = СоздатьДиалогВыбораОбъектов("00bf7ef0-6080-4edd-a548-95b44df465c4"); //Типы данных 
        диалог.МножественныйВыбор = false;
        диалог.ПоказатьПанельКнопок = false;
        диалог.Вид = "Типы НСИ";
        if (!диалог.Показать())
            return;

        _справочникЭксперта =
            диалог.ФокусированныйОбъект["9067725a-a9a7-48c0-b530-7bc0c02f6758"].ToString(); //Справочник эталона

        _меткаСправочникаЭксперта = диалог.ФокусированныйОбъект.ToString();
        ОбновитьЭлементыУправления();
    }

    public void ЗадатьАналог()
    {
        var диалог = СоздатьДиалогВыбораОбъектов("7e7be674-84c8-45d2-bed4-f58376cddd50"); //Источники данных
        диалог.МножественныйВыбор = false;
        диалог.ПоказатьПанельКнопок = false;
        if (!диалог.Показать())
            return;

        _справочникАналог = диалог.ФокусированныйОбъект["bd2f577a-874f-4376-9e2e-e4b06e9f3123"]; //Справочник аналог
        _меткаАналог = диалог.ФокусированныйОбъект.ToString();
        ОбновитьЭлементыУправления();
    }

    public void ЗадатьВременный()
    {
        var диалог = СоздатьДиалогВыбораОбъектов("a169916f-fa02-417b-b52f-63de54b06a59"); //Документы НСИ
        диалог.МножественныйВыбор = false;
        диалог.ПоказатьПанельКнопок = false;
        if (!диалог.Показать())
            return;

        _справочникВременный =
            диалог.ФокусированныйОбъект["79d316f0-0f13-4ee9-9316-e41f27389333"]; //Guid сгенерированного справочника
        _меткаВременный = диалог.ФокусированныйОбъект.ToString();
        ПоказатьДиалогСвойствНаРабочейСтранице(диалог.ФокусированныйОбъект, "item99");
        ОбновитьЭлементыУправления();
    }
    
    public void ПоказатьДиалогСвойствНаРабочейСтранице(Объект объект, string itemName)
    {
        var window = Context.GetCurrentWindow();

        if (!(window?.FindItem(itemName) is PropertyWindowLayoutItemViewModel propertyWindowItem))
            return;

        propertyWindowItem.SetContentObject((ReferenceObject)объект);
    }
    
    /// <summary>
    /// АРМ Эксперта - Создание эталонов - Работа с эталонами - Утвердить созданные эталоны
    /// </summary>
    public void ИзменитьСтадиюОбъектов_РабочаяСтраница()
    {
        //ВыполнитьМакрос("MDM. АРМ. Взять объект ячейку")
        Объекты выбранныеОбъекты = ВыполнитьМакрос(TakeCellObjectMacroGuid, "ПолучитьВыбраныеОбъекты",
            "EtalonsRecordsControl1");

        if (выбранныеОбъекты.Count == 0)
            throw new MacroException("Не найден элемент управления: EtalonsRecordsControl1");

        ChangeStageObjects(выбранныеОбъекты);
    }

    /// <summary>
    /// АРМ Эксперта - Создание эталонов - Работа с документами НСИ - Утвердить созданные эталоны
    /// </summary>
    public void ИзменитьСтадиюОбъектов_РабочаяСтраницаАналоги()
    {
        Объекты выбранныеОбъекты = ВыполнитьМакрос(TakeCellObjectMacroGuid, "ПолучитьВыбраныеОбъекты", "Записи");
        if (выбранныеОбъекты.Count == 0)
            throw new MacroException("Нет выбранных объектов");

        var эталоны = new Объекты();
        foreach (var объект in выбранныеОбъекты)
        {
            var эталон = объект.СвязанныйОбъект["Эталон"];
            if (эталон != null)
                эталоны.Add(эталон);
        }

        if (эталоны.Count == 0)
            throw new MacroException("У выбранных объектов не заданы эталоны");

        ChangeStageObjects(эталоны);
    }


    /// <summary>
    /// Кнопка утвердить на любом эталонном справочнике
    /// </summary>
    public void ИзменитьСтадиюОбъектов_Справочник() => ChangeStageObjects(ВыбранныеОбъекты);

    private void ChangeStageObjects(Объекты выбранныеОбъекты)
    {
        if (выбранныеОбъекты.Count == 0)
            return;

        var errors = new StringBuilder();

        //Словарь ошибок в котором хранятся описания ошибочных объектов 
        var errorsTypeList = CreateDictionary();

        var неПрошедныеПроверку = new List<Объект>();
        int countNotApproved = 0;
        foreach (var эталон in выбранныеОбъекты)
        {
            try
            {
                string checkEtalonResult = ВыполнитьМакрос("75b95647-4432-457f-aaf9-44afae6575c3", "ПроверитьЭталонМакрос", эталон);
                if (!String.IsNullOrEmpty(checkEtalonResult))
                {
                    неПрошедныеПроверку.Add(эталон);
                    continue;
                }

                if (!УтвердитьЭталон(errorsTypeList, эталон))
                    countNotApproved++;
            }
            catch (Exception e)
            {
                errors.Append(e.Message + Environment.NewLine);
            }
        }
        if (неПрошедныеПроверку.Count > 0)
        {
            if (Вопрос($"Эталоны: {String.Join(", ", неПрошедныеПроверку)} не прошли проверку, хотите их утвердить?"))
            {
                foreach (var эталон in неПрошедныеПроверку)
                {
                    try
                    {
                        if (!УтвердитьЭталон(errorsTypeList, эталон))
                            countNotApproved++;
                    }
                    catch (Exception e)
                    {
                        errors.Append(e.Message + Environment.NewLine);
                    }
                }
            }
        }

        errors.Insert(0, AddAllErrors(errorsTypeList));
        if (countNotApproved > 0)
            errors.Insert(0, $"Утверждено: {выбранныеОбъекты.Count - countNotApproved}, не утверждено: {countNotApproved}{Environment.NewLine}");

        if (errors.Length == 0)
            Сообщение("Сообщение", "Операция успешно выполнена");
        else
            Сообщение("Предупреждение", errors.ToString());
    }

    private static bool УтвердитьЭталон(Dictionary<int, List<string>> errorsTypeList, Объект эталон)
    {
        var документНСИ = эталон.СвязанныйОбъект["Документ НСИ"];
        if (эталон.ВзятНаРедактирование)
        {
            errorsTypeList[1].Add(эталон.ToString()); //Объект не сохранен на сервер
            return false;
        }
        else if (документНСИ == null)
        {
            errorsTypeList[0].Add(эталон.ToString()); //У объекта отсутствует Документ НСИ
            return false;
        }
        else if (!эталон.ИзменитьСтадию(ApprovedStageGuid.ToString())) // Стадия Утверждено
        {
            errorsTypeList[2].Add(эталон.ToString()); //Объект невозможно перевести из за схемы
            return false;
        }

        var классификаторы = документНСИ.СвязанныеОбъекты[ClassifierSections];
        if (классификаторы.Count == 0)
            return true;

        try
        {
            эталон.Изменить();
        }
        catch
        {
            return true;
        }

        foreach (var классификатор in классификаторы)
        {
            эталон.Подключить("Разделы Классификатора НСИ", классификатор);
        }

        эталон.Сохранить();
        return true;
    }

    private string AddAllErrors(Dictionary<int, List<string>> errorsTypeList)
    {
        string errors = String.Empty;

        foreach (var keyValue in errorsTypeList.Where(keyValue => keyValue.Value.Count != 0))
            errors += GetErrorText(keyValue);

        return errors;
    }

    private static string GetErrorText(KeyValuePair<int, List<string>> keyValue)
        => keyValue.Key switch
        {
            0 => String.Format("Следующие объекты: {0} не были утверждены по причине отсутствия связи с Документом НСИ.{1}", String.Join(", ", keyValue.Value), Environment.NewLine),
            1 => String.Format("Следующие объекты: {0} не были утверждены. Необходимо сохранить изменения.{1}", String.Join(", ", keyValue.Value), Environment.NewLine),
            2 => String.Format("Следующие объекты: {0} не были утверждены.{1}Проверьте возможность смены стадий.", String.Join(", ", keyValue.Value), Environment.NewLine),
            _ => String.Empty,
        };

    private Dictionary<int, List<string>> CreateDictionary()
        => new Dictionary<int, List<string>>
        {
            {0, new List<string>()}, //"Отсутствует документ НСИ"
            {1, new List<string>()}, //Объект не сохранен на сервер
            {2, new List<string>()} //Ошибка при переходе стадий
        };
}

