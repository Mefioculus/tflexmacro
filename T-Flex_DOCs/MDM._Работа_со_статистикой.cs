using System;
using System.Diagnostics;
using System.Linq;
using TFlex.DOCs.Model.Macros;
using TFlex.DOCs.Model.Macros.ObjectModel;
using TFlex.DOCs.Model.References;

public class MacroStatistic : MacroProvider
{
    public MacroStatistic(MacroContext context)
        : base(context)
    {
        if (Context.Connection.ClientView.HostName == "MOSINS")
            if (Вопрос("Хотите запустить в режиме отладки?"))
            {
                System.Diagnostics.Debugger.Launch();
                System.Diagnostics.Debugger.Break();
            }
    }

    private static readonly Guid СтадияРазработка = new Guid("527f5234-4c94-43d1-a38d-d3d7fd5d15af");
    private static readonly Guid СтадияУтверждено = new Guid("a5ea2e1c-d441-42fd-8f92-49840351d6c1");
    private static readonly Guid СтадияДобавлено = new Guid("b253355a-29ea-4db3-9e5c-2ffe8c27a0ee");
    private static readonly Guid СтадияНормализацияДанных = new Guid("02b8bbcd-e24d-4aac-8853-ba93cbbec5f0");
    private static readonly Guid СтадияОбработано = new Guid("6e3af1a3-e9b7-4ab6-b045-ecaa6e763001");
    private static readonly Guid СтадияКорректировка = new Guid("18df455a-0dc8-43a9-b256-c0fd6898df1b");
    private static readonly Guid СтадияПодготовлено = new Guid("cd731353-9912-43a2-b569-e56c9be84ea9");

    public override void Run()
    {
    }

    public void ЗаполнитьСтатистикуДокументНСИ()
    {
        Объект выбранныйОбъект = ТекущийОбъект;
        if (!Guid.TryParse(выбранныйОбъект["Guid сгенерированного справочника"], out Guid гуидСгенерированногоСправочника))
            Ошибка("Для выбранного документа не создан справочник");

        var etalonObjects = ПолучитьОбъектыАналогаИлиЭталона(гуидСгенерированногоСправочника).ToList();
        var объектСтатистика = СоздатьИПодключитьСтатистику(выбранныйОбъект, "Статистика Документа НСИ", "Документ НСИ", "Текущий Документ НСИ");
        try
        {
            var statistic = new StatisticData()
            {
                Create = etalonObjects.OrderBy(ob => ob.SystemFields.CreationDate).LastOrDefault().SystemFields.CreationDate,
                Edit = etalonObjects.OrderBy(ob => ob.SystemFields.EditDate).LastOrDefault().SystemFields.EditDate,
                CountWork = etalonObjects.FindAll(ob => ob.SystemFields.Stage?.Guid == СтадияРазработка || ob.SystemFields.Stage?.Guid == СтадияКорректировка).Count,
                Count = etalonObjects.Count(),
                CountComplete = etalonObjects.FindAll(ob => ob.SystemFields.Stage?.Guid == СтадияОбработано).Count,
                CountRaw = etalonObjects.FindAll(ob => ob.SystemFields.Stage?.Guid == СтадияПодготовлено).Count
            };

            ЗаполнитьпараметрыСтатистики(объектСтатистика, statistic);
        }
        catch (Exception e)
        {
            Ошибка(e.Message);
        }
    }

    public void ЗаполнитьСтатистикуИсточникДанных()
    {
        var выбранныйОбъект = ТекущийОбъект;
        ReferenceObjectCollection analodObjects =
            ПолучитьОбъектыАналогаИлиЭталона(выбранныйОбъект["bd2f577a-874f-4376-9e2e-e4b06e9f3123"]);

        var объектСтатистика = СоздатьИПодключитьСтатистику(выбранныйОбъект, "Статистика Источника данных", "Источник данных", "Источник данных текущий");
        try
        {
            var statistic = new StatisticData(analodObjects, СтадияНормализацияДанных, СтадияОбработано, СтадияДобавлено);

            ЗаполнитьпараметрыСтатистики(объектСтатистика, statistic);
        }
        catch (Exception e)
        {
            Ошибка(e.Message);
        }
    }

    public void ЗаполнитьСтатистикуТипНСИ()
    {
        Объект выбранныйОбъект = ТекущийОбъект;
        var etalonObjects = ПолучитьОбъектыАналогаИлиЭталона(выбранныйОбъект["Справочник эталона"]);
        var объектСтатистика = СоздатьИПодключитьСтатистику(выбранныйОбъект, "Статистика Типа НСИ", "Тип НСИ", "Тип НСИ текущий");

        try
        {
            var statistic = new StatisticData(etalonObjects, СтадияРазработка, СтадияУтверждено);
            ЗаполнитьпараметрыСтатистики(объектСтатистика, statistic);
        }
        catch (Exception e)
        {
            Ошибка(e.Message);
        }
    }

    private ReferenceObjectCollection ПолучитьОбъектыАналогаИлиЭталона(Guid гуидСправочникаАналога)
    {
        if (гуидСправочникаАналога == Guid.Empty)
            Ошибка("Не указан справочник");

        var analodReference = FindReference(гуидСправочникаАналога);
        if (analodReference == null)
            Ошибка("Не найден справочник");

        if (!analodReference.ParameterGroup.SupportsStages)
            Ошибка("В справочнике не включена поддержка стадий");

        var analodObjects = analodReference.Objects;
        if (analodObjects.Count() == 0)
            Ошибка("Не найдены объекты в справочнике");

        return analodObjects;
    }

    private static void ЗаполнитьпараметрыСтатистики(Объект объектСтатистика, StatisticData statistic)
    {
        объектСтатистика.Изменить();
        объектСтатистика["7df18271-8d01-4f55-af12-109516c95085"] = statistic.Create;//Дата последнего добавления эталона
        объектСтатистика["5bedf396-0eef-4aa2-94d9-2e95daa6a221"] = statistic.Edit;//Дата последнего изменения данных
        объектСтатистика["6f5fcf49-d55c-47f2-b30c-8c64d221c5a8"] = statistic.CountWork;//Количество записей в работе
        объектСтатистика["5a4385c4-84bb-4d9b-908b-997fb4e6a5b9"] = statistic.Count;//Количество записей всего
        объектСтатистика["bccb5344-783d-45a7-b56f-e35b7d41b188"] = statistic.CountComplete;//Количество обработанных записей
        объектСтатистика["724f6a08-6580-48ef-8798-82539337aceb"] = statistic.CountRaw;//Количество необработанных записей
        объектСтатистика.Сохранить();
    }

    private Объект СоздатьИПодключитьСтатистику(Объект выбранныйОбъект, string тип, string связьВсяСтатистика, string связьТекущаяСтатистика)
    {
        Объект статистика = СоздатьОбъект("Статистика НСИ", тип);
        статистика["Наименование"] = "Статистика: " + выбранныйОбъект.ToString();
        статистика.СвязанныйОбъект[связьТекущаяСтатистика] = выбранныйОбъект;
        статистика.СвязанныйОбъект[связьВсяСтатистика] = выбранныйОбъект;
        статистика.Сохранить();
        return статистика;
    }

    private Reference FindReference(Guid referenceGuid)
    {
        var parameterInfo = Context.Connection.ReferenceCatalog.Find(referenceGuid);
        if (parameterInfo == null)
            return null;

        return parameterInfo.CreateReference();
    }

    public class StatisticData
    {
        public StatisticData() { }

        public StatisticData(ReferenceObjectCollection objects, Guid workStage, Guid completeStage)
        {
            Create = objects.OrderBy(ob => ob.SystemFields.CreationDate).LastOrDefault().SystemFields.CreationDate;
            Edit = objects.OrderBy(ob => ob.SystemFields.EditDate).LastOrDefault().SystemFields.EditDate;
            CountWork = objects.Where(ob => ob.SystemFields.Stage?.Guid == workStage).Count();
            Count = objects.Count();
            CountComplete = objects.Where(ob => ob.SystemFields.Stage?.Guid == completeStage).Count();
        }

        public StatisticData(ReferenceObjectCollection objects, Guid workStage, Guid completeStage, Guid rawStage)
            : this(objects, workStage, completeStage)
        {
            CountRaw = objects.Where(ob => ob.SystemFields.Stage?.Guid == rawStage).Count();
        }

        public DateTime Create;
        public DateTime Edit;
        public int CountWork;
        public int Count;
        public int CountComplete;
        public int CountRaw = 0;
    }
}
