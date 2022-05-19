using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using TFlex.DOCs.Model.Macros;
using TFlex.DOCs.Model.Macros.ObjectModel;
using TFlex.DOCs.Model.References;

public class MacroMDMReestr : MacroProvider
{
    private StringBuilder _errors = new StringBuilder();

    public MacroMDMReestr(MacroContext context)
        : base(context)
    {
        if (Context.Connection.ClientView.HostName == "MOSINS")
            if (Вопрос("Хотите запустить в режиме отладки?"))
            {

                Debugger.Launch();
                Debugger.Break();
            }
    }

    private new Объекты ВыбранныеОбъекты
    {
        get
        {
            return ВыполнитьМакрос("ea0d8d3c-395b-48d5-9d42-dab98371522c", "ПолучитьВыбраныеОбъекты", "Записи");
        }
    }

    public override void Run()
    {
        Объекты выбранныеОбъекты = ВыбранныеОбъекты;
        if (выбранныеОбъекты.Count == 0)
            return;

        var filtredObjects = выбранныеОбъекты.Where(ob => (Guid)ob["Guid записи аналога"] != Guid.Empty).ToList();
        if (filtredObjects.Count == 0)
            Ошибка("У выбранных объектов не заполнены записи аналога");

        var errorsCountObjects = выбранныеОбъекты.Count - filtredObjects.Count;
        if (errorsCountObjects != 0)
        {
            WriteLog(String.Format("{0} выбранных записей не указана запись аналога", errorsCountObjects));
        }

        //Будут обрабатываться после основной обработки
        List<ReestData> отложенныеОбъекты = new List<ReestData>();

        var groupAnalogReference = filtredObjects.GroupBy(ob => (Guid)ob["Guid справочника аналога"]);
        foreach (var group in groupAnalogReference)
        {
            var reestData = new ReestData(this);

            if (group.Key == Guid.Empty)
            {
                WriteLog(String.Format("В {0} объектах, не указан гуид справочника аналога", group.Count()));
                continue;
            }

            var analogReference = FindReference(group.Key);
            if (analogReference == null)
            {
                WriteLog(String.Format("Не найден спраовчник аналог с гуидом: {0}", group.Key));
                continue;
            }

            var allAnalogObject = analogReference.Objects;

            var источникДанных = НайтиОбъект("Источники данных", "Справочник-аналог", group.Key);
            if (источникДанных == null)
            {
                WriteLog(String.Format("Не найден источник данных для справочника аналога с гуидом: {0}", group.Key));
                continue;
            }

            foreach (Объект временнаяЗапись in group)
            {
                var эталон = временнаяЗапись.СвязанныйОбъект["Эталон"];
                if (эталон == null)
                {
                    WriteLog(String.Format("Не найдена запись эталона для объекта {0}", временнаяЗапись));
                    continue;
                }

                if (эталон["Стадия"] != "Утверждено")
                {
                    WriteLog(String.Format("Запись эталона для объекта {0} не находится на стадии 'Утверждено'", временнаяЗапись));
                    continue;
                }

                Guid analogObjectGuid = временнаяЗапись["Guid записи аналога"];
                if (analogObjectGuid == Guid.Empty)
                    continue;

                var analogObject = allAnalogObject.FirstOrDefault(ob => ob.SystemFields.Guid == analogObjectGuid);
                if (analogObject == null)
                {
                    WriteLog(String.Format("Не найдена запись аналога для объекта {0}", временнаяЗапись));
                    continue;
                }

                reestData.ВнешнийИдентификатор = analogObject.SystemFields.Guid.ToString();
                reestData.ИсточникДанных = источникДанных["Guid"];
                reestData.ОбъектАналога = analogObjectGuid;
                reestData.ОбъектЭталона = эталон["Guid"];
                reestData.СправочникЭталона = эталон.Справочник.Guid;
                reestData.СправочникАналога = group.Key;

                string reestrFilter =
                    String.Format("[Источник данных] = '{0}' И [Внешний идентификатор источника данных] = '{1}'",
                    reestData.ИсточникДанных, reestData.ВнешнийИдентификатор);

                reestData.Дубль = НайтиОбъект("Реестр соответствия эталону",
            String.Format("[Источник данных] = '{0}' И [Объект эталона] = '{1}'", reestData.ИсточникДанных, reestData.ОбъектЭталона)) != null;

                reestData.Аналог = Объект.CreateInstance(analogObject, Context);
                reestData.Эталон = эталон;
                reestData.ВременныйОбъект = временнаяЗапись;

                var записьРеестр = НайтиОбъект("Реестр соответствия эталону", reestrFilter);
                if (записьРеестр != null)
                {
                    отложенныеОбъекты.Add(reestData);
                    continue;
                }
                reestData.СоздатьЗаписьРегистра();
                reestData.ИзменитьАналог();
                reestData.Аналог.ИзменитьСтадию("Обработано", true);
                reestData.ИзменитьСтадиюВременногоСправочника();
            }
        }

        if (отложенныеОбъекты.Count != 0)
        {
            if (Вопрос("Для нескольких объектов уже есть записи реестра, все равно хотите создать записи?"))
            {
                foreach (var reestData in отложенныеОбъекты)
                {
                    reestData.СоздатьЗаписьРегистра();
                    reestData.ИзменитьАналог();
                    reestData.Аналог.ИзменитьСтадию("Обработано", true);
                    reestData.ИзменитьСтадиюВременногоСправочника();
                }
            }
        }

        if (_errors.Length > 0)
            Сообщение("Предепреждение", _errors.ToString());
    }

    public void СоздатьРеестрЧерезНоменклатуру()
    {
        Объект аналог = ТекущийОбъект;
        if (аналог == null)
            return;

        if (аналог["Стадия"] != "Обработано")
            return;

        Объект ЭСИ = аналог.СвязанныйОбъект["Номенклатура и изделия"];
        if (ЭСИ == null)
            return;

        var источникДанных = НайтиОбъект("Источники данных", "Справочник-аналог", ЭСИ.Справочник.Guid);
        if (источникДанных == null)
            return;
        //Ошибка(String.Format("Не найден источник данных для справочника аналога с гуидом: {0}", гуидСправочникАналог));

        var reestData = new ReestData(this);

        reestData.ВнешнийИдентификатор = ЭСИ["Guid"].ToString();
        reestData.ИсточникДанных = источникДанных["Guid"];
        reestData.ОбъектАналога = ЭСИ["Guid"];
        reestData.ОбъектЭталона = аналог["GUID эталона"];
        reestData.СправочникЭталона = аналог["Справочник эталона"];
        reestData.СправочникАналога = ЭСИ.Справочник.Guid;

        string reestrFilter =
                    String.Format("[Источник данных] = '{0}' И [Внешний идентификатор источника данных] = '{1}'",
                    reestData.ИсточникДанных, reestData.ВнешнийИдентификатор);

        var записьРеестр = НайтиОбъект("Реестр соответствия эталону", reestrFilter);
        if (записьРеестр == null)
        {
            записьРеестр = СоздатьОбъект("Реестр соответствия эталону", "Запись реестра");
        }
        else
        {
            try
            //f0128a49-b6d8-406e-814b-eacd83cf5182
            {
                записьРеестр.Изменить();
            }
            catch (Exception e)
            {
                return;
                //Ошибка(e.Message);
            }
        }
        reestData.ИзменитьПараметрыРегистра(записьРеестр);
    }

    private Reference FindReference(Guid referenceGuid)
    {
        var referenceInfo = Context.Connection.ReferenceCatalog.Find(referenceGuid);
        if (referenceInfo == null)
            return null;
        return referenceInfo.CreateReference();
    }

    private void WriteLog(string message)
    {
        _errors.Append(message + Environment.NewLine + "--------------" + Environment.NewLine);
    }

    public class ReestData
    {
        public MacroMDMReestr Context;
        public ReestData(MacroMDMReestr context)
        {
            Context = context;
        }

        public string ВнешнийИдентификатор { get; set; }
        public bool Дубль = false;
        public Guid ИсточникДанных { get; set; }
        public Guid ОбъектАналога { get; set; }
        public Guid ОбъектЭталона { get; set; }
        public Guid СправочникАналога { get; set; }
        public Guid СправочникЭталона { get; set; }

        public Объект ВременныйОбъект { get; set; }
        public Объект Аналог { get; set; }
        public Объект Эталон { get; set; }


        public void СоздатьЗаписьРегистра()
        {
            Объект записьРеестр = Context.СоздатьОбъект("Реестр соответствия эталону", "Запись реестра");
            ИзменитьПараметрыРегистра(записьРеестр);
        }

        public void ИзменитьПараметрыРегистра(Объект записьРеестр)
        {
            записьРеестр["Внешний идентификатор источника данных"] = ВнешнийИдентификатор;
            записьРеестр["Источник данных"] = ИсточникДанных;
            записьРеестр["Объект эталона"] = ОбъектЭталона;
            записьРеестр["Объект аналога"] = ОбъектАналога;
            записьРеестр["Справочник аналог"] = СправочникАналога;
            записьРеестр["Справочник эталона"] = СправочникЭталона;
            записьРеестр.Сохранить();
        }

        public void ИзменитьАналог()
        {
            if (Аналог == null || Эталон == null)
                return;

            try
            {
                Аналог.Изменить();
                Аналог["Guid эталона"] = ОбъектЭталона;
                Аналог["Дубль"] = Дубль;
                Аналог["Состояние эталона"] = Эталон["Состояние"];
                Аналог["Справочник эталона"] = Эталон.Справочник.Guid;
                Аналог.Сохранить();
                //Аналог.ИзменитьСтадию("Обработано");
            }
            catch (Exception e)
            {
                Context.WriteLog(String.Format("Ошибка при изменении аналога: {0}{1}{2}", Аналог, Environment.NewLine, e.Message));
            }
        }

        internal void ИзменитьСтадиюВременногоСправочника()
        {
            ВременныйОбъект.ИзменитьСтадию("Обработано");
        }
    }
}
