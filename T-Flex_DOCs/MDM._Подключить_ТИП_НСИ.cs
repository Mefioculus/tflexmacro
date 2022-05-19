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

public class MacroMDMSelectTypeNSI : MacroProvider
{
    public MacroMDMSelectTypeNSI(MacroContext context)
        : base(context)
    {
        if (Context.Connection.ClientView.HostName == "MOSINS")
            if (Вопрос("Хотите запустить в режиме отладки?"))
            {

                Debugger.Launch();
                Debugger.Break();
            }
    }

    public override void Run()
    {
        StringBuilder errors = new StringBuilder();
        foreach (var объект in ВыбранныеОбъекты)
        {
            int типНоменклатуры = объект["3c7a075f-0b53-4d68-8242-9f76ca7b2e97"];
            var объектТипНСИ = ПолучитьГуидОбъектаНСИ(типНоменклатуры);
            if (объектТипНСИ == Guid.Empty)
                continue;

            try
            {
                объект.Изменить();
                объект["4b415737-1575-4eec-80b2-d466d47ce85d"] = объектТипНСИ;
                объект.Сохранить();
            }
            catch
            {
                errors.Append(String.Format("Ошибка при изменении объекта: {0} возможно объект заблокированн другим пользователем", объект));
            }
        }
    }

    private Guid ПолучитьГуидОбъектаНСИ(int typeNom)
    {
        switch (typeNom)
        {
            case 1://Сборочная единица
            case 4://Изделие
            case 5://Деталь
                return new Guid("b244663a-433e-436f-aea4-89bac5fff8d1"); // Объект продукция ДСЕ
            case 2://Стандартное изделие
                return new Guid("06c996ee-afb0-4279-bb43-5e812abf48f2");// Объект Стандартное изделие
            case 3://Прочие изделия
            case 6://Электронные компоненты
                return new Guid("a4b3c959-70c4-4429-b5f6-aebcec5fda8b");// Объект Покупные и комплектующие изделия
            case 7://Материалы
                return new Guid("759925fc-9270-42ed-8d02-0c48b15a6255");// Объект Материалы		
            default:
                return Guid.Empty;
        }
    }

    public void ПодключитьТипИДокументНСИ()
    {
        Объекты выбранныеОбъекты = ВыполнитьМакрос("ea0d8d3c-395b-48d5-9d42-dab98371522c", "ПолучитьВыбраныеОбъекты", "item13");
        if (выбранныеОбъекты.Count == 0)
            return;

        var группы = выбранныеОбъекты.GroupBy(ob => ob["Тип НСИ"]);
        if (группы.Count() > 1)
            Ошибка("Невозможно выполнить операцию для объектов с разными Типами НСИ. Выберите группу объектов с одинаковым Типом НСИ");

        Объект выбранныйОбъект = выбранныеОбъекты.FirstOrDefault();

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

        Guid гуидДокументаНСИ = ПолучитьДокументНСИ(гуидТипаНСИ, фильтрДокументовНСИ);
        if (гуидДокументаНСИ == Guid.Empty)
            return;

        if (!Вопрос("Параметры: Документ НСИ, Тип НСИ будут перезаписаны." + Environment.NewLine + "Продолжить выполнение?"))
            return;

        StringBuilder errors = new StringBuilder();
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
            Сообщение("Предепреждение", errors.ToString());
    }

    private Guid ПолучитьДокументНСИ(Guid гуидТипаНСИ, string фильтрДокументаНСИ)
    {
        var диалог = СоздатьДиалогВыбораОбъектов("Документы НСИ");
        диалог.Фильтр = фильтрДокументаНСИ;
        диалог.ПоказатьПанельКнопок = false;
        диалог.МножественныйВыбор = false;
        if (диалог.Показать())
            return диалог.ФокусированныйОбъект["Guid"];
        return Guid.Empty;
    }

    /// <summary>
    /// Находит тип НСИ для указанног справочника показывает диалог с выбором типа НСИ
    /// </summary>
    /// <param name="referenceGuid"></param>
    /// <returns>Вохзвращает гуид типа НСИ+</returns>
    public Guid ПолучитьТипНСИ(Guid referenceGuid)
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
        if (диалог.Показать())
            return диалог.ФокусированныйОбъект["Guid"];

        return Guid.Empty;
    }
}
