using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Forms = System.Windows.Forms;
using TFlex.DOCs.Model.Desktop;
using TFlex.DOCs.Model.Macros;
using TFlex.DOCs.Model.Macros.ObjectModel;
using TFlex.DOCs.Model.References;
// using TFlex.DOCs.Model.References.Files; // Для работы с файловым справочником
// using TFlex.DOCs.Model.Classes; // Для работы с классами
using TFlex.DOCs.Model.Stages; // Для работы со стадиями

using TFlex.DOCs.Model.FilePreview.CADExchange;
using TFlex.DOCs.Model.FilePreview.CADService;

using TFlex.DOCs.Model.References.Reporting;
using TFlex.DOCs.Common;

//using TFlex.DOCs.UI.Utils.Helpers;


/*
Для работы макроса так же потрубеются дополнительные ссылки на внешние библиотеки
TFlex.DOCs.UI.Utils.dll
TFlex.DOCs.Common.dll
TFlex.Reporting.dll
*/


public class Macro : MacroProvider
{
    string directoryWithPDF;
    string[] listOfFilesOnProcessing;
    Reference reportReference;
    Reference archiveReference;
    Report reportMI;
    Report reportFMI;
    
    #region Guids

    private static class Guids {
        public static class Reports {
            public static Guid Свидетельство = new Guid("a9e67d98-ce6c-4ad4-8c94-328e71089099");
        }

        public static class Types {
            public static Guid ТипPDF = new Guid("58e7a26a-cf5f-445b-b08a-885f1bcf7f12");
            public static Guid ПротоколМеталлографическойЛаборатории = new Guid("80debbb1-d94d-4926-a7d8-da6e22cd277f");
            public static Guid ПротоколФизикомеханическойЛаборатории = new Guid("9233e1f0-6584-49a5-9ffb-cedabbc579af");
            public static Guid ПротоколСпектральнойЛаборатории = new Guid("f674fa70-afe0-4be5-aa63-4729c5676dbc");
            public static Guid ПротоколХимическойЛаборатории = new Guid("0a98f565-a37d-4115-b7dc-547c051cd3e6");
            public static Guid ПротоколГальваническойЛаборатории = new Guid("08f5dc3f-7760-4098-8e20-20e12011ac27");
            public static Guid ПротоколМагнитнойЛаборатории = new Guid("3a2d970e-99b5-4cf7-8813-7c86609aaf6b");
            public static Guid ПротоколЭлектрическойЛаборатории = new Guid("c7ae898a-5e8a-4891-9dc2-9bd5710ae9fb");
            
        }

        public static class References {
            public static Guid АрхивЦЗЛ = new Guid("ca66f59f-e077-4ae4-91fe-645dd47c2300");
            public static Guid Отчеты = new Guid("d3396686-2cb9-44ff-994b-d446c0a42515");
        }

        public static class ListsOfObjects {
            public static Guid Образцы = new Guid("97d5d377-48ee-41d0-b6fc-ab43b7abdd6e");
            public static Guid Показатели = new Guid("7d2616a0-73c2-4f4a-933b-10c0bcc1f37a");
            public static Guid КонтролируемыеПараметры = new Guid("93b19910-3553-4f36-9dd7-f81e42e2739d");
            public static Guid ФактическиеПоказания = new Guid("a6f224ab-6044-472b-8e32-5780d2148e39");
        }

        public static class Props {
            public static Guid НомерПротокола = new Guid("d662eed7-c2a2-41fc-9b35-05e527349cc7");
            public static Guid ДатаПротокола = new Guid("3ba03308-123b-4840-bc78-bc0fdcb1de4f");
            public static Guid СводноеНаименованиеПротокола = new Guid("7b4b4de4-70b0-4a1c-83ce-20d2adf9f4f6");
            // Параметры для передачи данных в отчет
            public static Guid ПараметрВидТаблицы = new Guid("5ee7e25b-e56a-4f62-abb0-f092fc0bdb27");
            //
            // Параметры типа Образец
            // Для расчета сводного размера
            public static Guid ФормаОбразца = new Guid("03b50c57-5856-4761-807e-8bf07ef6c50f");
            public static Guid Длина = new Guid("6b5e98e3-81a5-4ec8-908a-ff73c8cad31d");
            public static Guid Ширина = new Guid("b391cbe3-0012-4f52-9c58-76e6e90fe30b");
            public static Guid Толщина = new Guid("5d3e8b14-bc35-408d-bcc2-53ec81d51d1e");
            public static Guid Диаметр = new Guid("50c7f6c1-e8da-496a-9327-f32bbbbee29f");
            public static Guid СводныйРазмер = new Guid("f91d8f9c-0d56-414f-b452-1032cc3cd0cd");
            // Для передачи в отчет
            public static Guid Испытание = new Guid("4cfe5d7a-93ad-4db2-a9d4-ab2fda219a36");
            public static Guid ПределПрочности = new Guid("d205d64d-ba0e-43e9-b7b3-52b250a5c8f8");
            public static Guid ПределТекучести = new Guid("20fb0b1a-be6d-4c0d-b78f-f8fc7cf0fe9c");
            public static Guid ОтносительноеУдлинение = new Guid("86bc650f-101e-47c2-9686-2ef8f5ccacee");
            public static Guid ПрочностьПриИзгибе = new Guid("885eb736-578a-427e-924e-044bafea03f6");

            // Параметры типа Показатель
            public static Guid Наименование = new Guid("759d291e-5a1c-4cb0-bc6b-6794cf1f3cb2");
            public static Guid ДанныеПоТУ = new Guid("6d25bb62-1975-4bae-82ae-20bfee8ecaf9");
            public static Guid ФактическиеДанные = new Guid("534a9e41-174b-46b9-9ba6-a12d336897a4");

            // Параметры типа Контролируемый параметр
            public static Guid КонтролируемыйПараметр = new Guid("a24c5b02-374f-40aa-be54-f5103f46ef77");
            public static Guid ПунктТребованийНТД = new Guid("971dbda8-a091-4e8f-ad82-4dcc37f3fcbb");
            public static Guid Норма = new Guid("39a0b103-2964-4e64-b8d1-754b2f477c5d");
            
            // Параметры типа Фактическое значение
            public static Guid НаименованиеИзмеряемогоОбъекта = new Guid("8e1c6e2e-d885-4b43-8dd7-6d5eba283783");
            public static Guid ФактическоеЗначение = new Guid("878df0b8-f018-41e6-a5a4-5f673dedf4ad");
        }

        public static class Stages {
            public static Guid Разработка = new Guid("527f5234-4c94-43d1-a38d-d3d7fd5d15af");
            public static Guid Согласование = new Guid("c89e6ee8-f060-4091-b576-00025be0846a");
            public static Guid Корректировка = new Guid("18df455a-0dc8-43a9-b256-c0fd6898df1b");
            public static Guid Хранение = new Guid("9826e84a-a0a7-404b-bf0e-d61b902e346a");
        }
    }

    #endregion Guids

    #region Properties

    Reference ReportReferenceInstance {
        get {
            if (reportReference == null) {
                reportReference = Context.Connection.ReferenceCatalog.Find(Guids.References.Отчеты).CreateReference();
            }
            return reportReference;
        }
    }

    Reference ArchiveReferenceInstance {
        get {
            if (archiveReference == null) {
                archiveReference = Context.Connection.ReferenceCatalog.Find(Guids.References.АрхивЦЗЛ).CreateReference();
            }
            return archiveReference;
        }
    }

    #endregion Properties

    public Macro(MacroContext context)
        : base(context)
    {
    }

    public override void Run()
    {
    }
    
    #region Формирование отчета из БП по окончанию согласования
    
    public void СформироватьОтчетДляБП(Объекты объекты) {

        foreach (Объект вложение in объекты) {
            ReferenceObject attachment = (ReferenceObject)вложение;
            
            if (attachment == null) {
                Message("Ошибка", "Не получилось обработать вложение");
                continue;
            }

            // Получаем контекст формирования отчета
            ReportGenerationContext reportContext = new ReportGenerationContext(attachment, null);
            reportContext.OpenFile = false;

            Report report = ReportReferenceInstance.Find(Guids.Reports.Свидетельство) as Report;
            if (report == null)
                Error("Не удалось найти объект отчета в справочнике 'Отчеты'. Обратитесь к системному администратору");

            // Перед началом формирования отчета переводим протокол на стадию корректировка для того, чтобы сформированный отчет можно было прикрепить к
            // обрабатываемому объекту

            ChangeStage(attachment, Guids.Stages.Корректировка);

            report.Generate(reportContext);

            // Автоматическое применение изменений файла отчета
            Desktop.CheckIn(reportContext.ReportFileObject, string.Format("Автоматическое формирование отчета для свидетельства №{0} архива ЦЗЛ", attachment[Guids.Props.СводноеНаименованиеПротокола].Value.ToString()), false);
        }
    }

    private void ChangeStage(ReferenceObject refObj, Guid guidOfStage) {
        // Метод для изменения стадии объекта
        
        Stage stage = Stage.Find(Context.Connection, guidOfStage);

        if (stage == null) {
            Error("При изменении стадии объекта возникла ошибка. Указанная стадия не найдена");
        }

        stage.Change(new List<ReferenceObject>() {refObj}); // Охренеть логика. Я думал записимость идет от объекта, а они идет от стадии
    }

    #endregion Формирование отчета из БП по окончанию согласования

    #region Формирование отчета для предварительного просмотра

    public void СформироватьОтчетДляПредварительногоПросмотра() {
        // Получаем текущий объект
        ReferenceObject currentObject = Context.ReferenceObject;
        if (currentObject == null) {
            Message("Ошибка", "Не получилось обратиться к текущему объекту для формирования отчета");
            return;
        }

        // Получаем контекст формирования отчета
        ReportGenerationContext reportContext = new ReportGenerationContext(currentObject, null);
        reportContext.OpenFile = true;

        // Получаем объект отчета
        Report report = ReportReferenceInstance.Find(Guids.Reports.Свидетельство) as Report;
        if (report == null)
            Error("Не удалось найти объект отчета в справочнике 'Отчеты'. Обратитесь к системному администратору");
        report.Generate(reportContext);

        Desktop.CheckIn(reportContext.ReportFileObject, "Предварительный просмотр", false);
    }

    #endregion Формирование отчета для предварительного просмотра

    #region генерация сводного наименования протокола

    // Данный метод используется по событию, при создании объекта в справочнике Архив ЦЗЛ
    public void GenerateSummaryName() {
        // Данный метод на основа порядкового номера протокола, его даты, а так же типа протокола будет
        // формировать сводное наименование протокола вида Мет-2-02/20

        // Данные, из которых будет формироваться сводное наименование
        int orderNumber;
        DateTime date;
        string typeOfRecord;

        // Получаем текущий объект
        ReferenceObject currentObject = Context.ReferenceObject;

        // Получаем у текущего объекта его порядковый номер, дату, а так же его тип
        orderNumber = (int)currentObject[Guids.Props.НомерПротокола].Value;
        date = (DateTime)currentObject[Guids.Props.ДатаПротокола].Value;
        typeOfRecord = GetStringOfType(currentObject);

        //currentObject.BeginChanges();
        currentObject[Guids.Props.СводноеНаименованиеПротокола].Value = string.Format("{0}-{1}-{2}/{3}", typeOfRecord, orderNumber, date.Month, date.Year);
        //currentObject.EndChanges();
        
    }

    private string GetStringOfType(ReferenceObject referenceObject) {
        // Получаем тип переданного объекта и возвращаем его сокращенное название
        if (referenceObject.Class.IsInherit(Guids.Types.ПротоколМеталлографическойЛаборатории))
            return "МЕТ";
        else if (referenceObject.Class.IsInherit(Guids.Types.ПротоколФизикомеханическойЛаборатории))
            return "ФИЗ";
        else if (referenceObject.Class.IsInherit(Guids.Types.ПротоколСпектральнойЛаборатории))
            return "СП";
        else if (referenceObject.Class.IsInherit(Guids.Types.ПротоколХимическойЛаборатории))
            return "ХИМ";
        else if (referenceObject.Class.IsInherit(Guids.Types.ПротоколГальваническойЛаборатории))
            return "ГАЛ";
        else if (referenceObject.Class.IsInherit(Guids.Types.ПротоколМагнитнойЛаборатории))
            return "МАГ";
        else if (referenceObject.Class.IsInherit(Guids.Types.ПротоколЭлектрическойЛаборатории))
            return "ЭЛ";
        else {
            Message("Ошибка", "Тип данного объекта не входит в список предусмотренных для формирования сводного наименования");
            return "ОШИБКА";
        }
    }

    #endregion генерация сводного наименования протокола

    #region Формирование сводного размера образца

    // Данный метод используется по событию (при любом изменении параметров объета образец)
    public void ФормированиеСводногоРазмераОбразца () {
        ReferenceObject sample = Context.ReferenceObject;
        
        const int бесформенный = 0;
        const int плоский = 1;
        const int прямоугольный = 2;
        const int круглый = 3;
        // Получаем у текущего объекта параметры размеров и вида формы
        int formType = (int)sample[Guids.Props.ФормаОбразца].Value;
        double area = 0.0;
        double length = 0.0;
        double width = 0.0;
        double thickness = 0.0;
        double diameter = 0.0;

        switch (formType) {
            case плоский:
                // Получаем значение длины и ширины, остальные значения обнуляем
                length = (double)sample[Guids.Props.Длина].Value;
                width = (double)sample[Guids.Props.Ширина].Value;

                if ((length != 0.0) && (width != 0.0))
                    area = length * width;
                
                sample[Guids.Props.СводныйРазмер].Value = string.Format("{0}Х{1} (S={2})", length, width, area);
                // Обнуляем остальные значения
                sample[Guids.Props.Толщина].Value = 0.0;
                sample[Guids.Props.Диаметр].Value = 0.0;
                break;
            case прямоугольный:
                length = (double)sample[Guids.Props.Длина].Value;
                width = (double)sample[Guids.Props.Ширина].Value;
                thickness = (double)sample[Guids.Props.Толщина].Value;

                if ((length != 0.0) && (width != 0.0))
                    area = length * width;

                sample[Guids.Props.СводныйРазмер].Value = string.Format("{0}Х{1}Х{2} (S={3})", length, width, thickness, area);
                // Обнуляем остальные значения
                sample[Guids.Props.Диаметр].Value = 0.0;
                break;
            case круглый:
                diameter = (double)sample[Guids.Props.Диаметр].Value;
                
                if (diameter != 0.0) {
                    double radius = diameter / 2;
                    area = System.Math.PI * radius * radius;
                }
                sample[Guids.Props.СводныйРазмер].Value = string.Format("d{0} (S={1})", diameter, area);
                // Обнуляем остальные значения
                sample[Guids.Props.Длина].Value = 0.0;
                sample[Guids.Props.Ширина].Value = 0.0;
                sample[Guids.Props.Толщина].Value = 0.0;
                break;
            case бесформенный:
                sample[Guids.Props.СводныйРазмер].Value = string.Empty;
                sample[Guids.Props.Длина].Value = 0.0;
                sample[Guids.Props.Ширина].Value = 0.0;
                sample[Guids.Props.Толщина].Value = 0.0;
                sample[Guids.Props.Диаметр].Value = 0.0;
                break;
            default:
                Error("При попытке сформировать сводный размер возникла ошибка. Выбран неизвестный тип формы образца");
                break;
        }
    }

    #endregion Формирование сводного размера образца

    #region Получение данных для отчета

    // Данный метод создан для использлования в "Выполнить макрос", который выполняется при формировании отчета
    // (вернее, при попытке получить параметр DataTable)
    // (для получения всех данных в вложенных в объект списках в виде единой строки)
    public string ПолучитьДанныеДляОтчета() {
        // Получаем текущий объект
        ReferenceObject record = Context.ReferenceObject;

        const int безТаблицы = 0;
        const int таблицаПоОбразцам = 1;
        const int таблицаПоПоказателям = 2;
        const int таблицаПоКонтролируемымПараметрам = 3;

        string result = string.Empty;
        // Класс для генерации строки, которая будет содержать табличные данные
        DataClass resultDataClass = new DataClass();

        switch ((int)record[Guids.Props.ПараметрВидТаблицы].Value) {
            case безТаблицы:
                break;
            case таблицаПоОбразцам:
                // код для получения выходных данным с вложенного списка "Образцы"
                resultDataClass.Add("Испытание");
                resultDataClass.Add("Сводный размер, мм");
                resultDataClass.Add("Предел прочности, МПа");
                resultDataClass.Add("Предел текучести, МПа");
                resultDataClass.Add("Относительное удлинение, %");
                resultDataClass.Add("Прочность при изгибе, Дж/см\u00B2");
                resultDataClass.EndRow();

                foreach (ReferenceObject sample in record.GetObjects(Guids.ListsOfObjects.Образцы)) {
                    resultDataClass.Add(sample[Guids.Props.Испытание].Value.ToString());
                    resultDataClass.Add(sample[Guids.Props.СводныйРазмер].Value.ToString());
                    resultDataClass.Add(sample[Guids.Props.ПределПрочности].Value.ToString());
                    resultDataClass.Add(sample[Guids.Props.ПределТекучести].Value.ToString());
                    resultDataClass.Add(sample[Guids.Props.ОтносительноеУдлинение].Value.ToString());
                    resultDataClass.Add(sample[Guids.Props.ПрочностьПриИзгибе].Value.ToString());
                    resultDataClass.EndRow();
                }

                result = resultDataClass.GenerateString();
                break;
            case таблицаПоПоказателям:
                // Код для получения данных с вложенного списка "Показатели"
                resultDataClass.Add("Наименование");
                resultDataClass.Add("Данные по ТУ");
                resultDataClass.Add("Фактические данные");
                resultDataClass.EndRow();

                foreach (ReferenceObject param in record.GetObjects(Guids.ListsOfObjects.Показатели)) {
                    resultDataClass.Add(param[Guids.Props.Наименование].Value.ToString());
                    resultDataClass.Add(param[Guids.Props.ДанныеПоТУ].Value.ToString());
                    resultDataClass.Add(param[Guids.Props.ФактическиеДанные].Value.ToString());
                    resultDataClass.EndRow();
                }
                result = resultDataClass.GenerateString();
                break;
            case таблицаПоКонтролируемымПараметрам:
                // код для получения данных с вложенного списка "Контролируемые параметры"
                resultDataClass.Add("Контролируемый параметр");
                resultDataClass.Add("Пункт требований НТД");
                resultDataClass.Add("Норма");
                resultDataClass.Add("Фактические показатели");
                resultDataClass.EndRow();
                
                bool isFirstValue;
                foreach (ReferenceObject controlParam in record.GetObjects(Guids.ListsOfObjects.КонтролируемыеПараметры)) {
                    isFirstValue = true;
                    foreach (ReferenceObject factVal in controlParam.GetObjects(Guids.ListsOfObjects.ФактическиеПоказания)) {
                        if (isFirstValue) {
                            resultDataClass.Add(controlParam[Guids.Props.КонтролируемыйПараметр].Value.ToString());
                            resultDataClass.Add(controlParam[Guids.Props.ПунктТребованийНТД].Value.ToString());
                            resultDataClass.Add(controlParam[Guids.Props.Норма].Value.ToString());
                            resultDataClass.Add(string.Format("{0}: {1}",
                                                            factVal[Guids.Props.НаименованиеИзмеряемогоОбъекта].Value.ToString(),
                                                            factVal[Guids.Props.ФактическоеЗначение].Value.ToString()));
                        }
                        else {
                            resultDataClass.Add("---");
                            resultDataClass.Add("---");
                            resultDataClass.Add("---");
                            resultDataClass.Add(string.Format("{0}: {1}",
                                                            factVal[Guids.Props.НаименованиеИзмеряемогоОбъекта].Value.ToString(),
                                                            factVal[Guids.Props.ФактическоеЗначение].Value.ToString()));
                        }
                        isFirstValue = false;
                        resultDataClass.EndRow();
                    }
                }
                result = resultDataClass.GenerateString();
                break;
            default:
                System.Windows.Forms.MessageBox.Show("При попытке передать в отчет табличные данные возникла ошибка. Выбран неизвестный вид таблицы");
                break;
        }

        return result;
    }

    private class DataClass {
        private string valueSplitter = "^";
        private string rowSplitter = ";";
        private List<string> intermediateResult = new List<string>();
        private List<string> result = new List<string>();

        public DataClass() {
        }

        public DataClass(string valueSplitter, string rowSplitter) {
            this.valueSplitter = valueSplitter;
            this.rowSplitter = rowSplitter;
        }

        public void Add(string value) {
            this.intermediateResult.Add(value);
        }

        public void EndRow() {
            result.Add(string.Join(this.valueSplitter, this.intermediateResult));
            this.intermediateResult.Clear();
        }

        public string GenerateString() {
            string result = string.Join(this.rowSplitter, this.result);
            this.result.Clear();
            return result;
        }
    }

    #endregion Получение данных для отчета

    #region Шаблонное заполнение данных

    // Код, который будет заполнять новые объекты в зависимости от того, какой протокол создается

    #endregion Шаблонное заполнение данных

    #region Присвоение дефолтного вида таблицы в зависимости от типа протокола

    // Данный код будет выполняться при любом создании новых объектов в справочнике Архив ЦЗЛ
    public void НазначитьВидТаблицы() {
        ReferenceObject currentObject = Context.ReferenceObject;
        // Вычисляем тип объекта
        if (currentObject.Class.IsInherit(Guids.Types.ПротоколМеталлографическойЛаборатории)) {
            currentObject[Guids.Props.ПараметрВидТаблицы].Value = 0;
        }
        else if (currentObject.Class.IsInherit(Guids.Types.ПротоколФизикомеханическойЛаборатории)) {
            currentObject[Guids.Props.ПараметрВидТаблицы].Value = 1;
        }
        else if (currentObject.Class.IsInherit(Guids.Types.ПротоколСпектральнойЛаборатории)) {
            currentObject[Guids.Props.ПараметрВидТаблицы].Value = 0;
        }
        else if (currentObject.Class.IsInherit(Guids.Types.ПротоколХимическойЛаборатории)) {
            currentObject[Guids.Props.ПараметрВидТаблицы].Value = 0;
        }
        else if (currentObject.Class.IsInherit(Guids.Types.ПротоколГальваническойЛаборатории)) {
            currentObject[Guids.Props.ПараметрВидТаблицы].Value = 0;
        }
        else if (currentObject.Class.IsInherit(Guids.Types.ПротоколМагнитнойЛаборатории)) {
            currentObject[Guids.Props.ПараметрВидТаблицы].Value = 3;
        }
        else if (currentObject.Class.IsInherit(Guids.Types.ПротоколЭлектрическойЛаборатории)) {
            currentObject[Guids.Props.ПараметрВидТаблицы].Value = 0;
        }
        else {
            Error(string.Format("Не получилось определить вариант таблицы для объекта '{0}'", currentObject.ToString()));
        }
    }

    #endregion Присвоение дефолтного вида таблицы в зависимости от типа протокола

}
