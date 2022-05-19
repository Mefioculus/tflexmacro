/* Дополнительные файлы
MdmNsi.dll
ClosedXML.dll
FastMember.Signed.dll
ExcelNumberFormat.dll
DocumentFormat.OpenXml.dll
*/

using System;
using System.Text;
using TFlex.DOCs.Macro.NormativeReferenceInfo.Handlers.DocumentParameterNsi;
using TFlex.DOCs.Macro.NormativeReferenceInfo.Handlers.DocumentParameterNsi.DescriptionParametersNsi;
using TFlex.DOCs.Macro.NormativeReferenceInfo.Handlers.DocumentParameterNsi.DescriptionParametersNsi.ListOfValuesNsi;
using TFlex.DOCs.Macro.NormativeReferenceInfo.Handlers.DocumentsNsi;
using TFlex.DOCs.Macro.NormativeReferenceInfo.Handlers.NewNsiReference;
using TFlex.DOCs.Macro.NormativeReferenceInfo.Handlers.ParameterTemplatesNsi;
using TFlex.DOCs.Model.Macros;
using TFlex.DOCs.Model.Macros.ObjectModel;
using TFlex.DOCs.Model.References;


namespace Macroses
{
    /// <summary>
    /// Макрос для обработки событий справочников (MDM и НСИ)
    /// </summary>
    public class MdmNsiEventHandlersMacro : MacroProvider
    {
        /// <summary> Guid макроса - 'MDM. АРМ. Взять объект ячейку' </summary>
        private const string TakeCellObjectMacroGuid = "ea0d8d3c-395b-48d5-9d42-dab98371522c";

        public MdmNsiEventHandlersMacro(MacroContext context)
            : base(context)
        {
        }

        #region 'Документы НСИ' - Генератор новых справочников НСИ

        public void ДокументыНСИ_СоздатьСправочник()
        {
            new GeneratorNewNsiReferenceCommand(this).СоздатьСправочник();
        }

        public void ДокументыНСИ_СоздатьСправочник_БП(Объекты объекты)
        {
            if (объекты == null)
                throw new ArgumentNullException("объекты");

            foreach (var объект in объекты)
            {
                var referenceObject = (ReferenceObject) объект;
                new GeneratorNewNsiReferenceCommand(this, referenceObject, true).СоздатьСправочник();
            }
        }

        #endregion

        #region 'Документы НСИ' - Генератор объектов для нового справочника НСИ из Excel файла

        /// <summary>
        /// Заполнить новый справочник из файла по связи 'Файл шаблон справочника' объекта 'Документ НСИ'
        /// </summary>
        public void ДокументыНСИ_ЗаполнитьСправочник()
        {
            new ObjectGeneratorFromExcelFileCommand(this).ЗаполнитьСправочник();
        }

        /// <summary>
        /// Заполнить новый справочник из файла по связи 'Файл шаблон справочника' объекта 'Документ НСИ'
        /// </summary>
        public void ДокументыНСИ_ЗаполнитьСправочникИзФайла_РабочаяСтраница()
        {
            var result = ВыполнитьМакрос(TakeCellObjectMacroGuid, "ПолучитьGuidСправочника", "Записи");

            if (!(result is Guid guid))
                throw new MacroException("В окно не загружен справочник");

            new ObjectGeneratorFromExcelFileCommand(this).ЗаполнитьСправочникИзФайла(guid);
        }

        /// <summary>
        /// Заполнить новый справочник из файла, который выбран из диалога выбора файла 
        /// </summary>
        public void ДокументыНСИ_ЗаполнитьСправочникИзФайла()
        {
            new ObjectGeneratorFromExcelFileCommand(this).ЗаполнитьСправочникИзФайла();
        }

        #endregion

        #region 'Документы НСИ' - Обновление новых справочников

        public void ДокументыНСИ_ОбновитьСправочники()
            => new UpdaterNewNsiReferences(this).ОбновитьСправочники();

        public void ДокументыНСИ_УстановитьСтадиюКорректировка()
            => new UpdaterNewNsiReferences(this).УстановитьСтадиюКорректировка();

        #endregion

        #region 'Параметры документов НСИ' - Обновление новых справочников

        public void ПараметрыДокументовНСИ_ОбновитьСправочники()
            => new UpdaterNewNsiReferences(this).ОбновитьСправочники();

        public void ПараметрыДокументовНСИ_УстановитьСтадиюКорректировка()
            => new UpdaterNewNsiReferences(this).УстановитьСтадиюКорректировка();

        #endregion

        #region 'Описание параметров' - Системные события

        public void ОписаниеПараметров_СозданиеОбъекта()
        {
            new DescParametersNsiSystemHandlers(this).СозданиеОбъекта();
        }

        public void ОписаниеПараметров_СохранениеОбъекта()
        {
            new DescParametersNsiSystemHandlers(this).СохранениеОбъекта();
        }

        public void ОписаниеПараметров_ИзменениеПараметраОбъекта()
        {
            new DescParametersNsiSystemHandlers(this).ИзменениеПараметраОбъекта();
        }

        public void ОписаниеПараметров_ИзменениеСвязиОбъекта()
        {
            new DescParametersNsiSystemHandlers(this).ИзменениеСвязиОбъекта();
        }

        #endregion

        #region 'Описание параметров' - Команды для импорта параметров

        public void ОписаниеПараметров_ВыбратьИзШаблоновПараметров()
        {
            new DescParametersNsiImportCommands(this).ВыбратьИзШаблоновПараметров();
        }

        public void ОписаниеПараметров_ДобавитьГруппыПараметров()
        {
            new DescParametersNsiImportCommands(this).ДобавитьГруппыПараметров();
        }

        public void ОписаниеПараметров_КопироватьИзПараметровДокументаНСИ()
        {
            new DescParametersNsiImportCommands(this).КопироватьИзПараметровДокументаНСИ();
        }

        public void ОписаниеПараметров_КопироватьИзДокументаНСИ()
        {
            new DescParametersNsiImportCommands(this).КопироватьИзДокументаНСИ();
        }

        public void ОписаниеПараметров_ДобавитьВшаблоны()
        {
            new DescParametersNsiImportCommands(this).ДобавитьВшаблоны();
        }

        #endregion

        #region 'Список значений' - Команды для списка объектов 'Список значений'

        public void СписокЗначений_ДобавитьСписок()
        {
            new ListOfValuesNsiAddListValuesCommand(this).ДобавитьСписок();
        }

        #endregion

        #region 'Новый справочник НСИ' - Генератор объектов справочника 'Эталоны СтИ'

        public void НовыйСправочникНСИ_СоздатьЭталоныСти()
        {
            new GeneratorNewObjectsEtalonStiCommand(this).СоздатьЭталоныСти();
        }

        #endregion

        #region 'Шаблоны параметров' - Системные события

        public void ШаблоныПараметров_ЗавершениеИзмененияСвязиОбъекта()
        {
            new ParameterTemplatesNsiSystemHandlers(this).ЗавершениеИзмененияСвязиОбъекта();
        }

        #endregion

        #region Команды для обновления структуры новых справочников НСИ

        public string ОбновитьСправочники_БП(Объекты объекты)
        {
            if (объекты == null)
                throw new ArgumentNullException(nameof(объекты));

            var messageBuilder = new StringBuilder();

            messageBuilder.AppendLine("Результат обновления структуры справочников:");

            foreach (var объект in объекты)
            {
                var referenceObject = (ReferenceObject) объект;
                var updaterNewNsiReferences = new UpdaterNewNsiReferences(this, referenceObject, true);
                updaterNewNsiReferences.ОбновитьСправочники();
                messageBuilder.AppendLine(updaterNewNsiReferences.GetMessageLog());
            }

            return messageBuilder.ToString();
        }

        public void УстановитьСтадиюКорректировка_БП(Объекты объекты)
        {
            if (объекты == null)
                throw new ArgumentNullException(nameof(объекты));

            foreach (var объект in объекты)
            {
                var referenceObject = (ReferenceObject) объект;
                new UpdaterNewNsiReferences(this, referenceObject, true).УстановитьСтадиюКорректировка();
            }
        }

        #endregion

    }
}

