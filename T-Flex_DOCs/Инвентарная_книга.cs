using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using TFlex.DOCs.Model.Macros;
using TFlex.DOCs.Model.Macros.ObjectModel;
using TFlex.DOCs.Model.References;
using TFlex.DOCs.Model.References.Documents;
using TFlex.DOCs.Model.References.Files;
using TFlex.DOCs.Model.References.Nomenclature;
using TFlex.DOCs.Model.Search;
using TFlex.DOCs.Model.Structure;

namespace InventoryBook
{
    public class Macro : MacroProvider
    {
        /// <summary> Поддерживаемые расширения файлов </summary>
        private static readonly string[] SupportedExtensions = { "tif", "tiff", "pdf" };

        private static class Guids
        {
            public static readonly Guid DocumentInventoryCard = new Guid("d708e1b4-2a1a-499c-aaaf-be5828e6377e");
            public static readonly Guid DocumentRevisionInventoryCard = new Guid("a193acc6-805f-4e8e-92c8-b7a029792951");
            public static readonly Guid InventoryCardReference = new Guid("f1d28c5d-6b08-4061-b81f-cde490088e1f");
            public static readonly Guid InventoryCardKDType = new Guid("d9f5cb08-8f8b-4e1f-947f-d1250a19a7b6");

            public static readonly Guid InventoryCardFormatsList = new Guid("2fc14743-2311-4ed8-9553-8c8715cdbf64");
            public static readonly Guid DocumentRevisionFormatLink = new Guid("773c9b18-ee96-48a4-954b-74d89face72e");
            public static readonly Guid CountFormatsParameter = new Guid("66723207-ccf6-445c-9512-014e83f3ae5a");

            public static readonly Guid InventoryCardFormatParameter = new Guid("124e3a9b-8d1f-419a-b030-bbc888c7fac6");
            public static readonly Guid InventoryCardNumberOfSheetsParameter = new Guid("6861340b-3b1f-4fdf-8736-3f5ccf7b48a4");
            public static readonly Guid InventoryCardDocumentDenotationParameter = new Guid("a0aedf52-cba9-4bd7-afe6-50ee5c100a6d");
            public static readonly Guid InventoryCardDocumentNameParameter = new Guid("795785c4-f1d5-4075-b846-b3abc964b7eb");
            public static readonly Guid InventoryCardLogicalObjectGuidParameter = new Guid("e2372665-9260-4bbe-b63f-580e31e3f8d8");

            public static readonly Guid FormatType = new Guid("bb574240-0e6a-4016-bf8c-c860996bed3b");
        }

        public Macro(MacroContext context)
        : base(context)
        {
        }

        public override void Run()
        {
        }

        public void СохранениеИнвентарнойКарточки()
        {
            СформироватьИнвентарныйНомерКарточки();
        }

        private void СформироватьИнвентарныйНомерКарточки()
        {
            string инвентарныйНомер = ТекущийОбъект["Инвентарный номер"];
            if (!string.IsNullOrEmpty(инвентарныйНомер))
                return;

            int порядковыйНомер = 1;

            string кодЖурнала = ТекущийОбъект["Код журнала"];
            Объект счётчик = НайтиОбъект("Нумератор инвентарной книги", "Код журнала", кодЖурнала);
            if (счётчик == null)
            {
                счётчик = СоздатьОбъект("Нумератор инвентарной книги", "Запись");
                счётчик["Код журнала"] = кодЖурнала;
            }
            else
            {
                порядковыйНомер = счётчик["Номер"];
                порядковыйНомер += 1;
                счётчик.Изменить();
            }

            счётчик["Номер"] = порядковыйНомер;
            счётчик.Сохранить();

            ТекущийОбъект["Инвентарный номер"] = string.Format("{0}{1}{2}", кодЖурнала, string.IsNullOrEmpty(кодЖурнала) ? string.Empty : ".", порядковыйНомер.ToString("D7"));
        }

        public ReferenceObject FindOrCreateDocumentInventoryCard(EngineeringDocumentObject document, bool needEndChanges)
        {
            if (document is null)
                return null;

            var nomenclatureReference = new NomenclatureReference(Context.Connection);
            var nomenclatureObject = document.GetLinkedNomenclatureObject();

            var revisionFilter = new Filter(nomenclatureReference.ParameterGroup);
            revisionFilter.Terms.AddTerm(
                nomenclatureReference.ParameterGroup.SystemParameters.Find(SystemParameterType.LogicalObjectGuid),
                ComparisonOperator.Equal,
                nomenclatureObject.SystemFields.LogicalObjectGuid);

            var existingRevisions = nomenclatureReference.Find(revisionFilter);

            // Ищем карточку по связи
            ReferenceObject inventoryCard = null;

            foreach (var revision in existingRevisions.OfType<NomenclatureObject>())
            {
                if (revision.LinkedObject is EngineeringDocumentObject revisionDocument)
                {
                    revisionDocument.TryGetObject(Guids.DocumentRevisionInventoryCard, out inventoryCard);
                    if (inventoryCard is null)
                        revisionDocument.TryGetObject(Guids.DocumentInventoryCard, out inventoryCard);
                }

                if (inventoryCard != null)
                    break;
            }

            if (inventoryCard is null)
            {
                var cardReferenceInfo = Context.Connection.ReferenceCatalog.Find(Guids.InventoryCardReference);
                var cardReference = cardReferenceInfo.CreateReference();
                var cardClassKD = cardReference.Classes.Find(Guids.InventoryCardKDType);
                if (cardClassKD is null)
                    return null;

                // Ищем в справочнике карточку
                using (var filter = new Filter(cardReferenceInfo))
                {
                    // условие по типу карточки
                    filter.Terms.AddTerm(
                        cardReference.ParameterGroup.ClassParameterInfo,
                        ComparisonOperator.IsInheritFrom,
                        cardClassKD);

                    // условие по гуиду логического объекта
                    filter.Terms.AddTerm(
                        cardReference.ParameterGroup[Guids.InventoryCardLogicalObjectGuidParameter],
                        ComparisonOperator.Equal,
                        document.SystemFields.LogicalObjectGuid);

                    inventoryCard = cardReference.Find(filter).FirstOrDefault();
                }

                if (inventoryCard is null)
                {
                    using (var filter = new Filter(cardReferenceInfo))
                    {
                        // условие по типу карточки
                        filter.Terms.AddTerm(cardReference.ParameterGroup.ClassParameterInfo, ComparisonOperator.IsInheritFrom, cardClassKD);

                        // условие по отсутствию связанного документа у карточки
                        filter.Terms.AddTerm(
                            Guids.DocumentInventoryCard.ToString(),
                            ComparisonOperator.IsNull,
                            null);

                        string denotation = document[SpecificationFields.Denotation].Value.ToString();
                        // условие по обозначению документа
                        filter.Terms.AddTerm
                            (cardReference.ParameterGroup[Guids.InventoryCardDocumentDenotationParameter],
                            ComparisonOperator.Equal,
                            denotation);

                        inventoryCard = cardReference.Find(filter).FirstOrDefault();
                    }
                }

                // Если карточка по фильтру не найдена
                if (inventoryCard is null)
                {
                    inventoryCard = cardReference.CreateReferenceObject(cardClassKD);

                    inventoryCard[Guids.InventoryCardDocumentNameParameter].Value = document[EngineeringDocumentFields.Name];
                    inventoryCard[Guids.InventoryCardDocumentDenotationParameter].Value = document[SpecificationFields.Denotation];

                    //подключить документ к карточке

                    var documentLink = inventoryCard.FindRelation(Guids.DocumentInventoryCard);
                    if (documentLink != null)
                        inventoryCard.SetLinkedObject(documentLink, document);
                }
                else
                    inventoryCard.BeginChanges();

                //подключить ревизию к карточке

                var revisionLink = inventoryCard.FindRelation(Guids.DocumentRevisionInventoryCard);
                if (revisionLink != null)
                    inventoryCard.AddLinkedObject(revisionLink, document);
            }
            else
            {
                inventoryCard.BeginChanges();

                //подключить документ к карточке
                var revisionLink = inventoryCard.FindRelation(Guids.DocumentRevisionInventoryCard);
                if (revisionLink != null)
                    inventoryCard.AddLinkedObject(revisionLink, document);
            }

            if (inventoryCard.Changing)
            {
                inventoryCard[Guids.InventoryCardLogicalObjectGuidParameter].Value = document.SystemFields.LogicalObjectGuid;

                if (needEndChanges)
                    inventoryCard.EndChanges();
            }

            return inventoryCard;
        }

        public void ProcessInventoryCard(ReferenceObject inventoryCard, EngineeringDocumentObject document, FileObject file, bool clearFormatList, bool useRevision)
        {
            if (document is null || file is null || inventoryCard is null)
                return;

            string fileExtension = file.Class.Extension.ToLower();
            if (!SupportedExtensions.Contains(fileExtension))
                return;

            try
            {
                string tempFolder = Path.Combine(Path.GetTempPath(), "Temp DOCs");
                if (!Directory.Exists(tempFolder))
                    Directory.CreateDirectory(tempFolder);

                string pathToFile = String.Format(@"{0}\{1}.{2}", tempFolder, Guid.NewGuid(), fileExtension);

                file.GetHeadRevision(pathToFile);

                DisassembleFormats(
                    pathToFile,
                    Объект.CreateInstance(inventoryCard, Context),
                    clearFormatList,
                    useRevision ? Объект.CreateInstance(document, Context) : null);

                DeleteTempFile(pathToFile);
            }
            catch (Exception)
            {
                // %%TODO                
            }
        }

        /// <summary>
        /// Вычисление форматов и кол-ва листов подлинника для заполнения данных инвентарной карточки
        /// </summary>
        /// <param name="путьКФайлу"></param>
        /// <param name="инвентарнаяКарточка"></param>
        public void DisassembleFormats(string путьКФайлу, Объект инвентарнаяКарточка, bool очиститьСписокФорматов, Объект ревизия)
        {
            bool needEditAndSave = !инвентарнаяКарточка.Changing;
            if (needEditAndSave)
                инвентарнаяКарточка.BeginChanges();

            try
            {
                // очистка списка листов
                if (очиститьСписокФорматов)
                {
                    foreach (Объект лист in инвентарнаяКарточка.СвязанныеОбъекты[Guids.InventoryCardFormatsList.ToString()])
                    {
                        if (лист.Тип.СодержитСвязь(Guids.DocumentRevisionFormatLink.ToString()))
                        {
                            if (лист.СвязанныйОбъект[Guids.DocumentRevisionFormatLink.ToString()] != null)
                                continue;
                        }

                        лист.Удалить();
                    }
                }

                // если задан путь к подлиннику, разбираем его по форматам
                if (!String.IsNullOrEmpty(путьКФайлу))
                {
                    Объекты списокФорматов = ВыполнитьМакрос("a00bcfd7-91b9-4c5d-9ef7-1538d448a31a", "ПолучитьСписокФорматов", путьКФайлу);
                    if (списокФорматов.Any())
                    {
                        Dictionary<Объект, int> списокФорматовСКоличеством = new Dictionary<Объект, int>();

                        int количествоНеопределенныхФорматов = 0;
                        foreach (Объект форматСправочника in списокФорматов)
                        {
                            if (форматСправочника == null)
                            {
                                количествоНеопределенныхФорматов++;
                                continue;
                            }

                            Объект формат = списокФорматовСКоличеством.Keys.FirstOrDefault(объектФормата => объектФормата["Наименование"] == форматСправочника["Наименование"]);
                            if (формат == null)
                            {
                                списокФорматовСКоличеством.Add(форматСправочника, 1);
                            }
                            else
                            {
                                списокФорматовСКоличеством[формат]++;
                            }
                        }

                        foreach (KeyValuePair<Объект, int> формат in списокФорматовСКоличеством)
                        {
                            Объект лист = инвентарнаяКарточка.СоздатьОбъектСписка(Guids.InventoryCardFormatsList.ToString(), Guids.FormatType.ToString());
                            лист.СвязанныйОбъект["Формат"] = формат.Key;
                            лист["Наименование"] = формат.Key["Наименование"];
                            лист["Количество"] = формат.Value;

                            if (ревизия != null && лист.Тип.СодержитСвязь(Guids.DocumentRevisionFormatLink.ToString()))
                                лист.Подключить(Guids.DocumentRevisionFormatLink.ToString(), ревизия);

                            лист.Сохранить();
                        }

                        if (количествоНеопределенныхФорматов > 0)
                        {
                            Объект лист = инвентарнаяКарточка.СоздатьОбъектСписка(Guids.InventoryCardFormatsList.ToString(), Guids.FormatType.ToString());
                            лист["Наименование"] = "Формат не определён";
                            лист["Количество"] = количествоНеопределенныхФорматов;

                            if (ревизия != null && лист.Тип.СодержитСвязь(Guids.DocumentRevisionFormatLink.ToString()))
                                лист.Подключить(Guids.DocumentRevisionFormatLink.ToString(), ревизия);

                            лист.Сохранить();
                        }
                    }
                }

                RecalcFormats((ReferenceObject)инвентарнаяКарточка, ревизия is null ? null : (ReferenceObject)ревизия);

                if (needEditAndSave && инвентарнаяКарточка.Changing)
                    инвентарнаяКарточка.Save();
            }
            catch
            {
                if (needEditAndSave && инвентарнаяКарточка.Changing)
                    инвентарнаяКарточка.CancelChanges();
            }
        }

        public void RecalcFormats(ReferenceObject inventoryCard, ReferenceObject documentRevision)
        {
            string formats = String.Empty;
            int allPagesCount = 0;

            // список Форматы
            var formatObjects = documentRevision is null
                ? inventoryCard.Links.ToMany[Guids.InventoryCardFormatsList].ToList()
                : inventoryCard.Links.ToMany[Guids.InventoryCardFormatsList]
                .Where(list => list.GetObject(Guids.DocumentRevisionFormatLink) == documentRevision).ToList();

            foreach (var format in formatObjects)
            {
                int count = (int)format[Guids.CountFormatsParameter].Value;
                allPagesCount += count;
                formats = String.Format("{0}{1}; ", formats, format.ToString());
            }

            inventoryCard[Guids.InventoryCardFormatParameter].Value = formats;
            inventoryCard[Guids.InventoryCardNumberOfSheetsParameter].Value = allPagesCount;
        }

        private void DeleteTempFile(string path)
        {
            try
            {
                var fileInfo = new FileInfo(path);
                if (fileInfo.Exists && !fileInfo.IsReadOnly)
                    fileInfo.Delete();
            }
            catch
            {
            }
        }
    }
}

