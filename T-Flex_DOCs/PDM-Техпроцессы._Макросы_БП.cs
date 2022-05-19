using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using TFlex.DOCs.Common;
using TFlex.DOCs.Model;
using TFlex.DOCs.Model.Classes;
using TFlex.DOCs.Model.Desktop;
using TFlex.DOCs.Model.Macros.ObjectModel;
using TFlex.DOCs.Model.Macros.Processes;
using TFlex.DOCs.Model.Processes.Events.Contexts;
using TFlex.DOCs.Model.References;
using TFlex.DOCs.Model.References.Files;
using TFlex.DOCs.Model.References.Nomenclature.ModificationNotices;
using TFlex.DOCs.Model.Search;
using TFlex.DOCs.Model.Stages;

namespace TechnologicalPDM_BPMacros
{
    public class Macro : ProcessActionMacroProvider
    {
        #region Guids

        private static class Guids
        {
            public static class References
            {
                public static readonly Guid ИнвентарнаяКнига = new Guid("f1d28c5d-6b08-4061-b81f-cde490088e1f");
            }

            public static class Classes
            {
                public static readonly Guid ТехнологическийПроцесс = new Guid("3e93d599-c214-48c8-854f-efe4b475c4d8");
                public static readonly Guid ТехнологическаяОперация = new Guid("f53c9d73-18bb-4c59-a260-61fea65f6ed9");
                public static readonly Guid КарточкаУчетаТехнологическихДокументов = new Guid("d2297aab-a159-45bf-8601-7a7f1f27a38c");
                public static readonly Guid ТехнологическийКомплект = new Guid("dc1cf2a0-6c01-400d-9a42-9642b7496404");
                public static readonly Guid ИзвещениеОбИзменении = new Guid("52ccb35c-67c5-4b82-af4f-e8ceac4e8d02");
            }

            public static class Links
            {
                public static readonly Guid ТПДокументация = new Guid("cc38caed-f747-45ce-9fbf-771566841796");
                public static readonly Guid ИзготавливаемыеДСЕ = new Guid("e1e8fa07-6598-444d-8f57-3cfd1a3f4360");
                public static readonly Guid ТехнологическийКомплектПодлинник = new Guid("148a64ed-3906-4da9-95fc-14bb018669f2");
                public static readonly Guid ИзмененияАктуальныйВариантТЭ = new Guid("254c4753-4b42-454e-84cc-f5abc82b2448");
                public static readonly Guid ИзмененияИсходныйВариантТЭ = new Guid("87dab3c7-c8f5-40a4-91dd-da6734ee1f3b");
                public static readonly Guid ИзмененияЦелевойВариантТЭ = new Guid("737f68ad-9038-4585-b944-428662256f18");
                public static readonly Guid ИзвещенияИзменения = new Guid("5e46670a-400c-4e36-bb37-d4d651bdf692");
                public static readonly Guid КарточкаТехнологическийДокумент = new Guid("a5a8963d-18f6-47c2-8767-70adc2e1694b");
                public static readonly Guid ФайлДокумента = new Guid("6b18c3fc-7cd1-4ece-a526-cacad8101f09");
                public static readonly Guid ТехнологическийКомплектПапка = new Guid("9f22f919-09f8-416c-876c-c52f4fbb36cd");
            }

            public static class Parameters
            {
                public static readonly Guid ТехнологическийДокументОбозначение = new Guid("d650d3f8-f423-4854-8777-472b2fd93921");
                public static readonly Guid ТехнологическийДокументНаименование = new Guid("11b85954-867b-4ca2-9b76-374e3c8c3ae9");
                public static readonly Guid ТПНаименование = new Guid("f97e40ea-3c79-4013-b1ea-383a2f09454d");
                public static readonly Guid ТПОбозначение = new Guid("c0d3d63b-e1aa-422e-abb7-27c8ab0e0b3e");
                public static readonly Guid ТПВариант = new Guid("4ae9f4a6-49a1-4a11-8075-50e2a403d214");
                public static readonly Guid НомерИзменения = new Guid("91486563-d044-4045-814b-3432b67812f1");
                public static readonly Guid КарточкаОбозначениеДокумента = new Guid("a0aedf52-cba9-4bd7-afe6-50ee5c100a6d");
                public static readonly Guid КарточкаНаименованиеДокумента = new Guid("795785c4-f1d5-4075-b846-b3abc964b7eb");
                public static readonly Guid ОбозначениеИзвещения = new Guid("b03c9129-7ac3-46f5-bf7d-fdd88ef1ff9a");
            }
        }

        #endregion

        private static string _папкаАрхивТД = "Архив ТД";
        private static string _папкаАрхивТДПодлинники = Path.Combine(_папкаАрхивТД, "Подлинники");
        private static string _папкаАрхивТДАннулировано = Path.Combine(_папкаАрхивТД, "Аннулировано");

        private ReferenceInfo _cardReferenceInfo;
        private Reference _cardReference;
        private ClassObject _cardClass;
        private FileReference _fileReference;

        public Macro(EventContext context)
            : base(context)
        {
        }

        private ReferenceInfo CardReferenceInfo
        {
            get
            {
                if (_cardReferenceInfo == null)
                    _cardReferenceInfo = Context.Connection.ReferenceCatalog.Find(Guids.References.ИнвентарнаяКнига);

                return _cardReferenceInfo;
            }
        }

        private Reference CardReference
        {
            get
            {
                if (_cardReference == null)
                    _cardReference = CardReferenceInfo.CreateReference();

                return _cardReference;
            }
        }

        private ClassObject CardClass
        {
            get
            {
                if (_cardClass == null)
                    _cardClass = CardReference.Classes.Find(Guids.Classes.КарточкаУчетаТехнологическихДокументов);

                return _cardClass;
            }
        }

        private FileReference FileReferenceInstance
        {
            get
            {
                if (_fileReference == null)
                    _fileReference = new FileReference(Context.Connection);

                return _fileReference;
            }
        }

        public void ПроверкаТДПередСогласованием()
        {
            var processObjects = Objects.Select(объект => (ReferenceObject)объект).ToList();

            string result;
            List<ReferenceObject> documentFiles = CheckTD(processObjects, out result);
            if (documentFiles.Any())
                ПодключитьОбъекты(documentFiles.Select(referenceObject => Объект.CreateInstance(referenceObject, Context)));

            Переменные["РезультатПроверкиТД"] = result;
        }

        private List<ReferenceObject> CheckTD(List<ReferenceObject> referenceObjects, out string result)
        {
            result = string.Empty;

            StringBuilder errors = new StringBuilder();
            List<ReferenceObject> documentFiles = new List<ReferenceObject>();

            foreach (ReferenceObject tp in referenceObjects)
            {
                if (!tp.Class.IsInherit(Guids.Classes.ТехнологическийПроцесс))
                    continue;

                var rootDocuments = tp.GetObjects(Guids.Links.ТПДокументация);
                if (!rootDocuments.Any())
                {
                    errors.AppendLine(string.Format("Технологический процесс '{0}' не содержит комплект документов.", tp));
                    continue;
                }

                var setOfDocuments = rootDocuments.FirstOrDefault(document => document.Class.IsInherit(Guids.Classes.ТехнологическийКомплект));
                documentFiles.Add(setOfDocuments);

                // проверка на заполнение параметров комплекта документов
                string denotation = setOfDocuments[Guids.Parameters.ТехнологическийДокументОбозначение].Value.ToString();
                if (string.IsNullOrEmpty(denotation))
                    errors.AppendLine("Не задано обозначение для комплекта документов.");

                ReferenceObject folder = setOfDocuments.GetObject(Guids.Links.ТехнологическийКомплектПапка);
                if (folder == null)
                    errors.AppendLine("Не задана папка для файла комплекта документов.");

                var originalFile = setOfDocuments.GetObject(Guids.Links.ФайлДокумента) as FileObject;
                if (originalFile != null)
                    errors.AppendLine("Файл комплекта документов уже существует.");

                setOfDocuments.Children.Reload();
                foreach (ReferenceObject document in setOfDocuments.Children)
                {
                    ReferenceObject file = document.GetObject(Guids.Links.ФайлДокумента) as FileObject;
                    if (file == null)
                    {
                        errors.AppendLine(string.Format("Отсутствует файл технологического документа '{0}'.", document.ToString()));
                        continue;
                    }

                    if (file.IsCheckedOut)
                    {
                        errors.AppendLine(string.Format("К файлу '{0}' технологического документа '{1}' не применены изменения.", file.ToString(), document.ToString()));
                    }
                }
            }

            if (errors.Length == 0)
                errors.AppendLine("Ошибок нет.");

            result = errors.ToString();

            return documentFiles;
        }

        //public void ПодключитьПодлинникКТД()
        //{
        //    var processObjects = Objects.Select(объект => (ReferenceObject)объект).ToList();

        //    // подлинники pdf, которые были созданы и подключены к БП на предыдущем этапе
        //    List<FileObject> allRealFiles = processObjects.OfType<FileObject>().Where(file => Path.GetExtension(file.Name).ToLower() == ".pdf").ToList();

        //    List<FileObject> linkedRealFiles = new List<FileObject>();

        //    foreach (ReferenceObject referenceObject in processObjects)
        //    {
        //        if (referenceObject.Class.IsInherit(Guids.Classes.ТехнологическийПроцесс))
        //        {
        //            FileObject linkedRealFile = LinkRealFile(referenceObject, allRealFiles);
        //            if (linkedRealFile != null)
        //                linkedRealFiles.Add(linkedRealFile);
        //        }
        //        else if (referenceObject.Class.IsInherit(Guids.Classes.ИзвещениеОбИзменении))
        //        {
        //            var modifications = referenceObject.GetObjects(Guids.Links.ИзвещенияИзменения).ToList();
        //            foreach (var modification in modifications)
        //            {
        //                var tp = modification.GetObject(Guids.Links.ИзмененияЦелевойВариантТЭ);
        //                if (tp == null)
        //                    continue;

        //                if (tp.Class.IsInherit(Guids.Classes.ТехнологическийПроцесс))
        //                {
        //                    FileObject linkedRealFile = LinkRealFile(tp, allRealFiles);
        //                    if (linkedRealFile != null)
        //                        linkedRealFiles.Add(linkedRealFile);
        //                }
        //            }
        //        }
        //    }

        //    // переводим на стадию "Утверждено" подключенные подлинники
        //    ChangeStage(new List<ReferenceObject>(linkedRealFiles), "Утверждено");

        //    // %%TODO отключаем подлинники от БП
        //}

        //private FileObject LinkRealFile(ReferenceObject tp, List<FileObject> allRealFiles)
        //{
        //    var rootDocuments = tp.GetObjects(Guids.Links.ТПДокументация);
        //    if (!rootDocuments.Any())
        //        return null;

        //    var setOfDocuments = rootDocuments.FirstOrDefault(document => document.Class.IsInherit(Guids.Classes.ТехнологическийКомплект));
        //    if (setOfDocuments == null)
        //        return null;

        //    var originalFile = setOfDocuments.GetObject(Guids.Links.ОригиналТД) as FileObject;
        //    if (originalFile == null)
        //        return null;

        //    var realFile = allRealFiles.FirstOrDefault(fileObject => Path.GetFileNameWithoutExtension(fileObject.Name) == Path.GetFileNameWithoutExtension(originalFile.Name));
        //    if (realFile == null)
        //        return null;

        //    setOfDocuments.BeginChanges();
        //    setOfDocuments.SetLinkedObject(Guids.Links.ТехнологическийКомплектПодлинник, realFile);
        //    setOfDocuments.EndChanges();

        //    return realFile;
        //}

        public void ПрименитьИИдляТехнологическихИзменений()
        {
            var modificationNotice = GetModificationNotice();
            if (modificationNotice == null)
                throw new ArgumentNullException("Не найдено ИИ");

            ProcessModificationNotice(modificationNotice);
        }

        private ReferenceObject GetModificationNotice()
        {
            var objects = Объекты.Union(ВспомогательныеОбъекты).Select(объект => (ReferenceObject)объект).ToList();
            return objects.FirstOrDefault(obj => obj.Reference.ParameterGroup.Guid == ModificationNoticesReference.ModificationNoticesReferenceGuid);
        }

        private void ProcessModificationNotice(ReferenceObject modificationNotice)
        {
            var modifications = modificationNotice.GetObjects(Guids.Links.ИзвещенияИзменения).Where(modification => !CheckReferenceObjectInStage(modification, "Хранение")).ToList();
            foreach (var modification in modifications)
            {
                var actualTP = modification.GetObject(Guids.Links.ИзмененияАктуальныйВариантТЭ);
                if (actualTP == null || !actualTP.Class.IsInherit(Guids.Classes.ТехнологическийПроцесс))
                    continue;

                var targetTP = modification.GetObject(Guids.Links.ИзмененияЦелевойВариантТЭ);
                if (targetTP == null || !targetTP.Class.IsInherit(Guids.Classes.ТехнологическийПроцесс))
                    continue;

                // Формируем следующий номер варианта
                int lastVariantNumber = FindLastVariantNumberTP(actualTP);
                string newVariantName = string.Format("ИИ.{0}", lastVariantNumber + 1);

                // Текущий актуальный ТП, комплект документации и файлы
                List<ReferenceObject> allActualTechnologicalObjects = GetAllTechnologicalObjects(actualTP, false);

                // Обрабатываем актуальный ТЭ: присваиваем ему следующий номер варианта
                ChangeStage(new List<ReferenceObject>() { actualTP }, "Исправление");
                actualTP.BeginChanges();
                actualTP[Guids.Parameters.ТПВариант].Value = newVariantName;
                List<ReferenceObject> dseList = actualTP.GetObjects(Guids.Links.ИзготавливаемыеДСЕ);
                if (dseList.Any())
                {
                    // Отключаем "Изготавливаемые ДСЕ" от актуального варианта
                    actualTP.ClearLinks(Guids.Links.ИзготавливаемыеДСЕ);
                }
                actualTP.EndChanges();

                // обработка старого подлинника
                FileObject notUsedFile = allActualTechnologicalObjects.OfType<FileObject>().FirstOrDefault();
                if (notUsedFile != null)
                {
                    ChangeStage(new List<ReferenceObject>() { notUsedFile }, "Исправление");

                    var folder = GetFolder(_папкаАрхивТДАннулировано);
                    // перемещаем в папку "Аннулировано"
                    if (folder != null)
                    {
                        if (notUsedFile.Parent.Name.GetString() != _папкаАрхивТДАннулировано)
                        {
                            notUsedFile.MoveFileToFolder(folder);
                        }
                    }

                    // переименовываем подлинник
                    FileObject checkingOutFile = Desktop.CheckOut(notUsedFile, false).OfType<FileObject>().FirstOrDefault();
                    if (checkingOutFile != null)
                    {
                        checkingOutFile.BeginChanges();

                        string modificationNoticeNumber = modificationNotice[Guids.Parameters.ОбозначениеИзвещения].ToString();
                        string newFileName = string.Format("{0}_ИИ{1}{2}", Path.GetFileNameWithoutExtension(checkingOutFile.Name), modificationNoticeNumber, Path.GetExtension(checkingOutFile.Name));
                        checkingOutFile.Name.Value = newFileName;

                        checkingOutFile.EndChanges();

                        Desktop.CheckIn(checkingOutFile, string.Format("Автоматическое переименование и перемещение файла '{0}' в папку {1}", string.Join(Environment.NewLine, checkingOutFile.OfType<FileObject>().Select(file => file.Path)), _папкаАрхивТДАннулировано), false);
                    }
                }

                ChangeStage(allActualTechnologicalObjects, "Аннулировано");


                // Объекты, которые переводим на "Хранение": целевой вариант, все дочерние объекты, комплект документов, файлы документов
                // Обрабатываем инвентарные карточки технологических документов
                List<ReferenceObject> allTargetTechnologicalObjects = GetAllTechnologicalObjects(targetTP, true);
                // добавляем изменение для перевода на "Хранение"
                allTargetTechnologicalObjects.Add(modification);

                // Обрабатываем целевой ТЭ
                // Убираем номер варианта, подключаем ДСЕ и изменения, которые были у актуального варианта
                ChangeStage(new List<ReferenceObject>() { targetTP }, "Исправление");
                targetTP.BeginChanges();

                targetTP[Guids.Parameters.ТПВариант].Value = string.Empty;

                foreach (ReferenceObject dse in dseList)
                    targetTP.AddLinkedObject(Guids.Links.ИзготавливаемыеДСЕ, dse);

                List<ReferenceObject> modificationList = actualTP.GetObjects(Guids.Links.ИзмененияАктуальныйВариантТЭ);
                if (modificationList.Any())
                {
                    foreach (ReferenceObject mod in modificationList)
                        targetTP.AddLinkedObject(Guids.Links.ИзмененияАктуальныйВариантТЭ, mod);
                }

                targetTP.EndChanges();

                // перемещаем новый подлинник
                FileObject realFile = allTargetTechnologicalObjects.OfType<FileObject>().FirstOrDefault();
                if (realFile != null)
                {
                    var folder = GetFolder(_папкаАрхивТДПодлинники);
                    // перемещаем в папку "Подлинники"
                    if (folder != null)
                    {
                        if (realFile.Parent.Name.GetString() != _папкаАрхивТДПодлинники)
                        {
                            realFile.MoveFileToFolder(folder);
                        }
                    }
                }

                // Обрабатываем изменение
                // Меняем номер изменения (вариант), переподключаем варианты
                ChangeStage(new List<ReferenceObject>() { modification }, "Исправление");
                modification.BeginChanges();
                modification[Guids.Parameters.НомерИзменения].Value = newVariantName;
                // актуальный делаем исходным (он теперь с вариантом "ИИ.{0}")
                modification.SetLinkedObject(Guids.Links.ИзмененияИсходныйВариантТЭ, actualTP);
                // целевой делаем актуальным (он без варианта)
                modification.SetLinkedObject(Guids.Links.ИзмененияАктуальныйВариантТЭ, targetTP);
                // и отключаем его от связи целевой
                modification.SetLinkedObject(Guids.Links.ИзмененияЦелевойВариантТЭ, null);
                modification.EndChanges();

                ChangeStage(allTargetTechnologicalObjects, "Хранение");

                // %%TODO меняем у всех изменений, связанных с актуальным ТП, актуальный ТП на новый актуальный (который был целевым в данном изменении)
                //var allModifications = actualTP.GetObjects(Guids.Links.ИзмененияАктуальныйВариантТЭ);
                //foreach (ReferenceObject actualModification in allModifications)
                //{
                //    if (CheckReferenceObjectInStage(actualModification, "Хранение"))
                //        continue;

                //    ReferenceObject changingModification = ChangeStage(new List<ReferenceObject>() { modification }, "Исправление").FirstOrDefault();
                //    if (changingModification != null)
                //    {
                //        changingModification.BeginChanges();
                //        changingModification.SetLinkedObject(Guids.Links.ИзмененияАктуальныйВариантТЭ, targetTP);
                //        changingModification.EndChanges();
                //    }
                //}
            }

            ChangeStage(new List<ReferenceObject>() { modificationNotice }, "Хранение");
        }

        private List<ReferenceObject> GetAllTechnologicalObjects(ReferenceObject currentTP, bool needProcessInventoryCard)
        {
            List<ReferenceObject> allTechnologicalObjects = new List<ReferenceObject>()
            {
                currentTP
            };

            var rootDocuments = currentTP.GetObjects(Guids.Links.ТПДокументация);
            allTechnologicalObjects.AddRange(rootDocuments);

            // комплект документов
            var setOfDocuments = rootDocuments.FirstOrDefault(document => document.Class.IsInherit(Guids.Classes.ТехнологическийКомплект));
            if (setOfDocuments != null)
            {
                // подлинник
                var realFile = setOfDocuments.GetObject(Guids.Links.ТехнологическийКомплектПодлинник);
                if (realFile != null)
                    allTechnologicalObjects.Add(realFile);

                // обработка инвентарной карточки (переподключение комплекта, если требуется)
                if (needProcessInventoryCard)
                    ProcessTechnologyInventoryCard(setOfDocuments);
            }

            foreach (ReferenceObject operation in currentTP.Children)
            {
                allTechnologicalObjects.Add(operation);
                allTechnologicalObjects.AddRange(operation.Children);

                if (!operation.Class.IsInherit(Guids.Classes.ТехнологическаяОперация))
                    continue;

                allTechnologicalObjects.AddRange(operation.GetObjects(Guids.Links.ТПДокументация));
            }

            return allTechnologicalObjects;
        }

        private void ProcessTechnologyInventoryCard(ReferenceObject document)
        {
            ReferenceObject inventoryCard = null;

            // Ищем в справочнике карточку
            using (Filter filter = new Filter(CardReferenceInfo))
            {
                // условие по типу карточки
                filter.Terms.AddTerm(CardReference.ParameterGroup.ClassParameterInfo, ComparisonOperator.IsInheritFrom, CardClass);
                // условие по наличию связанного документа у карточки (т.е. у документа уже должна была быть карточка)
                filter.Terms.AddTerm(Guids.Links.КарточкаТехнологическийДокумент.ToString(), ComparisonOperator.IsNotNull, null);

                // условие по обозначению документа
                string denotation = document[Guids.Parameters.ТехнологическийДокументОбозначение].Value.ToString();
                filter.Terms.AddTerm(CardReference.ParameterGroup[Guids.Parameters.КарточкаОбозначениеДокумента], ComparisonOperator.Equal, denotation);

                // условие по наименованию документа
                string name = document[Guids.Parameters.ТехнологическийДокументНаименование].Value.ToString();
                filter.Terms.AddTerm(CardReference.ParameterGroup[Guids.Parameters.КарточкаНаименованиеДокумента], ComparisonOperator.Equal, name);

                inventoryCard = CardReference.Find(filter).FirstOrDefault();
            }

            // новую карточку не создаем
            if (inventoryCard == null)
                return;

            inventoryCard.BeginChanges();

            // переподключаем документ к карточке
            inventoryCard.SetLinkedObject(Guids.Links.КарточкаТехнологическийДокумент, document);
            inventoryCard[Guids.Parameters.КарточкаНаименованиеДокумента].Value = document[Guids.Parameters.ТехнологическийДокументНаименование];
            inventoryCard[Guids.Parameters.КарточкаОбозначениеДокумента].Value = document[Guids.Parameters.ТехнологическийДокументОбозначение];

            inventoryCard.EndChanges();
        }

        private FolderObject GetFolder(string folderRelativePath)
        {
            if (string.IsNullOrEmpty(folderRelativePath))
                return null;

            if (FileReferenceInstance == null)
                return null;

            var folder = FileReferenceInstance.FindByRelativePath(folderRelativePath) as FolderObject;
            if (folder == null)
            {
                folder = FileReferenceInstance.CreatePath(folderRelativePath, null);
                if (folder != null)
                    Desktop.CheckIn(folder, TFlex.DOCs.Resources.Strings.Texts.AutoCreateText, false);
            }

            return folder;
        }

        private void DeleteTempFile(string path)
        {
            FileInfo fileInf = new FileInfo(path);
            if (fileInf.Exists && !fileInf.IsReadOnly)
                fileInf.Delete();
        }

        private int FindLastVariantNumberTP(ReferenceObject currentTP)
        {
            List<ReferenceObject> allVariants = new List<ReferenceObject>();

            var referenceInfo = Context.Connection.ReferenceCatalog.Find(SystemParameterGroups.TechnologicalProcesses);
            using (Filter filter = new Filter(referenceInfo))
            {
                var reference = referenceInfo.CreateReference();
                string name = currentTP[Guids.Parameters.ТПНаименование].Value.ToString();
                string denotation = currentTP[Guids.Parameters.ТПОбозначение].Value.ToString();

                filter.Terms.AddTerm(reference.ParameterGroup[Guids.Parameters.ТПНаименование], ComparisonOperator.Equal, name);
                filter.Terms.AddTerm(Guids.Parameters.ТПОбозначение.ToString(), ComparisonOperator.Equal, denotation);
                allVariants = reference.Find(filter);
            }

            return FindLastVariantNumber(allVariants, Guids.Parameters.ТПВариант);
        }

        private int FindLastVariantNumber(List<ReferenceObject> allVariants, Guid variantParameterGuid)
        {
            if (!allVariants.Any())
                return 0;

            List<int> variantNumbers = new List<int>();

            foreach (var variant in allVariants)
            {
                string variantName = variant[variantParameterGuid].Value.ToString();

                if (string.IsNullOrEmpty(variantName)) // актуальный вариант
                    continue;

                string[] parts = variantName.Split('.');
                if (parts.Length != 2)
                    continue;

                if (parts[0] != "ИИ")
                    continue;

                int number = 0;
                if (int.TryParse(parts[1], out number))
                    variantNumbers.Add(number);
            }

            return variantNumbers.Any() ? variantNumbers.Max() : 0;
        }

        private bool CheckReferenceObjectInStage(ReferenceObject referenceObject, string stageName)
        {
            if (referenceObject == null)
                throw new ArgumentNullException("referenceObject");

            if (string.IsNullOrEmpty(stageName))
                throw new ArgumentNullException("stageName");

            var schemeStage = referenceObject.SystemFields.Stage;
            if (schemeStage == null)
                return false;

            var stage = schemeStage.Stage;
            if (stage == null)
                return false;

            return stage.Name == stageName;
        }

        private List<ReferenceObject> ChangeStage(List<ReferenceObject> referenceObjects, string stageName)
        {
            Stage stage = Stage.Find(Context.Connection, stageName);
            if (stage == null)
                return null;

            return stage.Set(referenceObjects);
        }
    }
}
