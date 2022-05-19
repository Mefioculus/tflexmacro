using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using TFlex.DOCs.Common;
using TFlex.DOCs.Model;
using TFlex.DOCs.Model.Classes;
using TFlex.DOCs.Model.Desktop;
using TFlex.DOCs.Model.Macros;
using TFlex.DOCs.Model.Macros.ObjectModel;
using TFlex.DOCs.Model.References;
using TFlex.DOCs.Model.References.Files;
using TFlex.DOCs.Model.References.Users;
using TFlex.DOCs.Model.Search;
using TFlex.DOCs.Resources.Strings;

namespace TechnologicalPDM_Changing
{
    public class Macro : MacroProvider
    {
        /// <summary>
        /// Шаблон для формирования наименования варианта при создании изменения ("КИ.Фамилия-Дата")
        /// </summary> 
        private static string _шаблонНаименованияВарианта = "КИ.{0}-{1}";

        /// <summary>
        /// Шаблон даты (добавляется к наименованию варианта, изменения, папки для хранения файлов изменения)
        /// </summary>
        private static string _шаблонДаты = "dd.MM.yy"; //"dd.MM.yy.HH.mm.ss";

        #region Папки

        private static string _папкаАрхивТД = "Архив ТД";
        private static string _папкаАрхивТДПодлинники = Path.Combine(_папкаАрхивТД, "Подлинники");

        #endregion

        #region Guids

        private static class Guids
        {
            public static class References
            {
                public static readonly Guid Изменения = new Guid("c9a4bb1b-cacb-4f2d-b61a-265f1bfc7fb9");
                public static readonly Guid ИнвентарнаяКнига = new Guid("f1d28c5d-6b08-4061-b81f-cde490088e1f");
            }

            public static class Classes
            {
                public static readonly Guid ТехнологическийПроцесс = new Guid("3e93d599-c214-48c8-854f-efe4b475c4d8");
                public static readonly Guid ТехнологическаяОперация = new Guid("f53c9d73-18bb-4c59-a260-61fea65f6ed9");
                public static readonly Guid КарточкаУчетаТехнологическихДокументов = new Guid("d2297aab-a159-45bf-8601-7a7f1f27a38c");
                public static readonly Guid ИзвещениеОбИзменении = new Guid("52ccb35c-67c5-4b82-af4f-e8ceac4e8d02");
                public static readonly Guid ТехнологическоеИзменение = new Guid("a8c61089-dc59-43c8-9296-f2629fec2c4e");
                public static readonly Guid ТехнологическийКомплект = new Guid("dc1cf2a0-6c01-400d-9a42-9642b7496404");
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
                public static readonly Guid ИзвещениеПрименяемость = new Guid("67e9ecca-8db6-4089-8897-29757c5ae46e");
                public static readonly Guid КарточкаТехнологическийКомплект = new Guid("a5a8963d-18f6-47c2-8767-70adc2e1694b");
                public static readonly Guid КарточкаОрганизация = new Guid("b45cbc68-3fff-4f3b-8d90-2f90bf37e9e5");
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
                public static readonly Guid ИИВыпущеноНа = new Guid("7ee36a71-a877-4327-87cc-c1d34c92d9e4");
                public static readonly Guid ОрганизацииКод = new Guid("f63b26ee-a724-4b56-a780-55d5b75566d0");
            }

            #endregion

        }

        private FileReference _fileReference;
        private ReferenceInfo _cardReferenceInfo;
        private Reference _cardReference;
        private ClassObject _cardClassTD;
        private List<FolderObject> _folders = new List<FolderObject>();

        public Macro(MacroContext context)
            : base(context)
        {
        }

        public override void Run()
        {
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

        private ClassObject CardClassTD
        {
            get
            {
                if (_cardClassTD == null)
                    _cardClassTD = CardReference.Classes.Find(Guids.Classes.КарточкаУчетаТехнологическихДокументов);

                return _cardClassTD;
            }
        }

        public void ПринятьНаХранениеТП()
        {
            ReferenceObject currentTP = Context.ReferenceObject;
            if (!currentTP.Class.IsInherit(Guids.Classes.ТехнологическийПроцесс))
                Error(string.Format("Текущий объект '{0}' не является технологическим процессом.", currentTP));

            var rootDocuments = currentTP.GetObjects(Guids.Links.ТПДокументация);
            if (!rootDocuments.Any())
                Error(string.Format("Технологический процесс '{0}' не содержит комплект документов.", currentTP));

            List<ReferenceObject> allTechnologyDocuments = new List<ReferenceObject>();
            allTechnologyDocuments.AddRange(rootDocuments);
            foreach (var rootDocument in rootDocuments)
            {
                var childDocuments = rootDocument.Children.ToList();
                if (childDocuments.Any())
                    allTechnologyDocuments.AddRange(childDocuments);
            }

            // комплект документов
            var setOfDocuments = rootDocuments.FirstOrDefault(document => document.Class.IsInherit(Guids.Classes.ТехнологическийКомплект));
            // проверка на заполнение параметров
            string checkResult = CheckTD(setOfDocuments);

            if (!Question(string.Format("Принять на хранение?{0}{0}{1}", Environment.NewLine, checkResult)))
                return;

            //Поиск учетной карточки для технологического комплекта документов
            ReferenceObject inventoryCard = GetInventoryCard(setOfDocuments);
            // заполнить данные инвентарной карточки
            FillInventoryCard(inventoryCard);
            if (!ShowPropertyDialog(RefObj.CreateInstance(inventoryCard, Context)))
                return;

            // все объекты, у которых меняем стадию на хранение
            List<ReferenceObject> storageObjects = new List<ReferenceObject>();

            // обрабатываем файлы комплекта документов
            if (!DoActionWithSetOfDocument(setOfDocuments, ref storageObjects))
                return;

            var operations = currentTP.Children.ToList();
            operations.ForEach(operation => storageObjects.AddRange(operation.Children.ToList()));

            storageObjects.Add(currentTP);
            storageObjects.AddRange(operations);
            storageObjects.AddRange(allTechnologyDocuments);

            try
            {
                //Перевод на стадию "Хранение" ТП и Комплект документов
                ChangeStage(storageObjects, "Хранение");
            }
            catch
            {
                //%%TODO записать в лог?
                //LogWriter
            }
            currentTP.Reload();
        }

        public void СоздатьИзменениеДляТехпроцесса()
        {
            ReferenceObject currentTP = Context.ReferenceObject;
            if (!currentTP.Class.IsInherit(Guids.Classes.ТехнологическийПроцесс))
                Error(string.Format("Текущий объект '{0}' не является технологическим процессом.", currentTP));

            List<ReferenceObject> modifications = currentTP.GetObjects(Guids.Links.ИзмененияАктуальныйВариантТЭ);
            ReferenceObject existingModification = modifications.FirstOrDefault(modification => !IsReferenceObjectInStage(modification, "Хранение"));
            if (existingModification != null)
            {
                if (Question(string.Format("Для технологического процесса '{0}' уже существует изменение '{1}'. Открыть его свойства?", currentTP, existingModification)))
                    ShowPropertyDialog(RefObj.CreateInstance(existingModification, Context));
                return;
            }

            if (!Question("Создать изменение?"))
                return;

            CreateChangingForTechnologicalProcess(currentTP);

            RefreshReferenceWindow();
        }

        public void СоздатьИзвещениеДляТехнологическогоИзменения()
        {
            if (!Question("Создать извещение?"))
                return;

            var changeNotificationsReferenceInfo = Context.Connection.ReferenceCatalog.Find(SystemParameterGroups.ModificationNotices);
            var changeNotificationsReference = changeNotificationsReferenceInfo.CreateReference();
            var changeNotificationsClass = changeNotificationsReference.Classes.Find(Guids.Classes.ИзвещениеОбИзменении);

            var errors = new StringBuilder();

            ReferenceObject changeNotification = changeNotificationsReference.CreateReferenceObject(changeNotificationsClass);

            var changingList = Context.GetSelectedObjects();
            //Заполняем параметр "Выпущено на", если объект один
            if (changingList.Length == 1)
            {
                ReferenceObject actualVariantTP = changingList[0].GetObject(Guids.Links.ИзмененияАктуальныйВариантТЭ);

                if (actualVariantTP != null)
                {
                    if (actualVariantTP.Class.IsInherit(Guids.Classes.ТехнологическийПроцесс))
                        changeNotification[Guids.Parameters.ИИВыпущеноНа].Value = actualVariantTP[Guids.Parameters.ТПОбозначение].Value.ToString();
                }
            }

            foreach (var changing in changingList)
            {
                //Связываем все выбранные изменения с созданным извещением
                changeNotification.AddLinkedObject(Guids.Links.ИзвещенияИзменения, changing);

                var endVariant = changing.GetObject(Guids.Links.ИзмененияЦелевойВариантТЭ);
                if (endVariant is null)
                {
                    errors.AppendLine($"Изменение '{changing}': не задан целевой вариант технологического элемента.");
                }
                else
                {
                    foreach (var dse in endVariant.GetObjects(Guids.Links.ИзготавливаемыеДСЕ))
                        changeNotification.AddLinkedObject(Guids.Links.ИзвещениеПрименяемость, dse);
                }

                // В целевой варианте изготавливаемых ДСЕ м.б. нет (очистили связь при создании изменения)
                // берем ДСЕ из актуального варианта
                var actualVariant = changing.GetObject(Guids.Links.ИзмененияАктуальныйВариантТЭ);
                if (endVariant is null)
                {
                    errors.AppendLine($"Изменение '{changing}': не задан актуальный вариант технологического элемента.");
                }
                else
                {
                    foreach (var dse in actualVariant.GetObjects(Guids.Links.ИзготавливаемыеДСЕ))
                        changeNotification.AddLinkedObject(Guids.Links.ИзвещениеПрименяемость, dse);
                }
            }

            string errorsText = errors.ToString();
            if (!String.IsNullOrEmpty(errorsText))
                Message(Texts.Error, errorsText);

            ShowPropertyDialog(RefObj.CreateInstance(changeNotification, Context));
        }

        private ReferenceObject GetInventoryCard(ReferenceObject document)
        {
            // Ищем карточку по связи 1:1
            ReferenceObject inventoryCard = document.GetObject(Guids.Links.КарточкаТехнологическийКомплект);
            if (inventoryCard == null)
            {
                if (CardClassTD == null)
                    return null;

                // Ищем в справочнике карточку
                using (Filter filter = new Filter(CardReferenceInfo))
                {
                    // условие по типу карточки
                    filter.Terms.AddTerm(CardReference.ParameterGroup.ClassParameterInfo, ComparisonOperator.IsInheritFrom, CardClassTD);

                    // условие по отсутствию связанного документа у карточки
                    filter.Terms.AddTerm(Guids.Links.КарточкаТехнологическийКомплект.ToString(), ComparisonOperator.IsNull, null);

                    string denotation = document[Guids.Parameters.ТехнологическийДокументОбозначение].Value.ToString();
                    // условие по обозначению документа
                    filter.Terms.AddTerm(CardReference.ParameterGroup[Guids.Parameters.КарточкаОбозначениеДокумента], ComparisonOperator.Equal, denotation);

                    inventoryCard = CardReference.Find(filter).FirstOrDefault();
                }

                // Если карточка по фильтру не найдена
                if (inventoryCard == null)
                {
                    inventoryCard = CardReference.CreateReferenceObject(CardClassTD);
                }
                else
                    inventoryCard.BeginChanges();

                //подключить документ к карточке
                inventoryCard.SetLinkedObject(Guids.Links.КарточкаТехнологическийКомплект, document);
            }
            else
                inventoryCard.BeginChanges();

            return inventoryCard;
        }

        private bool DoActionWithSetOfDocument(ReferenceObject setOfDocument, ref List<ReferenceObject> storageObjects)
        {
            FileObject realFile = setOfDocument.GetObject(Guids.Links.ТехнологическийКомплектПодлинник) as FileObject;
            if (realFile == null)
                return Question(string.Format("У технологического комплекта '{0}' отсутствуют подлинник.{1}Принять его на хранение?'", setOfDocument, Environment.NewLine));

            if (!IsReferenceObjectInStage(realFile, "Утверждено"))
                return Question(string.Format("У технологического комплекта '{0}' подлинник '{1}' не в стадии 'Утверждено'.{2}Принять его на хранение?'", setOfDocument, realFile, Environment.NewLine));

            List<Exception> movingExceptions = new List<Exception>();
            try
            {
                TFlex.DOCs.Model.References.Files.FolderObject folder = GetFolder(_папкаАрхивТДПодлинники);
                if (folder != null)
                {
                    realFile.MoveFileToFolder(folder);
                }
                storageObjects.Add(realFile);
            }
            catch (Exception e)
            {
                movingExceptions.Add(e);
            }

            if (movingExceptions.Any())
            {
                return Question(string.Format("При принятии на хранение произошли ошибки. Изменить стадию на 'Хранение' у объектов:{0}{0}{1}{0}?{2}",
                     Environment.NewLine,
                     string.Join(Environment.NewLine, storageObjects),
                     movingExceptions.Any() ? string.Format("{0}{0}Ошибки перемещения файлов:{0}{1}", Environment.NewLine, string.Join(Environment.NewLine, movingExceptions.Select(exception => exception.Message))) : string.Empty));
            }

            return true;
        }

        private void FillInventoryCard(ReferenceObject inventoryCard, bool fillOnButton = false)
        {
            //заполнение организации
            if (inventoryCard.GetObject(Guids.Links.КарточкаОрганизация) == null)
            {
                string organizationCode = Global["Код предприятия"];
                if (!string.IsNullOrEmpty(organizationCode))
                {
                    var referenceInfo = Context.Connection.ReferenceCatalog.Find(SystemParameterGroups.Companies);
                    using (Filter filter = new Filter(referenceInfo))
                    {
                        var reference = referenceInfo.CreateReference();
                        filter.Terms.AddTerm(reference.ParameterGroup[Guids.Parameters.ОрганизацииКод], ComparisonOperator.Equal, organizationCode);
                        ReferenceObject company = reference.Find(filter).FirstOrDefault();
                        if (company != null)
                            inventoryCard.SetLinkedObject(Guids.Links.КарточкаОрганизация, company);
                    }
                }
            }

            ReferenceObject document = inventoryCard.GetObject(Guids.Links.КарточкаТехнологическийКомплект);
            if (document == null)
            {
                Message("Внимание", "У инвентарной карточки отсутствует связанный технологический комплект.");

                inventoryCard[Guids.Parameters.КарточкаНаименованиеДокумента].Value = string.Empty;
                inventoryCard[Guids.Parameters.КарточкаОбозначениеДокумента].Value = string.Empty;
            }
            else
            {
                inventoryCard[Guids.Parameters.КарточкаНаименованиеДокумента].Value = document[Guids.Parameters.ТехнологическийДокументНаименование];
                inventoryCard[Guids.Parameters.КарточкаОбозначениеДокумента].Value = document[Guids.Parameters.ТехнологическийДокументОбозначение];

                if (!fillOnButton)
                {
                    FileObject realFile = document.GetObject(Guids.Links.ТехнологическийКомплектПодлинник) as FileObject;
                    if (realFile != null)
                    {
                        if (!IsReferenceObjectInStage(realFile, "Утверждено"))
                        {
                            Message("Внимание", string.Format("У технологического комплекта '{0}' подлинник '{1}' не находится в стадии 'Утверждено'.", document, realFile));
                        }
                    }
                    else
                    {
                        Message("Внимание", (string.Format("У технологического комплекта '{0}' отсутствует подлинник.", document)));
                    }
                }
            }
        }

        private string CheckTD(ReferenceObject setOfDocument)
        {
            StringBuilder errors = new StringBuilder();

            string denotation = setOfDocument[Guids.Parameters.ТехнологическийДокументОбозначение].Value.ToString();
            if (string.IsNullOrEmpty(denotation))
                errors.AppendLine("Не задано обозначение.");

            FileObject realFile = setOfDocument.GetObject(Guids.Links.ТехнологическийКомплектПодлинник) as FileObject;
            if (realFile == null)
                errors.AppendLine(string.Format("У технологического комплекта '{0}' отсутствуют подлинник.'", setOfDocument));
            else
            {
                if (IsReferenceObjectInStage(realFile, "Хранение"))
                    errors.AppendLine(string.Format("Файл подлинника '{0}' находится в стадии 'Хранение'.", realFile));
            }
            return errors.ToString();
        }

        private void CreateChangingForTechnologicalProcess(ReferenceObject currentTP)
        {
            string currentDate = DateTime.Now.ToString(_шаблонДаты);
            // Временное наименование варианта
            string newVariantName = string.Format(_шаблонНаименованияВарианта, CurrentUser[User.Fields.LastName.ToString()], currentDate);

            // Создаем вариант технологического процесса
            ReferenceObject newVariantTP = currentTP.CreateCopy(currentTP.Class);
            newVariantTP[Guids.Parameters.ТПВариант].Value = newVariantName;
            // очищаем связь "Изготавливаемые ДСЕ"
            newVariantTP.ClearLinks(Guids.Links.ИзготавливаемыеДСЕ);
            newVariantTP.EndChanges();

            // Копируем комплект документации для нового варианта

            // Копируем документацию для ТП
            ReferenceObject newRootDocumentSet = null;
            var tpDocuments = currentTP.GetObjects(Guids.Links.ТПДокументация);
            // комплект документов
            ReferenceObject rootDocumentSet = tpDocuments.FirstOrDefault(document => document.Parent == null && document.Class.IsInherit(Guids.Classes.ТехнологическийКомплект));
            if (rootDocumentSet != null)
            {
                newRootDocumentSet = rootDocumentSet.CreateCopy();
                newRootDocumentSet.SetLinkedObject(Guids.Links.ТПДокументация, newVariantTP);
                newRootDocumentSet.EndChanges();

                foreach (var document in tpDocuments)
                {
                    if (document.Parent == null)
                        continue;

                    // технологический документ
                    var newDocument = document.CreateCopy(document.Class, newRootDocumentSet);
                    newDocument.SetLinkedObject(Guids.Links.ТПДокументация, newVariantTP);
                    newDocument.EndChanges();
                }
            }

            // Копируем операции ТП
            foreach (var operation in currentTP.Children)
            {
                var variantOperation = operation.CreateCopy(operation.Class, newVariantTP);
                variantOperation.EndChanges();

                if (operation.Class.IsInherit(Guids.Classes.ТехнологическаяОперация))
                {
                    if (newRootDocumentSet != null)
                    {
                        // Копируем документацию для операций
                        foreach (var operationDocument in operation.GetObjects(Guids.Links.ТПДокументация))
                        {
                            var newDocument = operationDocument.CreateCopy(operationDocument.Class, newRootDocumentSet);
                            newDocument.SetLinkedObject(Guids.Links.ТПДокументация, variantOperation);
                            newDocument.EndChanges();
                        }
                    }
                }

                // Копируем переходы ТП
                foreach (var step in operation.Children)
                {
                    var variantStep = step.CreateCopy(step.Class, variantOperation);
                    variantStep.EndChanges();
                }
            }

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

            ReferenceObject actualVariantTP = allVariants.FirstOrDefault(variant => string.IsNullOrEmpty(variant[Guids.Parameters.ТПВариант].Value.ToString()));

            var changingReferenceInfo = Context.Connection.ReferenceCatalog.Find(Guids.References.Изменения);
            var changingReference = changingReferenceInfo.CreateReference();
            var changingClass = changingReference.Classes.Find(Guids.Classes.ТехнологическоеИзменение);

            ReferenceObject changing = changingReference.CreateReferenceObject(changingClass);
            changing[Guids.Parameters.НомерИзменения].Value = newVariantName;

            changing.SetLinkedObject(Guids.Links.ИзмененияЦелевойВариантТЭ, newVariantTP);
            if (actualVariantTP != null)
                changing.SetLinkedObject(Guids.Links.ИзмененияАктуальныйВариантТЭ, actualVariantTP);

            changing.EndChanges();
        }

        private TFlex.DOCs.Model.References.Files.FolderObject GetFolder(string folderRelativePath)
        {
            if (string.IsNullOrEmpty(folderRelativePath))
                return null;

            var haveFolder = _folders.FirstOrDefault(folderObject => folderObject.Path == folderRelativePath);
            if (haveFolder != null)
                return haveFolder;

            if (FileReferenceInstance == null)
                return null;

            var folder = FileReferenceInstance.FindByRelativePath(folderRelativePath) as TFlex.DOCs.Model.References.Files.FolderObject;
            if (folder == null)
            {
                folder = FileReferenceInstance.CreatePath(folderRelativePath, null);
                if (folder == null)
                    Message("Создание папки", TFlex.DOCs.Resources.Strings.Messages.UnableToCreateFolderMessage, folderRelativePath);
                else
                    Desktop.CheckIn(folder, TFlex.DOCs.Resources.Strings.Texts.AutoCreateText, false);
            }

            if (folder != null)
                _folders.Add(folder);

            return folder;
        }

        private List<ReferenceObject> ChangeStage(IEnumerable<ReferenceObject> referenceObjects, string stageName)
        {
            TFlex.DOCs.Model.Stages.Stage stage = TFlex.DOCs.Model.Stages.Stage.Find(Context.Connection, stageName);
            if (stage == null)
                return null;

            return stage.Change(referenceObjects);
        }

        private bool IsReferenceObjectInStage(ReferenceObject referenceObject, string stageName)
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
    }
}
