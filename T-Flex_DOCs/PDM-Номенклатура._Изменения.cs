using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using TFlex.DOCs.Common;
using TFlex.DOCs.Model.Desktop;
using TFlex.DOCs.Model.Macros;
using TFlex.DOCs.Model.Macros.ObjectModel;
using TFlex.DOCs.Model.References;
using TFlex.DOCs.Model.References.Documents;
using TFlex.DOCs.Model.References.Files;
using TFlex.DOCs.Model.References.Nomenclature;
using TFlex.DOCs.Model.References.Users;
using TFlex.DOCs.Model.Search;

namespace PDM_Changing
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

        private static char[] _пропускаемыеВариантыДляСборок = new char[] { 'И', 'К' };

        /// <summary>
        /// Флаг удаления объектов из корзины
        /// </summary>
        private static bool _deleteFromRecycleBin = true;

        /// <summary> Поддерживаемые расширения файлов </summary>
        private static readonly string[] SupportedExtensions = { "tif", "tiff", "pdf" };

        #region Папки

        private static string _папкаАрхив = "Архив";
        private static string _папкаПодлинники = Path.Combine(_папкаАрхив, "Подлинники");
        private static string _папкаОригиналы = Path.Combine(_папкаАрхив, "Оригиналы");
        private static string _папкаОригиналыTF = Path.Combine(_папкаОригиналы, "TF");
        private static string _папкаОригиналыMS = Path.Combine(_папкаОригиналы, "MS");
        private static string _папкаОригиналыДругое = Path.Combine(_папкаОригиналы, "Другое");

        private static string _папкаСдано = "Сдано";

        private static string _временнаяПапкаДляСравненияПодлинников = Path.Combine(Path.GetTempPath(), "Temp DOCs", "CompareTiff");

        #endregion

        #region Guids

        public static readonly string InventoryCardMacroGuid = "a184d20e-8f23-49fa-964f-c61b81b266a6";

        /// <summary> Стадия по умолчанию - "Хранение" </summary>
        private Guid _defaultStageGuid = new Guid("9826e84a-a0a7-404b-bf0e-d61b902e346a");

        private static class Guids
        {
            public static class References
            {
                public static readonly Guid Изменения = new Guid("c9a4bb1b-cacb-4f2d-b61a-265f1bfc7fb9");
                public static readonly Guid Согласование = new Guid("c285ad92-877f-47dc-ad9c-5b3c987f307a");
                public static readonly Guid ИнвентарнаяКнига = new Guid("f1d28c5d-6b08-4061-b81f-cde490088e1f");
            }

            public static class Classes
            {
                public static readonly Guid КарточкаУчетаКонструкторскихДокументов = new Guid("d9f5cb08-8f8b-4e1f-947f-d1250a19a7b6");
                public static readonly Guid ИзвещениеОбИзменении = new Guid("52ccb35c-67c5-4b82-af4f-e8ceac4e8d02");
                public static readonly Guid Изменение = new Guid("f40ea698-bfaa-4143-9534-6276ddec0955");
                public static readonly Guid Согласование = new Guid("dae8dd45-ba88-458d-a4e8-ca384bfeef90");
                public static readonly Guid Формат = new Guid("bb574240-0e6a-4016-bf8c-c860996bed3b");
            }

            public static class Links
            {
                public static readonly Guid ИзмененияАктуальныйВариант = new Guid("d962545d-bccb-40b5-9986-257b57032f6e");
                public static readonly Guid ИзмененияИсходныйВариант = new Guid("48b83092-a645-4dbd-83c0-a3ab0a02ee62");
                public static readonly Guid ИзмененияЦелевойВариант = new Guid("a0e64cef-bf5b-47b9-ae5d-12155c0db936");
                public static readonly Guid ИзмененияФайлы = new Guid("15f76619-7f52-4a56-8498-587dc381e808");
                public static readonly Guid ИзмененияРабочиеФайлы = new Guid("6b65a575-3ca4-4fb0-9bfc-4d1655c2d83e");
                public static readonly Guid ИзвещенияИзменения = new Guid("5e46670a-400c-4e36-bb37-d4d651bdf692");
                public static readonly Guid ИзвещениеПрименяемость = new Guid("67e9ecca-8db6-4089-8897-29757c5ae46e");
                public static readonly Guid ИзмененияСогласования = new Guid("85c0b50f-925c-4ec5-8a14-e53e2295db37");
                public static readonly Guid СогласованиеНоменклатура = new Guid("c218a037-4841-444f-9770-a348d91a5f04");
                public static readonly Guid СогласованиеОтветственный = new Guid("1fd0c38e-d3fd-482e-9124-9a7e603a5460");
                public static readonly Guid КарточкаДокумент = new Guid("d708e1b4-2a1a-499c-aaaf-be5828e6377e");
                public static readonly Guid КарточкаРевизии = new Guid("a193acc6-805f-4e8e-92c8-b7a029792951");
                public static readonly Guid КарточкаФорматы = new Guid("2fc14743-2311-4ed8-9553-8c8715cdbf64");
                public static readonly Guid КарточкаОрганизация = new Guid("b45cbc68-3fff-4f3b-8d90-2f90bf37e9e5");
                public static readonly Guid СвязьЦелеваяРевизия = new Guid("7898c148-6434-494a-bb27-f19f31a5baa2");
            }

            public static class Parameters
            {
                public static readonly Guid НомерИзменения = new Guid("91486563-d044-4045-814b-3432b67812f1");
                public static readonly Guid КарточкаОбозначениеДокумента = new Guid("a0aedf52-cba9-4bd7-afe6-50ee5c100a6d");
                public static readonly Guid КарточкаНаименованиеДокумента = new Guid("795785c4-f1d5-4075-b846-b3abc964b7eb");
                public static readonly Guid ИИВыпущеноНа = new Guid("7ee36a71-a877-4327-87cc-c1d34c92d9e4");
                public static readonly Guid ИИОбозначение = new Guid("b03c9129-7ac3-46f5-bf7d-fdd88ef1ff9a");
                public static readonly Guid ОрганизацииКод = new Guid("f63b26ee-a724-4b56-a780-55d5b75566d0");
            }

            #endregion

        }

        private FileReference _fileReference;

        private List<TFlex.DOCs.Model.References.Files.FolderObject> _folders = new List<TFlex.DOCs.Model.References.Files.FolderObject>();

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

        public void СоздатьИзменение()
        {
            if (!Question("Создать изменение?"))
                return;

            foreach (ReferenceObject referenceObject in Context.GetSelectedObjects())
            {
                CreateChangingForNomenclatureObject(referenceObject as NomenclatureObject);
            }

            RefreshReferenceWindow();
        }

        public void СоздатьИзменениеДляСборки()
        {
            if (!Question("Создать изменение?"))
                return;

            // объекты Согласования (сборки, в которые входит деталь) 
            foreach (ReferenceObject agreement in Context.GetSelectedObjects())
            {
                NomenclatureObject nomenclatureObject = agreement.GetObject(Guids.Links.СогласованиеНоменклатура) as NomenclatureObject;
                if (nomenclatureObject == null)
                    continue;

                if (IsReferenceObjectInStage(nomenclatureObject, "Хранение"))
                    CreateChangingForNomenclatureObject(nomenclatureObject);
            }
        }

        private string DeleteReferenceObjects(List<ReferenceObject> referenceObjects, bool deleteFromRecycleBin = true)
        {
            if (referenceObjects.IsNullOrEmpty())
                return String.Empty;

            var deletingObjects = referenceObjects.Where(referenceObject => referenceObject != null);

            var referenceObjectsWithSupportDesktop = deletingObjects.Where(ro => ro.Reference.ParameterGroup.SupportsDesktop).ToList();

            var referenceObjectsWithoutSupportDesktop = deletingObjects.Except(referenceObjectsWithSupportDesktop).Where(ro => ro.CanDelete).ToArray();
            Array.ForEach(referenceObjectsWithoutSupportDesktop, ro => ro.Delete());

            var addedObjectsWithSupportDesktop = referenceObjectsWithSupportDesktop.Where(ro => ro.LockState == ReferenceObjectLockState.CheckedOutForAdd).ToArray();
            Desktop.UndoCheckOut(addedObjectsWithSupportDesktop);

            var savedObjectsWithSupportDesktop = referenceObjectsWithSupportDesktop.Except(addedObjectsWithSupportDesktop).ToList();
            var checkedOutObjectsWithSupportDesktop = savedObjectsWithSupportDesktop.Where(ro => ro.IsCheckedOutByCurrentUser).ToList();
            Desktop.UndoCheckOut(checkedOutObjectsWithSupportDesktop);

            var canCheckOutObjects = savedObjectsWithSupportDesktop.Where(ro => ro.CanCheckOut).ToList();
            Desktop.CheckOut(canCheckOutObjects, true);

            var objectNames = new StringBuilder();
            foreach (var referenceObject in canCheckOutObjects)
            {
                if (referenceObject is FileObject)
                    objectNames.AppendLine(((FileObject)referenceObject).Path);
                else
                    objectNames.AppendLine(referenceObject.ToString());
            }

            var deletedObjects = Desktop.CheckIn(
                canCheckOutObjects,
                String.Format("Автоматическое удаление объектов:{0}{1}", Environment.NewLine, objectNames),
                false);

            string clearRecycleBinResult = String.Empty;

            if (deleteFromRecycleBin)
            {
                var objectsInRecycleBin = deletingObjects.Union(deletedObjects).Where(desktopObject => ((ReferenceObject)desktopObject).IsInRecycleBin);
                try
                {
                    Desktop.ClearRecycleBin(objectsInRecycleBin);
                }
                catch (EmptyRecycleBinError exception)
                {
                    clearRecycleBinResult = exception.Message;
                }
            }

            return clearRecycleBinResult;
        }

        private void CreateChangingForNomenclatureObject(NomenclatureObject nomenclatureObject)
        {
            if (nomenclatureObject == null)
                return;

            string name = nomenclatureObject[NomenclatureReferenceObject.FieldKeys.Name].GetString();
            string denotation = nomenclatureObject[NomenclatureReferenceObject.FieldKeys.Denotation].GetString();
            string currentDate = DateTime.Now.ToString(_шаблонДаты);

            string folderName = String.Format("{0} [{1}] {2}", denotation, currentDate, name).Trim();
            string folderPath = Path.Combine("Изменения", folderName);
            // делаем проверку, что папка для файлов изменения отсутствует
            TFlex.DOCs.Model.References.Files.FolderObject folderForChangingFiles = FileReferenceInstance.FindByRelativePath(folderPath) as TFlex.DOCs.Model.References.Files.FolderObject;
            if (folderForChangingFiles != null)
            {
                Message("Внимание", "Папка '{0}' уже существует. Невозможно создать изменение для '{1}'.", folderPath, nomenclatureObject.ToString());
                return;
            }

            folderForChangingFiles = FileReferenceInstance.CreatePath(folderPath, null);
            if (folderForChangingFiles == null)
            {
                Message("Ошибка создания папки изменения", TFlex.DOCs.Resources.Strings.Messages.UnableToCreateFolderMessage, folderPath);
                return;
            }
            else
                Desktop.CheckIn(folderForChangingFiles, TFlex.DOCs.Resources.Strings.Texts.AutoCreateText, false);

            // Временное наименование варианта
            string newVariantName = String.Format(_шаблонНаименованияВарианта, CurrentUser[User.Fields.LastName.ToString()], currentDate);

            // Создаем вариант ДСЕ    
            NomenclatureObject newVariant = nomenclatureObject.CreateVariant(newVariantName, true, false);

            // Отключаем все файлы для нового варианта
            ReferenceObject variantLinkedDocument = newVariant.LinkedObject;
            variantLinkedDocument.BeginChanges();
            variantLinkedDocument.ClearLinks(EngineeringDocumentFields.File);
            variantLinkedDocument.EndChanges();

            if (newVariant.Changing)
                newVariant.EndChanges();

            List<ReferenceObject> allVariants = new List<ReferenceObject>();

            var referenceInfo = Context.Connection.ReferenceCatalog.Find(SystemParameterGroups.Nomenclature);
            using (Filter filter = new Filter(referenceInfo))
            {
                var reference = referenceInfo.CreateReference();
                filter.Terms.AddTerm(reference.ParameterGroup[NomenclatureReferenceObject.FieldKeys.Name], ComparisonOperator.Equal, name);
                filter.Terms.AddTerm(reference.ParameterGroup[NomenclatureReferenceObject.FieldKeys.Denotation], ComparisonOperator.Equal, denotation);
                allVariants = reference.Find(filter);
            }

            ReferenceObject actualVariant = null;
            ReferenceObject lastVariant = FindLastAndActualVariant(allVariants, NomenclatureReferenceObject.FieldKeys.VariantName, out actualVariant);

            // Создаем изменение
            var changingReferenceInfo = Context.Connection.ReferenceCatalog.Find(Guids.References.Изменения);
            var changingReference = changingReferenceInfo.CreateReference();
            var changingClass = changingReference.Classes.Find(Guids.Classes.Изменение);

            ReferenceObject changing = changingReference.CreateReferenceObject(changingClass);
            changing[Guids.Parameters.НомерИзменения].Value = newVariantName;

            //устанавливаем целевой вариант - Связь "Целевой документ"
            changing.SetLinkedObject(Guids.Links.ИзмененияЦелевойВариант, newVariant);
            if (lastVariant != null)
                changing.SetLinkedObject(Guids.Links.ИзмененияИсходныйВариант, lastVariant);
            if (actualVariant != null)
                changing.SetLinkedObject(Guids.Links.ИзмененияАктуальныйВариант, actualVariant);

            List<int> idList = new List<int>();
            foreach (ReferenceObject assembly in nomenclatureObject.Parents)
            {
                int id = assembly.SystemFields.Id;
                if (idList.Contains(id))
                    continue;

                idList.Add(id);

                string assemblyVariant = assembly[NomenclatureReferenceObject.FieldKeys.VariantName].GetString();
                if (String.IsNullOrEmpty(assemblyVariant) || !_пропускаемыеВариантыДляСборок.Contains(assemblyVariant.FirstOrDefault()))
                {
                    //Справочник "Согласование", тип "Согласование"
                    var agreementReferenceInfo = Context.Connection.ReferenceCatalog.Find(Guids.References.Согласование);
                    var agreementReference = agreementReferenceInfo.CreateReference();
                    var agreementClass = agreementReference.Classes.Find(Guids.Classes.Согласование);

                    var agreement = agreementReference.CreateReferenceObject(agreementClass);
                    //Связь "Номенклатура"
                    agreement.SetLinkedObject(Guids.Links.СогласованиеНоменклатура, assembly);

                    // Автора последнего изменения сборки
                    User lastAssemblyEditor = assembly.SystemFields.Editor;
                    if (lastAssemblyEditor != null)
                    {
                        //заполняем ответственного для согласования, если это не Система
                        if (!lastAssemblyEditor.IsSystem)
                            agreement.SetLinkedObject(Guids.Links.СогласованиеОтветственный, lastAssemblyEditor);
                    }

                    agreement.EndChanges();

                    // подключить объект "Согласование" к изменению
                    changing.AddLinkedObject(Guids.Links.ИзмененияСогласования, agreement);
                }
            }

            // копируем файлы для изменения
            CopyFilesForChanging(nomenclatureObject.LinkedObject, changing, folderForChangingFiles);

            changing.EndChanges();
        }

        private void CopyFilesForChanging(ReferenceObject linkedReferenceObject, ReferenceObject changing, TFlex.DOCs.Model.References.Files.FolderObject folderForChangingFiles)
        {
            if (linkedReferenceObject == null || changing == null)
                return;

            var originalDocument = linkedReferenceObject as EngineeringDocumentObject;
            if (originalDocument == null)
                return;

            // Получаем файлы-оригиналы
            List<FileObject> originalFiles = originalDocument.GetFiles().Where(file => file.Path.ToString().ToLower().Contains(@"архив\оригиналы")).ToList();
            if (!originalFiles.Any())
                return;

            List<Exception> exceptions = new List<Exception>();
            List<FileObject> filesInChangingFolder = folderForChangingFiles.Children.OfType<FileObject>().ToList();
            List<FileObject> checkInFiles = new List<FileObject>();
            foreach (FileObject file in originalFiles)
            {
                string fileName = file.Name.GetString();

                switch (file.Class.Extension)
                {
                    case ("tiff"):
                    case ("tif"):
                    case ("pdf"):
                        continue;
                    default:
                        FileObject copiedFile = filesInChangingFolder.FirstOrDefault(fileObject => fileObject.Name.GetString().ToLower() == fileName.ToLower());
                        if (copiedFile == null)
                        {
                            try
                            {
                                copiedFile = file.CreateCopy(fileName, folderForChangingFiles, null);
                                copiedFile.EndChanges();
                                checkInFiles.Add(copiedFile);
                            }
                            catch (Exception e)
                            {
                                if (copiedFile != null && copiedFile.Changing)
                                    copiedFile.CancelChanges();

                                exceptions.Add(e);
                            }
                        }
                        continue;
                }
            }

            if (exceptions.Any())
            {
                Message("Внимание", "При копировании файлов для изменения '{0}' произошли ошибки.\r\n{1}", changing, String.Join("\r\n", exceptions.Select(exception => exception.Message)));
                return;
            }

            // Сохранение
            if (checkInFiles.Any())
            {
                try
                {
                    IEnumerable<DesktopObject> checkedInFiles = Desktop.CheckIn(checkInFiles,
                        String.Format("Автоматическое копирование файлов объекта '{0}' в папку '{1}' для изменения '{2}'",
                        linkedReferenceObject, folderForChangingFiles.Path, changing[Guids.Parameters.НомерИзменения].Value), false);

                    foreach (FileObject file in checkInFiles)
                    {
                        // подключить копию файла к изменению
                        changing.AddLinkedObject(Guids.Links.ИзмененияРабочиеФайлы, file);
                    }

                    Desktop.CheckOut(checkedInFiles, false);
                }
                catch (Exception e)
                {
                    Message("Внимание", e.Message);
                }
            }
        }

        private ReferenceObject FindLastAndActualVariant(List<ReferenceObject> allVariants, Guid variantParameterGuid, out ReferenceObject actualVariantTP)
        {
            actualVariantTP = null;

            if (!allVariants.Any())
                return null;

            ReferenceObject lastVariantTP = null;

            List<int> variantNumbers = new List<int>();

            foreach (var variant in allVariants)
            {
                string variantName = variant[variantParameterGuid].Value.ToString();

                if (String.IsNullOrEmpty(variantName))
                {
                    actualVariantTP = variant;
                    continue;
                }

                string[] parts = variantName.Split('.');
                if (parts.Length != 2)
                    continue;

                if (parts[0] != "ИИ")
                    continue;

                int number = 0;
                if (int.TryParse(parts[1], out number))
                    variantNumbers.Add(number);
            }

            if (variantNumbers.Any())
            {
                string lastVariantName = String.Format("ИИ.{0}", variantNumbers.Max());
                lastVariantTP = allVariants.FirstOrDefault(variant => variant[variantParameterGuid].Value.ToString() == lastVariantName);
            }

            return lastVariantTP;
        }

        public void ПринятьНаХранение()
        {
            var nomenclatureObject = Context.ReferenceObject as NomenclatureObject;
            if (nomenclatureObject == null)
                Error(String.Format("Текущий объект '{0}' не является объектом ЭСИ.", Context.ReferenceObject));

            //Поиск учетной карточки документа
            ReferenceObject inventoryCard = GetInventoryCard(nomenclatureObject);
            if (inventoryCard == null)
                Error(String.Format("Ошибка создания инвентарной карточки для документа '{0}'.", nomenclatureObject));

            //Заполнение данных инвентарной карточки
            FillInventoryCard(inventoryCard);

            if (!ShowPropertyDialog(RefObj.CreateInstance(inventoryCard, Context)))
            {
                return;
            }

            // все объекты, у которых меняем стадию на хранение
            List<ReferenceObject> storageObjects = new List<ReferenceObject> { nomenclatureObject };

            var document = nomenclatureObject.LinkedObject as EngineeringDocumentObject;
            var documentFiles = document.GetFiles();

            // обрабатываем файлы документа
            if (!DoActionWithDocumentFiles(document, documentFiles, ref storageObjects))
                return;

            //Перевод на стадию "Хранение" ДСЕ, Документа, файлов
            ChangeStage(storageObjects, "Хранение");

            document.Reload();
        }

        private ReferenceObject GetInventoryCard(NomenclatureObject nomenclatureObject)
        {
            var document = nomenclatureObject.LinkedObject as EngineeringDocumentObject;
            if (document is null)
                Error(String.Format("У объекта ЭСИ '{0}' отсутствует связанный документ.", nomenclatureObject));

            var inventoryCard = (ReferenceObject)RunMacro(InventoryCardMacroGuid, "FindOrCreateDocumentInventoryCard", document, false);
            return inventoryCard;
        }

        /// <summary>
        /// Обработка файлов документа (создание дубликатов, удаление рабочих файлов)
        /// </summary>
        /// <param name="document">Документ</param>
        /// <param name="documentType">Тип документа</param>
        /// <param name="documentFiles">Обрабатываемые файлы</param>
        /// <param name="storageObjects">Список объектов на хранение</param>
        /// <returns>Результат обработки файлов</returns>
        private bool DoActionWithDocumentFiles(ReferenceObject document, List<FileObject> documentFiles, ref List<ReferenceObject> storageObjects)
        {
            if (!documentFiles.Any())
                return Question(String.Format("У документа '{0}' отсутствуют файлы.{1}Принять его на хранение?'", document, Environment.NewLine));

            // берем все файлы в стадии Утверждено
            var approvedFiles = documentFiles.Where(file => IsReferenceObjectInStage(file, "Утверждено")).ToList();
            if (!approvedFiles.Any())
                return Question(String.Format("У документа '{0}' отсутствуют файлы в стадии 'Утверждено'.{1}Принять его на хранение?", document, Environment.NewLine));

            List<Exception> duplicationExceptions = new List<Exception>();
            List<Exception> movingExceptions = new List<Exception>();
            List<string> warnings = new List<string>();

            //Перенос утверждённых файлов в архив
            foreach (FileObject file in approvedFiles)
            {
                // выбор папки для хранения утвержденных файлов
                TFlex.DOCs.Model.References.Files.FolderObject folder = null;
                switch (file.Class.Extension.ToLower())
                {
                    case ("tif"):
                    case ("tiff"):
                    case ("pdf"):
                        folder = GetFolder(_папкаПодлинники);
                        break;

                    case ("grb"):
                        folder = GetFolder(_папкаОригиналыTF);
                        break;

                    case ("xls"):
                    case ("xlsx"):
                    case ("doc"):
                    case ("docx"):
                        folder = GetFolder(_папкаОригиналыMS);
                        break;

                    default:
                        folder = GetFolder(_папкаОригиналыДругое);
                        break;
                }

                FileObject duplicatedFile = null;

                try
                {
                    // делаем дубликат файла в соответствующей папке
                    duplicatedFile = file.CreateFileDuplicate(folder);
                    // добавляем в список объектов для смены стадии на "Хранение"
                    storageObjects.Add(duplicatedFile);
                }
                catch (Exception e)
                {
                    duplicationExceptions.Add(e);
                }

                // после создания дубликата файла определяем, что делать с исходным файлом
                if (duplicatedFile != null)
                {
                    string previousLocalPath = file.LocalPath;
                    try
                    {
                        // перемещаем файл в папку "Сдано", если он еще в ней не находится
                        if (file.Parent.Name.GetString() != _папкаСдано)
                        {
                            folder = GetFolder(Path.Combine(file.Parent.Path, _папкаСдано));

                            if (folder != null)
                            {
                                file.MoveFileToFolder(folder);
                            }
                        }
                    }
                    catch (IOException)
                    {
                        warnings.Add(previousLocalPath);
                    }
                    catch (Exception e)
                    {
                        movingExceptions.Add(e);
                    }
                }
            }

            if (duplicationExceptions.Any() || movingExceptions.Any())
            {
                return Question(String.Format("При принятии на хранение произошли ошибки. Изменить стадию на 'Хранение' у объектов:{0}{0}{1}{0}?{2}{3}",
                     Environment.NewLine,
                     String.Join(Environment.NewLine, storageObjects),
                     duplicationExceptions.Any() ? String.Format("{0}{0}Ошибки создания дубликатов:{0}{1}", Environment.NewLine, String.Join(Environment.NewLine, duplicationExceptions.Select(exception => exception.Message))) : String.Empty,
                     movingExceptions.Any() ? String.Format("{0}{0}Ошибки перемещения файлов:{0}{1}", Environment.NewLine, String.Join(Environment.NewLine, movingExceptions.Select(exception => exception.Message))) : String.Empty));
            }
            else if (warnings.Any())
            {
                Message("Внимание", "Не удалось удалить файлы:{0}{1}", Environment.NewLine, String.Join(Environment.NewLine, warnings));
            }

            // Обработка файлов из рабочей зоны при принятии на хранение
            DoActionWithWorkingFiles(approvedFiles);

            return true;
        }

        private void DoActionWithWorkingFiles(List<FileObject> files)
        {
            List<ReferenceObject> changingFiles = ChangeStage(files, "Корректировка");
            if (!changingFiles.Any())
                return;

            // По умолчанию удаляем все файлы
            string objectErrors = DeleteReferenceObjects(changingFiles, _deleteFromRecycleBin);
            if (!String.IsNullOrEmpty(objectErrors))
                Message("Сообщение", objectErrors);
        }

        public void ОткрытьСвойстваСвязаннойИнвентарнойКарточки()
        {
            var focusedNomenclatureObject = Context.ReferenceObject as NomenclatureObject;
            var document = focusedNomenclatureObject?.LinkedObject;
            if (document is null)
                Error(String.Format("У объекта ЭСИ '{0}' отсутствует связанный документ.", focusedNomenclatureObject));

            document.TryGetObject(Guids.Links.КарточкаРевизии, out var inventoryCard);
            if (inventoryCard is null)
                document.TryGetObject(Guids.Links.КарточкаДокумент, out inventoryCard);

            if (inventoryCard is null)
                Error(String.Format("Инвентарная карточка для объекта ЭСИ '{0}' не найдена.", focusedNomenclatureObject));

            if (ShowPropertyDialog(RefObj.CreateInstance(inventoryCard, Context)))
                focusedNomenclatureObject.Reload();
        }

        public void ЗаполнитьТекущуюИнвентарнуюКарточку()
        {
            if (Question("Заполнить инвентарную карточку автоматически?"))
                FillInventoryCard(Context.ReferenceObject, true);
        }

        private void FillInventoryCard(ReferenceObject inventoryCard, bool fillOnButton = false)
        {
            //заполнение организации
            if (inventoryCard.GetObject(Guids.Links.КарточкаОрганизация) == null)
            {
                string organizationCode = Global["Код предприятия"];
                if (!String.IsNullOrEmpty(organizationCode))
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

            inventoryCard.TryGetObjects(Guids.Links.КарточкаРевизии, out var revisions);
            if (!revisions.IsNullOrEmpty())
            {
                foreach (var revision in revisions.OfType<EngineeringDocumentObject>())
                {
                    var revisionFiles = revision.GetFiles();

                    // файл-подлинник ревизии
                    var realFile = revisionFiles
                           .Where(file => IsReferenceObjectInStage(file, "Утверждено"))
                           .FirstOrDefault(file => SupportedExtensions.Contains(file.Class.Extension.ToLower()));

                    if (realFile is null)
                        continue;

                    string pathToRealFile = String.Empty;

                    string fileExtension = realFile.Class.Extension.ToLower();

                    string tempFolder = Path.Combine(Path.GetTempPath(), "Temp DOCs");
                    if (!Directory.Exists(tempFolder))
                        Directory.CreateDirectory(tempFolder);

                    pathToRealFile = String.Format(@"{0}\{1}.{2}", tempFolder, Guid.NewGuid(), fileExtension);

                    realFile.GetHeadRevision(pathToRealFile);

                    try
                    {
                        РазборПоФорматам(pathToRealFile, Объект.CreateInstance(inventoryCard, Context), false, Объект.CreateInstance(revision, Context));
                    }
                    catch (Exception)
                    {
                        // %%TODO
                    }
                }
            }
            else
            {
                // файл-подлинник
                FileObject realFile = null;

                inventoryCard.TryGetObject(Guids.Links.КарточкаДокумент, out var referenceObject);

                if (!(referenceObject is EngineeringDocumentObject document))
                {
                    Message("Внимание", "У инвентарной карточки отсутствует связанный документ.");

                    inventoryCard[Guids.Parameters.КарточкаНаименованиеДокумента].Value = String.Empty;
                    inventoryCard[Guids.Parameters.КарточкаОбозначениеДокумента].Value = String.Empty;
                }
                else
                {
                    inventoryCard[Guids.Parameters.КарточкаНаименованиеДокумента].Value = document[EngineeringDocumentFields.Name];
                    inventoryCard[Guids.Parameters.КарточкаОбозначениеДокумента].Value = document[SpecificationFields.Denotation];

                    List<FileObject> documentFiles = document.GetFiles();
                    if (documentFiles.Any())
                    {
                        // все файлы в стадии Утверждено
                        var approvedFiles = documentFiles.Where(file => IsReferenceObjectInStage(file, "Утверждено"));
                        if (!approvedFiles.Any())
                        {
                            if (!fillOnButton)
                                Message("Внимание", String.Format("У документа '{0}' отсутствуют файлы в стадии 'Утверждено'.", document));
                        }
                        else
                        {
                            // определение файла подлинника
                            realFile = approvedFiles.FirstOrDefault(file => SupportedExtensions.Contains(file.Class.Extension.ToLower()));
                            if (realFile == null)
                            {
                                if (!fillOnButton)
                                    Message("Внимание", "Среди утверждённых файлов отсутствует файл подлинника в формате tiff, pdf.");
                            }
                        }
                    }
                    else
                    {
                        if (!fillOnButton)
                            Message("Внимание", (String.Format("У документа '{0}' отсутствуют файлы.", document)));
                    }
                }

                // обновить форматы согласно подлиннику
                bool updateFormatsByRealFile = true;

                // определяем, заполнен ли список форматов в карточке
                if (fillOnButton && inventoryCard.Links.ToMany[Guids.Links.КарточкаФорматы].Any())
                {
                    updateFormatsByRealFile =
                        realFile == null ?
                        Question("Отсутствует файл подлинника. Очистить список форматов?") :
                        Question(String.Format("Перезаписать список форматов согласно подлиннику '{0}'?", realFile.Path));
                }

                // согласно подлиннику будем полностью обновлять список форматов в карточке с дальнейшим их пересчётом
                // если путь к подлиннику не задан, очищаем список форматов
                if (updateFormatsByRealFile)
                {
                    string pathToRealFile = String.Empty;
                    if (realFile != null)
                    {
                        string fileExtension = realFile.Class.Extension.ToLower();

                        string tempFolder = Path.Combine(Path.GetTempPath(), "Temp DOCs");
                        if (!Directory.Exists(tempFolder))
                            Directory.CreateDirectory(tempFolder);

                        pathToRealFile = String.Format(@"{0}\{1}.{2}", tempFolder, Guid.NewGuid(), fileExtension);

                        realFile.GetHeadRevision(pathToRealFile);
                    }

                    try
                    {
                        РазборПоФорматам(pathToRealFile, Объект.CreateInstance(inventoryCard, Context));
                    }
                    catch (Exception e)
                    {
                        //throw e;
                        throw new Exception(String.Format("Ошибка определения форматов подлинника{0}'{1}'{2}{3}.",
                            Environment.NewLine,
                            pathToRealFile,
                            Environment.NewLine,
                            e.InnerException != null ? e.InnerException.Message : e.Message), e);
                    }
                }
                else
                {
                    // просто формируем строку форматов и пересчитываем кол-во листов по списку форматов в карточке
                    RecalcFormats(inventoryCard);
                }
            }
        }

        private TFlex.DOCs.Model.References.Files.FolderObject GetFolder(string folderRelativePath)
        {
            if (String.IsNullOrEmpty(folderRelativePath))
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

            if (String.IsNullOrEmpty(stageName))
                throw new ArgumentNullException("stageName");

            var schemeStage = referenceObject.SystemFields.Stage;
            if (schemeStage == null)
                return false;

            var stage = schemeStage.Stage;
            if (stage == null)
                return false;

            return stage.Name == stageName;
        }

        /// <summary>
        /// Вычисление форматов и кол-ва листов подлинника для заполнения данных инвентарной карточки
        /// </summary>
        /// <param name="путьКФайлу"></param>
        /// <param name="инвентарнаяКарточка"></param>
        public void РазборПоФорматам(string путьКФайлу, Объект инвентарнаяКарточка, bool очиститьСписокФорматов = true, Объект ревизия = null)
            => RunMacro(InventoryCardMacroGuid, "DisassembleFormats", путьКФайлу, инвентарнаяКарточка, очиститьСписокФорматов, ревизия);

        private void RecalcFormats(ReferenceObject inventoryCard)
            => RunMacro(InventoryCardMacroGuid, "RecalcFormats", inventoryCard, null);

        public void СоздатьИзвещение()
        {
            if (!Question("Создать извещение?"))
                return;

            //Справочник "Извещения об изменениях", тип "Извещение об изменении"
            var changeNotificationsReferenceInfo = Context.Connection.ReferenceCatalog.Find(SystemParameterGroups.ModificationNotices);
            var changeNotificationsReference = changeNotificationsReferenceInfo.CreateReference();
            var changeNotificationsClass = changeNotificationsReference.Classes.Find(Guids.Classes.ИзвещениеОбИзменении);

            ReferenceObject changeNotification = changeNotificationsReference.CreateReferenceObject(changeNotificationsClass);

            var changingList = Context.GetSelectedObjects();

            //Заполняем параметр "Выпущено на", если объект один
            if (changingList.Length == 1)
            {
                ReferenceObject actualVariant = changingList[0].GetObject(Guids.Links.ИзмененияАктуальныйВариант);
                if (actualVariant != null)
                {
                    changeNotification[Guids.Parameters.ИИВыпущеноНа].Value = actualVariant[NomenclatureReferenceObject.FieldKeys.Denotation].Value.ToString();
                }
            }

            List<int> idList = new List<int>();

            //Связываем все выбранные изменения с созданным извещением
            foreach (var changing in changingList)
            {
                changeNotification.AddLinkedObject(Guids.Links.ИзвещенияИзменения, changing);

                foreach (var agreement in changing.Links.ToMany[Guids.Links.ИзмененияСогласования])
                {
                    var assembly = agreement.GetObject(Guids.Links.СогласованиеНоменклатура);
                    if (assembly == null)
                        continue;

                    int id = assembly.SystemFields.Id;
                    if (idList.Contains(id))
                        continue;

                    idList.Add(id);

                    // подключаем к текущему извещению все изменения сборок, у которых еще нет извещения
                    foreach (var assemblyChanging in assembly.Links.ToMany[Guids.Links.ИзмененияАктуальныйВариант])
                    {
                        if (assemblyChanging.GetObject(Guids.Links.ИзвещенияИзменения) == null)
                            changeNotification.AddLinkedObject(Guids.Links.ИзвещенияИзменения, assemblyChanging);
                    }

                    // заполняем в ИИ связь "Применяемость" номенклатурными объектами из "Изменение - Согласования - Номенклатура"
                    if (IsReferenceObjectInStage(assembly, "Хранение") || IsReferenceObjectInStage(assembly, "Утверждено"))
                    {
                        changeNotification.AddLinkedObject(Guids.Links.ИзвещениеПрименяемость, assembly);
                    }
                }

                if (changing.TryGetObject(Guids.Links.СвязьЦелеваяРевизия, out var targetRevision))
                {
                    if (targetRevision != null)
                    {
                        var childDocs = new List<ReferenceObject>();
                        targetRevision.Children.Reload();
                        foreach (var child in targetRevision.Children)
                        {
                            if (child.LockState != ReferenceObjectLockState.CheckedOutForAdd)
                                continue;

                            var link = child.GetParentLink(targetRevision) as NomenclatureHierarchyLink;
                            string bomSection = link.BomSection.Value;
                            if (bomSection.ToLower() == "документация")
                                childDocs.Add(child);
                        }

                        if (childDocs.Any())
                        {
                            foreach (var child in childDocs)
                            {
                                if (child.TryGetObjects(Guids.Links.СвязьЦелеваяРевизия, out var linkedModifications))
                                {
                                    foreach (var modification in linkedModifications.Where(mod =>
                                    mod.SystemFields.Stage.Guid != _defaultStageGuid
                                    && mod.GetObject(Guids.Links.ИзвещенияИзменения) is null))
                                    {
                                        changeNotification.AddLinkedObject(Guids.Links.ИзвещенияИзменения, modification);
                                    }
                                }
                            }
                        }
                    }
                }
            }

            ShowPropertyDialog(RefObj.CreateInstance(changeNotification, Context));
        }

        public void СформироватьНомерИзвещения()
        {
            // Параметр "Обозначение" (в диалоге наименование "Извещение")
            if (!String.IsNullOrEmpty(Context.ReferenceObject[Guids.Parameters.ИИОбозначение].GetString()))
                return;

            int currentNumber = Global["Номер извещения"];
            currentNumber += 1;
            Global["Номер извещения"] = currentNumber;
            Context.ReferenceObject[Guids.Parameters.ИИОбозначение].Value = currentNumber.ToString().PadLeft(4, '0');
        }

        public Объекты ПолучитьФайлыИзДокументаДляСогласования(Объект согласование)
        {
            ReferenceObject agreement = (ReferenceObject)согласование;

            // Связь "Изменение"
            ReferenceObject changing = agreement.GetObject(Guids.Links.ИзмененияСогласования);
            if (changing == null)
                return null;

            // Связь "Номенклатура"
            NomenclatureObject nomenclatureObject = agreement.GetObject(Guids.Links.СогласованиеНоменклатура) as NomenclatureObject;
            if (nomenclatureObject == null)
                return null;

            // Связанный документ
            var documentObject = nomenclatureObject.LinkedObject as EngineeringDocumentObject;
            if (documentObject == null)
                return null;

            // Файлы документа из папки "архив\оригиналы"
            IEnumerable<FileObject> documentFiles = documentObject.GetFiles().Where(file => file.Path.Value.ToLower().Contains(_папкаОригиналы.ToLower()));
            if (!documentFiles.Any())
                return null;

            // Связь "Файлы" в изменении
            List<ReferenceObject> filesInChanging = changing.GetObjects(Guids.Links.ИзмененияРабочиеФайлы);
            ReferenceObject fileInChanging = filesInChanging.FirstOrDefault();
            if (fileInChanging == null)
                return null;

            // Папка, в которую нужно скопировать файлы сборки
            // Изменения\string.Format("{0} [{1}] {2}", документВариантаДСЕ["Обозначение"], DateTime.Now.Date.ToString("dd.MM.yy"), документВариантаДСЕ["Наименование"]);
            Объект папка = НайтиОбъект(SystemParameterGroups.Files.ToString(), String.Format("[Дочерние объекты].[ID] = '{0}'", fileInChanging.SystemFields.Id));
            if (папка == null)
                return null;

            TFlex.DOCs.Model.References.Files.FolderObject targetFolder = (TFlex.DOCs.Model.References.Files.FolderObject)папка;

            Объекты файлыДляСогласования = null;
            using (ReferenceObjectSaveSet copiedFilesSaveSet = new ReferenceObjectSaveSet())
            {
                try
                {
                    foreach (var documentFile in documentFiles)
                    {
                        documentFile.CreateCopy(documentFile.Name.GetString(), targetFolder, copiedFilesSaveSet);
                    }

                    // применить изменения ко всем скопированным файлам
                    if (copiedFilesSaveSet.Changing)
                        copiedFilesSaveSet.EndChanges();

                    List<ReferenceObject> copiedFiles = copiedFilesSaveSet.ToList();

                    // Сохранение
                    Desktop.CheckIn(copiedFiles, String.Format("Автоматическое копирование файлов объекта '{0}' из папки '{1}' в папку '{2}'", nomenclatureObject, _папкаОригиналы, targetFolder.Path), false);
                    //Перевод скопированных файлов на стадию Хранение
                    ChangeStage(copiedFiles, "Хранение");

                    // Подключаем скопированные файлы сборки к изменению
                    changing.BeginChanges();
                    foreach (ReferenceObject file in copiedFiles)
                    {
                        // Связь "Изменение->Файлы"
                        changing.AddLinkedObject(Guids.Links.ИзмененияРабочиеФайлы, file);
                    }
                    changing.EndChanges();

                    файлыДляСогласования = new Объекты(copiedFiles, Context);
                }
                catch
                {
                    if (copiedFilesSaveSet != null && copiedFilesSaveSet.Changing)
                        copiedFilesSaveSet.CancelChanges();

                    throw;
                }
            }

            return файлыДляСогласования;
        }

        public void СравнитьПодлинникИзмененияСАктуальнымВариантом()
        {
            CompareRealFiles(true);
        }

        public void СравнитьПодлинникИзмененияСИсходнымВариантом()
        {
            CompareRealFiles(false);
        }

        /// <summary>
        /// Сравнение подлинников
        /// </summary>
        /// <param name="compareWithActualVariant">Сравнить с актуальным вариантом</param>
        private void CompareRealFiles(bool compareWithActualVariant)
        {
            ReferenceObject changing = (ReferenceObject)ТекущийОбъект;

            FileObject changingFile = null;
            // если стадия изменения Хранение - берем файл по связи Изменения-Файлы
            // если стадия другая - берем файл по связи Изменения-Рабочие Файлы
            string commonFileName = String.Empty;
            bool getFromWorkingLink = false;
            if (IsReferenceObjectInStage(changing, "Хранение") || IsReferenceObjectInStage(changing, "Аннулировано"))
            {
                changingFile = FindTiffFile(changing, Guids.Links.ИзмененияФайлы, ref commonFileName);
            }
            else
            {
                getFromWorkingLink = true;
                changingFile = FindTiffFile(changing, Guids.Links.ИзмененияРабочиеФайлы, ref commonFileName);
            }

            if (changingFile == null)
            {
                Error("В изменении '{0}' не найден файл подлинника формата TIFF для сравнения.", changing);
            }

            // папка для сравнения файлов
            string changingFolder = changingFile.Parent != null ? changingFile.Parent.LocalPath : _временнаяПапкаДляСравненияПодлинников;
            if (!Directory.Exists(changingFolder))
                Directory.CreateDirectory(changingFolder);

            string changingFilePath = Path.Combine(changingFolder, "ПодлинникИзменения.tiff");
            string variantFilePath = Path.Combine(changingFolder, "ПодлинникВарианта.tiff");

            // Получаем версию файла изменения
            if (changingFile.IsAdded || changingFile.IsCheckedOutByCurrentUser)
            {
                // берем локальный файл
                File.Copy(changingFile.LocalPath, changingFilePath, true);
            }
            else
            {
                // загружаем с сервера версию файла изменения
                changingFile.GetFileVersion(changingFilePath, changingFile.SystemFields.Version);
            }

            new FileInfo(changingFilePath).IsReadOnly = false;

            FileObject variantFile = null;
            int fileVersion = 1;

            // если сравнение с актуальным вариантом или сравнение с предыдущим вариантом и файл изменения взят из связи "Рабочие файлы"
            // сравниваем с актуальным вариантом
            if (compareWithActualVariant || (!compareWithActualVariant && getFromWorkingLink))
            {
                NomenclatureObject actualVariant = changing.GetObject(Guids.Links.ИзмененияАктуальныйВариант) as NomenclatureObject;
                if (actualVariant == null)
                    Error("Отсутствует актуальный вариант.");

                ReferenceObject actualDocument = actualVariant.LinkedObject;
                if (actualDocument == null)
                    Error("Отсутствует связанный документ номенклатурного объекта '{0}'.", actualVariant);

                // Для поиска файла подлинника будем проверять стадию ("Хранение")
                variantFile = FindTiffFile(actualDocument, EngineeringDocumentFields.File, ref commonFileName, true);
                if (variantFile == null)
                    Error("В актуальном варианте '{0}' не найден файл подлинника формата TIFF для сравнения.", actualVariant);

                fileVersion = variantFile.SystemFields.Version;
                variantFile.GetHeadRevision(variantFilePath);
            }
            else
            {
                // сравнение с предыдущим вариантом и файл изменения взят из связи "Файлы" 
                variantFile = changingFile;

                // берем предыдущую версию файла
                fileVersion = changingFile.SystemFields.Version - 1;

                changingFile.GetFileVersion(variantFilePath, fileVersion);
            }

            new FileInfo(variantFilePath).IsReadOnly = false;

            //CompareTiff.exe
            string directoryPath = Path.GetDirectoryName(typeof(MacroContext).Assembly.Location);
            string tiffComparatorSourcePath = Path.Combine(directoryPath, "CompareTiff.exe");
            if (!File.Exists(tiffComparatorSourcePath))
                Error("Файл '{0}' не найден.", tiffComparatorSourcePath);

            if (!Question(String.Format("Сравнить файлы?{0}{0}'{1}', версия {2}, файл варианта{0}{0}'{3}', версия {4}, файл изменения", Environment.NewLine, variantFile.Path, fileVersion, changingFile.Path, changingFile.SystemFields.Version)))
                return;

            Process compareTiffProcess = Process.Start(tiffComparatorSourcePath, String.Format("-multi \"{0}\" \"{1}\"", variantFilePath, changingFilePath));
            compareTiffProcess.WaitForExit();
        }

        private FileObject FindTiffFile(ReferenceObject referenceObject, Guid filesLinkGuid, ref string commonName, bool checkStage = false)
        {
            commonName = String.Empty;
            foreach (FileObject file in referenceObject.GetObjects(filesLinkGuid).OfType<FileObject>())
            {
                string fileName = file.Name.GetString();
                if (IsTiffFile(fileName))
                {
                    if (checkStage)
                    {
                        if (!IsReferenceObjectInStage(file, "Хранение"))
                            continue;
                    }

                    fileName = Path.GetFileNameWithoutExtension(fileName);

                    if (String.IsNullOrEmpty(commonName))
                        commonName = GetCommonFileName(fileName);

                    if (!String.IsNullOrEmpty(commonName))
                    {
                        if (fileName == commonName)
                            return file;
                    }
                }
            }

            return null;
        }

        private bool IsTiffFile(string filePath)
        {
            if (String.IsNullOrEmpty(filePath))
                return false;

            string extension = Path.GetExtension(filePath).TrimStart('.').ToLower();
            if (extension == "tif" || extension == "tiff")
                return true;

            return false;
        }

        private string GetCommonFileName(string fileName)
        {
            // д.б. формат: наименованиеФайла_номер
            int indexSeparator = fileName.LastIndexOf('_');
            if (indexSeparator == -1)
                return fileName;

            return indexSeparator > 0 ? fileName.Substring(0, indexSeparator) : String.Empty;
        }
    }
}
