using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using TFlex.DOCs.Common;
using TFlex.DOCs.Model.Desktop;
using TFlex.DOCs.Model.Macros.ObjectModel;
using TFlex.DOCs.Model.Macros.Processes;
using TFlex.DOCs.Model.Processes.Events.Contexts;
using TFlex.DOCs.Model.References;
using TFlex.DOCs.Model.References.Documents;
using TFlex.DOCs.Model.References.Files;
using TFlex.DOCs.Model.References.Nomenclature;
using TFlex.DOCs.Model.References.Nomenclature.ModificationNotices;
using TFlex.DOCs.Model.Search;
using TFlex.DOCs.Model.Stages;
using TFiles = TFlex.DOCs.Model.References.Files;

namespace PDM_ChangNotification
{
    public static class Guids
    {
        public static readonly Guid СвязьИзменения = new Guid("5e46670a-400c-4e36-bb37-d4d651bdf692");
        public static readonly Guid СвязьЦелевойДокумент = new Guid("a0e64cef-bf5b-47b9-ae5d-12155c0db936");
        public static readonly Guid СвязьИсходныйДокумент = new Guid("48b83092-a645-4dbd-83c0-a3ab0a02ee62");
        public static readonly Guid СвязьДокумент = new Guid("d962545d-bccb-40b5-9986-257b57032f6e");
        public static readonly Guid СвязьФайлыИзменения = new Guid("15f76619-7f52-4a56-8498-587dc381e808");
        public static readonly Guid СвязьРабочиеФайлыИзменения = new Guid("6b65a575-3ca4-4fb0-9bfc-4d1655c2d83e");
        public static readonly Guid СвязьФайлыДокумента = new Guid("9eda1479-c00d-4d28-bd8e-0d592c873303");
        public static readonly Guid СвязьКарточкаУчета = new Guid("d708e1b4-2a1a-499c-aaaf-be5828e6377e");
        public static readonly Guid СвязьЭскиз = new Guid("180e553f-7248-42d8-be38-472386f54ff9");
        public static readonly Guid СвязьТипПокупногоИзделия = new Guid("c215b492-5145-49cf-b849-bbc696edd635");
        public static readonly Guid СвязьОсновнойМатериал = new Guid("2167290d-faa1-4c55-a5cb-32bcd205502a");
        public static readonly Guid СвязьСвязанныеДокументы = new Guid("b840c7cc-bc01-48db-84a0-3706b7aba745");

        public static readonly Guid ПараметрНомерИзменения = new Guid("91486563-d044-4045-814b-3432b67812f1");
        public static readonly Guid ПараметрНазваниеВарианта = new Guid("dab2c5d2-2063-4018-ba88-abe038f52557");
        public static readonly Guid ПараметрОбозначениеИзвещения = new Guid("b03c9129-7ac3-46f5-bf7d-fdd88ef1ff9a");
        public static readonly Guid ПараметрОбозначениеДокумента = new Guid("ae35e329-15b4-4281-ad2b-0e8659ad2bfb");

        public static readonly Guid ГруппаПараметровДанныеДляСпецификации = new Guid("aa5a9c14-85b8-45a5-8fb9-be72286fb4db");

        public static readonly Guid guidESI = new Guid("853d0f07-9632-42dd-bc7a-d91eae4b8e83");
        public static readonly Guid guidProductsClassifier = new Guid("07e0dfaa-305e-468b-acd7-b6dad340da8f");
    }

    [Serializable]
    public class CancelException : ApplicationException
    {
        public CancelException() { }
        public CancelException(string message) : base(message) { }
        public CancelException(string message, Exception ex) : base(message, ex) { }
        protected CancelException(System.Runtime.Serialization.SerializationInfo info,
            System.Runtime.Serialization.StreamingContext contex)
            : base(info, contex) { }
    }

    [Serializable]
    public class UserFriendlyException : ApplicationException
    {
        public UserFriendlyException() { }
        public UserFriendlyException(string message) : base(message) { }
        public UserFriendlyException(string message, Exception ex) : base(message, ex) { }
        protected UserFriendlyException(System.Runtime.Serialization.SerializationInfo info,
            System.Runtime.Serialization.StreamingContext contex)
            : base(info, contex) { }
    }

    public static class ReferenceObjectExtensions
    {
        public static void StartUpdate(this ReferenceObject ro)
        {
            if (ro.IsCheckedOut && !ro.IsCheckedOutByCurrentUser)
                throw new InvalidOperationException(String.Format(TFlex.DOCs.Resources.Strings.Messages.CantCheckOutExceptionMessage, ro));

            if (ro.CanCheckOut)
                Desktop.CheckOut(ro, false);

            if (!ro.Changing)
                ro.BeginChanges();
        }

        public static void CancelUpdate(this ReferenceObject ro)
        {
            if (ro.Changing)
                ro.CancelChanges();

            if (ro.CanUndoCheckOut)
                Desktop.UndoCheckOut(ro);
        }
    }

    public class Macro : ProcessActionMacroProvider
    {
        private Dictionary<string, Stage> _stagesCache;
        private List<ReferenceObject> _referenceObjectsForDelete;
        private ReferenceObject _modificationNotice;

        /// <summary> Поддерживаемые расширения файлов </summary>
        private static readonly string[] SupportedExtensions = { "tif", "tiff", "pdf" };

        public Macro(EventContext context)
            : base(context)
        {
        }

        public override void Run()
        {
        }

        #region  Проверка возможности применения ИИ

        public void ПроверитьВозможностьПримененияИзвещения()
        {
            var modificationNotice = GetModificationNotice();
            if (modificationNotice == null)
            {
                Переменные["Результат проверки возможности применения изменений"] = "Не найдено ИИ";
                return;
            }

            string result = GetReportAboutPossibilityOfApplyingNotice(modificationNotice);
            Переменные["Результат проверки возможности применения изменений"] = string.IsNullOrEmpty(result) ? "Готово к применению" : result;
        }

        private string GetReportAboutPossibilityOfApplyingNotice(ReferenceObject modificationNoticeRO)
        {
            List<ReferenceObject> objectsWithValidabilityChecking = new List<ReferenceObject>();
            StringBuilder stringBuilder = new StringBuilder();
            var modifications = modificationNoticeRO.GetObjects(Guids.СвязьИзменения).Where(mod => HasEstablishedStage(mod, "Утверждение")).ToList();
            if (modifications.IsNullOrEmpty())
            {
                stringBuilder.AppendLine(String.Format("Извещение '{0}' не содержит изменений, или все изменения проведены",
                    modificationNoticeRO));

                return stringBuilder.ToString();
            }

            foreach (var modification in modifications)
            {
                objectsWithValidabilityChecking.Add(modification);

                var targetDocument = modification.GetObject(Guids.СвязьЦелевойДокумент) as NomenclatureObject;
                if (targetDocument == null)
                    stringBuilder.AppendLine(String.Format("У изменения '{0}' отсутствует целевой вариант или к нему необходимо применить изменения.", modification));
                else
                    objectsWithValidabilityChecking.Add(targetDocument);

                EngineeringDocumentObject actualDocument = null;
                ReferenceObject actualNomenclature = modification.GetObject(Guids.СвязьДокумент);
                if (actualNomenclature == null)
                {
                    stringBuilder.AppendLine(String.Format("У изменения '{0}' отсутствует актуальный вариант.",
                        modification));
                }
                else
                {
                    actualDocument = (actualNomenclature as NomenclatureObject).LinkedObject as EngineeringDocumentObject;

                    var actualDocumentFiles = actualDocument.GetFiles();
                    if (actualDocumentFiles.IsNullOrEmpty())
                        stringBuilder.AppendLine(String.Format("У документа '{0}' отсутствуют связанные файлы", GetDocumentFullName(actualDocument)));

                    var archiveFiles = actualDocumentFiles.Where(file => file.Path.GetString().ToLower().Contains(@"архив\")).ToList();
                    if (archiveFiles.IsNullOrEmpty())
                        stringBuilder.AppendLine(String.Format("У документа '{0}' отсутствуют файлы в архиве", GetDocumentFullName(actualDocument)));

                    var modificationFiles = modification.GetObjects(Guids.СвязьРабочиеФайлыИзменения).OfType<FileObject>()
                        .Where(file => HasEstablishedStage(file, "Утверждение")).ToList();
                    if (modificationFiles.IsNullOrEmpty())
                        stringBuilder.AppendLine(String.Format("У изменения '{0}' отсутствуют файлы в стадии 'Утверждение'", modification));

                    var scriptExtensions = new string[] { "tiff", "tif", "pdf" };
                    var scripts = archiveFiles.Where(docFile => docFile.Path.GetString().ToLower().Contains(@"подлинники") && scriptExtensions.Contains(docFile.Class.Extension)).ToList();
                    if (scripts.IsNullOrEmpty())
                        stringBuilder.AppendLine(String.Format("У документа '{0}' отсутствуют подлинники", GetDocumentFullName(actualDocument)));
                    else
                    {
                        foreach (var script in scripts)
                        {
                            var correspondingFileInModification = modificationFiles.FirstOrDefault(mf => mf.Name.GetString() == script.Name.GetString());
                            if (correspondingFileInModification == null)
                                stringBuilder.AppendLine(String.Format("У документа '{0}' для файла-подлинника '{1}' не найден соответствующий файл в изменении '{2}'",
                                    GetDocumentFullName(actualDocument), script, modification));
                        }
                    }

                    var originals = archiveFiles.Where(docFile => docFile.Path.GetString().ToLower().Contains(@"оригиналы") && !scriptExtensions.Contains(docFile.Class.Extension)).ToList();
                    if (originals.IsNullOrEmpty())
                        stringBuilder.AppendLine(String.Format("У документа '{0}' отсутствуют оригиналы", GetDocumentFullName(actualDocument)));
                    else
                    {
                        foreach (var original in originals)
                        {
                            var correspondingFileInModification = modificationFiles.FirstOrDefault(mf => mf.Name.GetString() == original.Name.GetString());
                            if (correspondingFileInModification == null)
                                stringBuilder.AppendLine(String.Format("У документа '{0}' для файла-оригинала '{1}' не найден соответствующий файл в изменении '{2}'",
                                    GetDocumentFullName(actualDocument), original, modification));
                        }
                    }

                    objectsWithValidabilityChecking.Add(actualNomenclature);
                    objectsWithValidabilityChecking.Add(actualDocument);

                    if (!archiveFiles.IsNullOrEmpty())
                        objectsWithValidabilityChecking.AddRange(archiveFiles);

                    if (!modificationFiles.IsNullOrEmpty())
                        objectsWithValidabilityChecking.AddRange(modificationFiles);
                }

                ReferenceObject initialDocument = modification.GetObject(Guids.СвязьИсходныйДокумент);
                if (initialDocument != null)
                {
                    string initialDocumentName = initialDocument[Guids.ПараметрНазваниеВарианта].GetString();
                    if (initialDocumentName.ToLower() == "ии.0")
                    {
                        if (actualDocument != null)
                            stringBuilder.AppendLine(String.Format("Нельзя применить изм.1 для документа '{0} - {1}' т.к. вариант 'ИИ.0' уже существует.", actualDocument.Denotation, actualDocument.Name));
                        else
                            stringBuilder.AppendLine("Нельзя применить изм.1 т.к. вариант 'ИИ.0' уже существует.");
                    }
                    else if (!IsCorrectModificationNumber(initialDocumentName))
                    {
                        stringBuilder.AppendLine(String.Format("У исходного документа '{0}' некорректно заполнен параметр '{1}'.", modification,
                            initialDocument[Guids.ПараметрНазваниеВарианта].ParameterInfo.Name));
                    }
                }
                else
                {
                    if (actualDocument != null)
                    {
                        var documentsReference = actualDocument.Reference;
                        documentsReference.LoadSettings.LoadDeleted = true;
                        var pgDataForSpecification = documentsReference.ParameterGroup.FindRelation(Guids.ГруппаПараметровДанныеДляСпецификации);
                        using (Filter filter = new Filter(documentsReference.ParameterGroup))
                        {
                            filter.Terms.AddTerm(pgDataForSpecification.Parameters.Find(SpecificationFields.Denotation),
                                ComparisonOperator.Equal, actualDocument.Denotation);
                            filter.Terms.AddTerm(documentsReference.ParameterGroup[EngineeringDocumentFields.Name],
                                ComparisonOperator.Equal, actualDocument.Name);
                            filter.Terms.AddTerm(pgDataForSpecification.Parameters.Find(SpecificationFields.VariantName),
                               ComparisonOperator.Equal, "ИИ.0");

                            var initialDocuments = documentsReference.Find(filter);

                            if (!initialDocuments.IsNullOrEmpty())
                            {
                                if (initialDocuments.First().IsDeleted)
                                    stringBuilder.AppendLine(String.Format("Нельзя применить изм.1 для документа '{0} - {1}' т.к. вариант 'ИИ.0' не удален из корзины.", actualDocument.Denotation, actualDocument.Name));
                                else
                                    stringBuilder.AppendLine(String.Format("Нельзя применить изм.1 для документа '{0} - {1}' т.к. вариант 'ИИ.0' уже существует.", actualDocument.Denotation, actualDocument.Name));
                            }
                        }
                    }
                }

                var sketch = modification.GetObject(Guids.СвязьЭскиз) as FileObject;
                if (sketch != null)
                    objectsWithValidabilityChecking.Add(sketch);
            }

            stringBuilder.Append(CheckObjectsForEditing(objectsWithValidabilityChecking));
            return stringBuilder.ToString();
        }

        private StringBuilder CheckObjectsForEditing(IEnumerable<ReferenceObject> objects)
        {
            List<ReferenceObject> processedObjects = objects.ToList();
            StringBuilder stringBuilder = new StringBuilder();

            if (processedObjects.IsNullOrEmpty())
                return stringBuilder;

            var deletedObjects = processedObjects.Where(po => po.IsDeleted).ToList();
            if (!deletedObjects.IsNullOrEmpty())
                stringBuilder.AppendLine(String.Format("Следующие объекты являются удаленными: {0}.", GetObjectsTextRepresentation(deletedObjects)));

            processedObjects = processedObjects.Except(deletedObjects).ToList();
            var blockedObjects = processedObjects.Where(po => po.IsCheckedOut).ToList();
            if (!blockedObjects.IsNullOrEmpty())
                stringBuilder.AppendLine(String.Format("Следующие объекты взяты на редактирование: {0}.", GetObjectsTextRepresentation(blockedObjects)));

            processedObjects = processedObjects.Except(blockedObjects).ToList();
            if (processedObjects.IsNullOrEmpty())
                return stringBuilder;

            return stringBuilder;
        }

        private string GetObjectsTextRepresentation(List<ReferenceObject> objects)
        {
            if (objects.IsNullOrEmpty())
                return String.Empty;

            return String.Join(Environment.NewLine, objects);
        }

        #endregion

        public void ПрименениеИИ()
        {
            var modificationNotice = GetModificationNotice();
            if (modificationNotice == null)
            {
                var error = new UserFriendlyException("Не найдено ИИ");
                RaiseError(GetExceptionText(error, 0));
            }

            ProcessModificationNotice(modificationNotice);
        }

        private ReferenceObject GetModificationNotice()
        {
            var objects = Объекты.Union(ВспомогательныеОбъекты).Select(объект => (ReferenceObject)объект).ToList();
            return objects.FirstOrDefault(obj => obj.Reference.ParameterGroup.Guid == ModificationNoticesReference.ModificationNoticesReferenceGuid);
        }

        private bool HasEstablishedStage(ReferenceObject ro, string stageName)
        {
            if (ro == null)
                throw new ArgumentNullException("referenceObject");

            if (string.IsNullOrEmpty(stageName))
                throw new ArgumentNullException("stageName");

            var schemeStage = ro.SystemFields.Stage;
            if (schemeStage != null)
            {
                var stage = schemeStage.Stage;
                if (stage != null)
                    return stage.Name == stageName;
            }

            return false;
        }

        private bool IsCorrectModificationNumber(string name)
        {
            if (string.IsNullOrEmpty(name))
                return false;

            string[] parts = name.Split('.');
            if (parts.Length != 2)
                return false;

            int number = 0;
            return parts[0].ToLower() == "ии" && Int32.TryParse(parts[1], out number);
        }

        private string GetDocumentFullName(EngineeringDocumentObject document)
        {
            return string.Concat(document.Denotation, "^", document.Name);
        }

        private string GetNomenclatureFullName(NomenclatureObject nomenclature)
        {
            return string.Concat(nomenclature.Denotation, "^", nomenclature.Name);
        }

        private bool ChangeStage(ReferenceObject ro, string stageName)
        {
            if (ro == null)
                return false;

            stageName = stageName.ToLower();
            Stage stage = null;

            if (_stagesCache != null)
                _stagesCache.TryGetValue(stageName, out stage);
            else
                _stagesCache = new Dictionary<string, Stage>();

            if (stage == null)
            {
                stage = Stage.Find(Context.Connection, stageName);
                if (stage == null)
                    return false;

                _stagesCache.Add(stageName, stage);
            }

            return stage.Set(new List<ReferenceObject>() { ro }).FirstOrDefault() != null;
            //return (stage.Change(new List<ReferenceObject>() { ro }).FirstOrDefault() != null); // Если не игнорируем схему переходов 
        }

        private Dictionary<ReferenceObject, List<Exception>> _dictionaryErrors = new Dictionary<ReferenceObject, List<Exception>>();

        private void ProcessModificationNotice(ReferenceObject modificationNotice)
        {
            _modificationNotice = modificationNotice;
            _referenceObjectsForDelete = new List<ReferenceObject>();
            var modifications = modificationNotice.GetObjects(Guids.СвязьИзменения).Where(mod => !HasEstablishedStage(mod, "Хранение")).ToList();

            Stage modificationStage;
            List<Exception> exceptions = new List<Exception>();
            ReferenceObjectSaveSet saveSet;

            ReferenceObject typeStatusDesigning = (ReferenceObject)НайтиОбъект("Типы статусов", "Наименование", "Утверждено");
            var referenceInfoStatuses = Context.Connection.ReferenceCatalog.Find(new Guid("770d906c-5fc9-4bc7-a8aa-6d7f45dd5166"));
            var referenceStatuses = referenceInfoStatuses.CreateReference();
            var statusReferenceObject = referenceStatuses.CreateReferenceObject();
            statusReferenceObject[new Guid("b6c6f4cb-2548-4224-b8ed-f3abf82c3e8d")].Value = "Утверждено";
            statusReferenceObject[new Guid("c7496ed9-56d3-4489-a307-5017bbabccf4")].Value = (DateTime)modificationNotice[new Guid("fab79790-f88c-416f-bc5d-bc520eb89101")].Value;
            statusReferenceObject.SetLinkedObject(new Guid("baca443a-0fee-4dcf-bed4-063f32c88fd7"), typeStatusDesigning);
            statusReferenceObject.EndChanges();

            foreach (var modification in modifications)
            {
                var nomenclature = modification.GetObject(Guids.СвязьДокумент) as NomenclatureObject;
                var document = nomenclature.LinkedObject as EngineeringDocumentObject;
                var targetNomenclature = modification.GetObject(Guids.СвязьЦелевойДокумент) as NomenclatureObject;

                foreach (var parent in nomenclature.Parents)
                {
                    Объект выбранныеОбъект = Объект.CreateInstance(nomenclature, Context);
                    Объект родитель = Объект.CreateInstance(parent, Context);
                    ЗаполнитьСтатус(выбранныеОбъект, родитель);
                }

                modificationStage = modification.SystemFields.Stage == null ? null : modification.SystemFields.Stage.Stage; // TODO: добавить структуру, запоминающую обрабатываемые объекты и первоначальные стадии
                var actualNomenclatureStage = nomenclature.SystemFields.Stage == null ? null : nomenclature.SystemFields.Stage.Stage;
                var actualDocumentStage = nomenclature.LinkedObject.SystemFields.Stage == null ? null : nomenclature.LinkedObject.SystemFields.Stage.Stage;
                var targetNomenclatureStage = targetNomenclature.SystemFields.Stage == null ? null : targetNomenclature.SystemFields.Stage.Stage;


                if (targetNomenclature.LockState == ReferenceObjectLockState.None)
                    targetNomenclature.CheckOut(false);

                targetNomenclature.BeginChanges();
                targetNomenclature.AddLinkedObject(new Guid("13903cba-7b7d-464a-8481-56ef6a3f36dc"), statusReferenceObject);
                targetNomenclature.EndChanges();
                targetNomenclature.CheckIn("");

                ReferenceObject initialDocument;
                saveSet = new ReferenceObjectSaveSet();
                try
                {
                    ProcessModificationDocuments(saveSet, modification, out initialDocument);
                    ProcessModificationFiles(saveSet, modification, document, targetNomenclature);

                    try
                    {
                        ProcessSketchFile(modification, targetNomenclature);
                    }
                    catch (Exception ex)
                    {
                        exceptions.Add(ex);
                    }

                    if (saveSet.Count > 0)
                        saveSet.EndChanges();

                    var objectsForCheckIn = saveSet.Where(ro => ro.CanCheckIn).ToList();
                    Desktop.CheckIn(objectsForCheckIn, "Автоматическое принятие на хранение", false);
                    foreach (var obj in objectsForCheckIn)
                        ChangeStage(obj, "Хранение");

                    if (initialDocument != null)
                    {
                        if (initialDocument.CanCheckIn)
                            Desktop.CheckIn(initialDocument, "Автоматическое принятие на хранение", false);

                        ChangeStage(initialDocument, "Хранение");
                    }
                    ChangeStage(modification, "Хранение");

                    var files = objectsForCheckIn.OfType<FileObject>().ToList();
                    foreach (var file in files)
                    {
                        try
                        {
                            ProcessInventoryCard(file, document);
                        }
                        catch (Exception ex)
                        {
                            exceptions.Add(ex);
                        }
                    }
                }
                catch (Exception ex)
                {
                    exceptions.Add(ex);

                    if (saveSet.Changing)
                        saveSet.CancelChanges();

                    Desktop.UndoCheckOut(saveSet.Where(ro => ro.CanUndoCheckOut));
                    DeleteReferenceObjects(_referenceObjectsForDelete);
                    _referenceObjectsForDelete.Clear();

                    if (modificationStage != null)
                        ChangeStage(modification, modificationStage.Name);

                    if (actualNomenclatureStage != null)
                        ChangeStage(nomenclature, actualNomenclatureStage.Name);

                    if (actualDocumentStage != null)
                        ChangeStage(nomenclature.LinkedObject, actualDocumentStage.Name);

                    if (targetNomenclatureStage != null)
                        ChangeStage(targetNomenclature, targetNomenclatureStage.Name);
                }
            }

            if (exceptions.Count > 0)
                RaiseError(GetExceptionsText(exceptions));
            //СоздатьОшибку(GetAggregateExceptionText(new AggregateException(exceptions)));
            else
            {
                ChangeStage(modificationNotice, "Хранение");

                ClearWorkingFolders(modifications);
            }
            // TODO: Есть ошибки, при которых мы не откатываем изменения, но при этом ИИ считается примененным. Учесть данный момент.
        }

        // ----------------------------------------------------------------------------
        public void ЗаполнитьСтатус(Объект выбранныйОбъект, Объект родительскийОбъект)
        {
        	//var tst = НайтиОбъект("test", "Id", 1);
        	
            string мажорнаяРевизияВыбранногоОбъекта = ПолучитьМажорнуюРевизию(выбранныйОбъект);
            int минорнаяРевизияВыбранногоОбъекта = ПолучитьМинорнуюРевизию(выбранныйОбъект);
            var подключение = выбранныйОбъект.РодительскиеПодключения.FirstOrDefault(t => (int)t.РодительскийОбъект["Id"] == (int)родительскийОбъект["Id"]);
            if (подключение == null)
                return;

            var применяемостьРевизияЭСИ = подключение.СвязанныеОбъекты["Применяемость в изделиях"];
            var родительТекущегоПодключения = подключение.РодительскийОбъект;
            var ревизияСледующаяЭСИ = НайтиОбъект(Guids.guidESI.ToString(), string.Format("[Guid логического объекта] = '{0}' И [Ревизия] = '{1}.{2}'", выбранныйОбъект["Guid логического объекта"], мажорнаяРевизияВыбранногоОбъекта, минорнаяРевизияВыбранногоОбъекта + 1));
            if (ревизияСледующаяЭСИ == null)
                return;

            var подключениеРевизияСледующийЭСИ = ревизияСледующаяЭСИ.РодительскиеПодключения.FirstOrDefault(t => t.РодительскийОбъект["Id"].ToString() == родительТекущегоПодключения["Id"].ToString());
            if (подключениеРевизияСледующийЭСИ == null)
                return;
            
            var применяемостьИзделияРевизияСледующийЭСИ = подключениеРевизияСледующийЭСИ.СвязанныеОбъекты["Применяемость в изделиях"];
            if (применяемостьИзделияРевизияСледующийЭСИ == null)
                return;

            foreach (var применяемостьРевизияСледующийЭСИ in применяемостьИзделияРевизияСледующийЭСИ)
            {
                int начинаяС = применяемостьРевизияСледующийЭСИ["Начиная с изделия"];                                                              // Id
                if (начинаяС == 0)
                    continue;

                var начинаяСКлассификатор = НайтиОбъект(Guids.guidProductsClassifier.ToString(), string.Format("[Id] = '{0}'", начинаяС));
                var начинаяСКлассификаторРодитель = начинаяСКлассификатор.РодительскийОбъект;
                int порядковыйНомерНачинаяСКлассификатор = начинаяСКлассификатор["№"];
                int порядковыйНомерНачинаяСКлассификаторПредыдущий = порядковыйНомерНачинаяСКлассификатор - 1;
                
                /*tst.Изменить();
                tst["txt"] = начинаяСКлассификаторРодитель.ToString();
                tst.Сохранить();*/
                
                var начинаяСКлассификаторПредыдущий = начинаяСКлассификаторРодитель.ДочерниеОбъекты.FirstOrDefault(t => (int)t["№"] == порядковыйНомерНачинаяСКлассификаторПредыдущий);
                if (начинаяСКлассификаторПредыдущий == null)
                    continue;

                foreach (var применяемостьЭСИ in применяемостьРевизияЭСИ)
                {
                    int начинаяСТекущийОбъект = применяемостьЭСИ["Начиная с изделия"];                                                              // Id
                    if (начинаяСТекущийОбъект == 0)
                        continue;

                    var начинаяСКлассификаторТекущийОбъект = НайтиОбъект(Guids.guidProductsClassifier.ToString(), string.Format("[Id] = '{0}'", начинаяСТекущийОбъект));
                    var начинаяСКлассификаторТекущийОбъектРодитель = начинаяСКлассификаторТекущийОбъект.РодительскийОбъект;
                    if (начинаяСКлассификаторТекущийОбъектРодитель["Id"].ToString() != начинаяСКлассификаторРодитель["Id"].ToString())
                        continue;

                    применяемостьЭСИ.Изменить();
                    применяемостьЭСИ["Заканчивая изделием"] = (int)начинаяСКлассификаторПредыдущий["Id"];
                    применяемостьЭСИ.Сохранить();
                }
            }
        }

        private int ПолучитьМинорнуюРевизию(Объект объект)
        {
            string ревизияОбъекта = объект["Ревизия"].ToString();
            string минорнаяРевизияОбъекта = ревизияОбъекта.Substring(ревизияОбъекта.IndexOf(".") + 1);

            return Convert.ToInt32(минорнаяРевизияОбъекта);
        }

        private string ПолучитьМажорнуюРевизию(Объект объект)
        {
            string ревизияОбъекта = объект["Ревизия"].ToString();
            string мажорнаяРевизияОбъекта = ревизияОбъекта.Substring(0, ревизияОбъекта.IndexOf("."));

            return мажорнаяРевизияОбъекта;
        }

        // ----------------------------------------------------------------------------

        private void ClearWorkingFolders(List<ReferenceObject> modifications)
        {
            List<ReferenceObject> allWorkingFiles = new List<ReferenceObject>();
            List<int> folderIDs = new List<int>();

            foreach (var modification in modifications)
            {
                var modificationWorkingFiles = modification
                     .GetObjects(Guids.СвязьРабочиеФайлыИзменения);

                if (modificationWorkingFiles.IsNullOrEmpty())
                    continue;

                allWorkingFiles.AddRange(modificationWorkingFiles);
                folderIDs.AddRange(
                    modificationWorkingFiles
                    .OfType<FileReferenceObject>()
                    .Select(referenceObject => referenceObject.Parent.SystemFields.Id));
            }

            bool deleteFromRecycleBin = true;
            DeleteReferenceObjects(allWorkingFiles, deleteFromRecycleBin);

            var folders = FindObjects("Файлы",
                String.Format("[ID] Входит в список '{0}'", String.Join(", ", folderIDs.Distinct()))).ToList();

            DeleteReferenceObjects(folders.Select(refObj => (ReferenceObject)refObj).ToList(), deleteFromRecycleBin);
        }

        private void DeleteReferenceObjects(List<ReferenceObject> referenceObjects, bool deleteFromRecycleBin = true)
        {
            if (referenceObjects.IsNullOrEmpty())
                return;

            var referenceObjectsWithSupportDesktop = referenceObjects.Where(ro => ro.Reference.ParameterGroup.SupportsDesktop).ToList();

            var referenceObjectsWithoutSupportDesktop = referenceObjects.Except(referenceObjectsWithSupportDesktop).Where(ro => ro.CanDelete).ToArray();
            Array.ForEach(referenceObjectsWithoutSupportDesktop, ro => ro.Delete());

            var addedObjectsWithSupportDesktop = referenceObjectsWithSupportDesktop.Where(ro => ro.LockState == ReferenceObjectLockState.CheckedOutForAdd).ToArray();
            Desktop.UndoCheckOut(addedObjectsWithSupportDesktop);

            var savedObjectsWithSupportDesktop = referenceObjectsWithSupportDesktop.Except(addedObjectsWithSupportDesktop).ToList();
            var checkedOutObjectsWithSupportDesktop = savedObjectsWithSupportDesktop.Where(ro => ro.IsCheckedOutByCurrentUser).ToList();
            Desktop.UndoCheckOut(checkedOutObjectsWithSupportDesktop);

            var canCheckOutObjects = savedObjectsWithSupportDesktop.Where(ro => ro.CanCheckOut).ToList();
            Desktop.CheckOut(canCheckOutObjects, true);
            var deletedObjects = Desktop.CheckIn(canCheckOutObjects, "Удаление файлов", false);

            if (deleteFromRecycleBin)
            {
                var objectsInRecycleBin = referenceObjects.Union(deletedObjects).Where(desktopObject => ((ReferenceObject)desktopObject).IsInRecycleBin);
                Desktop.ClearRecycleBin(objectsInRecycleBin);
            }
        }

        private void ProcessModificationDocuments(ReferenceObjectSaveSet saveSet, ReferenceObject modification, out ReferenceObject initialDoc)
        {
        	
            NomenclatureObject initialDocument = modification.GetObject(Guids.СвязьИсходныйДокумент) as NomenclatureObject;
            var actualNomenclature = modification.GetObject(Guids.СвязьДокумент) as NomenclatureObject;
            var targetDocument = modification.GetObject(Guids.СвязьЦелевойДокумент) as NomenclatureObject;

            if (initialDocument == null)
            {
                initialDocument = actualNomenclature.CreateVariant("ИИ.0", true, false);
                Desktop.CheckIn(initialDocument, "Автоматическое принятие на хранение", false);
                _referenceObjectsForDelete.Add(initialDocument);
            }

            initialDoc = initialDocument;
            var actualNomenclatureStage = actualNomenclature.SystemFields.Stage == null ? null : actualNomenclature.SystemFields.Stage.Stage;
            var targetDocumentStage = targetDocument.SystemFields.Stage == null ? null : targetDocument.SystemFields.Stage.Stage;
            var actualDocumentStage = actualNomenclature.LinkedObject.SystemFields.Stage == null ? null : targetDocument.SystemFields.Stage.Stage;

            
            return;
            try
            {
                modification.StartUpdate();
                modification.SetLinkedObject(Guids.СвязьИсходныйДокумент, initialDocument);

                ChangeActualDocument(actualNomenclature, targetDocument, saveSet);

                targetDocument.StartUpdate();
                string variantName = initialDocument.VariantName.ToString();
                var number = System.Convert.ToInt32(variantName.Split('.').Last()) + 1;
                targetDocument.VariantName.Value = String.Format("ИИ.{0}", number);
                modification[Guids.ПараметрНомерИзменения].Value = String.Format("ИИ.{0}", number);

                saveSet.Add(targetDocument);
                saveSet.Add(actualNomenclature);
                saveSet.Add(modification);
            }
            catch
            {
                if (actualNomenclatureStage != null)
                    ChangeStage(actualNomenclature, actualNomenclatureStage.Name);

                if (targetDocumentStage != null)
                    ChangeStage(targetDocument, targetDocumentStage.Name);

                if (actualDocumentStage != null)
                    ChangeStage(actualNomenclature.LinkedObject, actualDocumentStage.Name);

                throw;
            }
        }

        private FileObject GetCorrespondingFileInDocuments(List<FileObject> documentFiles, FileObject modificationFile)
        {
            List<FileObject> neededFiles = documentFiles.Where(documentFile => documentFile.Class.Extension == modificationFile.Class.Extension).ToList();
            switch (modificationFile.Class.Extension)
            {
                case ("tiff"):
                case ("tif"):
                case ("pdf"):
                    neededFiles = documentFiles.Where(documentFile => documentFile.Path.GetString().ToLower().Contains(@"подлинники")).ToList();
                    break;
                default:
                    neededFiles = documentFiles.Where(documentFile => documentFile.Path.GetString().ToLower().Contains(@"оригиналы")).ToList();
                    break;
            }

            return neededFiles.FirstOrDefault(documentFile => documentFile.Name.GetString() == modificationFile.Name.GetString());
        }

        private void ProcessModificationFiles(ReferenceObjectSaveSet saveSet, ReferenceObject modification, EngineeringDocumentObject document, NomenclatureObject targetNomenclature)
        {
            var modificationFiles = modification.GetObjects(Guids.СвязьРабочиеФайлыИзменения).OfType<FileObject>().Where(f => HasEstablishedStage(f, "Утверждено")).ToList();
            var documentFiles = document.GetObjects(Guids.СвязьФайлыДокумента).OfType<FileObject>().Where(f => f.Path.ToString().ToLower().Contains(@"архив\")).ToList();

            if (modificationFiles.IsNullOrEmpty() || documentFiles.IsNullOrEmpty())
                return;

            List<FileObject> attachFiles = new List<FileObject>();
            foreach (var modificationFile in modificationFiles)
            {
                FileObject correspondingFileInDocuments = GetCorrespondingFileInDocuments(documentFiles, modificationFile);
                if (correspondingFileInDocuments == null) // могут быть файлы сборок
                    continue;

                //    throw new FileNotFoundException(String.Format("Для файла изменения '{0}' не найден файл-подлинник у документа '{1}'", modificationFile, GetDocumentFullName(document)));
                // TODO: а остальные файлы не обрабатываем??

                var originalModificationFile = (FileObject)НайтиОбъект("Файлы", String.Format("[ID] = '{0}'", modificationFile.SystemFields.Id));
                originalModificationFile.GetHeadRevision();
                if (!File.Exists(originalModificationFile.LocalPath))
                    continue;

                var stage = correspondingFileInDocuments.SystemFields.Stage == null ? null : correspondingFileInDocuments.SystemFields.Stage.Stage;
                try
                {
                    ChangeStage(correspondingFileInDocuments, "Исправление");
                    correspondingFileInDocuments.StartUpdate();

                    File.Copy(originalModificationFile.LocalPath, correspondingFileInDocuments.LocalPath, true);
                    new FileInfo(correspondingFileInDocuments.LocalPath).IsReadOnly = false;
                    correspondingFileInDocuments.AddLinkedObject(Guids.СвязьФайлыДокумента, targetNomenclature.LinkedObject);
                    attachFiles.Add(correspondingFileInDocuments);
                    saveSet.Add(correspondingFileInDocuments);
                }
                catch
                {
                    correspondingFileInDocuments.CancelUpdate();

                    if (stage != null)
                        ChangeStage(correspondingFileInDocuments, stage.Name);

                    throw;
                }
            }

            if (!attachFiles.IsNullOrEmpty())
            {
                modification.StartUpdate();

                foreach (var file in attachFiles)
                    modification.AddLinkedObject(Guids.СвязьФайлыИзменения, file);

                if (!saveSet.Contains(modification))
                    saveSet.Add(modification);
            }
        }

        private void ChangeActualDocument(NomenclatureObject actualNomenclature, ReferenceObject targetNomenclature, ReferenceObjectSaveSet saveSet)
        {
            ChangeStage(actualNomenclature, "Исправление");
            actualNomenclature.StartUpdate();
            foreach (var param in targetNomenclature.ParameterValues)
            {
                if (param.ParameterInfo.IsSystem || param.IsEmpty)
                    continue;

                actualNomenclature.ParameterValues[param.ParameterInfo].Value = param.Value;
            }

            CopyHierarchyLinkParameters(targetNomenclature as NomenclatureObject, actualNomenclature);

            var typeOfPurchase = targetNomenclature.GetObject(Guids.СвязьТипПокупногоИзделия);
            actualNomenclature.SetLinkedObject(Guids.СвязьТипПокупногоИзделия, typeOfPurchase);

            var actualDocument = actualNomenclature.LinkedObject;
            var targetDocument = (targetNomenclature as NomenclatureObject).LinkedObject;

            actualDocument.StartUpdate();

            if (targetDocument.Links.ToOne.LinkGroups.Find(Guids.СвязьОсновнойМатериал) != null)
            {
                var mainMaterial = targetDocument.GetObject(Guids.СвязьОсновнойМатериал);
                actualDocument.SetLinkedObject(Guids.СвязьОсновнойМатериал, mainMaterial);
            }

            var linkedDocuments = targetDocument.GetObjects(Guids.СвязьСвязанныеДокументы);
            if (linkedDocuments.Count > 0)
            {
                foreach (var ld in linkedDocuments)
                    actualDocument.AddLinkedObject(Guids.СвязьСвязанныеДокументы, ld);
            }
            else
            {
                var linkedDocumentsOfActualDocuments = actualDocument.GetObjects(Guids.СвязьСвязанныеДокументы);
                foreach (var ld in linkedDocumentsOfActualDocuments)
                    actualDocument.RemoveLinkedObject(Guids.СвязьСвязанныеДокументы, ld);
            }

            // Удалить связанные документы, если их нет у целевого.
            actualNomenclature.VariantName.Value = String.Empty;
            saveSet.Add(actualDocument);
        }

        private void CopyHierarchyLinkParameters(NomenclatureObject targetNomenclature, NomenclatureObject actualNomenclature)
        {
            if (targetNomenclature.Children.State != LoadState.Loaded)
                targetNomenclature.Children.Load();

            var childrenOfTargetNomenclature = targetNomenclature.Children.OfType<NomenclatureObject>().ToList();
            if (childrenOfTargetNomenclature.IsNullOrEmpty())
                return;

            if (actualNomenclature.Children.State != LoadState.Loaded)
                actualNomenclature.Children.Load();
            var childrenOfActualNomenclature = actualNomenclature.Children.OfType<NomenclatureObject>().ToList();

            var detachableChildren = childrenOfActualNomenclature.Where(an => !childrenOfTargetNomenclature.Any(tn => tn.SystemFields.Guid == an.SystemFields.Guid)).ToList();
            foreach (var dc in detachableChildren)
            {
                var link = actualNomenclature.GetChildLink(dc);
                actualNomenclature.DeleteLink(link);
            }

            var attachedChildren = childrenOfTargetNomenclature.Where(tn => !childrenOfActualNomenclature.Any(an => an.SystemFields.Guid == tn.SystemFields.Guid)).ToList();
            foreach (var ac in attachedChildren)
                actualNomenclature.CreateChildLink(ac).EndChanges();

            foreach (var tnChild in childrenOfTargetNomenclature)
            {
                var correspondingANChild = childrenOfActualNomenclature.FirstOrDefault(ch => ch.Name.GetString() == tnChild.Name.GetString() && ch.Denotation.GetString() == tnChild.Denotation.GetString());
                if (correspondingANChild == null)
                    continue;

                var tnComplexHierarchyLink = targetNomenclature.GetChildLink(tnChild);
                var anComplexHierarchyLink = actualNomenclature.GetChildLink(correspondingANChild);

                anComplexHierarchyLink.BeginChanges();

                foreach (var param in tnComplexHierarchyLink.Where(par => par.ParameterInfo.CanEdit))
                    anComplexHierarchyLink[param.ParameterInfo.Guid].Value = param.Value;

                anComplexHierarchyLink.EndChanges(); // Отмену изменений
            }
        }

        private void ProcessInventoryCard(FileObject fileInDocument, ReferenceObject document)
        {
            if (fileInDocument == null || document == null)
                return;

            var inventoryCard = document.GetObject(Guids.СвязьКарточкаУчета);
            if (inventoryCard == null)
                return;

            string fileExtension = fileInDocument.Class.Extension.ToLower();
            if (!SupportedExtensions.Contains(fileExtension))
                return;

            string tempFolder = Path.Combine(Path.GetTempPath(), "Temp DOCs");
            if (!Directory.Exists(tempFolder))
                Directory.CreateDirectory(tempFolder);

            string pathToFile = String.Format(@"{0}\{1}.{2}", tempFolder, Guid.NewGuid(), fileExtension);

            fileInDocument.GetHeadRevision(pathToFile);
            RunMacro("c59fe3ae-7f60-454b-be3e-68db2feb6c6b", "РазборПоФорматам", pathToFile, Объект.CreateInstance(inventoryCard, Context));
            DeleteTempFile(pathToFile);
        }

        private void DeleteTempFile(string path)
        {
            try
            {
                FileInfo fileInfo = new FileInfo(path);
                if (fileInfo.Exists && !fileInfo.IsReadOnly)
                    fileInfo.Delete();
            }
            catch
            {
            }
        }

        #region Файлы по связи "Эскиз"
        private void ProcessSketchFile(ReferenceObject modification, ReferenceObject actualDocument)
        {
            if (modification == null)
                return;

            var sketch = modification.GetObject(Guids.СвязьЭскиз) as FileObject;
            if (sketch == null)
                return;

            var folder = sketch.Reference.FindByRelativePath(@"Архив\Файлы изменений") as TFiles.FolderObject;
            if (folder == null)
                return;

            string modificationNoticeNumber = _modificationNotice[Guids.ПараметрОбозначениеИзвещения].ToString();
            string newName = String.Format("{0}_{1}_{2}_Э{3}", modificationNoticeNumber, actualDocument[Guids.ПараметрОбозначениеДокумента], modification[Guids.ПараметрНомерИзменения], Path.GetExtension(sketch.Name));
            ReferenceObjectCopySet copy = null;
            try
            {
                copy = sketch.Reference.CopyReferenceObject(sketch, folder);
                var copyFile = copy.GetNewObject(sketch);

                if (!String.IsNullOrWhiteSpace(newName))
                    copyFile[FileObject.FieldKeys.Name].Value = newName;

                modification.SetLinkedObject(Guids.СвязьЭскиз, copyFile);

                copy.EndChanges();

                Desktop.CheckIn(copyFile, "Применение изменений", false);
                ChangeStage(copyFile, "Хранение");
            }
            catch
            {
                if (copy != null && copy.Changing)
                    copy.CancelChanges();

                return;
            }

            DeleteReferenceObjects(new List<ReferenceObject>() { sketch });
        }
        #endregion

        private string GetExceptionsText(List<Exception> exceptions)
        {
            if (exceptions.IsNullOrEmpty())
                return String.Empty;

            if (exceptions.Count == 1)
                return GetExceptionText(exceptions.First(), 0);

            var builder = new StringBuilder(512);
            builder.AppendLine("При применении ИИ возникли следующие ошибки:");

            var userExceptions = exceptions.Where(exc => exc is UserFriendlyException).ToList();
            var othersExceptions = exceptions.Except(userExceptions).ToList();
            int i = 1;

            foreach (var exception in userExceptions)
            {
                builder.AppendLine(GetExceptionText(exception, i));
                i++;
            }

            foreach (var exception in othersExceptions)
            {
                builder.AppendLine(GetExceptionText(exception, i));
                i++;
            }

            return builder.ToString();
        }

        private string GetExceptionText(Exception exception, int errorNumber)
        {
            if (exception is UserFriendlyException)
                return String.Format("{0}. {1}", errorNumber, exception.Message) + Environment.NewLine;

            var builder = new StringBuilder(512);
            bool inner = false;
            int i = 0;
            int indentationMultiplier = 10;
            while (exception != null)
            {
                if (inner)
                    builder.AppendLine().Append(' ', i * indentationMultiplier).Append("Внутреннее исключение:").AppendLine();

                if (!inner && errorNumber > 0)
                    builder.Append(String.Format("{0}. ", errorNumber));

                builder.Append(' ', i * indentationMultiplier).Append("Сообщение: : ").Append(exception.Message);
                builder.AppendLine().Append(' ', i * indentationMultiplier).Append("Тип исключения: ").Append(exception.GetType().FullName);
                builder.AppendLine().Append(' ', i * indentationMultiplier).Append("Имя приложения или объекта, вызвавшего исключение: ").AppendLine(exception.Source);
                builder.AppendLine().Append(' ', i * indentationMultiplier).Append("Стек вызовов методов: ").Append(exception.StackTrace).AppendLine();

                exception = exception.InnerException;
                inner = true;
                i++;
            }

            builder.AppendLine();
            return builder.ToString();
        }
    }
}
