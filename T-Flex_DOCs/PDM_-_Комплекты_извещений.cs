using System;
using System.Collections.Generic;
using System.Linq;
using TFlex.DOCs.Client;
using TFlex.DOCs.Client.ViewModels.Base;
using TFlex.DOCs.Client.ViewModels.References.SelectionDialogs;
using TFlex.DOCs.Common;
using TFlex.DOCs.Model.Desktop;
using TFlex.DOCs.Model.Macros;
using TFlex.DOCs.Model.References;
using TFlex.DOCs.Model.References.Macros;
using TFlex.DOCs.Model.References.Nomenclature.ModificationNotices;
using TFlex.DOCs.Model.Search;

namespace PDM.Macros
{
    public class PDM_Комплекты_извещений : MacroProvider
    {
        private static class Guids
        {
            public static readonly Guid ModificationsType = new Guid("52ccb35c-67c5-4b82-af4f-e8ceac4e8d02");
            public static readonly Guid ModificationsComplexType = new Guid("322d371b-c672-40ea-9166-316140fc28cb");
            public static readonly Guid ModificationNoticesLink = new Guid("740644db-40a2-44d8-8fe9-841391985d46");

            public static readonly Guid Denotation = new Guid("b03c9129-7ac3-46f5-bf7d-fdd88ef1ff9a");
            public static readonly Guid SourceDenotation = new Guid("ae3186e1-b166-40bb-91ef-504e3eba03b3");
            public static readonly Guid ApplyDate = new Guid("94e9f388-479a-4380-a756-f5db9d9a044d");
            public static readonly Guid Number = new Guid("05a851e3-8633-4e7f-8ac6-cfffefaf1163");

            /// <summary> Стадия "Хранение" </summary>
            public static readonly Guid StorageStageGuid = new Guid("9826e84a-a0a7-404b-bf0e-d61b902e346a");

            /// <summary> Стадия "Разработка" </summary>
            public static readonly Guid DevelopmentStage = new Guid("527f5234-4c94-43d1-a38d-d3d7fd5d15af");
        }

        public PDM_Комплекты_извещений(MacroContext context) : base(context)
        {
        }

        public ButtonValidator ValidateAddModificationNoticeButton()
        {
            var validator = new ButtonValidator();
            if (!ValidateExecutionPlace())
            {
                validator.Enable = false;
                validator.Visible = false;
            }
            else
            {
                if (!ValidateCanEditModificationNoticeComplect())
                {
                    validator.Enable = false;
                    validator.Visible = false;
                }
            }

            return validator;
        }

        public ButtonValidator ValidateRemoveModificationNoticeButton()
        {
            var validator = new ButtonValidator();
            if (!ValidateExecutionPlace())
            {
                validator.Enable = false;
                validator.Visible = false;
            }
            else
            {
                if (ValidateCanEditModificationNoticeComplect())
                {
                    if (Context.Reference.Objects.IsNullOrEmpty())
                    {
                        validator.Enable = false;
                        validator.Visible = false;
                    }
                }
                else
                {
                    validator.Enable = false;
                    validator.Visible = false;
                }
            }
            return validator;
        }

        private bool ValidateExecutionPlace() => DynamicMacro.ExecutionPlace == ExecutionPlace.WpfClient;

        private bool ValidateCanEditModificationNoticeComplect()
        {
            if (!Context.Reference.IsSlave)
                return false;

            var modificationNoticeCompect = Context.Reference.LinkInfo.MasterObject;
            if (modificationNoticeCompect.SystemFields.Stage?.Stage?.Guid == Guids.StorageStageGuid
                || !modificationNoticeCompect.IsCheckedOutByCurrentUser)
            {
                return false;
            }

            return true;
        }

        public void ДобавитьИзвещение() => AddModificationNotice();

        private void AddModificationNotice()
        {
            var macro = Context[nameof(MacroKeys.MacroSource)] as CodeMacro;
            if (macro.DebugMode.Value == true)
            {
                System.Diagnostics.Debugger.Launch();
                System.Diagnostics.Debugger.Break();
            }

            if (!Context.Reference.IsSlave)
                return;

            var complect = Context.Reference.LinkInfo.MasterObject;
            if (!complect.Class.IsInherit(Guids.ModificationsComplexType))
                return;

            List<ReferenceObject> selectedObjects;

            var linkedObjects = Context.Reference.Objects.ToList();

            var reference = Context.Connection.ReferenceCatalog.Find(ModificationNoticesReference.ModificationNoticesReferenceGuid);

            using (var selectViewModel = new SelectReferenceObjectViewModel(reference, true))
            {
                var filter = new Filter(reference);
                filter.Terms.AddTerm(reference.Description.ClassParameterInfo, ComparisonOperator.IsInheritFrom, Guids.ModificationsType);
                selectViewModel.ContextFilter = filter;

                if (!ApplicationManager.OpenDialog(selectViewModel, default(IDialogOwnerWindowSource)))
                    return;

                selectedObjects = selectViewModel.GetSelectedReferenceObjects().ToList();
            }

            if (selectedObjects.IsNullOrEmpty())
                return;

            var checkOutSelectedObjects = new List<ReferenceObject>();

            foreach (var selectedObject in selectedObjects)
            {
                if (selectedObject.IsCheckedOut)
                {
                    if (selectedObject.IsCheckedOutByCurrentUser)
                    {
                        checkOutSelectedObjects.Add(selectedObject);
                    }
                    else
                    {
                        Error("Невозможно добавить, так как ИИ '{0}' заблокирован пользователем '{1}'. Для выполнения операции разблокируйте ИИ '{0}'.",
                            selectedObject, selectedObject.SystemFields.ClientView.UserName);
                    }
                }
            }

            var checkedInObjects = new List<ReferenceObject>();

            if (checkOutSelectedObjects.Any())
            {
                if (Question($"Применить изменения к ИИ {String.Join(", ", checkOutSelectedObjects.Select(obj => $"'{obj}'"))}?"))
                {
                    checkedInObjects = Desktop.CheckIn(checkOutSelectedObjects, String.Empty, true).OfType<ReferenceObject>().ToList();
                    if (checkedInObjects.IsNullOrEmpty())
                        return;
                }
                else
                    return;
            }

            checkedInObjects.AddRange(selectedObjects.Except(checkedInObjects));

            var addingObjects = Desktop.CheckOut(checkedInObjects, false).OfType<ReferenceObject>().ToList();
            if (addingObjects.IsNullOrEmpty())
                return;

            string complectDenotation = String.Empty;

            int addingObjectsCount = addingObjects.Count;

            if (linkedObjects.IsNullOrEmpty())
            {
                var modificationNotice = addingObjects.First();
                complectDenotation = (string)modificationNotice[Guids.Denotation].Value;
                var complectApplyDate = modificationNotice[Guids.ApplyDate].Value;

                int number = 1;
                foreach (var addingObject in addingObjects)
                {
                    addingObject.BeginChanges();

                    string numberString = $"{number++}/{addingObjectsCount}";
                    addingObject[Guids.Number].Value = numberString;
                    addingObject[Guids.SourceDenotation].Value = addingObject[Guids.Denotation];
                    addingObject[Guids.Denotation].Value = $"{complectDenotation} {numberString}";
                    addingObject[Guids.ApplyDate].Value = complectApplyDate;

                    addingObject.EndChanges();
                }

                var addedObjects = Desktop.CheckIn(addingObjects, $"ИИ внесено в комплект '{complectDenotation}'", false)
                    .OfType<ReferenceObject>().ToList();

                bool needEndChanges = false;
                if (!complect.Changing)
                {
                    complect.BeginChanges();
                    needEndChanges = true;
                }

                complect[Guids.Denotation].Value = complectDenotation;
                complect[Guids.ApplyDate].Value = complectApplyDate;

                foreach (var addedObject in addedObjects)
                {
                    complect.AddLinkedObject(Guids.ModificationNoticesLink, addedObject);
                }

                if (needEndChanges)
                    complect.EndChanges();

                RefreshControls("DenotationControl", "DateControl");
            }
            else
            {
                var checkOutObjects = new List<ReferenceObject>();

                foreach (var linkedObject in linkedObjects)
                {
                    if (linkedObject.IsCheckedOut)
                    {
                        if (linkedObject.IsCheckedOutByCurrentUser)
                        {
                            checkOutObjects.Add(linkedObject);
                        }
                        else
                        {
                            Error("Невозможно добавить, так как ИИ '{0}' заблокирован пользователем '{1}'. Для выполнения операции разблокируйте ИИ '{0}'.",
                                linkedObject, linkedObject.SystemFields.ClientView.UserName);
                        }
                    }
                }

                if (checkOutObjects.Any())
                {
                    if (Question($"Применить изменения к ИИ {String.Join(", ", checkOutObjects.Select(obj => $"'{obj}'"))}?"))
                    {
                        var checkInObjects = Desktop.CheckIn(checkOutObjects, String.Empty, true);
                        if (checkInObjects.IsNullOrEmpty())
                            return;
                    }
                    else
                        return;
                }

                int linkedObjectsCount = linkedObjects.Count;
                int allObjectsCount = linkedObjectsCount + addingObjectsCount;
                complectDenotation = (string)complect[Guids.Denotation].Value;

                foreach (var addingObject in addingObjects)
                {
                    addingObject.BeginChanges();

                    string numberString = $"{++linkedObjectsCount}/{allObjectsCount}";

                    addingObject[Guids.Number].Value = numberString;
                    addingObject[Guids.SourceDenotation].Value = addingObject[Guids.Denotation];
                    addingObject[Guids.Denotation].Value = $"{complectDenotation} {numberString}";
                    addingObject[Guids.ApplyDate].Value = complect[Guids.ApplyDate];

                    addingObject.EndChanges();
                }

                var addedObjects = Desktop.CheckIn(addingObjects, $"ИИ внесено в комплект '{complectDenotation}'", false)
                    .OfType<ReferenceObject>().ToList();

                bool needEndChanges = false;
                if (!complect.Changing)
                {
                    complect.BeginChanges();
                    needEndChanges = true;
                }

                foreach (var addedObject in addedObjects)
                {
                    complect.AddLinkedObject(Guids.ModificationNoticesLink, addedObject);
                }

                if (needEndChanges)
                    complect.EndChanges();

                var editingObjects = Desktop.CheckOut(linkedObjects, false).OfType<ReferenceObject>().ToList();
                editingObjects.AddRange(linkedObjects.Except(editingObjects));
                if (editingObjects.Any())
                {
                    foreach (var editingObject in editingObjects)
                    {
                        string number = (string)editingObject[Guids.Number].Value;
                        string[] parts = number.Split('/');

                        if (parts.Length == 2)
                        {
                            editingObject.BeginChanges();

                            string numberString = $"{parts[0]}/{allObjectsCount}";
                            editingObject[Guids.Number].Value = numberString;
                            editingObject[Guids.Denotation].Value = $"{complectDenotation} {numberString}";

                            editingObject.EndChanges();
                        }
                    }
                    Desktop.CheckIn(editingObjects, "Изменён номер и обозначение", false);
                }
            }

            RefreshReferenceWindow();
        }

        public void ОтключитьИзвещение() => RemoveModificationNotice();

        private void RemoveModificationNotice()
        {
            var macro = Context[nameof(MacroKeys.MacroSource)] as CodeMacro;
            if (macro.DebugMode.Value == true)
            {
                System.Diagnostics.Debugger.Launch();
                System.Diagnostics.Debugger.Break();
            }

            if (!Context.Reference.IsSlave)
                return;

            var complect = Context.Reference.LinkInfo.MasterObject;
            if (!complect.Class.IsInherit(Guids.ModificationsComplexType))
                return;

            var selectedObjects = Context.GetSelectedObjects();
            if (!selectedObjects.Any())
                return;

            var currentLinkedObjects = Context.Reference.Objects.ToList();
            var checkOutObjects = new List<ReferenceObject>();

            foreach (var linkedObject in currentLinkedObjects)
            {
                if (linkedObject.IsCheckedOut)
                {
                    if (linkedObject.IsCheckedOutByCurrentUser)
                    {
                        checkOutObjects.Add(linkedObject);
                    }
                    else
                    {
                        Error("Невозможно удалить, так как ИИ '{0}' заблокирован пользователем '{1}'. Для выполнения операции разблокируйте ИИ '{0}'.",
                            linkedObject, linkedObject.SystemFields.ClientView.UserName);
                    }
                }
            }

            if (checkOutObjects.Any())
            {
                if (Question($"Применить изменения к ИИ {String.Join(", ", checkOutObjects.Select(obj => $"'{obj}'"))}?"))
                {
                    var checkInObjects = Desktop.CheckIn(checkOutObjects, String.Empty, true);
                    if (checkInObjects.IsNullOrEmpty())
                        return;
                }
                else
                    return;
            }

            string currentComplectDenotation = (string)complect[Guids.Denotation].Value;

            var firstObject = selectedObjects.FirstOrDefault(obj => (string)obj[Guids.SourceDenotation].Value == currentComplectDenotation);
            if (firstObject is null) // Отключаемое извещение не является "первым", по которому назначено обозначение комплекта
            {
                var removingObjects = Desktop.CheckOut(selectedObjects, false).OfType<ReferenceObject>().ToList();
                removingObjects.AddRange(selectedObjects.Except(removingObjects));
                if (removingObjects.IsNullOrEmpty())
                    return;

                bool needEndChanges = false;

                if (!complect.Changing)
                {
                    complect.BeginChanges();
                    needEndChanges = true;
                }

                foreach (var removingObject in removingObjects)
                {
                    removingObject.BeginChanges();

                    removingObject[Guids.Number].Value = String.Empty;
                    removingObject[Guids.Denotation].Value = removingObject[Guids.SourceDenotation].Value;
                    removingObject[Guids.SourceDenotation].Value = String.Empty;

                    complect.RemoveLinkedObject(Guids.ModificationNoticesLink, removingObject);

                    removingObject.EndChanges();
                }
                if (needEndChanges)
                    complect.EndChanges();

                Desktop.CheckIn(removingObjects, $"ИИ удалено из комплекта '{currentComplectDenotation}'", false);


                currentLinkedObjects = currentLinkedObjects.Except(removingObjects).ToList();
                var editingObjects = Desktop.CheckOut(currentLinkedObjects, false).OfType<ReferenceObject>().ToList();
                editingObjects.AddRange(currentLinkedObjects.Except(editingObjects));
                if (editingObjects.Any())
                {
                    int editingObjectsCount = editingObjects.Count;

                    // Обработать первое извещение, потом остальные
                    var editingFirstObject = editingObjects.FirstOrDefault(obj => (string)obj[Guids.SourceDenotation].Value == currentComplectDenotation);
                    if (editingFirstObject != null) // %%TODO
                    {
                        editingFirstObject.BeginChanges();

                        string numberString = $"{1}/{editingObjectsCount}";
                        editingFirstObject[Guids.Number].Value = numberString;
                        editingFirstObject[Guids.Denotation].Value = $"{currentComplectDenotation} {numberString}";

                        editingFirstObject.EndChanges();

                        Desktop.CheckIn(editingFirstObject, $"Изменён номер и обозначение", false);

                        editingObjects.Remove(editingFirstObject);
                    }

                    int number = 2;
                    foreach (var editingObject in editingObjects)
                    {
                        editingObject.BeginChanges();

                        string numberString = $"{number++}/{editingObjectsCount}";
                        editingObject[Guids.Number].Value = numberString;
                        editingObject[Guids.Denotation].Value = $"{currentComplectDenotation} {numberString}";

                        editingObject.EndChanges();
                    }

                    Desktop.CheckIn(editingObjects, "Изменён номер и обозначение", false);
                }
            }
            else
            {
                ReferenceObject newFirstModificationNotice = null;
                bool needAddingFirstObject = false;

                if (Question($"Внимание! " +
                    $"Из комплекта отключается ИИ, по которому назначено обозначение комплекта. " +
                    $"Хотите продолжить операцию и выбрать другое ИИ для обозначения комплекта?"))
                {
                    var availableNoticesForSelection = currentLinkedObjects.Except(selectedObjects).ToList();

                    if (availableNoticesForSelection.Count >= 1)
                    {
                        using (var selectViewModel = new SelectReferenceObjectViewModel(Context.Reference, false))
                        {
                            var filter = new Filter(Context.Reference.ParameterGroup);
                            filter.Terms.AddTerm(
                                Context.Reference.ParameterGroup.SystemParameters.Find(TFlex.DOCs.Model.Structure.SystemParameterType.ObjectId),
                                ComparisonOperator.IsOneOf,
                                availableNoticesForSelection.Select(obj => obj.Id).ToList());
                            selectViewModel.ContextFilter = filter;

                            if (!ApplicationManager.OpenDialog(selectViewModel, default(IDialogOwnerWindowSource)))
                                return;

                            newFirstModificationNotice = selectViewModel.GetSelectedReferenceObject();
                        }
                    }
                    else
                    {
                        var reference = Context.Connection.ReferenceCatalog.Find(ModificationNoticesReference.ModificationNoticesReferenceGuid);

                        using (var selectViewModel = new SelectReferenceObjectViewModel(reference, false))
                        {
                            var filter = new Filter(reference);
                            filter.Terms.AddTerm(reference.Description.ClassParameterInfo, ComparisonOperator.IsInheritFrom, Guids.ModificationsType);
                            filter.Terms.AddTerm(reference.Description[Guids.SourceDenotation], ComparisonOperator.IsEmptyString, null);
                            filter.Terms.AddTerm(reference.Description[TFlex.DOCs.Model.Structure.SystemParameterType.Stage], ComparisonOperator.Equal, Guids.DevelopmentStage);

                            selectViewModel.ContextFilter = filter;

                            if (!ApplicationManager.OpenDialog(selectViewModel, default(IDialogOwnerWindowSource)))
                                return;

                            newFirstModificationNotice = selectViewModel.GetSelectedReferenceObject();
                            needAddingFirstObject = true;
                        }
                    }

                    if (newFirstModificationNotice is null)
                        return;

                    if (newFirstModificationNotice == firstObject)
                    {
                        Error("Выбрано извещение, по которому уже назначено обозначение комплекта.");
                    }
                }
                else
                    return;

                // Отключаем выбранные извещения из комплекта
                var removingObjects = Desktop.CheckOut(selectedObjects, false).OfType<ReferenceObject>().ToList();
                removingObjects.AddRange(selectedObjects.Except(removingObjects));
                if (removingObjects.IsNullOrEmpty())
                    return;

                bool needEndChanges = false;

                if (!complect.Changing)
                {
                    complect.BeginChanges();
                    needEndChanges = true;
                }

                foreach (var removingObject in removingObjects)
                {
                    removingObject.BeginChanges();

                    removingObject[Guids.Number].Value = String.Empty;
                    removingObject[Guids.Denotation].Value = removingObject[Guids.SourceDenotation].Value;
                    removingObject[Guids.SourceDenotation].Value = String.Empty;

                    complect.RemoveLinkedObject(Guids.ModificationNoticesLink, removingObject);

                    removingObject.EndChanges();
                }

                Desktop.CheckIn(removingObjects, $"ИИ удалено из комплекта '{currentComplectDenotation}'", false);


                currentLinkedObjects = currentLinkedObjects.Except(removingObjects).ToList();
                int allObjectsCount = needAddingFirstObject ? currentLinkedObjects.Count + 1 : currentLinkedObjects.Count;
                string newComplectDenotation = String.Empty;

                if (needAddingFirstObject) // Выбрано другое извещение из справочника
                {
                    newComplectDenotation = (string)newFirstModificationNotice[Guids.Denotation].Value;

                    var addingFirstObject = Desktop.CheckOut(newFirstModificationNotice, false).OfType<ReferenceObject>().FirstOrDefault();
                    if (addingFirstObject is null)
                        addingFirstObject = newFirstModificationNotice;

                    addingFirstObject.BeginChanges();

                    complect.AddLinkedObject(Guids.ModificationNoticesLink, addingFirstObject);

                    string numberString = $"{1}/{allObjectsCount}";
                    addingFirstObject[Guids.Number].Value = numberString;
                    addingFirstObject[Guids.SourceDenotation].Value = addingFirstObject[Guids.Denotation];
                    addingFirstObject[Guids.Denotation].Value = $"{newComplectDenotation} {numberString}";

                    addingFirstObject.EndChanges();

                    Desktop.CheckIn(addingFirstObject, $"ИИ внесено в комплект. " +
                        $"Номер комплекта изменён с '{currentComplectDenotation}' на '{newComplectDenotation}'.", false);
                }
                else
                {
                    // Выбранное извещение из уже подключенных делаем первым

                    var editingFirstObject = Desktop.CheckOut(newFirstModificationNotice, false).OfType<ReferenceObject>().FirstOrDefault();
                    if (editingFirstObject is null)
                        editingFirstObject = newFirstModificationNotice;

                    newComplectDenotation = (string)newFirstModificationNotice[Guids.SourceDenotation].Value;

                    editingFirstObject.BeginChanges();

                    string numberString = $"{1}/{allObjectsCount}";
                    editingFirstObject[Guids.Number].Value = numberString;
                    editingFirstObject[Guids.Denotation].Value = $"{newComplectDenotation} {numberString}";

                    editingFirstObject.EndChanges();

                    Desktop.CheckIn(editingFirstObject, $"Изменён номер и обозначение. " +
                        $"Номер комплекта изменён с '{currentComplectDenotation}' на '{newComplectDenotation}'.", false);
                }

                // Обновляем обозначение и срок изменения комплекта из параметров нового "первого" извещения
                complect[Guids.Denotation].Value = newComplectDenotation;
                complect[Guids.ApplyDate].Value = newFirstModificationNotice[Guids.ApplyDate];

                if (needEndChanges)
                    complect.EndChanges();

                RefreshControls("DenotationControl", "DateControl");

                // Редактируем номер, обозначение и срок изменения остальных подключенных извещений
                currentLinkedObjects.Remove(newFirstModificationNotice);
                var editingObjects = Desktop.CheckOut(currentLinkedObjects, false).OfType<ReferenceObject>().ToList();
                editingObjects.AddRange(currentLinkedObjects.Except(editingObjects));
                if (editingObjects.Any())
                {
                    int number = 2;

                    foreach (var editingObject in editingObjects)
                    {
                        editingObject.BeginChanges();

                        string numberString = $"{number++}/{allObjectsCount}";
                        editingObject[Guids.Number].Value = numberString;
                        editingObject[Guids.Denotation].Value = $"{newComplectDenotation} {numberString}";
                        editingObject[Guids.ApplyDate].Value = newFirstModificationNotice[Guids.ApplyDate];

                        editingObject.EndChanges();
                    }

                    Desktop.CheckIn(editingObjects, "Изменён номер и обозначение", false);
                }
            }

            RefreshReferenceWindow();
        }

        public void ЗавершениеРедактированияСвойствОбъекта() => OnPropertiesEndChanges();

        private void OnPropertiesEndChanges()
        {
            var macro = Context[nameof(MacroKeys.MacroSource)] as CodeMacro;
            if (macro.DebugMode.Value == true)
            {
                System.Diagnostics.Debugger.Launch();
                System.Diagnostics.Debugger.Break();
            }

            var setOfNotices = Context.ReferenceObject;
            if (!setOfNotices.Class.IsInherit(Guids.ModificationsComplexType))
                return;

            if (CancelWhenObjectEndEditProperties)
            {
                if (setOfNotices.IsAdded)
                {
                    RemoveModificationNoticesFromSet(setOfNotices);
                }
            }
            else
            {
                // Изменить параметр "Срок изменения" у всех подключенных извещений

                var currentLinkedObjects = setOfNotices.GetObjects(Guids.ModificationNoticesLink);
                var objectsForEditing = currentLinkedObjects.Where(
                    obj => (DateTime)obj[Guids.ApplyDate] != (DateTime)setOfNotices[Guids.ApplyDate]).ToList();

                if (objectsForEditing.IsNullOrEmpty())
                    return;

                var editingObjects = Desktop.CheckOut(objectsForEditing, false).OfType<ReferenceObject>().ToList();
                editingObjects.AddRange(objectsForEditing.Except(editingObjects));

                foreach (var editingObject in editingObjects)
                {
                    editingObject.BeginChanges();
                    editingObject[Guids.ApplyDate].Value = setOfNotices[Guids.ApplyDate];
                    editingObject.EndChanges();
                }

                Desktop.CheckIn(editingObjects, "Изменён срок изменения", false);
            }
        }

        public void ОтменаИзменений() => OnUndoCheckOut();

        private void OnUndoCheckOut()
        {
            var macro = Context[nameof(MacroKeys.MacroSource)] as CodeMacro;
            if (macro.DebugMode.Value == true)
            {
                System.Diagnostics.Debugger.Launch();
                System.Diagnostics.Debugger.Break();
            }

            var setOfNotices = Context.ReferenceObject;
            if (!setOfNotices.Class.IsInherit(Guids.ModificationsComplexType))
                return;

            if (setOfNotices.LockState == ReferenceObjectLockState.CheckedOutForAdd)
            {
                RemoveModificationNoticesFromSet(setOfNotices);
            }
        }

        private void RemoveModificationNoticesFromSet(ReferenceObject setOfNotices)
        {
            var linkedObjects = setOfNotices.GetObjects(Guids.ModificationNoticesLink);

            var removingObjects = Desktop.CheckOut(linkedObjects, false).OfType<ReferenceObject>().ToList();
            removingObjects.AddRange(linkedObjects.Except(removingObjects));
            if (removingObjects.IsNullOrEmpty())
                return;

            string currentComplectDenotation = (string)setOfNotices[Guids.Denotation].Value;

            foreach (var removingObject in removingObjects)
            {
                removingObject.BeginChanges();

                removingObject[Guids.Number].Value = String.Empty;
                removingObject[Guids.Denotation].Value = removingObject[Guids.SourceDenotation].Value;
                removingObject[Guids.SourceDenotation].Value = String.Empty;

                removingObject.EndChanges();
            }

            Desktop.CheckIn(removingObjects, $"ИИ удалено из комплекта '{currentComplectDenotation}'", false);
        }
    }
}

