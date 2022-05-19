using System;
using System.Collections.Generic;
using TFlex.DOCs.Client;
using TFlex.DOCs.Model.Macros;
using TFlex.DOCs.Model.Macros.ObjectModel;
using TFlex.DOCs.Model.References;
using TFlex.Technology.References;
using TFlex.DOCs.Model.References.Nomenclature;
using TFlex.Model.Technology.References.TechnologyElements;
using System.Linq;
using TFlex.DOCs.Client.ViewModels;
using TFlex.DOCs.Model;
using TFlex.DOCs.Model.Classes;
using TFlex.DOCs.Model.References.Links;
using TFlex.DOCs.Model.Search;
using TFlex.DOCs.UI.Client.Technology.Dialogs.TechnologicalVariableData.OperationsDialog;
using TFlex.Model.Technology.References.TechnologyElements.TechnologicalVariableData;
using TFlex.Technology;

namespace Technology
{
    public class GroupTypicalTechProcessMacros : MacroProvider
    {
        private static class RelationsGuidHelper
        {
            /// <summary>
            /// Связь тех.процесса с Материалами ТП (Материалы ТП)
            /// </summary>
            public static readonly Guid TechProcessLinksToMaterials = new Guid("bbdc2edb-55b8-4eb6-bdd8-e9d6397bc83c");

            /// <summary>
            /// Связь операции с Материалами ТП (Материалы операции)
            /// </summary>
            public static readonly Guid OperationLinksToMaterials = new Guid("beeab0ff-1598-44b5-a2d4-32fdf0e98e90");

            /// <summary>
            /// Связь объектов технологического процесса с оснащением (Оснащение ТП)
            /// </summary>
            public static readonly Guid TechnologyElementLinksToEquipment = new Guid("ad1b95c0-dfba-41b6-b7b6-876ea7125985");
        }

        public GroupTypicalTechProcessMacros(MacroContext context) : base(context)
        {
        }

        public override void Run()
        {
        }

        /// <summary>
        /// Добавление операций в переменные данные ТП
        /// </summary>
        public void AddOperations()
        {
            // 1. собрать добавленные в переменные данные операции в HashSet
            // 2. перейти к ТП, собрать все операции оттуда, отмечая, какие операции уже есть
            // 3. показать в диалоге ввода
            // 4. синхронизировать операции

            if (Context.ReferenceObject is not TechProcessVariableDataReferenceObject techProcessVariableData)
                return;

            var techProcess = techProcessVariableData.GetGroupTechProcess() as StructuredTechnologicalProcess;

            if (techProcess is null)
                Error("В переменных данных ТП не задан групповой / типовой тех. процесс");

            var product = techProcessVariableData.GetProduct() as MaterialObject;

            if (product is null)
                Error("К переменным данным ТП не привязано изготавливаемое изделие!");

            var dictionary = new Dictionary<Guid, ReferenceObject>();

            foreach (var operation in techProcessVariableData.GetOperations()
                .OfType<StructuredTechnologicalOperation>())
            {
                if (!dictionary.ContainsKey(operation.Guid))
                    dictionary.Add(operation.Guid, operation);
            }

            var owner = ((UIMacroContext)Context).OwnerViewModel;
            IEnumerable<ISelectedOperationInfo> operations = Array.Empty<ISelectedOperationInfo>();

            Context.RunOnUIThread(() =>
            {
                using ISelectTypicalGroupOperationsDialog dialog = VariableDataSelectOperationsDialogFactory.Create(dictionary, techProcess, owner, Context.Connection);
                if (!ApplicationManager.OpenDialog(dialog, owner))
                    return;

                operations = dialog.GetSelectedOperations();
            });

            techProcessVariableData.Modify(o =>
            {
                // синхронизация. Добавить то, что нужно добавить. Удалить то, что надо удалить
                foreach (var operationInfo in operations)
                {
                    if (operationInfo.NeedAdd)
                    {
                        if (dictionary.ContainsKey(operationInfo.Operation.Guid))
                            continue;

                        techProcessVariableData.AddOperation(operationInfo.Operation);

                        CreateVariableDataForOperationWithSteps(operationInfo.Operation, product);
                    }
                    else
                    {
                        if (!dictionary.ContainsKey(operationInfo.Operation.Guid))
                            continue;

                        techProcessVariableData.RemoveOperation(operationInfo.Operation);
                    }
                }
            });
        }

        /// <summary>
        /// Доступность кнопки Заимствовать переменные данные
        /// </summary>
        /// <returns></returns>
        public ButtonValidator CheckEnableCopyVariableData()
        {
            bool visible = Context.Reference.IsSlave &&
                           Context.Reference.LinkInfo.RootMasterReference is TechnologicalProcessReference;

            return new ButtonValidator() { Enable = true, Visible = visible };
        }

        /// <summary>
        /// Заимствование переменных данных ТП
        /// </summary>
        public void CloneVariableData()
        {
            if (Context.ReferenceObject is not TechProcessVariableDataReferenceObject techProcessVariableData)
                return;

            // Выбор изделия из ЭСИ
            // Создание копии выбранных переменных переменных данных ТП со всеми переменными данными операций и переходв
            // Подстановка выбранного изделия в связь ДСЕ

            ISelectObjectDialog dialog =
                Context.CreateSelectObjectDialog(new NomenclatureReference(Context.Connection));

            if (!dialog.Show())
                return;

            var selectedProduct = dialog.SelectedObjects.FirstOrDefault();

            var techProcessVariableDataClone = techProcessVariableData.CreateFullCopy(new List<Guid>()
                {
                    TechnologicalVariableDataReferenceObject.RelationKeys.Materials
                }) as TechProcessVariableDataReferenceObject;

            techProcessVariableDataClone.Modify(o =>
            {
                o.SetProduct(selectedProduct);
                o.Name.Value = selectedProduct[NomenclatureReferenceObject.FieldKeys.Object].GetString();
            });

            var search = new SearchService(Context.Connection);

            foreach (var groupOperation in techProcessVariableData.GetOperations().OfType<StructuredTypicalGroupOperation>())
            {
                string filterString = $"[Тип] = 'Данные операции' И [Операция]->[Guid] = '{groupOperation.Guid}'";
                var operationVariableData = search.Find(filterString).FirstOrDefault();

                if (operationVariableData is null)
                    continue;

                var operationVariableDataClone = operationVariableData
                    .CreateFullCopy(new List<Guid>()
                    {
                        TechnologicalVariableDataReferenceObject.RelationKeys.Materials,
                        TechnologicalVariableDataReferenceObject.RelationKeys.Equipment
                    }) as OperationVariableDataReferenceObject;

                operationVariableDataClone.Modify(o =>
                {
                    o.SetProduct(selectedProduct);
                    o.Name.Value = groupOperation.Name;
                });

                foreach (var groupStep in groupOperation.Children.OfType<StructuredTypicalGroupTechnologicalStep>())
                {
                    filterString =
                        $"[Тип] = 'Данные перехода' И [Технологический переход]->[Guid] = '{groupStep.Guid}'";

                    var stepVariableData = search.Find(filterString).FirstOrDefault();

                    // не нашли, тогда создаем
                    if (stepVariableData is null)
                        continue;

                    var stepVariableDataClone = stepVariableData.CreateFullCopy(new List<Guid>()
                    {
                        TechnologicalVariableDataReferenceObject.RelationKeys.Equipment
                    }) as StepVariableDataReferenceObject;

                    stepVariableDataClone.Modify(o =>
                    {
                        o.SetProduct(selectedProduct);
                        o.Name.Value = groupStep.Text;
                    });
                }
            }

            var techProcess = techProcessVariableData.GetGroupTechProcess() as StructuredTechnologicalProcess;
            techProcess.Modify(o => { ((OneToManyLink)o.ProducedDSEGroup).AddLinkedObject(selectedProduct); });
        }

        /// <summary>
        /// Создание единичного тех. процесса из переменных данных ТП
        /// </summary>
        public void CreateTechnologicalProcess() => CreateTechnologicalProcess(Context.ReferenceObject as TechProcessVariableDataReferenceObject);

        /// <summary>
        /// Обработчик события удаления переменных данных
        /// </summary>
        public void OnDeleteVariableData()
        {
            void OnDeleteTechProcessVariableData(TechnologicalVariableDataReferenceObject variableDataReferenceObject)
            {
                // при удалении переменных данных тех. процесса удаляем переменные данные операций, относящиеся к текущему изделию
                // отключить изготавливаемое изделие из связи
                bool result = Question("Будут удалены все переменные данные операций и переходов. Продолжить?");

                if (!result)
                    Cancel();

                var techProcessVariableData = variableDataReferenceObject as TechProcessVariableDataReferenceObject;
                var groupTechProcess =
                    techProcessVariableData.GetGroupTechProcess() as StructuredTypicalGroupTechnologicalProcess;

                if (groupTechProcess is null)
                    return;

                var product = techProcessVariableData.GetProduct();

                if (product is null)
                    return;

                groupTechProcess?.Modify(o =>
                    o.RemoveLinkedObject(TechnologicalProcess.OneToManyGroups.ProducedDSE, product));

                var technologicalVariableDataReference = new TechnologicalVariableDataReference(Context.Connection);
                var filter = Filter.Parse($"[Тип] = 'Данные операции' И [ДСЕ]->[Guid] = '{product.Guid}'",
                    technologicalVariableDataReference.ParameterGroup);

                using var saveSet = new ReferenceObjectSaveSet();

                foreach (var operationVariableData in technologicalVariableDataReference.Find(filter))
                {
                    saveSet.AddObjectToDelete(operationVariableData);
                }

                saveSet.EndChanges();
            }

            void OnDeleteOperationVariableData(TechnologicalVariableDataReferenceObject variableDataReferenceObject)
            {
                var product = variableDataReferenceObject.GetProduct();

                if (product is null)
                    return;

                var technologicalVariableDataReference = new TechnologicalVariableDataReference(Context.Connection);
                var techProcessFilter = Filter.Parse($"[Тип] = 'Данные ТП' И [ДСЕ]->[Guid] = '{product.Guid}'",
                    technologicalVariableDataReference.ParameterGroup);

                var techProcessVariableDataList = technologicalVariableDataReference.Find(techProcessFilter);

                // отключаем групповую операцию, присоединенную к удаляемым переменным данным операции, от переменных данных ТП 
                if (techProcessVariableDataList.FirstOrDefault() is TechProcessVariableDataReferenceObject
                    techProcessVariableData)
                {
                    var operation = ((OperationVariableDataReferenceObject)Context.ReferenceObject).GetGroupOperation();
                    techProcessVariableData.Modify(o => o.RemoveOperation(operation));
                }

                var saveSet = new ReferenceObjectSaveSet();
                var stepFilter = Filter.Parse($"[Тип] = 'Данные перехода' И [ДСЕ]->[Guid] = '{product.Guid}'",
                    technologicalVariableDataReference.ParameterGroup);

                // при удалении переменных данных операции удаляем переменные данные переходов
                foreach (var stepVariableData in technologicalVariableDataReference.Find(stepFilter))
                {
                    saveSet.AddObjectToDelete(stepVariableData);
                }

                saveSet.EndChanges();
            }

            Action<TechnologicalVariableDataReferenceObject> strategy = null;
            if (Context.ReferenceObject is TechProcessVariableDataReferenceObject)
            {
                strategy = OnDeleteTechProcessVariableData;
            }
            else if (Context.ReferenceObject is OperationVariableDataReferenceObject)
            {
                strategy = OnDeleteOperationVariableData;
            }

            strategy?.Invoke(Context.ReferenceObject as TechnologicalVariableDataReferenceObject);
        }

        /// <summary>
        /// Обработчик события завершения сохранения переменных данных операции
        /// </summary>
        public void OnSaveOperationVariableData()
        {
            if (Context.ReferenceObject is not OperationVariableDataReferenceObject operationVariableData)
                return;

            if (!Context.ReferenceObject.IsNew)
                return;

            // получить групповую операцию
            // создать переменные данные переходов для этой операции
            var operation = operationVariableData.GetGroupOperation() as StructuredTechnologicalOperation;

            if (operation is null)
                Error("Не задана типовая / групповая операция. Невозможно создать переменные данные для дочерних переходов");

            var product = operationVariableData.GetProduct() as MaterialObject;

            if (product is null)
                Error("Не задано изготавливаемое изделие");

            var techProcess = operation.Process;
            techProcess.Modify(o => { ((OneToManyLink)o.ProducedDSEGroup).AddLinkedObject(product); });

            using var saveSet = new ReferenceObjectSaveSet();
            var stepClass = operationVariableData.Reference.ParameterGroup.Classes.Find(TechnologicalVariableDataTypes.Keys.StepVariableData);

            CreateVariableDataForSteps(operation, product, stepClass, operationVariableData.Reference, saveSet);

            saveSet.EndChanges();

            // проверить, есть ли у связанного изготавливаемого изделия связь с переменными данными ТП
            var technologicalVariableDataReference = new TechnologicalVariableDataReference(Context.Connection);
            var filter = Filter.Parse($"[Тип] = 'Данные ТП' И [ДСЕ]->[Guid] = '{product.Guid}'",
                technologicalVariableDataReference.ParameterGroup);

            var tpVariableDataWithProduct =
                technologicalVariableDataReference.Find(filter).FirstOrDefault() as
                    TechProcessVariableDataReferenceObject;

            if (tpVariableDataWithProduct is null)
                tpVariableDataWithProduct = CreateTechnologicalProcessVariableDataObject(product);

            tpVariableDataWithProduct.Modify(o => o.AddOperation(operationVariableData.GetGroupOperation()));
        }

        /// <summary>
        /// Обработчик события добавления изделия по связи Изготавливаемые ДСЕ
        /// </summary>
        public void OnAddLinkedProductToTechProcessVariableData()
        {
            // проверить, что изменяемая связь - это Изготавливаемые ДСЕ
            // проверить, что объект в контексте - это тех. процесс
            // проверить, что таких переменных данных еще нет
            // создать соответствующие переменные данные

            if (Context.ChangedLink.LinkGroup.Guid != TechnologicalProcess.OneToManyGroups.ProducedDSE)
                return;

            if (Context.ReferenceObject is not StructuredTypicalGroupTechnologicalProcess techProcess)
                return;

            var product = ((ObjectLinkChangedEventArgs)Context.ModelChangedArgs).AddedObject as MaterialObject;

            if (product is null)
                return;

            var variableDataReference = new TechnologicalVariableDataReference(Context.Connection);
            var filter = Filter.Parse(
                    $"[Тип] = 'Данные ТП' И [Тех. процесс]->[Guid] = '{techProcess.Guid}' И [ДСЕ]->[Guid] = '{product.Guid}'",
                    variableDataReference.ParameterGroup);

            if (variableDataReference.Find(filter).Any())
                return;

            CreateTechnologicalProcessVariableDataObject(product);
        }

        private void CreateVariableDataForOperationWithSteps(StructuredTechnologicalOperation operation, MaterialObject product)
        {
            using var saveSet = new ReferenceObjectSaveSet();
            var variableDataReference = new TechnologicalVariableDataReference(Context.Connection);
            var classToCreate =
                variableDataReference.ParameterGroup.Classes.Find(TechnologicalVariableDataTypes.Keys.OperationVariableData);

            CreateOperationVariableData(operation, product, classToCreate, variableDataReference, saveSet);

            classToCreate = variableDataReference.ParameterGroup.Classes.Find(TechnologicalVariableDataTypes.Keys.StepVariableData);
            CreateVariableDataForSteps(operation, product, classToCreate, variableDataReference, saveSet);

            saveSet.EndChanges();
        }

        private void CreateOperationVariableData(StructuredTechnologicalOperation operation, MaterialObject product,
            ClassObject classObject, TechnologicalVariableDataReference variableDataReference, ReferenceObjectSaveSet saveSet)
        {
            var operationVariableDataReferenceObject =
                variableDataReference.CreateReferenceObject(classObject) as OperationVariableDataReferenceObject;

            operationVariableDataReferenceObject.SetProduct(product);
            operationVariableDataReferenceObject.SetGroupOperation(operation);
            operationVariableDataReferenceObject.Name.Value = operation.Name;

            saveSet.Add(operationVariableDataReferenceObject);
        }

        private void CreateVariableDataForSteps(StructuredTechnologicalOperation operation, MaterialObject product,
            ClassObject classObject, TechnologicalVariableDataReference variableDataReference, ReferenceObjectSaveSet saveSet)
        {
            foreach (var step in operation.Children.OfType<StructuredTypicalGroupTechnologicalStep>())
            {
                CreateStepVariableDataObject(step, product, classObject, variableDataReference, saveSet);
            }
        }

        private void CreateStepVariableDataObject(StructuredTypicalGroupTechnologicalStep step, MaterialObject product,
            ClassObject classObject, TechnologicalVariableDataReference variableDataReference, ReferenceObjectSaveSet saveSet)
        {
            var stepVariableDataReferenceObject =
                variableDataReference.CreateReferenceObject(classObject) as StepVariableDataReferenceObject;

            stepVariableDataReferenceObject.SetGroupTechnologicalStep(step);
            stepVariableDataReferenceObject.SetProduct(product);
            stepVariableDataReferenceObject.Name.Value = step.Text;

            saveSet.Add(stepVariableDataReferenceObject);
        }

        private void CreateTechnologicalProcess(
            TechProcessVariableDataReferenceObject techProcessVariableDataObject)
        {
            if (techProcessVariableDataObject is null)
                return;

            // получить изготавливаемое изделие
            var product = techProcessVariableDataObject.GetProduct() as MaterialObject;

            if (product is null)
                Error("Не указано изготавливаемое изделие!");

            var groupTechProcess = techProcessVariableDataObject.GetGroupTechProcess();

            if (groupTechProcess is null)
                Error("Не указан групповой / типовой тех. процесс");

            var technologicalProcessReference = new TechnologicalProcessReference(Context.Connection);
            var technologicalProcessClass =
                technologicalProcessReference.Classes.Find(Technology2012Classes.StructuredTechnologicalProcessType);

            var technologicalProcessObject = groupTechProcess.CreateFullCopy(new List<Guid>()
                    {
                        RelationsGuidHelper.TechProcessLinksToMaterials,
                        RelationsGuidHelper.TechnologyElementLinksToEquipment
                    }, technologicalProcessClass) as StructuredTechnologicalProcess;

            technologicalProcessObject.Modify(o =>
            {
                string caption = product[NomenclatureReferenceObject.FieldKeys.Object].GetString();
                technologicalProcessObject.Name.Value = caption;

                ((OneToManyLink)technologicalProcessObject.ProducedDSEGroup).RemoveAll();
                ((OneToManyLink)technologicalProcessObject.ProducedDSEGroup).AddLinkedObject(product);

                var materials = techProcessVariableDataObject.GetMaterials();

                foreach (var material in materials)
                {
                    technologicalProcessObject.AddLinkedObject(RelationsGuidHelper.TechProcessLinksToMaterials,
                        material);
                }
            });

            IDictionary<Guid, Guid> mappings = new Dictionary<Guid, Guid>();
            mappings.Add(Technology2012Classes.TypicalGroupOperation, Technology2012Classes.TechnologyOperationType);
            mappings.Add(Technology2012Classes.TypicalGroupTechnologicalStep, Technology2012Classes.TechnologyStepType);

            var operationClassObject =
                technologicalProcessReference.Classes.Find(Technology2012Classes.TechnologyOperationType);

            var searchService = new SearchService(Context.Connection);

            // единичный тех. процесс требуется наполнить только теми операциями, которые указаны в связи с переменными данными ТП
            foreach (var groupOperation in techProcessVariableDataObject.GetOperations())
            {
                var operation = groupOperation.CreateFullCopy(new List<Guid>()
                    {
                        RelationsGuidHelper.TechProcessLinksToMaterials,
                        RelationsGuidHelper.TechnologyElementLinksToEquipment
                    }, operationClassObject, technologicalProcessObject) as
                    StructuredTechnologicalOperation;

                string filterString =
                    $"[Тип] = 'Данные операции' И [ДСЕ]->[Guid] = '{product.Guid}' И [Операция]->[Guid] = '{groupOperation.Guid}'";

                var possibleOperationVariableDataList = searchService.Find(filterString);
                operation.Modify(op =>
                {
                    foreach (var operationVariableData in possibleOperationVariableDataList.OfType<OperationVariableDataReferenceObject>())
                    {
                        var equipment = operationVariableData.GetEquipment();

                        foreach (var tool in equipment)
                        {
                            op.AddLinkedObject(RelationsGuidHelper.TechnologyElementLinksToEquipment, tool);
                        }

                        var materials = operationVariableData.GetMaterials();

                        foreach (var material in materials)
                        {
                            op.AddLinkedObject(RelationsGuidHelper.OperationLinksToMaterials, material);
                        }
                    }
                });

                FillTechProcessRecursive(groupOperation, operation, mappings, searchService, product);
            }
        }

        private void FillTechProcessRecursive(ReferenceObject rootSender, ReferenceObject rootReciever, IDictionary<Guid, Guid> mappings,
            SearchService service, ReferenceObject product)
        {
            foreach (var child in rootSender.Children)
            {
                if (!mappings.ContainsKey(child.Class.Guid))
                    Error($"Невозможно подставить соответствие указанному типу {child.Class} для единичного тех. процесса");

                var classObject = rootSender.Reference.Classes.Find(mappings[child.Class.Guid]);
                var newChild = rootReciever.CreateFullCopy(new List<Guid>()
                {
                    RelationsGuidHelper.TechnologyElementLinksToEquipment
                }, classObject, rootReciever);

                newChild.BeginChanges();

                // найти переменные данные, относящиеся именно к этой групповой операции (групповому пр) и именно к этому изделию
                // переложить переменные данные по материалам, оснащению в нужные связи ТП, операций, переходов
                if (newChild is StructuredTechnologicalStep step)
                {
                    string filterString =
                        $"[Тип] = 'Данные перехода' И [ДСЕ]->[Guid] = '{product.Guid}' И [Технологический переход]->[Guid] = '{child.Guid}'";

                    var possibleStepVariableDataList = service.Find(filterString);
                    foreach (var stepVariableData in possibleStepVariableDataList
                        .OfType<StepVariableDataReferenceObject>())
                    {
                        var equipment = stepVariableData.GetEquipment();

                        foreach (var tool in equipment)
                        {
                            step.AddLinkedObject(RelationsGuidHelper.TechnologyElementLinksToEquipment, tool);
                        }
                    }
                }

                if (newChild.Changing)
                    newChild.EndChanges();

                FillTechProcessRecursive(child, newChild, mappings, service, product);
            }
        }

        private TechProcessVariableDataReferenceObject CreateTechnologicalProcessVariableDataObject(
            ReferenceObject product)
        {
            var variableDataReference = new TechnologicalVariableDataReference(Context.Connection);
            var techProcessVariableDataClass =
                variableDataReference.ParameterGroup.Classes.Find(TechnologicalVariableDataTypes.Keys.TechProcessVariableData);

            var variableDataReferenceObject =
                variableDataReference.CreateReferenceObject(techProcessVariableDataClass) as TechProcessVariableDataReferenceObject;

            variableDataReferenceObject.Name.Value = product[NomenclatureReferenceObject.FieldKeys.Object].GetString();
            variableDataReferenceObject.SetProduct(product);

            if (Context.ReferenceObject is StructuredTypicalGroupOperation operation)
            {
                variableDataReferenceObject.SetGroupTechProcess(operation.Parent);
            }
            else if (Context.ReferenceObject is StructuredTypicalGroupTechnologicalProcess techProcess)
            {
                variableDataReferenceObject.SetGroupTechProcess(techProcess);
            }

            variableDataReferenceObject.EndChanges();

            return variableDataReferenceObject;
        }
    }

    public class SearchService
    {
        private TechnologicalVariableDataReference _reference;

        public SearchService(ServerConnection connection)
        {
            _reference = new TechnologicalVariableDataReference(connection);
        }

        public IList<ReferenceObject> Find(string filterString)
        {
            var filter = Filter.Parse(filterString, _reference.ParameterGroup);
            return _reference.Find(filter);
        }
    }
}
