using System;
using System.Collections.Generic;
using System.Linq;
using TFlex.DOCs.Client;
using TFlex.DOCs.Client.ViewModels;
using TFlex.DOCs.Model.Macros;
using TFlex.DOCs.Model.Macros.ObjectModel;
using TFlex.DOCs.Model.References;
using TFlex.DOCs.Model.References.Links;
using TFlex.DOCs.Model.References.Nomenclature;
using TFlex.DOCs.UI.Client.Technology.Dialogs.AssemblyOperations.AssemblyNodeDialog;
using TFlex.DOCs.UI.Client.Technology.Dialogs.AssemblyOperations.PartsDialog;
using TFlex.Model.Technology.References.ParametersProvider;
using TFlex.Model.Technology.References.TechnologyElements;
using TFlex.Technology.References;

namespace Macros
{
    public static class GuidHelper
    {
        /// <summary>
        /// Связь Комплектующие от сборочной операции на подключение в ЭСИ
        /// </summary>
        public static readonly Guid LinkToParts = new Guid("25a393dc-8f97-4e25-aa68-30f8382cd756");

        /// <summary>
        /// Связь между Сборочным узлом и сборочной операцией
        /// </summary>
        public static readonly Guid AssemblyNodeToAssemblyOperationRelation =
            new Guid("1caaf34a-744e-42a7-b618-a90e2f8e34f4");

        /// <summary>
        /// Guid типа Сборочный узел справочника ЭСИ
        /// </summary>
        public static readonly Guid AssemblyNodeTypeGuid = new Guid("ba08bce5-931c-47f5-a813-ec5e374e7c8a");
    }

    public class AssemblyOperationMacros : MacroProvider
    {
        public AssemblyOperationMacros(MacroContext context) : base(context)
        {
        }

        /// <summary>
        /// Создание сборочного узла
        /// </summary>
        public void CreateAssemblyNode()
        {
            IAssemblyNodeCreationService service = new AssemblyNodeCreationService();
            if (!(Context.ReferenceObject is AssemblyTechnologicalOperation assemblyOperation))
                return;

            service.CreateAssemblyNode(new AssemblyOperationAdapter(assemblyOperation));
            Context.RefreshControls();
        }

        /// <summary>
        /// Подключить комплектующие
        /// </summary>
        public void AddPartsToAssemblyOperation()
        {
            if (Context.ReferenceObject is AssemblyTechnologicalOperation operation && operation.GetAssemblyNode() is null)
                Error("К сборочной операции не подключен сборочный узел!");

            StructuredTechnologicalProcess techProcess =
                ((AssemblyTechnologicalOperation)Context.ReferenceObject).Process;

            var owner = ((UIMacroContext)Context).OwnerViewModel;
            ISelectPartsDialog selectObjectDialog =
                PartsSelectionDialogFactory.Create(techProcess, owner, Context.Connection);

            if (!ApplicationManager.OpenDialog(selectObjectDialog, owner))
                return;

            IEnumerable<IMaterialObject> materialObjects = selectObjectDialog.GetSelectedCompositionObjects()
                .Select(obj => new MaterialObjectAdapter(obj))
                .ToList();

            if (materialObjects is null)
                return;

            IAssemblyNodeCreationService service = new AssemblyNodeCreationService();
            if (!(Context.ReferenceObject is AssemblyTechnologicalOperation assemblyOperation))
                return;

            service.RemoveParts(materialObjects);
            service.AddParts(new AssemblyOperationAdapter(assemblyOperation), materialObjects);
            Context.RefreshControls();
        }

        /// <summary>
        /// Подключить одиночный сборочный узел
        /// </summary>
        public void AddAssemblyAsAssemblyNode()
        {
            ITechnologicalProcessParametersProvider techProcess =
                ((IAssemblyTechnologicalOperationParametersProvider)Context.ReferenceObject).TechnologicalProcess;

            var owner = ((UIMacroContext)Context).OwnerViewModel;

            IAssemblyNode assemblyNode = null;

            using ISelectAssemblyNodeDialog dialog =
                    AssemblyNodeSelectionDialogFactory.Create(techProcess, owner, Context.Connection);

            if (!ApplicationManager.OpenDialog(dialog, owner))
                return;

            var materialObject = dialog.GetSelectedCompositionObject();       
            assemblyNode = new AssemblyNodeAdapter(materialObject);

            if (!assemblyNode.HasChildren)
                Error("Выбранный объект не содержит дочерних элементов");

            IAssemblyNodeCreationService service = new AssemblyNodeCreationService();
            if (!(Context.ReferenceObject is AssemblyTechnologicalOperation assemblyOperation))
                return;

            service.LinkAssemblyNode(new AssemblyOperationAdapter(assemblyOperation), assemblyNode);
            Context.RefreshControls();
        }

        public void SynchronizeParts()
        {
            IAssemblyOperation currentOperation =
                new AssemblyOperationAdapter(Context.ReferenceObject as AssemblyTechnologicalOperation);

            IAssemblyNodeCreationService service = new AssemblyNodeCreationService();
            service.SyncronizeParts(currentOperation, currentOperation.GetAssemblyNode());
            Context.RefreshControls();
        }

        /// <summary>
        /// Создать сборочные операции на основе узлов изготавливаемого изделия
        /// </summary>
        public void CreateAssemblyOperations()
        {
            if (!(Context.ReferenceObject is ITechnologicalProcessParametersProvider techProcess))
            {
                Error($"Указанный объект {Context.ReferenceObject} не является тех. процессом!");
                return;
            }

            IInputDialog inputDialog = Context.CreateInputDialog();
            inputDialog.AddFlagField("Создавать сборочную операцию для текущей сборки", true, true, true);

            if (!inputDialog.Show(Context))
                return;

            IAssemblyNodeCreationService service = new AssemblyNodeCreationService();
            service.CreateAssemblyOperations(new TechnologyProcessAdapter(techProcess));

            if ((bool)inputDialog.GetValue("Создавать сборочную операцию для текущей сборки"))
                service.CreateAssemblyOperationFromMainProduct(new TechnologyProcessAdapter(techProcess));
        }

        private ISelectObjectDialog CreateDialog(bool multiSelect)
        {
            var currentOperation = Context.ReferenceObject as ITechnologicalOperationParametersProvider;
            NomenclatureReferenceObject producedNomenclatureObject =
                currentOperation.TechnologicalProcess.ProducedDSE.FirstOrDefault();

            if (producedNomenclatureObject is null)
                Error("Нет ДСЕ, подключенных по связи Изготавливаемые ДСЕ");

            Reference reference = new NomenclatureReference(Context.Connection);

            ISelectObjectDialog dialog = Context.CreateSelectObjectDialog(reference);
            dialog.FocusedObject = producedNomenclatureObject;
            dialog.MultipleSelect = multiSelect;

            return dialog;
        }
    }

    #region Services

    /// <summary>
    /// Реализация сервиса для работы со сборочными операциями и сборочными узлами
    /// </summary>
    public class AssemblyNodeCreationService : IAssemblyNodeCreationService
    {
        /// <inheritdoc cref="IAssemblyNodeCreationService.CreateAssemblyNode"/>
        public void CreateAssemblyNode(IAssemblyOperation assemblyOperation)
        {
            assemblyOperation.ClearParts();
            assemblyOperation.CreateAssemblyNodeInProducedMaterialObject(assemblyOperation);
        }

        /// <inheritdoc cref="IAssemblyNodeCreationService.AddParts"/>
        public void AddParts(IAssemblyOperation assemblyOperation, IEnumerable<IMaterialObject> parts)
        {
            if (assemblyOperation.GetAssemblyNode()?.GetCoreObject() is null)
                throw new InvalidOperationException("К сборочной операции не подключен сборочный узел!");

            // отключить от изготавливаемого изделия
            // подключить в сборочный узел
            foreach (IMaterialObject part in parts)
            {
                IComplexHierarchyLink link = assemblyOperation.AddPartToAssemblyNode(part);
                assemblyOperation.AddPart(link);
            }
        }

        /// <inheritdoc cref="IAssemblyNodeCreationService.LinkAssemblyNode"/>
        public void LinkAssemblyNode(IAssemblyOperation assemblyOperation, IAssemblyNode assemblyNode)
        {
            assemblyOperation.LinkAssemblyNode(assemblyNode);
            assemblyOperation.ClearParts();

            foreach (IComplexHierarchyLink link in assemblyNode.GetComplexHierarchyLinksToChildren())
            {
                assemblyOperation.AddPart(link);
            }
        }

        /// <inheritdoc cref="IAssemblyNodeCreationService.SyncronizeParts"/>
        public void SyncronizeParts(IAssemblyOperation assemblyOperation, IAssemblyNode assemblyNode)
        {
            if (assemblyNode is null)
            {
                assemblyOperation.ClearParts();
                return;
            }
            
            List<IComplexHierarchyLink> materialObjectsListInAssemblyNode =
                assemblyNode.GetComplexHierarchyLinksToChildren().ToList();

            materialObjectsListInAssemblyNode.Sort();

            List<IComplexHierarchyLink> materialObjectsListInAssemblyOperation =
                assemblyOperation.GetMaterialObjects().ToList();

            materialObjectsListInAssemblyOperation.Sort();

            // на добавление
            foreach (IComplexHierarchyLink link in materialObjectsListInAssemblyNode)
            {
                int index = materialObjectsListInAssemblyOperation.BinarySearch(link);
                if (index < 0)
                    assemblyOperation.AddPart(link);
            }

            // на удаление
            foreach (IComplexHierarchyLink link in materialObjectsListInAssemblyOperation)
            {
                int index = materialObjectsListInAssemblyNode.BinarySearch(link);
                if (index < 0)
                    assemblyOperation.DeletePart(link);
            }
        }

        /// <inheritdoc cref="IAssemblyNodeCreationService.CreateAssemblyOperations"/>
        public void CreateAssemblyOperations(ITechnologicalProcess technologicalProcess)
        {
            IMaterialObject producedProduct = GetProducedProduct(technologicalProcess);

            foreach (IComplexHierarchyLink link in producedProduct.GetComplexHierarchyLinksToChildren().Where(l => l.Child.HasChildren))
            {
                technologicalProcess.CreateAssemblyOperation(link.Child);
            }
        }

        /// <inheritdoc cref="IAssemblyNodeCreationService.CreateAssemblyOperationFromMainProduct"/>
        public void CreateAssemblyOperationFromMainProduct(ITechnologicalProcess technologicalProcess)
        {
            IMaterialObject producedProduct = GetProducedProduct(technologicalProcess);
            technologicalProcess.CreateAssemblyOperation(producedProduct);
        }

        public void RemoveParts(IEnumerable<IMaterialObject> materialObjects)
        {
            foreach (IMaterialObject materialObject in materialObjects)
            {
                var childObject = materialObject.GetCoreObject();
                childObject.DeleteLink(materialObject.Link.Link);
            }
        }

        private IMaterialObject GetProducedProduct(ITechnologicalProcess technologicalProcess)
        {
            if (!technologicalProcess.CheckProductExistInConfiguration())
                throw new InvalidOperationException("Изготавливаемый объект не подходит под условия конфигурирования");

            IMaterialObject producedProduct = technologicalProcess.GetProducedProduct();
            if (producedProduct is null)
                throw new InvalidOperationException("Отсутствует изготавливаемое изделие!");

            return producedProduct;
        }
    }

    #endregion

    /// <summary>
    /// Объект справочника
    /// </summary>
    public interface IReferenceObject
    {
        /// <summary>
        /// Получить объект справочника
        /// </summary>
        ReferenceObject GetCoreObject();
    }

    /// <summary>
    /// Объект справочника Технологические процессы типа Технологический процесс
    /// </summary>
    public interface ITechnologicalProcess : IReferenceObject
    {
        /// <summary>
        /// Изготавливаемые ДСЕ
        /// </summary>
        IMaterialObject GetProducedProduct();

        /// <summary>
        /// Создать сборочную операцию в составе ТП
        /// </summary>
        /// <param name="materialObject">Объект состава изделия</param>
        void CreateAssemblyOperation(IMaterialObject materialObject);

        /// <summary>
        /// Получить операции из тех. процесса
        /// </summary>
        IEnumerable<IAssemblyOperation> GetOperations();

        /// <summary>
        /// Проверить, подходит ли изготавливаемый объект под условия конфигурирования
        /// </summary>
        bool CheckProductExistInConfiguration();
    }

    /// <summary>
    /// Сборочные операции в технологическом процессе реализуют описание технологии сборки и должны содержать перечень комплектующих,
    /// которые расходуются в процессе сборки.
    /// </summary>
    public interface IAssemblyOperation : IReferenceObject
    {
        /// <summary>
        /// Ссылка на ТП
        /// </summary>
        ITechnologicalProcessParametersProvider TechnologicalProcess { get; }

        /// <summary>
        /// Ссылка на Сборочный узел
        /// </summary>
        IAssemblyNode GetAssemblyNode();

        /// <summary>
        /// Комплектующие
        /// </summary>
        IEnumerable<IComplexHierarchyLink> GetMaterialObjects();

        /// <summary>
        /// Добавить объект состава изделия по связи Комплектующие на подключение
        /// </summary>
        /// <param name="materialObject"></param>
        void AddPart(IComplexHierarchyLink materialObject);

        /// <summary>
        /// Удалить объект состава изделия по связи Комплектующие на подключение
        /// </summary>
        /// <param name="link">Подключение в ЭСИ</param>
        void DeletePart(IComplexHierarchyLink link);

        /// <summary>
        /// Создать сборочный узел в изготавливаемой ДСЕ
        /// </summary>
        /// <param name="assemblyOperation">сборочная операция</param>
        void CreateAssemblyNodeInProducedMaterialObject(IAssemblyOperation assemblyOperation);

        /// <summary>
        /// Добавить объект к сборочному узлу
        /// </summary>
        /// <param name="part">объект для подключения к сборочному узлу</param>
        IComplexHierarchyLink AddPartToAssemblyNode(IMaterialObject part);

        /// <summary>
        /// Привязать сборочный узел к операции
        /// </summary>
        /// <param name="assemblyNode"></param>
        void LinkAssemblyNode(IAssemblyNode assemblyNode);

        /// <summary>
        /// Очистить все подключенные комплектующие
        /// </summary>
        void ClearParts();
    }

    /// <summary>
    /// Объект справочника ЭСИ типа Сборочный узел
    /// </summary>
    public interface IAssemblyNode : IMaterialObject
    {
        /// <summary>
        /// Добавить сборочную операцию
        /// </summary>
        /// <param name="assemblyOperation"></param>
        void AddAssemblyOperation(IAssemblyOperation assemblyOperation);
    }

    /// <summary>
    /// Объект справочника ЭСИ
    /// </summary>
    public interface IMaterialObject : IReferenceObject
    {
        /// <summary>
        /// Проверка, есть ли дочерние объекты
        /// </summary>
        bool HasChildren { get; }

        /// <summary>
        /// Подключенные объекты
        /// </summary>
        IEnumerable<IMaterialObject> GetMaterialObjects();

        /// <summary>
        /// Подключения дочерних объектов
        /// </summary>
        IEnumerable<IComplexHierarchyLink> GetComplexHierarchyLinksToChildren();

        /// <summary>
        /// Подключить материальный объект
        /// </summary>
        /// <param name="part">Подключаемый материальный объект</param>
        /// <returns>Возвращает подключение объекта к родителю</returns>
        IComplexHierarchyLink AddPart(IMaterialObject part);

        /// <summary>
        /// Проверить, подходит ли объект под выбранные условия конфигурации
        /// </summary>
        /// <returns></returns>
        bool CheckExistInConfiguration();

        /// <summary>
        /// Получить подключение
        /// </summary>
        /// <returns></returns>
        IComplexHierarchyLink Link { get; }

        /// <summary>
        /// Получить родителя
        /// </summary>
        /// <returns></returns>
        IMaterialObject GetParent();
    }

    public interface IComplexHierarchyLink : IComparable, IComparable<IComplexHierarchyLink>
    {
        Guid Guid { get; }

        /// <summary>
        /// Родитель
        /// </summary>
        IMaterialObject Parent { get; }

        /// <summary>
        /// Зависимый элемент
        /// </summary>
        IMaterialObject Child { get; }

        /// <summary>
        /// Подключение
        /// </summary>
        ComplexHierarchyLink Link { get; }
    }

    /// <summary>
    /// Сервис создания сборочных операций
    /// </summary>
    public interface IAssemblyOperationCreationService
    {
        /// <summary>
        /// Проверка, что всё соответствует
        /// </summary>
        /// <param name="technologyProcess"></param>
        /// <returns></returns>
        bool CheckConditions(ITechnologicalProcess technologyProcess);

        /// <summary>
        /// Создать сборочные операции в указанном ТП
        /// </summary>
        /// <param name="technologyProcess">Технологический процесс</param>
        void CreateAssemblyOperations(ITechnologicalProcess technologyProcess);
    }

    /// <summary>
    /// Сервис создания сборочных узлов
    /// </summary>
    public interface IAssemblyNodeCreationService
    {
        /// <summary>
        /// Создать сборочный узел на основе указанной сборочной операции из указанного тех. процесса
        /// </summary>
        /// <param name="assemblyOperation">Сборочная операция</param>
        void CreateAssemblyNode(IAssemblyOperation assemblyOperation);

        /// <summary>
        /// Добавить комплектующие
        /// </summary>
        /// <param name="assemblyOperation">Сборочная операция</param>
        /// <param name="parts">Объекты справочника ЭСИ</param>
        void AddParts(IAssemblyOperation assemblyOperation, IEnumerable<IMaterialObject> parts);

        /// <summary>
        /// Присоединить одиночный сборочный узел
        /// </summary>
        /// <param name="assemblyOperation">сборочная операция</param>
        /// <param name="assemblyNode">сборочный узел</param>
        void LinkAssemblyNode(IAssemblyOperation assemblyOperation, IAssemblyNode assemblyNode);

        /// <summary>
        /// Синхронизировать комплектующие
        /// </summary>
        /// <param name="assemblyOperation">сборочная операция</param>
        /// <param name="assemblyNode">сборочный узел</param>
        void SyncronizeParts(IAssemblyOperation assemblyOperation, IAssemblyNode assemblyNode);

        /// <summary>
        /// Создать сборочные операции
        /// </summary>
        /// <param name="technologicalProcess">Технологический процесс</param>
        void CreateAssemblyOperations(ITechnologicalProcess technologicalProcess);

        /// <summary>
        /// Создать сборочную операцию для изготавливаемого изделия
        /// </summary>
        /// <param name="technologicalProcess">Технологический процесс</param>
        void CreateAssemblyOperationFromMainProduct(ITechnologicalProcess technologicalProcess);

        /// <summary>
        /// Отключить комплектующие у указанного родителя
        /// </summary>
        /// <param name="materialObjects"></param>
        void RemoveParts(IEnumerable<IMaterialObject> materialObjects);
    }

    #region Adapters

    /// <summary>
    /// Адаптер над объектом технологического процесса
    /// </summary>
    public class TechnologyProcessAdapter : ITechnologicalProcess
    {
        private readonly ITechnologicalProcessParametersProvider _technologicalProcess;

        public TechnologyProcessAdapter(ITechnologicalProcessParametersProvider technologicalProcess)
        {
            _technologicalProcess = technologicalProcess ?? throw new ArgumentNullException(nameof(technologicalProcess));
        }

        /// <inheritdoc cref="ITechnologicalProcess.GetProducedProduct"/>
        public IMaterialObject GetProducedProduct()
        {
            if (!(_technologicalProcess is ITechnologicalProcessParametersProvider techProcess))
                return null;

            return !(techProcess.ProducedDSE.FirstOrDefault() is MaterialObject producedProduct) ? null : new MaterialObjectAdapter(producedProduct);
        }

        /// <inheritdoc cref="ITechnologicalProcess.CreateAssemblyOperation"/>
        public void CreateAssemblyOperation(IMaterialObject materialObject)
        {
            if (materialObject.GetCoreObject() is not MaterialObject coreObject)
                return;

            AssemblyTechnologicalOperation operation = CreateOperation(coreObject);

            IAssemblyOperation adapter = new AssemblyOperationAdapter(operation);

            ProductsApplicabilityReference productsApplicability =
                ((ReferenceObject)_technologicalProcess).Reference.Connection.References.ProductsApplicability;

            productsApplicability.CopyIntervals(coreObject.Guid, operation.Guid);
            productsApplicability.CopyConditions(coreObject.Guid, operation.Guid);

            foreach (IComplexHierarchyLink child in materialObject.GetComplexHierarchyLinksToChildren())
            {
                adapter.AddPart(child);
            }
        }

        public IEnumerable<IAssemblyOperation> GetOperations()
        {
            return Enumerable.Empty<IAssemblyOperation>();
        }

        public bool CheckProductExistInConfiguration()
        {
            IMaterialObject producedProduct = GetProducedProduct();
            
            return producedProduct?.CheckExistInConfiguration() ?? false;
        }

        private AssemblyTechnologicalOperation CreateOperation(MaterialObject materialObject)
        {
        	System.Diagnostics.Debugger.Launch();
        	System.Diagnostics.Debugger.Break();
        	
            var techProcessObject = (ReferenceObject)_technologicalProcess;

            TechnologicalProcessReference reference = techProcessObject.Reference as TechnologicalProcessReference;

            var assemblyOperation = reference.CreateReferenceObject(techProcessObject,
                reference.Classes.AssemblyOperation) as AssemblyTechnologicalOperation;

            var assemblyOperationAdapter = new AssemblyOperationAdapter(assemblyOperation);
            assemblyOperationAdapter.LinkAssemblyNode(new AssemblyNodeAdapter(materialObject));

            assemblyOperation.Name.Value = materialObject.Name;
            assemblyOperation.EndChanges();

            return assemblyOperation;
        }

        public ReferenceObject GetCoreObject() => _technologicalProcess as ReferenceObject;
    }

    /// <summary>
    /// Адаптер материального объекта
    /// </summary>
    public class MaterialObjectAdapter : IMaterialObject
    {
        private readonly MaterialObject _materialObject;

        public MaterialObjectAdapter(MaterialObject materialObject)
        {
            _materialObject = materialObject ?? throw new ArgumentNullException(nameof(materialObject));
        }

        public MaterialObjectAdapter(ReferenceObjectWithLink referenceObjectWithLink)
        {
            _materialObject = referenceObjectWithLink.ReferenceObject as MaterialObject ?? throw new ArgumentNullException(nameof(referenceObjectWithLink));
            Link = referenceObjectWithLink.Link is null
                ? null
                : new ComplexHierarchyLinkAdapter(referenceObjectWithLink.Link);
        }

        /// <inheritdoc cref="IMaterialObject.HasChildren"/>
        public bool HasChildren => _materialObject.HasChildren;

        /// <inheritdoc cref="IMaterialObject.GetMaterialObjects"/>
        public IEnumerable<IMaterialObject> GetMaterialObjects()
        {
            _materialObject.Children.Reload();
            return _materialObject.Children.Select(obj =>
                new MaterialObjectAdapter(obj as MaterialObject));
        }

        /// <inheritdoc cref="IMaterialObject.GetComplexHierarchyLinksToChildren"/>
        public IEnumerable<IComplexHierarchyLink> GetComplexHierarchyLinksToChildren()
        {
            _materialObject.Children.Reload();
            return _materialObject.Children.GetHierarchyLinks()
                .Where(link => link.ChildObject is MaterialObject)
                .Select(link => new ComplexHierarchyLinkAdapter(link));
        }

        /// <inheritdoc cref="IMaterialObject.AddPart"/>
        public IComplexHierarchyLink AddPart(IMaterialObject part) => CreateChildLink(part);

        public bool CheckExistInConfiguration()
        {
            var nomenclatureReference = new NomenclatureReference(_materialObject.Reference.Connection)
                {
                    ConfigurationSettings = _materialObject.Reference.Connection.ConfigurationSettings
                };

            return !(nomenclatureReference.Find(_materialObject.Id) is null);
        }

        public IComplexHierarchyLink Link { get; }

        public IComplexHierarchyLink GetParentLink(IMaterialObject parent)
        {
            var link = _materialObject.GetChildLink(parent.GetCoreObject());

            return link is null ? null : new ComplexHierarchyLinkAdapter(link);
        }

        public IMaterialObject GetParent()
        {
            var parent = _materialObject.Parent as MaterialObject;
            return _materialObject.Parent is null ? null : new MaterialObjectAdapter(parent);
        }

        public ReferenceObject GetCoreObject() => _materialObject;


        private IComplexHierarchyLink CreateChildLink(IMaterialObject part)
        {
            var coreObject = part.GetCoreObject();
            if (coreObject is null)
                return null;

            var hierarchyLink = _materialObject.CreateChildLink(coreObject) as NomenclatureHierarchyLink;
            hierarchyLink.EndChanges();

            return new ComplexHierarchyLinkAdapter(hierarchyLink);
        }

        public override string ToString() => _materialObject.ToString();
    }

    /// <summary>
    /// Адаптер над сборочным узлом
    /// </summary>
    public class AssemblyNodeAdapter : MaterialObjectAdapter, IAssemblyNode
    {
        private readonly MaterialObject _materialObject;

        public AssemblyNodeAdapter(MaterialObject materialObject) : base(materialObject)
        {
            _materialObject = materialObject ?? throw new ArgumentNullException(nameof(materialObject));
        }

        /// <inheritdoc cref="IAssemblyNode.AddAssemblyOperation"/>
        public void AddAssemblyOperation(IAssemblyOperation assemblyOperation)
        {
            var coreObject = assemblyOperation.GetCoreObject();

            if (coreObject is null)
                return;

            _materialObject.Modify(material =>
            {
                material.AddLinkedObject(GuidHelper.AssemblyNodeToAssemblyOperationRelation, coreObject);
            });
        }
    }

    /// <summary>
    /// Адаптер над сборочной операцией
    /// </summary>
    public class AssemblyOperationAdapter : IAssemblyOperation
    {
        private readonly AssemblyTechnologicalOperation _assemblyOperationObject;

        public AssemblyOperationAdapter(AssemblyTechnologicalOperation assemblyOperationObject)
        {
            _assemblyOperationObject = assemblyOperationObject ?? throw new ArgumentNullException(nameof(assemblyOperationObject));
        }

        /// <inheritdoc cref="IAssemblyOperation.TechnologicalProcess"/>
        public ITechnologicalProcessParametersProvider TechnologicalProcess
        {
            get
            {
                if (_assemblyOperationObject is ITechnologicalOperationParametersProvider provider)
                    return provider.TechnologicalProcess;

                return null;
            }
        }

        /// <inheritdoc cref="IAssemblyOperation.GetAssemblyNode"/>
        public IAssemblyNode GetAssemblyNode()
        {
            var assemblyNode = _assemblyOperationObject.GetAssemblyNode();
            return assemblyNode is null ? null : new AssemblyNodeAdapter(assemblyNode);
        }

        /// <inheritdoc cref="IAssemblyOperation.GetMaterialObjects"/>
        public IEnumerable<IComplexHierarchyLink> GetMaterialObjects()
        {
            OneToManyLinkToComplexHierarchy link =
                _assemblyOperationObject.Links.ToManyToComplexHierarchy.Find(GuidHelper.LinkToParts);

            link.Objects.Reload();

            return link.Objects
                .GetHierarchyLinks()
                .Select(l => new ComplexHierarchyLinkAdapter(l)).Cast<IComplexHierarchyLink>()
                .ToList();
        }

        /// <inheritdoc cref="IAssemblyOperation.AddPart"/>
        public void AddPart(IComplexHierarchyLink complexHierarchyLink)
        {
            _assemblyOperationObject.Modify(operation =>
            {
                OneToManyLinkToComplexHierarchy linkToComplexHierarchy =
                    _assemblyOperationObject.Links.ToManyToComplexHierarchy.Find(GuidHelper.LinkToParts);

                linkToComplexHierarchy.AddLinkedComplexLink(complexHierarchyLink.Link);
            });
        }

        /// <inheritdoc cref="IAssemblyOperation.DeletePart"/>
        public void DeletePart(IComplexHierarchyLink complexHierarchyLink)
        {
            _assemblyOperationObject.Modify(operation =>
            {
                OneToManyLinkToComplexHierarchy linkToComplexHierarchy =
                    _assemblyOperationObject.Links.ToManyToComplexHierarchy.Find(GuidHelper.LinkToParts);

                linkToComplexHierarchy.RemoveLinkedComplexLink(complexHierarchyLink.Link);
            });
        }

        /// <inheritdoc cref="IAssemblyOperation.CreateAssemblyNodeInProducedMaterialObject"/>
        public void CreateAssemblyNodeInProducedMaterialObject(IAssemblyOperation assemblyOperation)
        {
            var producedDSE = TechnologicalProcess.ProducedDSE.FirstOrDefault() as MaterialObject;

            if (producedDSE is null)
                throw new InvalidOperationException("К тех. процессу нет привязанного объекта по связи Изготавливаемые ДСЕ");

            var nomenclatureReference = new NomenclatureReference(_assemblyOperationObject.Reference.Connection);
            NomenclatureType assemblyNodeType = nomenclatureReference.Classes.Find(GuidHelper.AssemblyNodeTypeGuid);

            if (assemblyNodeType.LinkedClass is null)
                throw new InvalidOperationException(String.Format("{0}: {1}", "Не найден тип", assemblyNodeType));

            var documentReference = assemblyNodeType.LinkedClass.Classes.Owner.ReferenceInfo.CreateReference();
            if (documentReference is null)
                throw new InvalidOperationException("Не найден справочник Документы");

            ReferenceObject documentParent = null;
            if (assemblyNodeType.Attributes.DefaultNewObjectFolder != null)
                documentParent = assemblyNodeType.Attributes.DefaultNewObjectFolder.FindDefaultParent(documentReference);

            var linkedDocument = documentReference.CreateReferenceObject(documentParent, assemblyNodeType.LinkedClass);

            if (linkedDocument == null)
                throw new InvalidOperationException("Не создан документ!");

            var assemblyNodeObject = nomenclatureReference.CreateNomenclatureObject(linkedDocument) as MaterialObject;
            assemblyNodeObject.Name.Value = "Сборочный узел";

            var assemblyNodeAdapter = new AssemblyNodeAdapter(assemblyNodeObject);

            var producedDSEObjectAdapter = new MaterialObjectAdapter(producedDSE);

            var coreObject = assemblyOperation.GetCoreObject();
            if (coreObject is not null)
                assemblyNodeObject.AddLinkedObject(GuidHelper.AssemblyNodeToAssemblyOperationRelation, coreObject);

            ProductsApplicabilityReference productsApplicability =
                _assemblyOperationObject.Reference.Connection.References.ProductsApplicability;

            productsApplicability.CopyIntervals(_assemblyOperationObject.Guid, assemblyNodeObject.Guid);
            productsApplicability.CopyConditions(_assemblyOperationObject.Guid, assemblyNodeObject.Guid);

            assemblyNodeObject.SaveSet.EndChanges();

            producedDSEObjectAdapter.AddPart(assemblyNodeAdapter);
        }

        /// <inheritdoc cref="IAssemblyOperation.AddPartToAssemblyNode"/>
        public IComplexHierarchyLink AddPartToAssemblyNode(IMaterialObject part)
        {
            var assemblyNode = GetAssemblyNode();
            if (assemblyNode is null)
                throw new InvalidOperationException("Отсутствует сборочный узел!");

            if (part.GetCoreObject() == assemblyNode.GetCoreObject())
                throw new InvalidOperationException("Нельзя добавить сборочный узел еще и как комплектующее!");

            return assemblyNode?.AddPart(part);
        }

        /// <inheritdoc cref="IAssemblyOperation.LinkAssemblyNode"/>
        public void LinkAssemblyNode(IAssemblyNode assemblyNode)
        {
            var coreObject = assemblyNode.GetCoreObject();

            if (coreObject is null)
                return;

            _assemblyOperationObject.Modify(op =>
            {
                op.SetLinkedObject(GuidHelper.AssemblyNodeToAssemblyOperationRelation, coreObject);
            });
        }

        /// <inheritdoc cref="IAssemblyOperation.GetCoreObject"/>
        public void ClearParts()
        {
            OneToManyLinkToComplexHierarchy link =
                _assemblyOperationObject.Links.ToManyToComplexHierarchy.Find(GuidHelper.LinkToParts);

            _assemblyOperationObject.Modify(op =>
            {
                link.RemoveAll();
            });
        }

        /// <inheritdoc cref="IAssemblyOperation.GetCoreObject"/>
        public ReferenceObject GetCoreObject() => _assemblyOperationObject;
    }

    /// <summary>
    /// Адаптер над подключением
    /// </summary>
    public class ComplexHierarchyLinkAdapter : IComplexHierarchyLink
    {
        private readonly ComplexHierarchyLink _link;

        public ComplexHierarchyLinkAdapter(ComplexHierarchyLink link)
        {
            _link = link;
        }

        public Guid Guid => _link.Guid;

        /// <inheritdoc cref="IComplexHierarchyLink.Parent"/>
        public IMaterialObject Parent => new MaterialObjectAdapter(_link.ParentObject as MaterialObject);

        /// <inheritdoc cref="IComplexHierarchyLink.Child"/>
        public IMaterialObject Child => new MaterialObjectAdapter(_link.ChildObject as MaterialObject);

        /// <inheritdoc cref="IComplexHierarchyLink.Link"/>
        public ComplexHierarchyLink Link => _link;

        #region IComparable

        public int CompareTo(object obj)
        {
            return CompareTo((IComplexHierarchyLink) obj);
        }

        public int CompareTo(IComplexHierarchyLink other)
        {
            return Guid.CompareTo(other.Guid);
        }

        #endregion

    }

    #endregion
}
