using System;
using System.Linq;
using TFlex.DOCs.Model.Macros;
using TFlex.DOCs.Model.References;
using TFlex.DOCs.Model.References.Nomenclature;
using TFlex.Model.Technology.References.AssemblyOperationSet;
using TFlex.Model.Technology.References.TechnologyElements;

namespace TechnologyMacros
{
    public class TechnologicalProcessProductStructureMacroProvider : MacroProvider
    {
        public TechnologicalProcessProductStructureMacroProvider(MacroContext context)
            : base(context)
        {
        }

        public override void Run()
        {
        }

        public void ЗавершениеИзмененияСвязиОбъекта()
        {
            if (Context.ChangedLink == null)
                return;

            if (Context.ChangedLink.LinkGroup.Guid == AssemblyTechnologicalOperation.AssemblyOperationRelations.Assembly)
                СоздатьКомплектНаТекущуюСборочнуюОперацию();
        }

        public void СоздатьКомплектНаТекущуюСборочнуюОперацию()
        {
            // текущая сборочная операция
            var assemblyOperation = (ReferenceObject)CurrentObject as AssemblyTechnologicalOperation;
            if (assemblyOperation == null)
                return;

            assemblyOperation.ClearObjectList(AssemblyTechnologicalOperation.AssemblyOperationRelations.OperationSet);

            var addedObject = assemblyOperation.Assembly as NomenclatureReferenceObject;
            if (addedObject == null)
                return;

            var dse = assemblyOperation.FindObjectInProcessProductStructure(addedObject);
            if (dse == null)
                throw new ArgumentException("Подключенный объект не найден в технологическом составе техпроцесса");

            bool noAmount = assemblyOperation.Process
                .GetOperations(true)
                .OfType<AssemblyTechnologicalOperation>()
                .Any(op => op != assemblyOperation && op.Assembly == dse);

            foreach (var child in dse.Children.OfType<MaterialObject>())
            {
                var newComplete = assemblyOperation.CreateListObject(
                    AssemblyTechnologicalOperation.AssemblyOperationRelations.OperationSet,
                    new Guid("2ff6381a-c649-4462-ac38-848e2b6cb599")) as AssemblyComponent;

                newComplete.Name = child.Name;
                newComplete.Denotation = child is NomenclatureObject ? child.Denotation : String.Empty;

                var nomenclatureHierarchyLink = child.GetParentLink(dse) as NomenclatureHierarchyLink;

                newComplete.Amount = noAmount ? 0 : nomenclatureHierarchyLink.Amount.Value;
                newComplete.Rest = newComplete.Amount;
                newComplete.Unit = nomenclatureHierarchyLink.Unit.Value;
                newComplete.Position = nomenclatureHierarchyLink.Position.Value;
                newComplete.Product = child;

                newComplete.EndChanges();
            }
        }

        public void ОбновитьКомплектНаТекущуюСборочнуюОперацию()
        {
            UpdateByAssemblyNode();
        }

        private void UpdateByAssemblyNode()
        {
            // список "Комплект на операцию"
            var operationSet = Context.Reference as OperationSetReference;
            if (operationSet == null)
                throw new ArgumentException(String.Format("Справочник '{0}' не является комплектом на операцию", Context.Reference));

            // текущая сборочная операция
            var assemblyOperation = operationSet.LinkInfo.MasterObject as AssemblyTechnologicalOperation;
            if (assemblyOperation == null)
                throw new ArgumentException(String.Format("Операция '{0}' не является сборочной", operationSet.LinkInfo.MasterObject));

            // узел сборки
            var assemblyNode = assemblyOperation.Assembly as NomenclatureReferenceObject;
            if (assemblyNode == null)
                throw new ArgumentException("Не задан узел сборки");

            var productStructure = assemblyOperation.Process.ProductStructure;
            if (productStructure == null)
                throw new ArgumentException("У техпроцесса не задан технологический состав");

            productStructure.Reload();

            var assemblyDSE = assemblyOperation.FindObjectInProcessProductStructure(assemblyNode);
            if (assemblyDSE == null)
                throw new ArgumentException(String.Format("Подключенный узел сборки '{0}' не найден в технологическом составе '{1}'",
                    assemblyNode, productStructure.ToString()));

            if (!Question(String.Format("Обновить комплект на операцию по узлу сборки '{0}'?", assemblyNode)))
                return;

            assemblyDSE.Reload();

            assemblyOperation.BeginChanges();

            foreach (var component in operationSet.Objects.AsList.ToList())
            {
                bool delete = false;
                if (component.Product == null)
                {
                    delete = Question(String.Format("У компонента '{0}' отсутствует связанный номенклатурный объект. Удалить его из комплекта?",
                        component));
                }
                else
                {
                    assemblyDSE.Children.Reload();
                    var dse = assemblyDSE.Children.OfType<NomenclatureReferenceObject>().FirstOrDefault(child => child == component.Product);
                    if (dse == null)
                    {
                        delete = Question(String.Format("Компонент '{0}' не найден в составе '{1}'. Удалить его из комплекта?",
                            component.Product.ToString(), assemblyDSE.ToString()));
                    }
                }

                if (delete)
                    component.Delete();
            }

            bool noAmount = assemblyOperation.Process
                .GetOperations(true)
                .OfType<AssemblyTechnologicalOperation>()
                .Any(operation => operation != assemblyOperation && operation.Assembly == assemblyDSE);

            foreach (var child in assemblyDSE.Children.OfType<MaterialObject>())
            {
                var existComponent = operationSet.Objects
                    .OfType<AssemblyComponent>()
                    .FirstOrDefault(component => component.Product == child) as AssemblyComponent;

                if (existComponent == null)
                {
                    existComponent = assemblyOperation.CreateListObject(
                        AssemblyTechnologicalOperation.AssemblyOperationRelations.OperationSet,
                        new Guid("2ff6381a-c649-4462-ac38-848e2b6cb599")) as AssemblyComponent;

                    existComponent.Product = child;
                }
                else
                    existComponent.BeginChanges();

                existComponent.Name = child.Name;
                existComponent.Denotation = child is NomenclatureObject ? child.Denotation : String.Empty;

                var nomenclatureHierarchyLink = child.GetParentLink(assemblyDSE) as NomenclatureHierarchyLink;
                existComponent.Amount = noAmount ? 0 : nomenclatureHierarchyLink.Amount.Value;
                existComponent.Unit = nomenclatureHierarchyLink.Unit.Value;
                existComponent.Position = nomenclatureHierarchyLink.Position.Value;

                // для перерасчёта остатка при вызове EndChanges()
                existComponent.IsRestRecalculated = false;

                existComponent.EndChanges();
            }

            assemblyOperation.EndChanges();

            Context.RefreshReferenceWindow();
        }
    }
}

