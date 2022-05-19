using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using TFlex.DOCs.Model.Classes;
using TFlex.DOCs.Model.Macros;
using TFlex.DOCs.Model.Macros.ObjectModel;
using TFlex.DOCs.Model.References;
using TFlex.DOCs.Model.References.Links;
using TFlex.DOCs.Model.References.Nomenclature;
using TFlex.DOCs.Model.References.Users;
using TFlex.DOCs.Model.Search;
using TFlex.Model.Technology.References.TechnologyElements;
using TFlex.Technology;
using TFlex.Technology.References;

namespace Macros
{
    public class TechRoutesMacros : MacroProvider
    {
        /// <summary>
        /// Параметр наименование для объектов справочника Технологические процессы
        /// </summary>
        private static readonly Guid NameParameterGuid = new Guid("f97e40ea-3c79-4013-b1ea-383a2f09454d");

        private static class RouteGuids
        {
            /// <summary>
            /// Параметр Основной
            /// </summary>
            public static readonly Guid IsMainRouteParameter = new Guid("a3394032-b99e-48c4-8e6f-87db51950405");

            /// <summary>
            /// Параметр Маршрут
            /// </summary>
            public static readonly Guid ParameterGuid = new Guid("da5c50c3-4add-4aea-9e4a-ef2ba4cb3573");

            /// <summary>
            /// Изготавливаемое изделие
            /// </summary>
            public static readonly Guid ProducedProductRelation = new Guid("e1e8fa07-6598-444d-8f57-3cfd1a3f4360");

            /// <summary>
            /// Продолжительность
            /// </summary>
            public static readonly Guid DurationParameterGuid = new Guid("5f6fb9f2-18e5-4d46-a59c-ac41f484b365");
        }

        private static class ProductionUnitStepGuids
        {
            /// <summary>
            /// Guid класса Цехопереход
            /// </summary>
            public static readonly Guid ProductionUnitStepClass = new Guid("25fad9d1-be23-4f4b-9afc-581b6d96b992");

            /// <summary>
            /// Продолжительность
            /// </summary>
            public static readonly Guid DurationParameterGuid = new Guid("66536b01-dabb-418c-830c-1d396c1d1b24");

            /// <summary>
            /// Номер по порядку
            /// </summary>
            public static readonly Guid OrderNumberParameterGuid = new Guid("10276E9F-79F3-4046-9368-18484C83FEFA");

            /// <summary>
            /// Связь с цехоперехода с технологическим процессом
            /// </summary>
            public static readonly Guid TechProcessRelationGuid = new Guid("2c0aed62-4ad9-4152-8138-18e94c4ffbe6");

            /// <summary>
            /// Связь цехоперехода с производственным подразделением
            /// </summary>
            public static readonly Guid ProductionUnitRelation = new Guid("30888ac1-d215-478f-aaf2-915be9aa9066");
        }

        private static class UserAndGroupsGuids
        {
            /// <summary>
            /// [Параметры производственных подразделений].[Номер]
            /// </summary>
            public static readonly Guid ProductionUnitNumberGuid = new Guid("1ff481a8-2d7f-4f41-a441-76e83728e420");
        }

        public TechRoutesMacros(MacroContext context) : base(context)
        {
        }

        /// <summary>
        /// Проверка видимости кнопки "По аналогу"
        /// </summary>
        /// <returns></returns>
        public ButtonValidator GetCreateRouteByTemplateButtonValidator()
        {
            bool visible = Context.Reference.IsSlave && Context.Reference.LinkInfo.RootMasterReference is NomenclatureReference;
            return new ButtonValidator()
            {
                Enable = true,
                Visible = visible
            };
        }

        /// <summary>
        /// Создать набор цехозаходов из строки параметра Строка маршрута
        /// </summary>
        public void СоздатьЦехопереходы()
        {
            string строкаМаршрута = ТекущийОбъект["[Параметры маршрута].[Маршрут]"];
            if (String.IsNullOrEmpty(строкаМаршрута))
                return;

            СоздатьЦехопереходы(строкаМаршрута);
        }

        /// <summary>
        /// Обновить строку из набора цехозаходов
        /// </summary>
        public void СформироватьСтрокуМаршрута()
        {
            var routeStringBuilder = new StringBuilder();

            if (Context.ReferenceObject.Reference.IsSlave)
                Context.ReferenceObject.Children.Reload();
            
            foreach (var productionUnitStep in Context.ReferenceObject.Children.OrderBy(o =>
                o[ProductionUnitStepGuids.OrderNumberParameterGuid].GetInt32()))
            {
                routeStringBuilder.Append($"{productionUnitStep[NameParameterGuid]}-");
            }

            Context.ReferenceObject.Modify(
                o => o[RouteGuids.ParameterGuid].Value = routeStringBuilder.ToString().TrimEnd('-'));
        }

        /// <summary>
        /// Присвоить имя маршруту
        /// </summary>
        public void SetName()
        {
            var relationWithProducedProduct =
                Context.ReferenceObject.Links.ToMany.Find(RouteGuids.ProducedProductRelation);
            if (relationWithProducedProduct is null)
                Error("Не найдена связь Изготавливаемые ДСЕ");

            var producedProduct = relationWithProducedProduct.Objects.FirstOrDefault();

            if (producedProduct is null)
                Error("Не найдено изготавливаемое изделие! Невозможно присвоить имя маршруту");

            Context.ReferenceObject.Modify(o => o[NameParameterGuid].Value = producedProduct.ToString());
        }

        /// <summary>
        /// Посчитать общую продолжительность
        /// </summary>
        public void CalculateTotalDuration()
        {
            double total = 0;
            
            if (Context.ReferenceObject.Reference.IsSlave)
                Context.ReferenceObject.Children.Reload();
            
            foreach (var productionUnitStep in Context.ReferenceObject.Children)
            {
                double currentDurationValue = productionUnitStep
                    .ParameterValues[ProductionUnitStepGuids.DurationParameterGuid].GetDouble();
                if (currentDurationValue < 0)
                    Error("Продолжительность не может быть меньше нуля!");

                total += currentDurationValue;
            }

            Context.ReferenceObject.Modify(o => o[RouteGuids.DurationParameterGuid].Value = total);
        }

        /// <summary>
        /// Проверить, можно ли устанавливать текущему маршруту параметр Основной
        /// </summary>
        public void CheckIsMain()
        {
            if (!Context.ReferenceObject[RouteGuids.IsMainRouteParameter].GetBoolean())
                return;

            MaterialObject producedProduct = FindProducedProduct(Context.ReferenceObject);

            // найти все маршруты на изготавливаемое изделие

            string filterString =
                $"[Тип] = 'Маршрут' И [Изготавливаемые ДСЕ]->[Guid] = '{producedProduct.Guid}' И [Guid] != '{Context.ReferenceObject.Guid}'";

            var filter = Filter.Parse(filterString,
                Context.ReferenceObject.Reference.ParameterGroup);

            var technologicalProcessReference =
                new TechnologicalProcessReference(Context.Connection);

            var possibleRoutesList = technologicalProcessReference.Find(filter);
            ReferenceObject registeredMainRoute = null;
            foreach (var route in possibleRoutesList)
            {
                if (!route[RouteGuids.IsMainRouteParameter].GetBoolean())
                    continue;

                registeredMainRoute = route;
                break;
            }

            if (registeredMainRoute is null)
                return;

            var sb = new StringBuilder();
            sb.AppendLine("Вы хотите изменить основной маршрут для следующих ДСЕ?");
            sb.AppendLine(producedProduct.ToString());

            if (Question(sb.ToString()))
            {
                registeredMainRoute.Modify(o => o[RouteGuids.IsMainRouteParameter].Value = false);
                return;
            }

            // убрать у текущего
            Context.ReferenceObject.Modify(o => o[RouteGuids.IsMainRouteParameter].Value = false);
        }

        /// <summary>
        /// Создать тех. процессы на каждый цехозаход
        /// </summary>
        public void CreateTechnologicalProcesses()
        {
            var producedProduct = FindProducedProduct(Context.ReferenceObject);
            var technologicalProcessClass =
                Context.ReferenceObject.Reference.ParameterGroup.Classes.Find(Technology2012Classes
                    .StructuredTechnologicalProcessType);

            var registeredTechProcessDesignations = PrepareRegisteredProductionUnits(Context.ReferenceObject);

            if (Context.ReferenceObject.Reference.IsSlave)
                Context.ReferenceObject.Children.Reload();
            
            foreach (var productionUnitStep in Context.ReferenceObject.Children)
            {
                var productionUnit = productionUnitStep.GetObject(ProductionUnitStepGuids.ProductionUnitRelation);
                var relationWithTechProcess =
                    productionUnitStep.Links.ToOne.Find(ProductionUnitStepGuids.TechProcessRelationGuid);

                if (relationWithTechProcess.LinkedObject is StructuredTechnologicalProcess structuredTechProcess)
                {
                    if (!registeredTechProcessDesignations.Contains(structuredTechProcess.Denotation))
                        registeredTechProcessDesignations.Add(structuredTechProcess.Denotation);

                    continue;
                }
                
                var technologicalProcess = CreateTechnologicalProcess(productionUnit, productionUnitStep, producedProduct, 
                    registeredTechProcessDesignations, technologicalProcessClass);
                
                productionUnitStep.Modify(o=> o.SetLinkedObject(ProductionUnitStepGuids.TechProcessRelationGuid,
                    technologicalProcess));

                if (!registeredTechProcessDesignations.Contains(technologicalProcess.Denotation))
                    registeredTechProcessDesignations.Add(technologicalProcess.Denotation);
            }
        }

        /// <summary>
        /// Создать тех. процессы на каждый цехозаход
        /// </summary>
        public void CreateTechnologicalProcesseForSelectedProductionUnitStep()
        {
            var route = Context.ReferenceObject.Parent;
            var producedProduct = FindProducedProduct(route);
            var productionUnit = Context.ReferenceObject.GetObject(ProductionUnitStepGuids.ProductionUnitRelation);
            if (productionUnit is null)
                Error("Нет связанного с цехопереходом производственного подразделения");

            var technologicalProcessClass =
                Context.ReferenceObject.Reference.ParameterGroup.Classes.Find(Technology2012Classes
                    .StructuredTechnologicalProcessType);

            var registeredUnitsSet = PrepareRegisteredProductionUnits(route);

            var relationWithTechProcess =
                Context.ReferenceObject.Links.ToOne.Find(ProductionUnitStepGuids.TechProcessRelationGuid);

            if (!(relationWithTechProcess.LinkedObject is null))
                Error($"Уже есть связь с технологическим процессом {relationWithTechProcess.LinkedObject}");

            var technologicalProcess = CreateTechnologicalProcess(productionUnit, Context.ReferenceObject, producedProduct,
                registeredUnitsSet, technologicalProcessClass);

            Context.ReferenceObject.Modify(o=> o.SetLinkedObject(ProductionUnitStepGuids.TechProcessRelationGuid,
                technologicalProcess));
        }

        private HashSet<string> PrepareRegisteredProductionUnits(ReferenceObject routeObject)
        {
            var registeredWorkshops = new HashSet<string>();
            foreach (var productionUnitStep in routeObject.Children)
            {
                var productionUnit = productionUnitStep.GetObject(ProductionUnitStepGuids.ProductionUnitRelation);

                if (productionUnit is null)
                    Error($"Нет связанного с цехопереходом {productionUnitStep} производственного подразделения");

                // требуется отобрать только те, у которых уже есть тех.процессы
                var relationWithTechProcess = productionUnitStep.Links.ToOne.Find(ProductionUnitStepGuids.TechProcessRelationGuid);
                if (relationWithTechProcess.LinkedObject is null)
                    continue;

                string techProcessName = ((StructuredTechnologicalProcess)relationWithTechProcess.LinkedObject).Denotation;

                if (!registeredWorkshops.Contains(techProcessName))
                    registeredWorkshops.Add(techProcessName);
            }

            return registeredWorkshops;
        }


        private StructuredTechnologicalProcess CreateTechnologicalProcess(ReferenceObject productionUnit,
            ReferenceObject productionUnitStep,
            MaterialObject producedProduct, HashSet<string> registeredProductionUnits,
            ClassObject technologicalProcessClass)
        {
            string designation = producedProduct.Denotation;
            string productionUnitNumber = productionUnit[UserAndGroupsGuids.ProductionUnitNumberGuid].GetString();

            var technologicalProcess =
                productionUnitStep.Reference.CreateReferenceObject(technologicalProcessClass) as
                    StructuredTechnologicalProcess;

            technologicalProcess.Name.Value = producedProduct.Name;

            string techProcessDesignation = String.Empty;
            bool successToFindUniqueName = false;
            int value = 0;
            while (!successToFindUniqueName)
            {
                techProcessDesignation = value == 0
                    ? $"{designation}_{productionUnitNumber}"
                    : $"{designation}_{productionUnitNumber}_{value}";

                if (registeredProductionUnits.Contains(techProcessDesignation))
                {
                    value++;
                    continue;
                }

                successToFindUniqueName = true;
            }

            technologicalProcess.Denotation.Value = techProcessDesignation;
            technologicalProcess.AddLinkedObject(TechnologicalProcess.OneToManyGroups.ProducedDSE, producedProduct);
            technologicalProcess.ProductionUnit.SetLinkedObject(productionUnit);
            technologicalProcess.DseMass.Value = producedProduct.Mass;

            var list = producedProduct.GetAllLinkedFiles();
            list.AddRange(producedProduct.LinkedObject.GetAllLinkedFiles());

            technologicalProcess.SketchFileLink.SetLinkedObject(list.FirstOrDefault());

            technologicalProcess.EndChanges();

            return technologicalProcess;
        }

        /// <summary>
        /// Создать набор цехозаходов из выбранных производственных единиц из справочника Группы и пользователи
        /// </summary>
        public void CreateGroupOfWorkshopSteps()
        {
            var dialog = CreateUserDialog(true);

            if (!dialog.Show())
                return;

            var route = Context.ReferenceObject;
            var technologicalProcessReference = new TechnologicalProcessReference(Context.Connection);
            var productionUnitStepClassObject =
                technologicalProcessReference.Classes.Find(ProductionUnitStepGuids.ProductionUnitStepClass);

            if (Context.ReferenceObject.IsNew)
                Context.ReferenceObject.ApplyChanges();

            foreach (var productionUnit in dialog.SelectedObjects)
            {
                var productionUnitStep =
                    technologicalProcessReference.CreateReferenceObject(route, productionUnitStepClassObject);
                productionUnitStep[NameParameterGuid].Value = productionUnit
                    .ParameterValues[UserAndGroupsGuids.ProductionUnitNumberGuid].GetString();
                productionUnitStep.SetLinkedObject(ProductionUnitStepGuids.ProductionUnitRelation, productionUnit);

                productionUnitStep.EndChanges();
            }
        }

        /// <summary>
        /// Создать маршрут по выбранным аналогам
        /// </summary>
        public void CreateRoutesByTemplate()
        {
            if (!Context.Reference.IsSlave)
                Error("Требуется запуск из ЭСИ!");

            // изготавливаемое изделие в ЭСИ
            var product = Context.Reference.LinkInfo.MasterObject;

            if (product is null)
                Error("Не выбрано ни одно изделие!");

            var dialog = CreateTechProcessOnlyRoutesDialog(true);

            if (!dialog.Show())
                return;

            foreach (ReferenceObject route in dialog.SelectedObjects)
            {
                var saveSet = route.Reference.CopyReferenceObject(route, route.Class, true);

                foreach (ReferenceObject referenceObject in saveSet)
                {
                    if (!referenceObject.Class.IsInherit(new Guid("c02f6d42-1a50-48b2-ab35-fef5a165cde3")))
                    {
                        // Отключение тех. процесса
                        referenceObject.SetLinkedObject(ProductionUnitStepGuids.TechProcessRelationGuid, null);
                        continue;
                    }

                    var relation =
                        referenceObject.Links.ToMany[TechnologicalProcess.OneToManyGroups.ProducedDSE] as OneToManyLink;
                    relation.RemoveAll();

                    referenceObject[NameParameterGuid].Value = product.ToString();
                    referenceObject.AddLinkedObject(TechnologicalProcess.OneToManyGroups.ProducedDSE, product);
                }

                if (saveSet.Changing)
                    saveSet.EndChanges();
            }
        }

        private ISelectObjectDialog CreateTechProcessOnlyRoutesDialog(bool multiSelect)
        {
            var technologicalProcessReference = new TechnologicalProcessReference(Context.Connection);
            var filter = Filter.Parse("[Тип] = 'Маршрут'", technologicalProcessReference.ParameterGroup);
            ISelectObjectDialog dialog = Context.CreateSelectObjectDialog(technologicalProcessReference);
            dialog.Filter = filter;
            dialog.MultipleSelect = multiSelect;

            return dialog;
        }

        private ISelectObjectDialog CreateUserDialog(bool multiSelect)
        {
            Reference reference = new UserReference(Context.Connection);

            ISelectObjectDialog dialog = Context.CreateSelectObjectDialog(reference);
            dialog.MultipleSelect = multiSelect;

            return dialog;
        }

        private MaterialObject FindProducedProduct(ReferenceObject routeObject)
        {
            var relationWithProducedProduct =
                routeObject.Links.ToMany.Find(RouteGuids.ProducedProductRelation);

            if (relationWithProducedProduct is null)
                Error("Не найдена связь Изготавливаемые ДСЕ");

            var producedProduct = relationWithProducedProduct.Objects.FirstOrDefault() as MaterialObject;

            if (producedProduct is null)
                Error("Не найдено изготавливаемое изделие! Невозможно присвоить имя маршруту");

            return producedProduct;
        }

        private void СоздатьЦехопереходы(string строкаМаршрута)
        {
            string[] parts = строкаМаршрута.Split(new char[] { '-' }, StringSplitOptions.RemoveEmptyEntries);

            if (Context.ReferenceObject.IsNew)
                Context.ReferenceObject.ApplyChanges();

            foreach (string part in parts)
            {
                СоздатьЦехозаход(part, ТекущийОбъект);
            }

            ОбновитьЭлементыУправления("Цехопереходы");
        }

        private void СоздатьЦехозаход(string номер, Объект родитель)
        {
            var productionUnit = НайтиОбъект("Группы и пользователи", "Номер", номер);
            if (productionUnit is null)
                Error("Не найдено производственное подразделение");

            var workshopToCreate = СоздатьОбъект("Технологические процессы", "Цехопереход", родитель);
            workshopToCreate["Наименование"] = номер;
            workshopToCreate.СвязанныйОбъект["Производственное подразделение"] = productionUnit;

            workshopToCreate.Сохранить();
        }
    }
}
