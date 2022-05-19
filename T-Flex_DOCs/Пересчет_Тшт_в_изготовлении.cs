using System;
using System.Collections.Generic;
using System.Linq;
using TFlex.DOCs.Model.Macros;
using TFlex.DOCs.Model.Macros.ObjectModel;
using TFlex.Model.Technology.References.TechnologyElements;
using TFlex.DOCs.Model;
using TFlex.DOCs.Model.References;
using TFlex.DOCs.Model.References.Units;
using TFlex.DOCs.Model.References.Nomenclature;

public class Macro_AEM_TP : MacroProvider
{
    public Macro_AEM_TP(MacroContext context)
        : base(context)
    {
        //System.Diagnostics.Debugger.Launch();
        //System.Diagnostics.Debugger.Break();

    }

    public override void Run()
    {
        var цехопереходы = ВыбранныеОбъекты.SelectMany(тп => тп.ДочерниеОбъекты)
            .Where(цп => цп.Тип.ПорожденОт("Цехопереход")).ToArray();

        ДиалогОжидания.Показать("Подождите", false);
        ДиалогОжидания.СледующийШаг("Подождите, идет импорт: " + "Выгрузка маршрутов в FoxPro");
        ОбменДанными.ИмпортироватьОбъекты("Выгрузка маршрутов в FoxPro", цехопереходы);
        ДиалогОжидания.СледующийШаг("Подождите, идет импорт: " + "Выгрузка ведомости материалов в FoxPro 2");
        ОбменДанными.ИмпортироватьОбъекты("Выгрузка ведомости материалов в FoxPro 2", ВыбранныеОбъекты);
    }

    public void StepChanging()
    {
        var step = Context.ReferenceObject as StructuredTechnologicalStep;
        if (step == null)
            return;
        if (!step.Version.IsEmpty)
            return;

        var operation = step.Parent as StructuredTechnologicalOperation;
        if (operation == null)
            return;
        CalculateOperationTimes(operation);
    }

    public void CalculateOperationTimes(StructuredTechnologicalOperation operation)
    {	
        bool endChanges = false;
        if (!operation.Changing)
        {
            operation.BeginChanges();
            endChanges = true;
        }

        var steps = operation.GetSteps(true).Where(o => o.Version.IsEmpty).ToArray();

        if (operation.PieceTimeUnit == null && operation.PieceTimeUnitLink != null)
            operation.PieceTimeUnitLink.SetLinkedObject(operation.PieceTime.ParameterInfo.Unit);

        var operationPieceTimeUnit = operation.PieceTimeUnit ?? operation.PieceTime.ParameterInfo.Unit;
        if (operationPieceTimeUnit == null)
            operation.PieceTime.Value = steps.Aggregate(0.0, (total, next) => total + next.BaseTime + next.AdditionalTime);
        else
        {
        	try
        	{
                operation.PieceTime.Value = steps.Aggregate(0.0, (total, next) =>
                {
                    var baseTimeUnit = next.BaseTimeUnit ?? next.BaseTime.ParameterInfo.Unit;
                    total += baseTimeUnit == null ? next.BaseTime : operationPieceTimeUnit.Convert(next.BaseTime, baseTimeUnit);
                    var addTimeUnit = next.AdditionalTimeUnit ?? next.AdditionalTime.ParameterInfo.Unit;
                    total += addTimeUnit == null ? next.AdditionalTime : operationPieceTimeUnit.Convert(next.AdditionalTime, addTimeUnit);
                    return total;
                });
        	}
        	catch
        	{}
        }
        if (endChanges)
            operation.EndChanges();

        var route = operation.Parent;
        if (route == null)
            return;
        if (route.Class.IsInherit(new Guid("459ae48b-165b-44fd-8b3e-890298f2c3d7")))
        {
            CalculateRouteTimes(route);
            route = route.Parent;
        }
        ReferenceObject process = route;
        if (process.Class.IsInherit("Технологический процесс"))
            CalculateProcessTimes(process as StructuredTechnologicalProcess);
    }

    private void CalculateRouteTimes(ReferenceObject route)
    {
        bool endChanges = false;
        if (!route.IsChanged)
        {
            route.BeginChanges();
            endChanges = true;
        }

        var operations = route.Children.Where(o => o[new Guid("4ae9f4a6-49a1-4a11-8075-50e2a403d214")].IsEmpty).Cast<StructuredTechnologicalOperation>().ToArray();     // Вариант
        var pieceTimeUnitLink = route.GetObject(new Guid("ab7c9be3-f31e-40b5-9204-035a74a1bdf0")) as Unit;                              // ЕИ суммарного штучного времени

        if (pieceTimeUnitLink != null)
            route.SetLinkedObject(new Guid("ab7c9be3-f31e-40b5-9204-035a74a1bdf0"), route[new Guid("94ccb12d-e61c-46a7-ba4f-157d3b32f229")].ParameterInfo.Unit);    // Тшт

        var routePieceTimeUnit = pieceTimeUnitLink ?? route[new Guid("94ccb12d-e61c-46a7-ba4f-157d3b32f229")].ParameterInfo.Unit;
        var unitHour = НайтиОбъект("Единицы измерения", "Наименование", "Час");
        
        if (routePieceTimeUnit == null)
        	routePieceTimeUnit = (ReferenceObject)unitHour as Unit;
        
        route[new Guid("94ccb12d-e61c-46a7-ba4f-157d3b32f229")].Value = operations.Aggregate(0.0, (total, next) =>
            {
                var baseTimeUnit = next.PrepTimeUnit ?? next.PrepTime.ParameterInfo.Unit;
                total += baseTimeUnit == null ? next.PrepTime : routePieceTimeUnit.Convert(next.PrepTime, baseTimeUnit);
                var addTimeUnit = next.PieceTimeUnit ?? next.PieceTime.ParameterInfo.Unit;
                total += addTimeUnit == null ? next.PieceTime : routePieceTimeUnit.Convert(next.PieceTime, addTimeUnit);
                return total;
            });
        
        if (endChanges)
            route.EndChanges();
    }

    private void CalculateProcessTimes(StructuredTechnologicalProcess process)
    {
        process.Reload();
    	
        if (process.StandAloneChanging)
            process = process.EditableObject as StructuredTechnologicalProcess;
        
        if (!process.Changing)
            process.BeginChanges();

        var operations = process.GetOperations(true).Where(o => o.Version.IsEmpty).ToArray();
        if (process.SumPieceTimeUnit == null && process.SumPieceTimeUnitLink != null)
            process.SumPieceTimeUnitLink.SetLinkedObject(process.SumPieceTime.ParameterInfo.Unit);
        if (process.SumPrepTimeUnit == null && process.SumPrepTimeUnitLink != null)
            process.SumPrepTimeUnitLink.SetLinkedObject(process.SumPrepTime.ParameterInfo.Unit);

        var sumPieceUnit = process.SumPieceTimeUnit ?? process.SumPieceTime.ParameterInfo.Unit;
        var sumPrepUnit = process.SumPrepTimeUnit ?? process.SumPrepTime.ParameterInfo.Unit;

        if (process.StandAloneChanging)
            process = process.EditableObject as StructuredTechnologicalProcess;

        process.SumPieceTime.Value = operations.Aggregate(0.0, (total, next) =>
        {
            var pieceTimeUnit = next.PieceTimeUnit ?? next.PieceTime.ParameterInfo.Unit;
            total += (sumPieceUnit == null || pieceTimeUnit == null) ?
                next.PieceTime
                : sumPieceUnit.Convert(next.PieceTime, pieceTimeUnit);
            return total;
        });
        process.SumPrepTime.Value = operations.Aggregate(0.0, (total, next) =>
        {
            var prepTimeUnit = next.PrepTimeUnit ?? next.PrepTime.ParameterInfo.Unit;
            total += (sumPrepUnit == null || prepTimeUnit == null) ?
                next.PrepTime
                : sumPrepUnit.Convert(next.PrepTime, prepTimeUnit);
            return total;
        });
        
        if (process.Changing)
            process.EndChanges();
    }

    private void CalculateProcessTimes(ReferenceObject process)
    {
        //var original = process;
        //System.Windows.Forms.MessageBox.Show("");
        // Вариант - пусто; тип - порожден от "Технологическая операция" 

        try
        {
            var operations = process.Children.RecursiveLoad().Where(o => o[new Guid("4ae9f4a6-49a1-4a11-8075-50e2a403d214")].IsEmpty &&
                                                                    o.Class.IsInherit(new Guid("f53c9d73-18bb-4c59-a260-61fea65f6ed9"))).Cast<StructuredTechnologicalOperation>().ToArray();

            var routes = process.Children.RecursiveLoad().Where(o => o.Class.IsInherit(new Guid("459ae48b-165b-44fd-8b3e-890298f2c3d7"))).ToArray();       // Цехопереходы

            var pieceTimeUnitLink = process.GetObject(new Guid("ab7c9be3-f31e-40b5-9204-035a74a1bdf0")) as Unit;            // ЕИ суммарного штучного времени
            if (pieceTimeUnitLink != null)
                process.SetLinkedObject(new Guid("ab7c9be3-f31e-40b5-9204-035a74a1bdf0"), process[new Guid("9f192442-d2eb-43b5-818e-31bcec840574")].ParameterInfo.Unit);    // Суммарное Тшт

            var processPieceTimeUnit = pieceTimeUnitLink ?? process[new Guid("9f192442-d2eb-43b5-818e-31bcec840574")].ParameterInfo.Unit;

            var prepTimeUnitLink = process.GetObject(new Guid("44aed101-c8d9-4774-85a5-f2e4c2c8d36c")) as Unit;             // ЕИ суммарного подготовительно-заключительного времени
            if (prepTimeUnitLink != null)
                process.SetLinkedObject(new Guid("44aed101-c8d9-4774-85a5-f2e4c2c8d36c"), process[new Guid("2bc359aa-bff1-477e-8ec4-afe26ed765cb")].ParameterInfo.Unit);        // Суммарное Тпз

            var processPrepTimeUnit = prepTimeUnitLink ?? process[new Guid("2bc359aa-bff1-477e-8ec4-afe26ed765cb")].ParameterInfo.Unit;

            if (process.StandAloneChanging)
                process = process.EditableObject;

            bool changing = process.Changing;

            if (!changing)
                process.BeginChanges(false);

            process[new Guid("9f192442-d2eb-43b5-818e-31bcec840574")].Value = operations.Aggregate(0.0, (total, next) =>
            {
                var pieceTimeUnit = next.PieceTimeUnit ?? next.PieceTime.ParameterInfo.Unit;
                total += (processPieceTimeUnit == null || pieceTimeUnit == null) ?
                    next.PieceTime
                    : processPieceTimeUnit.Convert(next.PieceTime, pieceTimeUnit);
                return total;
            });
            process[new Guid("2bc359aa-bff1-477e-8ec4-afe26ed765cb")].Value = operations.Aggregate(0.0, (total, next) =>
            {
                var prepTimeUnit = next.PrepTimeUnit ?? next.PrepTime.ParameterInfo.Unit;
                total += (processPrepTimeUnit == null || prepTimeUnit == null) ?
                    next.PrepTime
                    : processPrepTimeUnit.Convert(next.PrepTime, prepTimeUnit);
                return total;
            });

            if (routes.Any())
            {
                process[new Guid("1d54cd34-4692-47d5-95fc-ed4e8f1293fc")].Value =
                    routes.Aggregate(0.0, (total, next) => total + (double)next[new Guid("08a8275c-901f-49d3-9ec2-e129471030d5")].Value); // Опытное время
                process[new Guid("2bc359aa-bff1-477e-8ec4-afe26ed765cb")].Value =
                    routes.Aggregate(0.0, (total, next) => total + (double)next[new Guid("29ecd770-f9a9-49ad-887d-e27ddb5f2f59")].Value); // Трудоемкость

                List<string> Executors = new List<string>();
                foreach (var route in routes)
                {
                    if (route.GetObject(new Guid("30888ac1-d215-478f-aaf2-915be9aa9066")) != null)
                        Executors.Add(route.GetObject(new Guid("30888ac1-d215-478f-aaf2-915be9aa9066"))[new Guid("1ff481a8-2d7f-4f41-a441-76e83728e420")].ToString());
                    else
                        Executors.Add(" ");
                }
                var lastRoute = routes.LastOrDefault();
                if (lastRoute != null)
                {
                	string potrebitel = lastRoute[new Guid("a3c05ec0-304c-42b8-a541-ea3a7feb9bc3")].Value.ToString();
                	if (!string.IsNullOrEmpty(potrebitel))
                		Executors.Add(potrebitel);
                }
                process[new Guid("cf0eb573-a7e1-4025-b05d-699e2ce69a1a")].Value = string.Join("-", Executors);
            }

            if (!changing)
                process.EndChanges();
        }
        catch // Почему в лог не записывается???
        { }
    }

    public void CalculateProcessTimes()
    {
        var process = Context.ReferenceObject;
        var unitHour = НайтиОбъект("Единицы измерения", "Наименование", "Час");
        //System.Windows.Forms.MessageBox.Show("");
        // Вариант - пусто; тип - порожден от "Технологическая операция"

        foreach (var route in process.Children)
        {
            foreach (var operation in route.Children)
                CalculateOperationTimes(operation as StructuredTechnologicalOperation);

            CalculateRouteTimes(route);
        }

        try
        {
        	process.Reload();
        	
        	if (!process.Changing)
        		process.BeginChanges();
        	
            var operations = process.Children.RecursiveLoad().Where(o => o[new Guid("4ae9f4a6-49a1-4a11-8075-50e2a403d214")].IsEmpty &&
                                                                    o.Class.IsInherit(new Guid("f53c9d73-18bb-4c59-a260-61fea65f6ed9"))).Cast<StructuredTechnologicalOperation>().ToArray();

            var routes = process.Children.RecursiveLoad().Where(o => o.Class.IsInherit(new Guid("459ae48b-165b-44fd-8b3e-890298f2c3d7"))).ToArray();       // Цехопереходы

            var pieceTimeUnitLink = process.GetObject(new Guid("ab7c9be3-f31e-40b5-9204-035a74a1bdf0")) as Unit;            // ЕИ суммарного штучного времени
            if (pieceTimeUnitLink == null)
            	pieceTimeUnitLink = (ReferenceObject)unitHour as Unit;
            
            if (pieceTimeUnitLink != null)
                process.SetLinkedObject(new Guid("ab7c9be3-f31e-40b5-9204-035a74a1bdf0"), process[new Guid("9f192442-d2eb-43b5-818e-31bcec840574")].ParameterInfo.Unit);    // Суммарное Тшт

            var processPieceTimeUnit = pieceTimeUnitLink ?? process[new Guid("9f192442-d2eb-43b5-818e-31bcec840574")].ParameterInfo.Unit;

            var prepTimeUnitLink = process.GetObject(new Guid("44aed101-c8d9-4774-85a5-f2e4c2c8d36c")) as Unit;             // ЕИ суммарного подготовительно-заключительного времени
            if (prepTimeUnitLink != null)
                prepTimeUnitLink = (ReferenceObject)unitHour as Unit;
            
            if (prepTimeUnitLink != null)
                process.SetLinkedObject(new Guid("44aed101-c8d9-4774-85a5-f2e4c2c8d36c"), process[new Guid("2bc359aa-bff1-477e-8ec4-afe26ed765cb")].ParameterInfo.Unit);        // Суммарное Тпз

            var processPrepTimeUnit = prepTimeUnitLink ?? process[new Guid("2bc359aa-bff1-477e-8ec4-afe26ed765cb")].ParameterInfo.Unit;

            if (process.StandAloneChanging)
                process = process.EditableObject;

            process[new Guid("9f192442-d2eb-43b5-818e-31bcec840574")].Value = operations.Aggregate(0.0, (total, next) =>
            {
                var pieceTimeUnit = next.PieceTimeUnit ?? next.PieceTime.ParameterInfo.Unit;
                total += (processPieceTimeUnit == null || pieceTimeUnit == null) ?
                    next.PieceTime
                    : processPieceTimeUnit.Convert(next.PieceTime, pieceTimeUnit);
                return total;
            });
            process[new Guid("2bc359aa-bff1-477e-8ec4-afe26ed765cb")].Value = operations.Aggregate(0.0, (total, next) =>
            {
                var prepTimeUnit = next.PrepTimeUnit ?? next.PrepTime.ParameterInfo.Unit;
                total += (processPrepTimeUnit == null || prepTimeUnit == null) ?
                    next.PrepTime
                    : processPrepTimeUnit.Convert(next.PrepTime, prepTimeUnit);
                return total;
            });

            if (routes.Any())
            {
                process[new Guid("1d54cd34-4692-47d5-95fc-ed4e8f1293fc")].Value =
                    routes.Aggregate(0.0, (total, next) => total + (double)next[new Guid("08a8275c-901f-49d3-9ec2-e129471030d5")].Value); // Опытное время
                process[new Guid("2bc359aa-bff1-477e-8ec4-afe26ed765cb")].Value =
                    routes.Aggregate(0.0, (total, next) => total + (double)next[new Guid("29ecd770-f9a9-49ad-887d-e27ddb5f2f59")].Value); // Трудоемкость

                List<string> Executors = new List<string>();
                foreach (var route in routes)
                {
                    if (route.GetObject(new Guid("30888ac1-d215-478f-aaf2-915be9aa9066")) != null)
                        Executors.Add(route.GetObject(new Guid("30888ac1-d215-478f-aaf2-915be9aa9066"))[new Guid("1ff481a8-2d7f-4f41-a441-76e83728e420")].ToString());
                    else
                        Executors.Add(" ");
                }
                process[new Guid("cf0eb573-a7e1-4025-b05d-699e2ce69a1a")].Value = string.Join("-", Executors);
            }
            
            if (process.Changing)
        		process.EndChanges();
        }
        catch
        { }
    }

    public void RouteChanging()
    {
        var route = Context.ReferenceObject;
        var process = route.Parent;

        CalculateProcessTimes(process);
    }

    public void WorkerChanging()
    {
        var worker = Context.ReferenceObject;
        if (worker == null)
            return;
        var operation = worker.MasterObject;
        var route = operation.Parent;
        var process = route.Parent;

        route.BeginChanges();
        double ExpirTime = 0;
        foreach (var oper in route.Children)
        {
            foreach (var w in oper.GetObjects(new Guid("c8ebd75f-c5e0-43bf-93a8-267a58c15384")))            // Исполнители
                ExpirTime += (double)w[new Guid("dd143e10-9cf9-4c07-8c0c-cd2221965ee2")].Value;                     // Трудоемкость	
        }
        route[new Guid("29ecd770-f9a9-49ad-887d-e27ddb5f2f59")].Value = ExpirTime;                          // Трудоемкость
        route.EndChanges();

        CalculateProcessTimes(process);
    }

    public void AddMaterialsToTechnologicalProcess(ReferenceObject technologicalProcess)
        => AddMaterialsToTechnologicalProcessCore(technologicalProcess, false);

    public void AddMaterialsToTp()
        => AddMaterialsToTechnologicalProcessCore(Context.ReferenceObject, true);

    private void AddMaterialsToTechnologicalProcessCore(ReferenceObject technologicalProcess, bool useDialogs)
    {
        var nomenclature = technologicalProcess.GetObject(new Guid("ba824125-2d20-4b50-b14f-0e5bfe9b4db4")) as NomenclatureObject; // Изготавливаемая ДСЕ

        if (nomenclature == null)
        {
            if (useDialogs)
                Message("Подключение материала", "К данному техпроцессу не подключен номенклатурный объект.");

            return;
        }

        var document = nomenclature.LinkedObject;
        ReferenceObject mainMaterial;
        document.TryGetObject(new Guid("2167290d-faa1-4c55-a5cb-32bcd205502a"), out mainMaterial); // Основной материал

        if (useDialogs)
            nomenclature.Children.Reload();

        var childNomMaterialHierarchyLinks = nomenclature.Children.GetHierarchyLinks()
            .Where(hl => hl.ChildObject.Class.IsInherit(new Guid("f7f45e16-ceba-4d26-a9af-f099a2e2fca6")));   // Дочерние объекты - материалы

        if (mainMaterial != null)
        {
            string name = mainMaterial[new Guid("23cfeee6-57f3-4a1e-9cf0-9040fed0e90c")].GetString();
            string denotation = mainMaterial[new Guid("d0441280-01ea-43b5-8726-d2d02e4d996f")].GetString();

            var existingMaterials = technologicalProcess.GetObjects(new Guid("8f505469-c1b4-4caf-a7c1-3bc7ea8c2bbe"));

            var material = existingMaterials.FirstOrDefault(m => m[new Guid("1339cab7-bfb6-41af-9363-a5950a2cfb0d")].GetString() == name
                && m[new Guid("71be2019-9498-4e19-acad-5dccc4ede0df")].GetString() == denotation);

            if (material == null)
            {
                material = technologicalProcess.CreateListObject(new Guid("8f505469-c1b4-4caf-a7c1-3bc7ea8c2bbe"), new Guid("35ee20c9-d771-4f90-8573-0505c2a7e398"));
                material[new Guid("1339cab7-bfb6-41af-9363-a5950a2cfb0d")].Value = name;            // Наименование
                material[new Guid("71be2019-9498-4e19-acad-5dccc4ede0df")].Value = denotation;            // Обозначение
            }
            else
            {
                material.BeginChanges(false);
            }

            //material.SetLinkedObject(new Guid("29826468-ded6-4fe6-b934-96812e91269f"), );
            material[new Guid("0f78f4c3-69ef-47da-baae-16084911a136")].Value = true;  // Основной

            ReferenceInfo referenceInfoUnit = Context.Connection.ReferenceCatalog.Find(new Guid("01c51d4c-e07d-4f31-9346-5697399a09fb"));
            Reference referenceUnit = referenceInfoUnit.CreateReference();

            ReferenceObject unitKG = referenceUnit.Find(new Guid("89b1bc71-7ff3-4d96-8164-c6a05f9aa1d7"));
            material.SetLinkedObject(new Guid("5f5d28ce-d269-4e3a-b7d1-97b96cc8e189"), unitKG);                                                                 // Единицы измерения

            material[new Guid("3388f183-f62c-43cb-942d-e8e0b0ca762d")].Value = nomenclature[new Guid("ee3cbb2b-3c92-4fef-85e9-d5bc3c9ce206")].Value;                  // Количество

            material.EndChanges();
        }

        foreach (var childNomMaterialHierarchyLink in childNomMaterialHierarchyLinks)
        {
            var childNomMaterial = childNomMaterialHierarchyLink.ChildObject;
            string name = childNomMaterial[new Guid("45e0d244-55f3-4091-869c-fcf0bb643765")].GetString();
            string denotation = childNomMaterial[new Guid("ae35e329-15b4-4281-ad2b-0e8659ad2bfb")].GetString();

            var existingMaterials = technologicalProcess.GetObjects(new Guid("8f505469-c1b4-4caf-a7c1-3bc7ea8c2bbe"));

            var material = existingMaterials.FirstOrDefault(m => m[new Guid("1339cab7-bfb6-41af-9363-a5950a2cfb0d")].GetString() == name
                && m[new Guid("71be2019-9498-4e19-acad-5dccc4ede0df")].GetString() == denotation);

            if (material == null)
            {
                material = technologicalProcess.CreateListObject(new Guid("8f505469-c1b4-4caf-a7c1-3bc7ea8c2bbe"), new Guid("35ee20c9-d771-4f90-8573-0505c2a7e398"));
                material[new Guid("1339cab7-bfb6-41af-9363-a5950a2cfb0d")].Value = name;            // Наименование
                material[new Guid("71be2019-9498-4e19-acad-5dccc4ede0df")].Value = denotation;            // Обозначение
            }
            else
            {
                material.BeginChanges(false);
            }

            material[new Guid("083d341a-6293-4f15-bf3d-6b64bf09d3b3")].Value = childNomMaterialHierarchyLink[new Guid("45df5f0b-dc53-494c-ae2e-54dfb8d64bd9")].Value; // Площадь покрытия
            material[new Guid("9a0eeb41-b16e-48a7-9a97-7eb280775c98")].Value = childNomMaterialHierarchyLink[new Guid("c024de94-d58c-4a15-b9a9-ccb2e57b246e")].Value; // Толщина покрытия
            material[new Guid("566fa92a-7068-4c0b-ba66-24fe5f47d997")].Value = childNomMaterialHierarchyLink[new Guid("a750a217-32f2-4438-bb32-60a547da50df")].Value; // Возратные отходы
            material[new Guid("22f28d56-7bc6-40f4-85f9-6ec82a217acc")].Value = childNomMaterialHierarchyLink[new Guid("3245305b-d5b6-419c-ad4d-17b317357272")].Value; // Потери
            material[new Guid("67057764-1640-464c-a561-f0434b9221fe")].Value = childNomMaterialHierarchyLink[new Guid("67732ea3-d9a3-448f-bdf9-6b2dee0c2403")].Value; // Чистый вес

            material.SetLinkedObject(new Guid("29826468-ded6-4fe6-b934-96812e91269f"), childNomMaterial); // Связь Номенклатура
            material[new Guid("0f78f4c3-69ef-47da-baae-16084911a136")].Value = false; // Основной

            string unit = childNomMaterialHierarchyLink[new Guid("444377c3-f73b-439a-b997-9b20c5707231")].GetString(); // Единица измерения
            var amount = childNomMaterialHierarchyLink[new Guid("3f5fc6c8-d1bf-4c3d-b7ff-f3e636603818")].Value; // Количество

            Объект ЕИ = String.IsNullOrWhiteSpace(unit) ? null : НайтиОбъект("Единицы измерения", "[Сокращённое наименование] = '" + unit + "'");
            if (ЕИ != null)
                material.SetLinkedObject(new Guid("5f5d28ce-d269-4e3a-b7d1-97b96cc8e189"), (ReferenceObject)ЕИ);

            material[new Guid("3388f183-f62c-43cb-942d-e8e0b0ca762d")].Value = amount; // Количество

            material.EndChanges();
        }
    }

}

