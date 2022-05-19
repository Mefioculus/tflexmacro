using System;
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using TFlex.DOCs.Model.Classes;
using TFlex.DOCs.Model.Macros;
using TFlex.DOCs.Model.References;
using TFlex.DOCs.Model.References.Nomenclature;
using TFlex.DOCs.Model.References.Units;
using TFlex.DOCs.Model.References.Users;
using TFlex.DOCs.Model.Search;
using TFlex.DOCs.Model.Structure;
using TFlex.Model.Technology.References.ParametersProvider;
using TFlex.Model.Technology.References.TechnologicalEquipment;

public class OperationalSchedulingHelper : MacroProvider
{
    private int _colorStep;
    private Color _color;
    private volatile int _processedPositionsCount;
    private volatile int _totalPositionsCount;
    private static bool _isAlreadyRunning = false;

    public OperationalSchedulingHelper(MacroContext context)
        : base(context)
    {
    }

    private void AddAssembliesToOrder(ReferenceObject order, List<ReferenceObject> selectedNomenclatureObjects)
    {
        selectedNomenclatureObjects.ForEach(referenceObject => AddAssemblyToOrder(order, null, referenceObject as NomenclatureReferenceObject, 1));
        WaitingDialog.NextStep("Сохранение...");
    }

    private void AddAssemblyToOrder(ReferenceObject order, ReferenceObject parent, NomenclatureReferenceObject nomenclatureObject,
        int amount)
    {
        if (nomenclatureObject == null || !nomenclatureObject.Class.IsMaterialObject)
            return;

        var positionsReference = Context.Connection.ReferenceCatalog.Find(OperationalSchedulingGuids.PositionsReference).CreateReference();
        ClassObject positionClass = positionsReference.Classes.Find(OperationalSchedulingGuids.Position);
        ReferenceObject position = positionsReference.CreateReferenceObject(parent, positionClass);
        if (position == null)
            return;

        position.SetLinkedObject(OperationalSchedulingGuids.NomenclatureLink, nomenclatureObject);
        string positionName = CreatePositionName(nomenclatureObject);
        position[OperationalSchedulingGuids.PositionName].Value = positionName;
        position[OperationalSchedulingGuids.PositionAmount].Value = amount;
        position[OperationalSchedulingGuids.PositionColor].Value = _color.ToArgb();
        position.EndChanges();
        order.AddLinkedObject(OperationalSchedulingGuids.PositionsLink, position);
        SetNextColor();

        if (!WaitingDialog.NextStep(String.Format("Добавление {0}", positionName)))
            return;

        foreach (ComplexHierarchyLink link in nomenclatureObject.Children.GetHierarchyLinks())
        {
            if (!WaitingDialog.NextStep())
                return;

            NomenclatureHierarchyLink nomenclatureLink = link as NomenclatureHierarchyLink;
            AddAssemblyToOrder(order, position,
                link.ChildObject as NomenclatureReferenceObject,
                ((int)nomenclatureLink.Amount.Value) * amount);
        }
    }

    private void AddOrderToPlan(ReferenceObject plan, ReferenceObject order)
    {
        if (plan == null || order == null)
            return;

        var positionsReference = Context.Connection.ReferenceCatalog.Find(OperationalSchedulingGuids.PositionsReference).CreateReference();
        positionsReference.LoadSettings.AddRelation(OperationalSchedulingGuids.BatchPositions);
        positionsReference.LoadSettings.AddRelation(OperationalSchedulingGuids.NomenclatureLink);

        var filter = new Filter(positionsReference.ParameterGroup);
        ReferenceObjectCollection orderPositionsCollection =
            order.Links.ToMany[OperationalSchedulingGuids.PositionsLink].Objects;

        filter.Terms.AddTerm(order.Reference.ParameterGroup[SystemParameterType.ObjectId],
            ComparisonOperator.Equal, order.SystemFields.Id, orderPositionsCollection.Reference.LinkInfo.LinkGroup);

        var parentTerm = new ReferenceObjectTerm(positionsReference.ParameterGroup);
        parentTerm.Path.AddParentObject();
        parentTerm.Operator = ComparisonOperator.Equal;
        parentTerm.Value = null;
        filter.Terms.Add(parentTerm);

        var addedBatches = new List<int>();
        var positions = positionsReference.Find(filter);
        if (positions.Count == 0)
            return;

        var allPositions = new List<ReferenceObject>(100);
        allPositions.AddRange(positions);

        _processedPositionsCount = 0;
        _totalPositionsCount = positions.Count;

        if (!ProcessPositionsNextStep("Загрузка позиций заказа"))
            return;

        positions.AsParallel().ForAll(position =>
        {
            if (WaitingDialog.NextStep())
            {
                List<ReferenceObject> positionTree = position.Children.RecursiveLoad();
                lock (((ICollection)allPositions).SyncRoot)
                    allPositions.AddRange(positionTree);

                ProcessPositionsNextStep("Загрузка позиций заказа");
            }
        });

        _processedPositionsCount = 0;
        _totalPositionsCount = allPositions.Count;

        if (!ProcessPositionsNextStep("Загрузка данных позиций"))
            return;

        allPositions.AsParallel().ForAll(position =>
        {
            ITechnologicalProcessParametersProvider technologicalProcess = GetTechnologicalProcess(position);
            if (technologicalProcess != null && WaitingDialog.NextStep())
                technologicalProcess.LoadShedulingParameters();

            ProcessPositionsNextStep("Загрузка данных позиций");
        });

        positions.AsParallel().ForAll(position => AddPositionToPlan(plan, position, addedBatches));
    }

    private void AddOrdersToPlan(ReferenceObject plan, List<ReferenceObject> orders)
    {
        plan.BeginChanges();
        orders.ForEach(order => AddOrderToPlan(plan, order));
        if (WaitingDialog.NextStep("Сохранение..."))
            plan.EndChanges();
        else
            plan.CancelChanges();
    }

    private ReferenceObject AddPositionToPlan(ReferenceObject plan, ReferenceObject position, List<int> addedBatches)
    {
        var positionBatch = position.GetObject(OperationalSchedulingGuids.BatchPositions);

        if (positionBatch != null)
            if (addedBatches.Contains(positionBatch.SystemFields.Id))
                return null;
            else
                addedBatches.Add(positionBatch.SystemFields.Id);

        var prevTasks = new List<ReferenceObject>();
        foreach (var child in position.Children)
        {
            if (!WaitingDialog.NextStep())
                return null;

            var task = AddPositionToPlan(plan, child, addedBatches);
            if (task != null)
                prevTasks.Add(task);
        }

        if (position[OperationalSchedulingGuids.State].GetInt32() != 0)
            return null;

        var positionTechnologicalProcess = GetTechnologicalProcess(position);

        return CreateTasks(plan, position, positionTechnologicalProcess, prevTasks);
    }

    private static TimeSpan CalcTaskDuration(ITechnologicalOperationParametersProvider operation, ReferenceObject task)
    {
        var isRecalcDuration = task[OperationalSchedulingGuids.TaskRecalcDuration].GetBoolean();

        if (isRecalcDuration && operation != null)
        {
            int taskAmount = task[OperationalSchedulingGuids.TaskAmount].GetInt32();
            double taskSameTimeCoff = task[OperationalSchedulingGuids.TaskSameTimeCoff].GetDouble();
            int amount = (int)(taskAmount / taskSameTimeCoff);
            if (amount * taskSameTimeCoff < taskAmount)
                amount++;

            return GetTimeSpanInBaseUnits(operation.PrepTime, operation.PrepTimeUnit) +
                   GetTimeSpanInBaseUnits(operation.PieceTime * amount, operation.PieceTimeUnit);
        }

        double durationParameter = task[OperationalSchedulingGuids.TaskDuration].GetDouble();
        return new TimeSpan(0, 0, (int)(durationParameter * 3600));
    }

    private static string CreatePositionName(NomenclatureReferenceObject referenceObject)
    {
        NomenclatureObject nomenclatureObject = referenceObject as NomenclatureObject;
        return nomenclatureObject != null ? nomenclatureObject.Denotation + " " + nomenclatureObject.Name : referenceObject.Name;
    }

    private ReferenceObject CreateTasks(ReferenceObject plan, ReferenceObject position,
        ITechnologicalProcessParametersProvider technologicalProcess, List<ReferenceObject> prevTasks)
    {
        if (technologicalProcess == null)
            return null;

        DateTime lastTime = DateTime.Now;
        ReferenceObject prevTask = null;
        foreach (var operation in technologicalProcess.Operations)
        {
            if (operation == null)
                continue;

            string positionName = position[OperationalSchedulingGuids.PositionName].GetString();

            string name = String.Format("{0} {1} {2} {3}", positionName, operation.Code, operation.Name, operation.Order);
            if (!WaitingDialog.NextStep(String.Format("Добавление {0}", name)))
                return null;

            var tasksReference = Context.Connection.ReferenceCatalog
                .Find(OperationalSchedulingGuids.TasksReference).CreateReference();

            var taskType = tasksReference.Classes.Find(OperationalSchedulingGuids.Task);

            ReferenceObject task = tasksReference.CreateReferenceObject(taskType);
            task[OperationalSchedulingGuids.TaskName].Value = name;
            task[OperationalSchedulingGuids.TaskColor].Value =
                position[OperationalSchedulingGuids.PositionColor].GetInt32();

            task.SetLinkedObject(OperationalSchedulingGuids.TaskPosition, position);
            task.SetLinkedObject(OperationalSchedulingGuids.TaskTechnologicalProcess, technologicalProcess as ReferenceObject);
            task[OperationalSchedulingGuids.TaskTPOperationGuid].Value =
                operation.AsReferenceObject().SystemFields.Guid;

            task[OperationalSchedulingGuids.TaskAlternativeEquipment].Value = operation.AlternativeEquipmentGroup;
            TimeSpan duration = CalcTaskDuration(operation, task);
            if (task[OperationalSchedulingGuids.TaskRecalcDuration].GetBoolean())
                task[OperationalSchedulingGuids.TaskDuration].Value = duration.TotalSeconds / 3600;

            task[OperationalSchedulingGuids.TaskStartTime].Value = lastTime;
            lastTime += duration;
            task[OperationalSchedulingGuids.TaskEndTime].Value = lastTime;
            lastTime += new TimeSpan(1, 30, 0);
            if (prevTask == null)
            {
                foreach (var prev in prevTasks)
                {
                    ComplexHierarchyLink link = task.CreateChildLink(prev);
                    if (link != null)
                        link.EndChanges();
                }
            }
            else
            {
                ComplexHierarchyLink link = task.CreateChildLink(prevTask);
                if (link != null)
                    link.EndChanges();
            }

            foreach (var equipmentObject in operation.Equipments)
            {
                if (equipmentObject.Equipment == null)
                    continue;

                EquipmentObject equipment = equipmentObject.Equipment as EquipmentObject;
                if (equipment == null)
                {
                    EquipmentComplete equipmentSet = equipmentObject.Equipment as EquipmentComplete;
                    if (equipmentSet == null)
                        continue;

                    var equipmentsInSet = new List<ReferenceObject>();
                    if (equipmentSet.TryGetObjects(EquipmentComplete.EquipmentCompleteRelations.Equipment, out equipmentsInSet))
                    {
                        foreach (ReferenceObject ro in equipmentsInSet)
                        {
                            equipment = ro as EquipmentObject;
                            if (equipment != null)
                                break;
                        }
                    }

                    if (equipment == null)
                        continue;
                }

                task.SetLinkedObject(OperationalSchedulingGuids.TaskEquipment, equipment);
                task.SetLinkedObject(OperationalSchedulingGuids.TaskProductionUnit, equipment.ProductionUnit);
                break;
            }

            var taskEquipment = task.GetObject(OperationalSchedulingGuids.TaskEquipment) as EquipmentObject;
            if (taskEquipment == null)
            {
                foreach (ReferenceObject unitObject in operation.ProductionUnits)
                {
                    ProductionUnit productUnit = unitObject as ProductionUnit;
                    if (productUnit == null)
                        continue;

                    task.SetLinkedObject(OperationalSchedulingGuids.TaskProductionUnit, productUnit);
                    break;
                }
            }

            task.EndChanges();
            plan.AddLinkedObject(OperationalSchedulingGuids.TaskPlan, task);
            prevTask = task;
        }

        return prevTask;
    }

    private static ITechnologicalProcessParametersProvider GetTechnologicalProcess(ReferenceObject position)
    {
        var nomenclature = position.GetObject(OperationalSchedulingGuids.NomenclatureLink) as MaterialObject;

        if (nomenclature == null || nomenclature.LinkedTPReference == null)
            return null;

        return nomenclature.LinkedTPReference.Objects.FirstOrDefault() as ITechnologicalProcessParametersProvider;
    }

    private static TimeSpan GetTimeSpanInBaseUnits(double sourceValue, Unit sourceUnit)
    {
        int seconds = 0;
        if (sourceUnit == null)
            seconds = (int)sourceValue * 3600;
        else
            seconds = (int)sourceUnit.Class.GetBaseUnit().Convert(sourceValue, sourceUnit);

        return new TimeSpan(0, 0, seconds);
    }

    private static void UpdateChild(ReferenceObject childPosition, ReferenceObject parent)
    {
        if (parent == null)
            return;

        var parentStateParameter = parent[OperationalSchedulingGuids.State];
        if (parentStateParameter.IsModified)
            childPosition[OperationalSchedulingGuids.State].Value = parentStateParameter.GetInt32();

        var childPositionNomenclature =
            childPosition.GetObject(OperationalSchedulingGuids.NomenclatureLink) as MaterialObject;

        var parentNomenclature =
            parent.GetObject(OperationalSchedulingGuids.NomenclatureLink) as MaterialObject;

        if (childPositionNomenclature == null || parentNomenclature == null)
            return;

        NomenclatureHierarchyLink nomenclatureLink =
            parentNomenclature.Children.GetHierarchyLinks()
                .FirstOrDefault(x => x.ChildObject == childPositionNomenclature) as NomenclatureHierarchyLink;

        if (nomenclatureLink == null)
            return;

        var nomenclatureAmount = nomenclatureLink.Amount.Value;
        childPosition[OperationalSchedulingGuids.Amount].Value =
            (int)(nomenclatureAmount * parent[OperationalSchedulingGuids.Amount].GetInt32());
    }

    private static bool NeedUpdate(ReferenceObject childPosition)
    {
        var parent = childPosition.Parent;
        if (childPosition[OperationalSchedulingGuids.AutoRecalc].GetBoolean() && parent[OperationalSchedulingGuids.Amount].IsModified)
            return true;

        if (parent[OperationalSchedulingGuids.State].IsModified && parent[OperationalSchedulingGuids.State].GetInt32() != childPosition[OperationalSchedulingGuids.State].GetInt32())
            return true;

        return false;
    }

    private bool ProcessPositionsNextStep(string text)
    {
        _processedPositionsCount++;
        return WaitingDialog.NextStep(text + String.Format(" {0} / {1}", _processedPositionsCount, _totalPositionsCount));
    }

    private void SetNextColor()
    {
        _colorStep %= 7;
        int R = _color.R;
        int G = _color.G;
        int B = _color.B;
        switch (_colorStep)
        {
            case 0:
                R = (R + 95) % 255;
                break;
            case 1:
                G = (G + 95) % 255;
                break;
            case 2:
                R = (R + 160) % 255;
                break;
            case 3:
                B = (B + 95) % 255;
                break;
            case 4:
                G = (G + 160) % 255;
                break;
            case 5:
                R = (R + 95) % 255;
                break;
            case 6:
                G = (G + 95) % 255;
                break;
        }

        _color = Color.FromArgb(_color.A, R, G, B);
        _colorStep++;
    }

    private static void UpdateChildren(ReferenceObject position)
    {
        var positionChildren =
            position.Reference.RecursiveLoad(new[] { position }, RelationLoadSettings.RecursiveLoadDirection.Children);

        if (position.SaveSet == null)
            position.CreateSaveSet();

        foreach (var child in positionChildren)
        {
            if (!NeedUpdate(child))
                continue;

            child.BeginChanges();
            position.SaveSet.Add(child);

            UpdateChild(child, child.Parent);
        }
    }

    public void ДобавитьЗаказ()
    {
        var plans = Context.GetSelectedObjects();
        if (plans.Length == 0)
            return;

        var plan = plans.First();
        var dialog = CreateSelectObjectsDialog(OperationalSchedulingGuids.OrderReference.ToString());
        dialog.Caption = "Выбор объектов";
        dialog.MultipleSelect = true;
        if (dialog.Show() != true)
            return;

        WaitingDialog.Show("Добавление объектов", true);
        AddOrdersToPlan(plan, dialog.SelectedObjects.Select(x => (ReferenceObject)x).ToList());
        WaitingDialog.Hide();

        ОбновитьОкноСправочника();
    }

    public void ДобавитьСборку()
    {
        var orders = Context.GetSelectedObjects();
        if (orders.Length == 0)
            return;

        var order = orders.First();
        _color = Color.Green;
        var dialog = CreateSelectObjectsDialog(OperationalSchedulingGuids.NomenclatureReference.ToString());
        dialog.Caption = "Выбор объектов";
        dialog.MultipleSelect = true;
        if (dialog.Show() != true)
            return;

        if (!order.CanEdit)
            return;

        order.BeginChanges();
        try
        {
            WaitingDialog.Show("Добавление объектов", true);
            var nomenclatureObjects =
                dialog.SelectedObjects.Select(x => (ReferenceObject)x).ToList();

            AddAssembliesToOrder(order, nomenclatureObjects);
            order.EndChanges();
            WaitingDialog.Hide();
        }
        catch
        {
            if (order.Changing)
                order.CancelChanges();

            throw;
        }

        ОбновитьОкноСправочника();
    }

    public void РасчетКоличестваПозиций()
    {
        if (_isAlreadyRunning)
            return;

        _isAlreadyRunning = true;
        var position = Context.ReferenceObject;
        var isAutoRecalc = position[OperationalSchedulingGuids.AutoRecalc].GetBoolean();
        var isAutoRecalcModified = position[OperationalSchedulingGuids.AutoRecalc].IsModified;
        if (isAutoRecalcModified && isAutoRecalc)
            UpdateChild(position, position.Parent);

        UpdateChildren(position);
        _isAlreadyRunning = false;
    }

    private static class OperationalSchedulingGuids
    {
        //References Guids
        public static readonly Guid OrderReference = new Guid("8bc6f591-576a-40ec-a3c6-b9bb70b15d4d");
        public static readonly Guid NomenclatureReference = new Guid("853d0f07-9632-42dd-bc7a-d91eae4b8e83");
        public static readonly Guid TasksReference = new Guid("fae849c0-9e9c-4a1f-95bd-23da8a254d16");

        //Type Guids
        public static readonly Guid Task = new Guid("43cce8cc-01c2-4924-b65e-237b14ace356");

        //Batch Reference Guids
        public static readonly Guid BatchPositions = new Guid("199cc8bc-3404-479f-b5d9-a255f0311133");

        //Positions Reference Guids
        public static readonly Guid PositionsLink = new Guid("9f5a9b4a-0354-46c2-9460-2af58782308d");
        public static readonly Guid PositionsReference = new Guid("a8bdff2f-d493-456d-b48e-489fa473d631");
        public static readonly Guid Position = new Guid("57192522-23c9-4e81-a79b-193f1469cd17");
        public static readonly Guid PositionName = new Guid("b01f9066-9ed5-4eb4-a82c-5073cd912958");
        public static readonly Guid PositionAmount = new Guid("ff4884ab-e1cf-48bb-b2ab-66e8c02f96fe");
        public static readonly Guid PositionColor = new Guid("3216876c-6e77-44a6-bd06-49a4d3455c32");
        public static readonly Guid NomenclatureLink = new Guid("c5223b64-4716-4598-95d7-3f713df40ad9");
        public static readonly Guid AutoRecalc = new Guid("f28a0508-c365-4635-a824-68359bdb9611");
        public static readonly Guid Amount = new Guid("ff4884ab-e1cf-48bb-b2ab-66e8c02f96fe");

        //public static readonly Guid IncludeToPlanning = new Guid("39cfdcdb-b6e3-40ce-8736-bd9b18328da4");
        public static readonly Guid State = new Guid("4518a94d-416f-4fba-a6ec-04f2d218af15");

        //Tasks Reference Guids
        public static readonly Guid TaskName = new Guid("a1df8c57-57f0-4016-8e93-98f7f302cddc");
        public static readonly Guid TaskColor = new Guid("c3a83272-d6d8-4b0a-aaa2-7368edee3243");
        public static readonly Guid TaskPosition = new Guid("3c74da62-5894-443b-b4a2-9495c9a5da2d");
        public static readonly Guid TaskTechnologicalProcess = new Guid("ff33be56-f5fa-4153-a6ad-5aa34664864c");
        public static readonly Guid TaskTPOperationGuid = new Guid("52da1017-8230-4441-8835-5066f7de762d");
        public static readonly Guid TaskAlternativeEquipment = new Guid("b8293cfb-f5f3-4d63-b270-f72ced208eed");
        public static readonly Guid TaskRecalcDuration = new Guid("51462700-660a-45fd-9834-a50c49d6c348");
        public static readonly Guid TaskAmount = new Guid("fd0d532c-fc59-412a-8273-08c39aa2801f");
        public static readonly Guid TaskSameTimeCoff = new Guid("d1d34c05-1aee-4126-a914-037d02358343");
        public static readonly Guid TaskDuration = new Guid("b6eaf8a0-7d00-489c-8c69-1f3eebd20c84");
        public static readonly Guid TaskStartTime = new Guid("611413a7-863b-44e7-a105-1aa1650b0555");
        public static readonly Guid TaskEndTime = new Guid("f1d079e6-596c-4ee4-835f-ab85e9c10a2d");
        public static readonly Guid TaskEquipment = new Guid("8626ec43-c467-41ca-91bf-55f0d01150a6");
        public static readonly Guid TaskProductionUnit = new Guid("f7459dc6-4bf4-4e48-8c3e-b4b24414d7a0");
        public static readonly Guid TaskPlan = new Guid("c4c97e67-4031-42fb-b576-1820a6799492");
    }
}
