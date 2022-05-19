using System;
using System.Windows.Forms;
using System.Collections.Generic;
using System.Linq;
using TFlex.DOCs.Scheduling;
using TFlex.DOCs.UI.Objects.Commands;
using TFlex.DOCs.UI.Objects;
using TFlex.DOCs.UI.Objects.ReferenceModel;
using TFlex.DOCs.References.WorkOrder;
using TFlex.DOCs.Model;
using TFlex.Technology.References;
using TFlex.DOCs.UI.Objects.Managers;
using TFlex.DOCs.UI.Objects.ReferenceModel.VisualRepresentation;
using TFlex.DOCs.Model.Diagnostics;
using TFlex.DOCs.Model.References;
using TFlex.DOCs.Model.References.Links;
using TFlex.DOCs.Model.Macros;
using TFlex.Model.Technology.References.ParametersProvider;

public class Macro : MacroProvider
{
    public Macro(MacroContext context)
        : base(context)
    {
    }

    public void IssueTaskOrder()
    {
            if (Context == null)
                 return;

            TaskReferenceObject task = Context == null || Context.ReferenceObject != null ? Context.ReferenceObject as TaskReferenceObject : null;
            if (task == null)
                return;

            task.WorkOrders.Clear();
            if (task.WorkOrders.OfType<WorkOrderReferenceObject>().FirstOrDefault(wo => wo.SystemFields.Stage.Guid != WorkOrderReferenceObject.StageKeys.Canceled) != null)
            {
                MessageBox.Show("На операцию не может быть выписано более одного актуального наряда");
                return;
            }

            WorkOrderReference orderReference = new WorkOrderReference();
            OperationOrderReferenceObject operationOrder = orderReference.CreateReferenceObject(orderReference.Classes.OperationOrder) as OperationOrderReferenceObject;
            if (operationOrder == null)
                return;

            try
            {
                operationOrder.Operation = task;
                operationOrder.IssueTime.Value = DateTime.Now;
                operationOrder.ClosingTime.Value = operationOrder.ClosingTime.EmptyValue;
                operationOrder.StartTime.Value = task.StartTime;
                operationOrder.EndTime.Value = task.EndTime;
                operationOrder.Owner = ClientView.Current.GetUser();
                operationOrder.ProductionUnit = task.ProductionUnit;
                operationOrder.Equipment = task.Equipment;

                PositionReferenceObject position = task.Position;
                if (position != null)
                {
                    operationOrder.Amount.Value = task.Amount;
                    operationOrder.Nomenclature = position.Nomenclature;

                    if (position.Order != null)
                    {
                        operationOrder.Order.Value = position.Order.Name;
                        operationOrder.OrderCode.Value = position.Order.Code;
                    }

                    if (position.Batch != null)
                        operationOrder.Batch.Value = position.Batch.Name;
                }

                operationOrder.OperationGuid.Value = task.OperationGuid;
                if (task.TechnologicalProcess != null)
                {
                    operationOrder.TechnologicalProcess.Value = task.TechnologicalProcess.Name;
                    var technologicalOperation = task.TechnologicalProcess.Operations.OfType<ReferenceObject>()
                    	.FirstOrDefault(o => o.SystemFields.Guid == task.OperationGuid) as ITechnologicalOperationParametersProvider;
                    if (technologicalOperation != null)
                    {
                        operationOrder.OperationName.Value = technologicalOperation.Name;
                        operationOrder.OperationCode.Value = technologicalOperation.Code;
                        int order;
                        if (int.TryParse(technologicalOperation.Order, out order))
                        	operationOrder.OperationNumber.Value = order;
                        operationOrder.PieceTime.Value = technologicalOperation.PieceTime;
                        operationOrder.PrepTime.Value = technologicalOperation.PrepTime;
                        operationOrder.PiceTimeUnit = technologicalOperation.PieceTimeUnit;
                        operationOrder.PrepTimeUnit = technologicalOperation.PrepTimeUnit;
                    }
                }

                ReferenceObjectEditManager.Instance.ShowReferenceObjectPropertiesDialog(operationOrder, null, null, null, null, null);
            }
            catch
            {
                operationOrder.CancelChanges();
            }
    }
}
