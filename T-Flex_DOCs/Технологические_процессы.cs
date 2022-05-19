using System.Linq;
using System.Windows.Forms;
using TFlex.DOCs.Common;
using TFlex.DOCs.Model.Macros;
using TFlex.DOCs.Model.Macros.ObjectModel;
using TFlex.DOCs.Model.Plugins.Technology;
using TFlex.DOCs.Model.References;
using TFlex.DOCs.Model.Structure;
using TFlex.DOCs.UI.Controls.References.ObjectStructure;
using TFlex.DOCs.UI.Objects.Managers;
using TFlex.Model.Technology.References.TechnologyElements;
using TFlex.Model.Technology.References.TechnologyElements.VariableData;
using TFlex.Technology.References;
using TFlex.Technology.UI.Dialogs;

public class TechnologicalMacro : MacroProvider
{
    public TechnologicalMacro(MacroContext context)
        : base(context)
    {
    }

    public bool СкрытьМатериалы()
    {
        Объект элементТехнологии = ТекущийОбъект.РодительскийОбъект;
        while (элементТехнологии != null)
        {
            //Если элементТехнологии является Технологическим процессом
            if (элементТехнологии.Тип.ПорожденОт("3e93d599-c214-48c8-854f-efe4b475c4d8"))
            {
                //Возвращаем значение флага "Запретить ввод материала в операциях и переходах"
                return (bool)элементТехнологии.Параметр["9cb66d93-af18-43a3-85e9-06997828faf1"];
            }
            элементТехнологии = элементТехнологии.РодительскийОбъект;
        }

        //Если не нашли техпроцесс, например нет доступа, то вкладку не скрываем
        return false;
    }

    public void СоздатьСборочныеОперации()
    {
        var process = Context.ReferenceObject as StructuredTechnologicalProcess;
        if (process == null)
            return;
        process.CreateAssemblyOperations();
    }

    public void РасчётСуммарногоВремениТехпроцесса()
    {
        if (Context.ReferenceObject is TechnologicalProcess)
        {
            ((TechnologicalProcess)Context.ReferenceObject).SumTimeByOperations();
        }
        else if (Context.ReferenceObject is StructuredTechnologicalProcess)
        {
            CalculateProcessTimes(Context.ReferenceObject as StructuredTechnologicalProcess);
        }
        else if (Context.ReferenceObject is StructuredTechnologicalOperation)
        {
            CalculateOperationTimes(Context.ReferenceObject as StructuredTechnologicalOperation);
        }
    }

    private void CalculateProcessTimes(StructuredTechnologicalProcess process)
    {
        if (process == null)
            return;

        var operations = process.GetOperations(true).Where(o => o.Version.IsEmpty).ToArray();

        if (process.SumPieceTimeUnit == null && process.SumPieceTimeUnitLink != null)
            process.SumPieceTimeUnitLink.SetLinkedObject(process.SumPieceTime.ParameterInfo.Unit);
        if (process.SumPrepTimeUnit == null && process.SumPrepTimeUnitLink != null)
            process.SumPrepTimeUnitLink.SetLinkedObject(process.SumPrepTime.ParameterInfo.Unit);

        var sumPieceUnit = process.SumPieceTimeUnit ?? process.SumPieceTime.ParameterInfo.Unit;
        var sumPrepUnit = process.SumPrepTimeUnit ?? process.SumPrepTime.ParameterInfo.Unit;

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
    }

    private void CalculateOperationTimes(StructuredTechnologicalOperation operation)
    {
        if (operation == null)
            return;

        bool endChanges = false;
        if (!operation.IsChanged)
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
            operation.PieceTime.Value = steps.Aggregate(0.0, (total, next) =>
            {
                var baseTimeUnit = next.BaseTimeUnit ?? next.BaseTime.ParameterInfo.Unit;
                total += baseTimeUnit == null ? next.BaseTime : operationPieceTimeUnit.Convert(next.BaseTime, baseTimeUnit);
                var addTimeUnit = next.AdditionalTimeUnit ?? next.AdditionalTime.ParameterInfo.Unit;
                total += addTimeUnit == null ? next.AdditionalTime : operationPieceTimeUnit.Convert(next.AdditionalTime, addTimeUnit);
                return total;
            });
        }
        if (endChanges)
            operation.EndChanges();
    }

    public void ВыборКвалитетаТехпроцесса()
    {
        UIMacroContext uiContext = Context as UIMacroContext;
        Tolerance selectKvalitet = new Tolerance();

        //selectKvalitet.data.posadka = _kvalitetButtonEdit.Text;
        if (selectKvalitet.ShowDialog() == System.Windows.Forms.DialogResult.OK)//TFlex.DOCs.UI.Common.DialogOpenResult.Ok)
        {
            Параметр["Неуказываемый квалитет точности"] = selectKvalitet.data.posadka;
            //_kvalitetButtonEdit.Text = selectKvalitet.data.posadka;
        }
    }


    public void РасчётВремениОперации1()
    {
        TechnologicalOperation operation = Context.ReferenceObject as TechnologicalOperation;
        if (operation == null)
            return;
        operation.SumTimeBySteps();
    }

    public void РасчётВремениОперации2()
    {
        TechnologicalOperation operation = Context.ReferenceObject as TechnologicalOperation;
        if (operation == null)
            return;
        operation.SumTimeByTP();
    }
    public void ЗаданиеНомераОперации()
    {
        TechnologicalOperation operation = Context.ReferenceObject as TechnologicalOperation;
        ВводНомера ввод = new ВводНомера(operation.SystemFields.Order.Value);
        if (ввод.ShowDialog() != DialogResult.OK)
            return;

        if (operation.SystemFields.Order != (int)ввод.numericUpDown1.Value)
        {
            try
            {
                operation.ValidateOrder((int)ввод.numericUpDown1.Value);
            }
            catch
            {
                MessageBox.Show("Номер операции " + ввод.numericUpDown1.Value.ToString() + " уже существует", "Ошибка");
                return;
            }
            operation.SystemFields.Order = (int)ввод.numericUpDown1.Value;
            ParameterInfo parameterOrder = operation.Reference.ParameterGroup[SystemParameterType.Order];

            operation.OnParameterChanged(operation.ParameterValues[parameterOrder]);
        }
    }

    public void Нумеровать()
    {
        var process = Context.ReferenceObject as StructuredTechnologicalProcess;
        if (process == null)
            return;

        var operationStartNumber = GetGlobalValue("OperationStartNumber", 5);
        var operationInterval = GetGlobalValue("OperationInterval", 5);
        var operationLength = GetGlobalValue("OperationNumberOfDigits", 3);

        var stepStartNumber = GetGlobalValue("StepStartNumber", 5);
        var stepInterval = GetGlobalValue("StepInterval", 5);
        var stepLength = GetGlobalValue("StepNumberOfDigits", 3);

        TechnologyPlugin.UpdateOperationNumbers(process, operationStartNumber, operationInterval, operationLength);

        foreach (var operation in process.GetOperations(true))
            TechnologyPlugin.UpdateStepNumbers(operation, stepStartNumber, stepInterval, stepLength);
    }

    private int GetGlobalValue(string name, int defaultValue = 0)
    {
        var globalParameter = Context.Connection.References.GlobalParameters[name];
        if (globalParameter != null && globalParameter.Class.IsInt)
            return globalParameter.Value.GetInt32();
        return defaultValue;
    }

    public ButtonValidator ValidateRenumber()
    {
        var uiContext = Context as UIMacroContext;
        return new ButtonValidator
        {
            Visible = uiContext != null && uiContext.OwnerWindow is ObjectStructureVisualRepresentationControl,
            Enable = true
        };
    }

    public void ВОперации()
    {
        var variables = Context.GetSelectedObjects()
            .OfType<TPVariableDataObject>()
            .ToArray();

        if (variables.IsNullOrEmpty())
            return;

        var current = (ReferenceObject)CurrentObject;
        if (!(current.MasterObject is TypicalProcess process))
            return;

        var operations = process.GetOperations(true).OfType<ITypicalOperation>();
        if (operations.IsNullOrEmpty())
            return;

        var userDialog = GetUserDialog("Выбор технологических операций");
        if (userDialog is null)
            return;

        userDialog.BeginChanges();
        foreach (var linkedObject in userDialog.LinkedObjects["27273d69-aa89-4b73-9c66-6f25fc263092"])
            linkedObject.Delete();

        foreach (var operation in operations.OfType<ReferenceObject>())
        {
            var newRefObj = userDialog.CreateListObject("27273d69-aa89-4b73-9c66-6f25fc263092", "4c7f8415-2a2f-4bd3-ad53-1afc46878ed1");
            newRefObj["fd7f34c5-ec51-4ebe-ab34-740b4243097b"] = operation.ToString();
            newRefObj.AddLink("0bf1f778-2392-40a1-bcf9-2c60a3addd34", RefObj.CreateInstance(operation, Context));
            newRefObj.Save();
        }
        userDialog.Save();

        if (!userDialog.Show())
            return;

        var selectedObjects = userDialog
            .LinkedObjects["27273d69-aa89-4b73-9c66-6f25fc263092"]
            .Where(listObject => (bool)listObject["d8beef51-69df-4a36-8457-188341ebef30"] == true);

        foreach (var selectedObject in selectedObjects)
        {
            var linkedObject = selectedObject.LinkedObject["0bf1f778-2392-40a1-bcf9-2c60a3addd34"];
            if (linkedObject is null)
                continue;

            var referenceObject = (ReferenceObject)linkedObject;
            if (referenceObject is ITypicalOperation operation)
                operation.ModifyOperation(o => o.CreateVariableData(variables));
        }
    }
}

class ВводНомера : Form
{
    public ВводНомера(int номер)
    {
        InitializeComponent();
        numericUpDown1.Value = номер;
    }

    private void InitializeComponent()
    {
        this.label1 = new System.Windows.Forms.Label();
        this.numericUpDown1 = new System.Windows.Forms.NumericUpDown();
        this.button1 = new System.Windows.Forms.Button();
        this.button2 = new System.Windows.Forms.Button();
        //((System.ComponentModel.ISupportInitialize)(this.numericUpDown1)).BeginInit();
        this.SuspendLayout();
        // 
        // label1
        // 
        this.label1.AutoSize = true;
        this.label1.Location = new System.Drawing.Point(12, 12);
        this.label1.Name = "label1";
        this.label1.Size = new System.Drawing.Size(55, 17);
        this.label1.TabIndex = 0;
        this.label1.Text = "Номер:";
        // 
        // numericUpDown1
        // 
        this.numericUpDown1.Increment = 5;
        this.numericUpDown1.Location = new System.Drawing.Point(73, 10);
        this.numericUpDown1.Minimum = 1;
        this.numericUpDown1.Maximum = 1000;
        this.numericUpDown1.Name = "numericUpDown1";
        this.numericUpDown1.Size = new System.Drawing.Size(121, 22);
        this.numericUpDown1.TabIndex = 2;
        // 
        // button1
        // 
        this.button1.DialogResult = System.Windows.Forms.DialogResult.OK;
        this.button1.Location = new System.Drawing.Point(38, 41);
        this.button1.Name = "button1";
        this.button1.Size = new System.Drawing.Size(75, 29);
        this.button1.TabIndex = 3;
        this.button1.Text = "ОК";
        this.button1.UseVisualStyleBackColor = true;
        // 
        // button2
        // 
        this.button2.DialogResult = System.Windows.Forms.DialogResult.Cancel;
        this.button2.Location = new System.Drawing.Point(119, 41);
        this.button2.Name = "button2";
        this.button2.Size = new System.Drawing.Size(75, 29);
        this.button2.TabIndex = 4;
        this.button2.Text = "Отмена";
        this.button2.UseVisualStyleBackColor = true;
        // 
        // Form1
        // 
        this.AcceptButton = this.button1;
        this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
        this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
        this.CancelButton = this.button2;
        this.ClientSize = new System.Drawing.Size(204, 80);
        this.ControlBox = false;
        this.Controls.Add(this.button2);
        this.Controls.Add(this.button1);
        this.Controls.Add(this.numericUpDown1);
        this.Controls.Add(this.label1);
        this.Name = "Form1";
        this.ShowIcon = false;
        this.ShowInTaskbar = false;
        this.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide;
        this.Text = "№ операции";
        this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;

        //((System.ComponentModel.ISupportInitialize)(this.numericUpDown1)).EndInit();
        this.ResumeLayout(false);
        this.PerformLayout();

    }


    private System.Windows.Forms.Label label1;
    public System.Windows.Forms.NumericUpDown numericUpDown1;
    private System.Windows.Forms.Button button1;
    private System.Windows.Forms.Button button2;
}
