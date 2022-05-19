using System;
using System.Linq;
using System.Windows.Forms;
using TFlex.DOCs.Model;
using TFlex.DOCs.Model.Macros;
using TFlex.DOCs.Model.Macros.ObjectModel;
using TFlex.DOCs.Model.Plugins;
using TFlex.DOCs.Model.Structure;
using TFlex.DOCs.UI.Common;
using TFlex.DOCs.UI.Common.Templates;
using TFlex.DOCs.UI.Controls;
using TFlex.DOCs.UI.Controls.Templates;
using TFlex.DOCs.UI.Messages;
using TFlex.DOCs.UI.Objects.Administration;
using TFlex.DOCs.UI.Objects.Administration.VisualRepresentation;
using TFlex.DOCs.UI.Objects.Managers;
using TFlex.DOCs.UI.Objects.ReferenceModel;
using TFlex.DOCs.UI.Objects.ReferenceModel.VisualRepresentation;
using TFlex.DOCs.Model.References;
using System.Collections.Generic;


public class SelectParameterDialog : SelectDialog
{
    private bool _canSelectOnlyLinks;
    private bool _canSelectOnlyParameter;
    private System.ComponentModel.IContainer components = null;


    public SelectParameterDialog()
    {
        InitializeComponent();
    }

    public SelectParameterDialog(bool canSelectOnlyLikns, bool canSelectOnlyParameter)
        : this()
    {
        _canSelectOnlyLinks = canSelectOnlyLikns;
        _canSelectOnlyParameter = canSelectOnlyParameter;
    }

    protected override void ShowInvalidSelectedMessage()
    {
        MessageDialog.MessageInformationShow(this, _canSelectOnlyLinks ? "Связь не выбрана" : "Параметр не выбран");
    }

    public new VisualRepresentationControl DialogControl
    {
        get
        {
            return base.DialogControl as VisualRepresentationControl;
        }
    }

    public override bool ButtonOkClicked()
    {
        if (!_canSelectOnlyLinks && GetSelectedObject() == null)
            return false;
        var view = DialogControl.VisualRepresentation as ReferenceGroupsVisualRepresentation;
        if (_canSelectOnlyLinks && view != null && !view.GetSelectedParameterGroup().IsLinkToOne)
            return false;

        return true;
    }

    public ParameterInfo GetSelectedObject()
    {
        var view = this.DialogControl.VisualRepresentation as ReferenceGroupsVisualRepresentation;
        if (view != null)
            return (view.FocusedObject as ParameterInfoUIObject).ParameterInfo;

        return null;
    }


    protected override void Dispose(bool disposing)
    {
        if (disposing && (components != null))
        {
            components.Dispose();
        }
        base.Dispose(disposing);
    }

    private void InitializeComponent()
    {
        this.SuspendLayout();
        // 
        // SelectParameterDialog
        // 
        this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
        this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
        this.ClientSize = new System.Drawing.Size(315, 161);
        this.Name = "SelectParameterDialog";
        this.ShowIcon = false;
        this.Text = "StringPatternObjectDialog";
        this.ResumeLayout(false);
    }
}

public  class TestForm : Form
{
    private ParameterGroup _parameterGroup;
    private ReferenceInfo _mainReference;
    private System.ComponentModel.IContainer components = null;

    public TestForm()
    {
        InitializeComponent();
    }

    public TestForm(ReferenceInfo reference)
        : this()
    {
        _mainReference = reference;
    }

    private void buttonEdit1_ButtonClick(object sender, DevExpress.XtraEditors.Controls.ButtonPressedEventArgs e)
    {
        using (SelectParameterDialog viewDialog = new SelectParameterDialog(true, false))
        {
            using (ReferenceGroupsVisualRepresentation viewGroupsRepresentation = new ReferenceGroupsVisualRepresentation(ObjectCreator.CreateObject<ITreeViewControl>()) { RootReference = null })
            {
                InitializeView(_mainReference, viewGroupsRepresentation, viewDialog, true);
                if (viewDialog.ShowDialog(this) == DialogOpenResult.Ok)
                {
                    _parameterGroup = viewGroupsRepresentation.GetSelectedParameterGroup();
                    InitializeEditControls();
                }
            }
        }
    }

    public ParameterInfo MainParameter
    {
        get
        {
            return _mainParameterButtonEdit.EditValue as ParameterInfo;
        }
    }

    public ParameterInfo SlaveParameter
    {
        get
        {
            return _slaveParameterButtonEdit.EditValue as ParameterInfo;
        }
    }

    public ParameterGroup LinkGroup
    {
        get
        {
            return _parameterGroup;
        }
    }

    private void InitializeEditControls()
    {
        buttonEdit1.EditValue = _parameterGroup;
        _slaveParameterButtonEdit.EditValue = null;
    }

    private void buttonEdit1_EditValueChanged(object sender, EventArgs e)
    {
        if (buttonEdit1.EditValue == null)
        {
            _mainParameterButtonEdit.EditValue = null;
            _slaveParameterButtonEdit.EditValue = null;
        }
    }

    private void buttonEdit2_ButtonClick(object sender, DevExpress.XtraEditors.Controls.ButtonPressedEventArgs e)
    {
        if (_parameterGroup == null)
            return;

        using (SelectParameterDialog viewDialog = new SelectParameterDialog(false, true))
        {
            using (ReferenceGroupsVisualRepresentation viewGroupsRepresentation = new ReferenceGroupsVisualRepresentation(ObjectCreator.CreateObject<ITreeViewControl>()) { RootReference = null })
            {
                InitializeView(_mainReference, viewGroupsRepresentation, viewDialog, false);
                if (viewDialog.ShowDialog(this) == DialogOpenResult.Ok)
                {
                    _mainParameterButtonEdit.EditValue = viewGroupsRepresentation.GetSelectedParameter();
                    InitializeEditControls();
                }
            }
        }
    }

    private void buttonEdit3_ButtonClick(object sender, DevExpress.XtraEditors.Controls.ButtonPressedEventArgs e)
    {
        using (SelectParameterDialog viewDialog = new SelectParameterDialog(false, true))
        {
            using (ReferenceGroupsVisualRepresentation viewGroupsRepresentation = new ReferenceGroupsVisualRepresentation(ObjectCreator.CreateObject<ITreeViewControl>()) { RootReference = null })
            {
                InitializeView(_parameterGroup.SlaveGroup.ReferenceInfo, viewGroupsRepresentation, viewDialog, false);
                if (viewDialog.ShowDialog(this) == DialogOpenResult.Ok)
                {
                    _slaveParameterButtonEdit.EditValue = viewGroupsRepresentation.GetSelectedParameter();
                }
            }
        }
    }

    private void InitializeView(ReferenceInfo rootReference, ReferenceGroupsVisualRepresentation viewGroupsRepresentation, SelectParameterDialog dlg, bool showOnlyToOneLinks)
    {
        IVisualRepresentationControl viewGroupsRepresentationControl = ObjectCreator.CreateObject<IVisualRepresentationControl>();
        viewGroupsRepresentation.RootReference = rootReference;
        if (showOnlyToOneLinks)
            viewGroupsRepresentation.ShowOnlyToOneLinks = true;
        else
            viewGroupsRepresentation.ShowOnlyToOneParameters = true;
        viewGroupsRepresentation.EnableToolBar = false;
        viewGroupsRepresentation.EnablePopUpMenu = false;
        viewGroupsRepresentationControl.SetVisualRepresentation(viewGroupsRepresentation);
        dlg.InitializeControl(viewGroupsRepresentationControl as VisualRepresentationControl, showOnlyToOneLinks ? "Выбор связи" : "Выбор параметра");
        viewGroupsRepresentationControl.EnablePopUpMenu = false;
        viewGroupsRepresentationControl.EnableToolBar = false;
        viewGroupsRepresentationControl.IsMultipleSelect = false;
        viewGroupsRepresentationControl.Initialize();
    }

    protected override void Dispose(bool disposing)
    {
        if (disposing && (components != null))
        {
            components.Dispose();
        }
        base.Dispose(disposing);
    }

    #region Windows Form Designer generated code

    private void InitializeComponent()
    {
        this.labelControl1 = new DevExpress.XtraEditors.LabelControl();
        this.labelControl2 = new DevExpress.XtraEditors.LabelControl();
        this.labelControl3 = new DevExpress.XtraEditors.LabelControl();
        this.buttonEdit1 = new DevExpress.XtraEditors.ButtonEdit();
        this._mainParameterButtonEdit = new DevExpress.XtraEditors.ButtonEdit();
        this._slaveParameterButtonEdit = new DevExpress.XtraEditors.ButtonEdit();
        this._okButton = new DevExpress.XtraEditors.SimpleButton();
        this._cancelButton = new DevExpress.XtraEditors.SimpleButton();
        ((System.ComponentModel.ISupportInitialize)(this.buttonEdit1.Properties)).BeginInit();
        ((System.ComponentModel.ISupportInitialize)(this._mainParameterButtonEdit.Properties)).BeginInit();
        ((System.ComponentModel.ISupportInitialize)(this._slaveParameterButtonEdit.Properties)).BeginInit();
        this.SuspendLayout();
        // 
        // labelControl1
        // 
        this.labelControl1.Location = new System.Drawing.Point(12, 50);
        this.labelControl1.Name = "labelControl1";
        this.labelControl1.Size = new System.Drawing.Size(53, 13);
        this.labelControl1.TabIndex = 1;
        this.labelControl1.Text = "Параметр:";
        // 
        // labelControl2
        // 
        this.labelControl2.Location = new System.Drawing.Point(12, 24);
        this.labelControl2.Name = "labelControl2";
        this.labelControl2.Size = new System.Drawing.Size(34, 13);
        this.labelControl2.TabIndex = 1;
        this.labelControl2.Text = "Связь:";
        // 
        // labelControl3
        // 
        this.labelControl3.Location = new System.Drawing.Point(12, 76);
        this.labelControl3.Name = "labelControl3";
        this.labelControl3.Size = new System.Drawing.Size(185, 13);
        this.labelControl3.TabIndex = 1;
        this.labelControl3.Text = "Параметр в связанном справочнике:";
        // 
        // buttonEdit1
        // 
        this.buttonEdit1.Location = new System.Drawing.Point(203, 21);
        this.buttonEdit1.Name = "buttonEdit1";
        this.buttonEdit1.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
            new DevExpress.XtraEditors.Controls.EditorButton()});
        this.buttonEdit1.Size = new System.Drawing.Size(351, 20);
        this.buttonEdit1.TabIndex = 2;
        this.buttonEdit1.ButtonClick += new DevExpress.XtraEditors.Controls.ButtonPressedEventHandler(this.buttonEdit1_ButtonClick);
        this.buttonEdit1.EditValueChanged += new System.EventHandler(this.buttonEdit1_EditValueChanged);
        // 
        // _mainParameterButtonEdit
        // 
        this._mainParameterButtonEdit.Location = new System.Drawing.Point(203, 47);
        this._mainParameterButtonEdit.Name = "_mainParameterButtonEdit";
        this._mainParameterButtonEdit.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
            new DevExpress.XtraEditors.Controls.EditorButton()});
        this._mainParameterButtonEdit.Size = new System.Drawing.Size(351, 20);
        this._mainParameterButtonEdit.TabIndex = 2;
        this._mainParameterButtonEdit.ButtonClick += new DevExpress.XtraEditors.Controls.ButtonPressedEventHandler(this.buttonEdit2_ButtonClick);
        // 
        // _slaveParameterButtonEdit
        // 
        this._slaveParameterButtonEdit.Location = new System.Drawing.Point(203, 73);
        this._slaveParameterButtonEdit.Name = "_slaveParameterButtonEdit";
        this._slaveParameterButtonEdit.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
            new DevExpress.XtraEditors.Controls.EditorButton()});
        this._slaveParameterButtonEdit.Size = new System.Drawing.Size(351, 20);
        this._slaveParameterButtonEdit.TabIndex = 2;
        this._slaveParameterButtonEdit.ButtonClick += new DevExpress.XtraEditors.Controls.ButtonPressedEventHandler(this.buttonEdit3_ButtonClick);
        // 
        // _okButton
        // 
        this._okButton.DialogResult = System.Windows.Forms.DialogResult.OK;
        this._okButton.Location = new System.Drawing.Point(398, 108);
        this._okButton.Name = "_okButton";
        this._okButton.Size = new System.Drawing.Size(75, 23);
        this._okButton.TabIndex = 3;
        this._okButton.Text = "OK";
        // 
        // _cancelButton
        // 
        this._cancelButton.DialogResult = System.Windows.Forms.DialogResult.Cancel;
        this._cancelButton.Location = new System.Drawing.Point(479, 108);
        this._cancelButton.Name = "_cancelButton";
        this._cancelButton.Size = new System.Drawing.Size(75, 23);
        this._cancelButton.TabIndex = 3;
        this._cancelButton.Text = "Отмена";
        // 
        // TestForm
        // 
        this.AcceptButton = this._okButton;
        this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
        this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
        this.CancelButton = this._cancelButton;
        this.ClientSize = new System.Drawing.Size(566, 143);
        this.Controls.Add(this._cancelButton);
        this.Controls.Add(this._okButton);
        this.Controls.Add(this._slaveParameterButtonEdit);
        this.Controls.Add(this._mainParameterButtonEdit);
        this.Controls.Add(this.buttonEdit1);
        this.Controls.Add(this.labelControl2);
        this.Controls.Add(this.labelControl3);
        this.Controls.Add(this.labelControl1);
        this.Name = "TestForm";
        ((System.ComponentModel.ISupportInitialize)(this.buttonEdit1.Properties)).EndInit();
        ((System.ComponentModel.ISupportInitialize)(this._mainParameterButtonEdit.Properties)).EndInit();
        ((System.ComponentModel.ISupportInitialize)(this._slaveParameterButtonEdit.Properties)).EndInit();
        this.ResumeLayout(false);
        this.PerformLayout();

    }

    #endregion

    private DevExpress.XtraEditors.LabelControl labelControl1;
    private DevExpress.XtraEditors.LabelControl labelControl2;
    private DevExpress.XtraEditors.LabelControl labelControl3;
    private DevExpress.XtraEditors.ButtonEdit buttonEdit1;
    private DevExpress.XtraEditors.ButtonEdit _mainParameterButtonEdit;
    private DevExpress.XtraEditors.ButtonEdit _slaveParameterButtonEdit;
    private DevExpress.XtraEditors.SimpleButton _okButton;
    private DevExpress.XtraEditors.SimpleButton _cancelButton;
}

public class Macro : MacroProvider
{
    public Macro(MacroContext context)
        : base(context)
    {
    }

    public override void Run()
    {
        UIMacroContext uiContext = Context as UIMacroContext;
        if (uiContext != null)
        {
            IReferenceCompositeVisualRepresentation[] views = uiContext.FindReferenceVisualRepresentations(TFlex.DOCs.UI.Objects.Managers.UIMacroContext.FindReferenceVisualRepresentationsType.CurrentApplicationWindow);
            if (views != null && views.Length != 0)
            {
                var selectedObjects = views[0].GetSelectedObjects().OfType<ReferenceUIObject>().ToArray();
                if (selectedObjects.Length != 0)
                {
                    using (TestForm form = new TestForm(selectedObjects[0].ReferenceObject.Reference.ParameterGroup.ReferenceInfo))
                    {
                        if (form.ShowDialog() == System.Windows.Forms.DialogResult.OK && form.LinkGroup != null)
                        {
                            Reference reference = form.LinkGroup.SlaveGroup.ReferenceInfo.CreateReference();
                            if (reference != null)
                            {
                                foreach (var uiObject in selectedObjects)
                                {
                                    List<ReferenceObject> objects = reference.Find(form.SlaveParameter, uiObject.ReferenceObject.ParameterValues[form.MainParameter].Value);
                                    if (objects.Count == 0)
                                        continue;
                                    uiObject.ReferenceObject.BeginChanges();    
                                    try
                                    {
                                              uiObject.ReferenceObject.SetLinkedObject(form.LinkGroup.Guid, objects[0]);
                                    }
                                    catch
                                   {
                                         uiObject.ReferenceObject.CancelChanges();
                                         throw;
                                    }
                                    uiObject.ReferenceObject.EndChanges();
                                    uiObject.Reset();
                                }
                            }
                        }
                    }
                }
            }
        }
    }
}
