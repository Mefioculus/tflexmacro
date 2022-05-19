using System;
using System.Linq;
using System.ComponentModel;
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
using DevExpress.XtraEditors.Controls;



public class TestForm : Form
{
    private ParameterGroup _firstParameterGroup, _secondParameterGroup;
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
                InitializeView(_mainReference, viewGroupsRepresentation, viewDialog, false);
                if (viewDialog.ShowDialog(this) == DialogOpenResult.Ok)
                {
                    _firstLinkButtonEdit.EditValue = _firstParameterGroup = viewGroupsRepresentation.GetSelectedParameterGroup();
                }
            }
        }
    }

    private void buttonEdit2_ButtonClick(object sender, DevExpress.XtraEditors.Controls.ButtonPressedEventArgs e)
    {
       if (_firstParameterGroup == null)
       {
                 MessageDialog.MessageInformationShow(this, "Необходимо сначала выбрать Связь 1");
                 return;
        }
   
        using (SelectParameterDialog viewDialog = new SelectParameterDialog(true, false))
        {
            using (ReferenceGroupsVisualRepresentation viewGroupsRepresentation = new ReferenceGroupsVisualRepresentation(ObjectCreator.CreateObject<ITreeViewControl>()) { RootReference = null })
            {
                InitializeView(_mainReference, viewGroupsRepresentation, viewDialog, _firstParameterGroup.IsLinkToMany ? true : false);
                if (viewDialog.ShowDialog(this) == DialogOpenResult.Ok)
                {
                    _secondLinkButtonEdit.EditValue = _secondParameterGroup = viewGroupsRepresentation.GetSelectedParameterGroup();
                }
            }
        }
    }

    public ParameterGroup LinkGroup1
    {
        get
        {
            return _firstParameterGroup;
        }
    }

    public ParameterGroup LinkGroup2
    {
        get
        {
            return _secondParameterGroup;
        }
    }

    private void buttonEdit1_EditValueChanged(object sender, EventArgs e)
    {
        if (_firstLinkButtonEdit.EditValue == null ||
        _firstParameterGroup.IsLinkToMany && _secondParameterGroup != null && _secondParameterGroup.IsLinkToOne)
            _secondLinkButtonEdit.EditValue = _secondParameterGroup = null;
    }


    private void InitializeView(ReferenceInfo rootReference, ReferenceGroupsVisualRepresentation viewGroupsRepresentation, SelectParameterDialog dlg, bool showOnlyToManyLinks)
    {
        IVisualRepresentationControl viewGroupsRepresentationControl = ObjectCreator.CreateObject<IVisualRepresentationControl>();
        viewGroupsRepresentation.RootReference = rootReference;
        if (showOnlyToManyLinks)
            viewGroupsRepresentation.ShowOnlyToManyLinks = showOnlyToManyLinks;
        else
            viewGroupsRepresentation.ShowOnlyLinks = true;
        viewGroupsRepresentation.EnableToolBar = false;
        viewGroupsRepresentation.EnablePopUpMenu = false;
        viewGroupsRepresentationControl.SetVisualRepresentation(viewGroupsRepresentation);
        dlg.InitializeControl(viewGroupsRepresentationControl as VisualRepresentationControl, "Выбор связи");
        viewGroupsRepresentationControl.EnablePopUpMenu = false;
        viewGroupsRepresentationControl.EnableToolBar = false;
        viewGroupsRepresentationControl.IsMultipleSelect = false;
        viewGroupsRepresentationControl.Initialize();
    }

    protected override void OnClosing(CancelEventArgs e)
    {
        if (DialogResult == System.Windows.Forms.DialogResult.OK)
            if (_firstParameterGroup == null)
            {
                MessageDialog.MessageInformationShow(this, "Значение связи 1 не заполнено");
                e.Cancel = true;
            }
            else if (_secondParameterGroup == null)
            {
                MessageDialog.MessageInformationShow(this, "Значение связи 2 не заполнено");
                e.Cancel = true;
            }
        base.OnClosing(e);
    }

    private void _okButton_Click(object sender, EventArgs e)
    {

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
        this.Text = "Объединение связей";
        this.labelControl2 = new DevExpress.XtraEditors.LabelControl();
        this._firstLinkButtonEdit = new DevExpress.XtraEditors.ButtonEdit();
        this._okButton = new DevExpress.XtraEditors.SimpleButton();
        this._cancelButton = new DevExpress.XtraEditors.SimpleButton();
        this.labelControl1 = new DevExpress.XtraEditors.LabelControl();
        this._secondLinkButtonEdit = new DevExpress.XtraEditors.ButtonEdit();
        ((System.ComponentModel.ISupportInitialize)(this._firstLinkButtonEdit.Properties)).BeginInit();
        ((System.ComponentModel.ISupportInitialize)(this._secondLinkButtonEdit.Properties)).BeginInit();
        this.SuspendLayout();
        // 
        // labelControl2
        // 
        this.labelControl2.Location = new System.Drawing.Point(12, 24);
        this.labelControl2.Name = "labelControl2";
        this.labelControl2.Size = new System.Drawing.Size(100, 13);
        this.labelControl2.TabIndex = 1;
        this.labelControl2.Text = "Связь 1 (источник):";
        // 
        // _firstLinkButtonEdit
        // 
        this._firstLinkButtonEdit.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
                    | System.Windows.Forms.AnchorStyles.Right)));
        this._firstLinkButtonEdit.Location = new System.Drawing.Point(128, 21);
        this._firstLinkButtonEdit.Name = "_firstLinkButtonEdit";
        this._firstLinkButtonEdit.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
            new DevExpress.XtraEditors.Controls.EditorButton()});
        this._firstLinkButtonEdit.Size = new System.Drawing.Size(605, 20);
        this._firstLinkButtonEdit.TabIndex = 2;
        this._firstLinkButtonEdit.Properties.TextEditStyle = TextEditStyles.DisableTextEditor;
        this._firstLinkButtonEdit.ButtonClick += new DevExpress.XtraEditors.Controls.ButtonPressedEventHandler(this.buttonEdit1_ButtonClick);
        this._firstLinkButtonEdit.EditValueChanged += new System.EventHandler(this.buttonEdit1_EditValueChanged);
        // 
        // _okButton
        // 
        this._okButton.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
        this._okButton.DialogResult = System.Windows.Forms.DialogResult.OK;
        this._okButton.Location = new System.Drawing.Point(577, 78);
        this._okButton.Name = "_okButton";
        this._okButton.Size = new System.Drawing.Size(75, 23);
        this._okButton.TabIndex = 3;
        this._okButton.Text = "Заменить";
        this._okButton.Click += new System.EventHandler(this._okButton_Click);
        // 
        // _cancelButton
        // 
        this._cancelButton.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
        this._cancelButton.DialogResult = System.Windows.Forms.DialogResult.Cancel;
        this._cancelButton.Location = new System.Drawing.Point(658, 78);
        this._cancelButton.Name = "_cancelButton";
        this._cancelButton.Size = new System.Drawing.Size(75, 23);
        this._cancelButton.TabIndex = 3;
        this._cancelButton.Text = "Отмена";
        // 
        // labelControl1
        // 
        this.labelControl1.Location = new System.Drawing.Point(12, 50);
        this.labelControl1.Name = "labelControl1";
        this.labelControl1.Size = new System.Drawing.Size(100, 13);
        this.labelControl1.TabIndex = 1;
        this.labelControl1.Text = "Связь 2 (приёмник):";
        // 
        // _secondLinkButtonEdit
        // 
        this._secondLinkButtonEdit.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
                    | System.Windows.Forms.AnchorStyles.Right)));
        this._secondLinkButtonEdit.Location = new System.Drawing.Point(128, 47);
        this._secondLinkButtonEdit.Name = "_secondLinkButtonEdit";
        this._secondLinkButtonEdit.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
            new DevExpress.XtraEditors.Controls.EditorButton()});
        this._secondLinkButtonEdit.Size = new System.Drawing.Size(605, 20);
        this._secondLinkButtonEdit.Properties.TextEditStyle = TextEditStyles.DisableTextEditor;
        this._secondLinkButtonEdit.TabIndex = 2;
        this._secondLinkButtonEdit.ButtonClick += new DevExpress.XtraEditors.Controls.ButtonPressedEventHandler(this.buttonEdit2_ButtonClick);
        // 
        // TestForm
        // 
        this.AcceptButton = this._okButton;
        this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
        this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
        this.CancelButton = this._cancelButton;
        this.ClientSize = new System.Drawing.Size(745, 113);
        this.Controls.Add(this._cancelButton);
        this.Controls.Add(this._okButton);
        this.Controls.Add(this._secondLinkButtonEdit);
        this.Controls.Add(this._firstLinkButtonEdit);
        this.Controls.Add(this.labelControl1);
        this.Controls.Add(this.labelControl2);
        this.Name = "TestForm";
        ((System.ComponentModel.ISupportInitialize)(this._firstLinkButtonEdit.Properties)).EndInit();
        ((System.ComponentModel.ISupportInitialize)(this._secondLinkButtonEdit.Properties)).EndInit();
        this.ResumeLayout(false);
        this.PerformLayout();

    }

    #endregion

    private DevExpress.XtraEditors.LabelControl labelControl2;
    private DevExpress.XtraEditors.ButtonEdit _firstLinkButtonEdit;
    private DevExpress.XtraEditors.SimpleButton _okButton;
    private DevExpress.XtraEditors.SimpleButton _cancelButton;
    private DevExpress.XtraEditors.LabelControl labelControl1;
    private DevExpress.XtraEditors.ButtonEdit _secondLinkButtonEdit;
}


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
        if (_canSelectOnlyLinks && view != null && !view.GetSelectedParameterGroup().IsLinkGroup)
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


public class RemovingObjectsForm : Form
{
    private System.ComponentModel.IContainer components = null;

    public RemovingObjectsForm()
    {
        InitializeComponent();
    }

    public RemovingObjectsForm(ReferenceObject[] objects)
        : this()
    {
        gridControl1.DataSource = objects.Select(o => new ObjectForGrid(o)).ToArray();
    }

    private class ObjectForGrid
    {
        ReferenceObject _ro;
        public ObjectForGrid(ReferenceObject ro)
        {
            _ro = ro;
        }

        public string Name
        {
            get
            {
                return _ro.ToString();
            }
        }
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
        this._okButton = new DevExpress.XtraEditors.SimpleButton();
        this._cancelButton = new DevExpress.XtraEditors.SimpleButton();
        this.gridControl1 = new DevExpress.XtraGrid.GridControl();
        this.gridView1 = new DevExpress.XtraGrid.Views.Grid.GridView();
        this._nameGridColumn = new DevExpress.XtraGrid.Columns.GridColumn();
        this.labelControl1 = new DevExpress.XtraEditors.LabelControl();
        ((System.ComponentModel.ISupportInitialize)(this.gridControl1)).BeginInit();
        ((System.ComponentModel.ISupportInitialize)(this.gridView1)).BeginInit();
        this.SuspendLayout();
        // 
        // _okButton
        // 
        this._okButton.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
        this._okButton.DialogResult = System.Windows.Forms.DialogResult.OK;
        this._okButton.Location = new System.Drawing.Point(577, 307);
        this._okButton.Name = "_okButton";
        this._okButton.Size = new System.Drawing.Size(75, 23);
        this._okButton.TabIndex = 3;
        this._okButton.Text = "Продолжить";
        // 
        // _cancelButton
        // 
        this._cancelButton.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
        this._cancelButton.DialogResult = System.Windows.Forms.DialogResult.Cancel;
        this._cancelButton.Location = new System.Drawing.Point(658, 307);
        this._cancelButton.Name = "_cancelButton";
        this._cancelButton.Size = new System.Drawing.Size(75, 23);
        this._cancelButton.TabIndex = 3;
        this._cancelButton.Text = "Отмена";
        // 
        // gridControl1
        // 
        this.gridControl1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
                    | System.Windows.Forms.AnchorStyles.Left)
                    | System.Windows.Forms.AnchorStyles.Right)));
        this.gridControl1.Location = new System.Drawing.Point(13, 32);
        this.gridControl1.MainView = this.gridView1;
        this.gridControl1.Name = "gridControl1";
        this.gridControl1.Size = new System.Drawing.Size(720, 269);
        this.gridControl1.TabIndex = 4;
        this.gridControl1.ViewCollection.AddRange(new DevExpress.XtraGrid.Views.Base.BaseView[] {
            this.gridView1});
        // 
        // gridView1
        // 
        this.gridView1.Columns.AddRange(new DevExpress.XtraGrid.Columns.GridColumn[] {
            this._nameGridColumn});
        this.gridView1.GridControl = this.gridControl1;
        this.gridView1.Name = "gridView1";
        this.gridView1.OptionsBehavior.Editable = false;
        this.gridView1.OptionsView.ShowGroupPanel = false;
        this.gridView1.OptionsView.ShowHorzLines = false;
        this.gridView1.OptionsView.ShowIndicator = false;
        this.gridView1.OptionsView.ShowVertLines = false;
        // 
        // _nameGridColumn
        // 
        this._nameGridColumn.Caption = "Наименование";
        this._nameGridColumn.FieldName = "Name";
        this._nameGridColumn.Name = "_nameGridColumn";
        this._nameGridColumn.Visible = true;
        this._nameGridColumn.VisibleIndex = 0;
        // 
        // labelControl1
        // 
        this.labelControl1.Location = new System.Drawing.Point(13, 13);
        this.labelControl1.Name = "labelControl1";
        this.labelControl1.Size = new System.Drawing.Size(304, 13);
        this.labelControl1.TabIndex = 5;
        this.labelControl1.Text = "Следующие объекты будут утеряны после замены связей:";
        // 
        // RemovingObjectsForm
        // 
        this.AcceptButton = this._okButton;
        this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
        this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
        this.CancelButton = this._cancelButton;
        this.ClientSize = new System.Drawing.Size(745, 342);
        this.Controls.Add(this.labelControl1);
        this.Controls.Add(this.gridControl1);
        this.Controls.Add(this._cancelButton);
        this.Controls.Add(this._okButton);
        this.Name = "RemovingObjectsForm";
        ((System.ComponentModel.ISupportInitialize)(this.gridControl1)).EndInit();
        ((System.ComponentModel.ISupportInitialize)(this.gridView1)).EndInit();
        this.ResumeLayout(false);
        this.PerformLayout();

    }

    #endregion

    private DevExpress.XtraEditors.SimpleButton _okButton;
    private DevExpress.XtraEditors.SimpleButton _cancelButton;
    private DevExpress.XtraGrid.GridControl gridControl1;
    private DevExpress.XtraGrid.Views.Grid.GridView gridView1;
    private DevExpress.XtraEditors.LabelControl labelControl1;
    private DevExpress.XtraGrid.Columns.GridColumn _nameGridColumn;
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
                using (TestForm form = new TestForm(views[0].Reference.ParameterGroup.ReferenceInfo))
                {
                    if (form.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                    {
                        int count = 0;
                        //только в этом случае возможно потеря данных
                        Reference mainReference = views[0].Reference;

                        if (form.LinkGroup1.IsLinkToOne)
                            mainReference.LoadSettings.GetLinkLoadSettings(form.LinkGroup1);
                        if (form.LinkGroup2.IsLinkToOne)
                            mainReference.LoadSettings.GetLinkLoadSettings(form.LinkGroup2);

                        Reference reference = form.LinkGroup1.SlaveGroup.ReferenceInfo.CreateReference();
                        var referenceObjects = mainReference.Objects.Where(o => (form.LinkGroup1.IsLinkToOne && o.Links.ToOne.LinkGroups.Contains(form.LinkGroup1) ||
                                                                                 form.LinkGroup1.IsLinkToMany && o.Links.ToMany.LinkGroups.Contains(form.LinkGroup1)) &&
                                                                                (form.LinkGroup2.IsLinkToOne && o.Links.ToOne.LinkGroups.Contains(form.LinkGroup2) ||
                                                                                 form.LinkGroup2.IsLinkToMany && o.Links.ToMany.LinkGroups.Contains(form.LinkGroup2))).ToArray();

                        if (referenceObjects.Length == 0)
                            return;

                        if (form.LinkGroup2.IsLinkToOne)
                        {
                            var objects = referenceObjects.Select(o => o.GetObject(form.LinkGroup2)).Where(ro => ro != null).ToArray();
                            bool canRemove = true;
                            if (objects.Length > 0)
                                using (RemovingObjectsForm rform = new RemovingObjectsForm(objects))
                                    if (rform.ShowDialog(uiContext.OwnerWindow) != System.Windows.Forms.DialogResult.OK)
                                        canRemove = false;

                            if (canRemove)
                                foreach (var obj in referenceObjects)
                                {
                                        var ro = obj.GetObject(form.LinkGroup1);
                                        if (ro != null)
                                        {
                                            //var res = reference.Find(ro.SystemFields.Guid);

                                            if (ro != null)
                                            {
                                                obj.BeginChanges();
                                                    obj.SetLinkedObject(form.LinkGroup2, ro);
                                                    obj.SetLinkedObject(form.LinkGroup1, null);
                                                    obj.EndChanges();
                                                    count++;
                                            }
                                        }
                                }
                        }
                        else
                        {
                            foreach (var obj in referenceObjects)
                            {
                                obj.BeginChanges();
                                    if (form.LinkGroup1.IsLinkToOne)
                                    {
                                        var ro = obj.GetObject(form.LinkGroup1);
                                        if (ro != null)
                                        {
                                            obj.AddLinkedObject(form.LinkGroup2, ro);
                                            obj.SetLinkedObject(form.LinkGroup1, null);
                                            count++;
                                        }

                                    }
                                    else
                                    {
                                        foreach (var o in obj.GetObjects(form.LinkGroup1))
                                        {
                                            obj.AddLinkedObject(form.LinkGroup2, o);
                                            obj.RemoveLinkedObject(form.LinkGroup1, o);
                                            count++;
                                        }
                                        obj.ClearLinks(form.LinkGroup1);
                                    }
                                    obj.EndChanges();
                            }
                        }

                        MessageDialog.MessageInformationShow(uiContext.OwnerWindow, string.Concat("Замена связей завершена. Заменено объектов: ", count));
                    }
                }
            }
        }
    }
}
