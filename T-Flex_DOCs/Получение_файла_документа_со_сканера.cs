/*
Список ссылок
System.Drawing.dll
System.Windows.Forms.dll
*/

using System;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using TFlex.DOCs.Model.Macros;
using TFlex.DOCs.Model.References;
using TFlex.DOCs.Model.References.Files;

public partial class DocumentScan : MacroProvider
{
    public DocumentScan(MacroContext context)
        : base(context)
    {
    }

    private static readonly Guid bmpClassGuid = new Guid("57b69cc9-07bd-4f1e-81fa-dd7422e2bc02");
    private static readonly Guid jpgClassGuid = new Guid("9e3337ed-a1fc-477e-be0d-a485e4e5370d");
    private static readonly Guid pngClassGuid = new Guid("b2aebbf9-84e3-4c03-8b2e-8ed7b9717cce");
    private static readonly Guid tiffClassGuid = new Guid("b5832684-994a-449b-93cb-f07f036304f9");
    private static readonly Guid pdfClassGuid = new Guid("58e7a26a-cf5f-445b-b08a-885f1bcf7f12");
    private static readonly Guid DocFileLinkID = new Guid("9eda1479-c00d-4d28-bd8e-0d592c873303");
    private static readonly Guid DocNameID = new Guid("d35e460c-5a0a-4f6a-8385-695ff5ad1119");
    private static readonly Guid DocDenotationID = new Guid("b8992281-a2c3-42dc-81ac-884f252bd062");

    public void Scan()
    {
        FileReference fReference = new FileReference(Context.Connection);
        if (fReference == null)
            throw new MacroException("Не найден справочник файлов");

        var диалогВыбораПапки = СоздатьДиалогВыбораОбъектов("Файлы");
        диалогВыбораПапки.Фильтр = "[Тип] = 'Папка'";
        if (!диалогВыбораПапки.Показать())
            return;

        FolderObject destFolder = ((ReferenceObject)диалогВыбораПапки.ФокусированныйОбъект) as FolderObject;
        if (destFolder == null)
        {
            MessageBox.Show("Необходимо выбрать папку");
            return;
        }

        ReferenceObject document = Context.ReferenceObject;
        string docName = System.Text.RegularExpressions.Regex.Replace(document.GetObjectValue("Наименование").ToString(), @"[\\|/:\*\?\""<>]+", " ");
        string docDenotation = System.Text.RegularExpressions.Regex.Replace(document.GetObjectValue("Обозначение").ToString(), @"[\\|/:\*\?\""<>]+", " ");
        InputNameDialog inputDialog = new InputNameDialog { FileName = GetUniqueDocName(destFolder, docName + "-" + docDenotation) };
        if (inputDialog.ShowDialog() != DialogResult.OK)
            return;

        if (String.IsNullOrWhiteSpace(inputDialog.FileName))
        {
            MessageBox.Show("Необходимо задать имя файла");
            return;
        }

        string path = Path.GetDirectoryName(typeof(MacroContext).Assembly.Location);
        string fileName = Path.Combine(path, "TFlex.DOCs.DocumentScan.exe");
        if (!File.Exists(fileName))
        {
            MessageBox.Show(String.Format("Файл '{0}' не найден", fileName));
            return;
        }

        string outputFileName = Path.ChangeExtension(Path.Combine(Path.GetTempPath(), Path.GetRandomFileName()), inputDialog.FileExtension);

        System.Diagnostics.ProcessStartInfo info = new System.Diagnostics.ProcessStartInfo
        {
            FileName = fileName,
            Arguments = String.Format("\"{0}\"", outputFileName),
            CreateNoWindow = true,
            WindowStyle = System.Diagnostics.ProcessWindowStyle.Hidden,
            UseShellExecute = false
        };

        using (var process = System.Diagnostics.Process.Start(info))
        {
            process.WaitForExit();
            if (process.ExitCode != 0)
                return;
        }

        Guid imageClassGuid;

        switch (inputDialog.FileExtension)
        {
            case ".pdf":
                imageClassGuid = pdfClassGuid;
                break;
            case ".jpg":
                imageClassGuid = jpgClassGuid;
                break;
            case ".bmp":
                imageClassGuid = bmpClassGuid;
                break;
            case ".tiff":
                imageClassGuid = tiffClassGuid;
                break;
            case ".png":
                imageClassGuid = pngClassGuid;
                break;
            default:
                imageClassGuid = jpgClassGuid;
                break;
        }

        FileType fileType = fReference.Classes.Find(imageClassGuid);
        if (fileType == null)
        {
            MessageBox.Show("Не найден тип файла.");
            return;
        }

        FileObject fileObject = destFolder.CreateFile(outputFileName, "Сканирование документа", GetUniqueDocName(destFolder, inputDialog.FileName, inputDialog.FileExtension) + inputDialog.FileExtension, fileType);
        if (fileObject == null)
            return;

        document.AddLinkedObject(DocFileLinkID, fileObject);

        FileInfo tmpFile = new FileInfo(outputFileName);
        if (tmpFile.Exists)
            tmpFile.Delete();
    }

    private string GetUniqueDocName(FolderObject folderObject, string defaultName = "Сканирование", string ext = ".jpg")
    {
        int index = defaultName.LastIndexOf(ext);
        if (index >= 0)
            defaultName = defaultName.Remove(index);
        string result = defaultName;
        int i = 1;
        while (folderObject.Children.AsList.Any(file => string.Compare(file.Name, result + ext, true) == 0))
        {
            result = String.Format("{0} ({1})", defaultName, i);
            i++;
        }
        return result;
    }
}

public class InputNameDialog : System.Windows.Forms.Form
{
    public InputNameDialog()
    {
        InitializeComponent();
    }

    private void cancelButton2_Click(object sender, EventArgs e)
    {
        DialogResult = System.Windows.Forms.DialogResult.Cancel;
    }

    private void okButton1_Click(object sender, EventArgs e)
    {
        DialogResult = System.Windows.Forms.DialogResult.OK;
    }

    public string FileName
    {
        get
        {
            return textEdit1.Text;
        }
        set
        {
            textEdit1.Text = value;
        }
    }

    public string FileExtension
    {
        get
        {
            return "." + comboBox1.Text.ToLower();
        }
    }
    /// <summary>
    /// Required designer variable.
    /// </summary>
    private System.ComponentModel.IContainer components = null;

    /// <summary>
    /// Clean up any resources being used.
    /// </summary>
    /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
    protected override void Dispose(bool disposing)
    {
        if (disposing && (components != null))
        {
            components.Dispose();
        }
        base.Dispose(disposing);
    }

    #region Windows Form Designer generated code

    /// <summary>
    /// Required method for Designer support - do not modify
    /// the contents of this method with the code editor.
    /// </summary>
    private void InitializeComponent()
    {
        this.textEdit1 = new System.Windows.Forms.TextBox();
        this.okButton1 = new System.Windows.Forms.Button();
        this.cancelButton2 = new System.Windows.Forms.Button();
        this.comboBox1 = new System.Windows.Forms.ComboBox();
        this.label1 = new System.Windows.Forms.Label();
        this.SuspendLayout();
        // 
        // textEdit1
        // 
        this.textEdit1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
                    | System.Windows.Forms.AnchorStyles.Right)));
        this.textEdit1.Location = new System.Drawing.Point(14, 14);
        this.textEdit1.Margin = new System.Windows.Forms.Padding(5);
        this.textEdit1.Name = "textEdit1";
        this.textEdit1.Size = new System.Drawing.Size(268, 20);
        this.textEdit1.TabIndex = 0;
        // 
        // okButton1
        // 
        this.okButton1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
        this.okButton1.Location = new System.Drawing.Point(126, 72);
        this.okButton1.Name = "okButton1";
        this.okButton1.Size = new System.Drawing.Size(75, 23);
        this.okButton1.TabIndex = 1;
        this.okButton1.Text = "OK";
        this.okButton1.Click += new System.EventHandler(this.okButton1_Click);
        // 
        // cancelButton2
        // 
        this.cancelButton2.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
        this.cancelButton2.DialogResult = System.Windows.Forms.DialogResult.Cancel;
        this.cancelButton2.Location = new System.Drawing.Point(207, 72);
        this.cancelButton2.Name = "cancelButton2";
        this.cancelButton2.Size = new System.Drawing.Size(75, 23);
        this.cancelButton2.TabIndex = 1;
        this.cancelButton2.Text = "Отмена";
        this.cancelButton2.Click += new System.EventHandler(this.cancelButton2_Click);
        // 
        // comboBox1
        // 
        this.comboBox1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
        this.comboBox1.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
        this.comboBox1.FormattingEnabled = true;
        this.comboBox1.Items.AddRange(new object[] {
            "JPG",
            "BMP",
            "PNG",
            "TIFF",
            "PDF"});
        this.comboBox1.Location = new System.Drawing.Point(226, 42);
        this.comboBox1.Name = "comboBox1";
        this.comboBox1.Size = new System.Drawing.Size(56, 21);
        this.comboBox1.SelectedIndex = 0;
        this.comboBox1.TabIndex = 2;
        // 
        // label1
        // 
        this.label1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
        this.label1.AutoSize = true;
        this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
        this.label1.Location = new System.Drawing.Point(87, 45);
        this.label1.Name = "label1";
        this.label1.Size = new System.Drawing.Size(120, 13);
        this.label1.TabIndex = 3;
        this.label1.Text = "Формат изображения";
        // 
        // InputNameDialog
        // 
        this.AcceptButton = this.okButton1;
        this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
        this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
        this.CancelButton = this.cancelButton2;
        this.ClientSize = new System.Drawing.Size(294, 107);
        this.Controls.Add(this.label1);
        this.Controls.Add(this.comboBox1);
        this.Controls.Add(this.cancelButton2);
        this.Controls.Add(this.okButton1);
        this.Controls.Add(this.textEdit1);
        this.MinimumSize = new System.Drawing.Size(310, 145);
        this.Name = "InputNameDialog";
        this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
        this.Text = "Имя документа";
        this.ResumeLayout(false);
        this.PerformLayout();

    }

    #endregion

    private System.Windows.Forms.TextBox textEdit1;
    private System.Windows.Forms.Button okButton1;
    private System.Windows.Forms.Button cancelButton2;
    private System.Windows.Forms.ComboBox comboBox1;
    private System.Windows.Forms.Label label1;
}

