using System;
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using TFlex.DOCs.Model.Macros;
using TFlex.DOCs.Model.Macros.ObjectModel;
using TFlex.DOCs.Model.References;
using TFlex.DOCs.Model.References.Nomenclature;

public class Macro : MacroProvider
{
    public Macro(MacroContext context)
        : base(context)
    {
        //if (Вопрос("Хотите запустить в режиме отладки?"))
        //{
        //System.Diagnostics.Debugger.Launch();
        //System.Diagnostics.Debugger.Break();
        //}
    }
    private Reference _refЭкземплярыИзделий;
    private static readonly Guid _guidRefЭкземплярыИзделий = new Guid("b8245afd-eb33-4eb8-a39a-e2699617fbee");
    private static readonly Guid _guidStrСерийныйНомер = new Guid("d8bc3a3f-030a-46ff-a10b-75e7fcde2610");
    private static readonly Guid _guidLnkИзделие = new Guid("bbe0955d-6334-4161-8045-c8c89d22e3d5");

    public override void Run()
    {
        _refЭкземплярыИзделий = Context.Connection.ReferenceCatalog.Find(_guidRefЭкземплярыИзделий)?.CreateReference() ?? throw new MacroException("Не найден Справочник \"Экземпляры изделий\"");
        ReferenceObject curObj = (ReferenceObject)CurrentObject;
        if (curObj == null) return;

        string диапазон = InputDialog.Show();
        if (string.IsNullOrWhiteSpace(диапазон)) return;

        CreateDiapazon(диапазон, curObj);
    }

    private List<ReferenceObject> CreateDiapazon(string diapazon, ReferenceObject curObj)
    {
        List<ReferenceObject> result = new List<ReferenceObject>();

        foreach (string diap in diapazon.Split(','))
        {
            if (string.IsNullOrWhiteSpace(diap)) continue;

            List<ReferenceObject> diapRef = null;
            if (diap.Contains('-'))
            {
                string[] diapStartEnd = diap.Split('-');
                if (int.TryParse(diapStartEnd[0], out int start) && int.TryParse(diapStartEnd[1], out int end)) diapRef = CreateCopiesOfProuctObjs(start, end, curObj);
            }
            else
            {
                if (int.TryParse(diap, out int strtEnd)) diapRef = CreateCopiesOfProuctObjs(strtEnd, strtEnd, curObj);
            }

            if (diapRef == null) continue;
            else result.AddRange(diapRef);
        }

        return result;
    }

    private List<ReferenceObject> CreateCopiesOfProuctObjs(int start, int end, ReferenceObject curObj)
    {
        if (start == 0 || end == 0 || start > end) return null;

        List<ReferenceObject> copies = new List<ReferenceObject>();
        for (int i = start; i <= end; i++)
        {
            copies.Add(CreateCopyOfProuctObj(i, curObj));
        }

        return copies;
    }

    private ReferenceObject CreateCopyOfProuctObj(int valСерийныйНомер, ReferenceObject curObj)
    {
        ReferenceObject copyOfProduct = _refЭкземплярыИзделий.CreateReferenceObject();
        copyOfProduct[_guidStrСерийныйНомер].Value = valСерийныйНомер.ToString("000000");
        copyOfProduct.SetLinkedObject(_guidLnkИзделие, curObj);
        copyOfProduct.EndChanges();
        return copyOfProduct;
    }
}

public static class InputDialog
{
    private static string _backup;
    private static readonly Regex _numsAndSeprsRegex = new Regex("[^0-9,-]+");

    public static string Show()
    {
        Form prompt = new Form()
        {
            Width = 330,
            Height = 133,
            MinimumSize = new Size(250, 125),
            Text = "Введите значения",
            StartPosition = FormStartPosition.CenterScreen,
            TopMost = true
        };

        TreeView treeView = new TreeView
        {
            CheckBoxes = true,
            Location = new System.Drawing.Point(0, 0),
            Name = "treeView1",
            Size = new System.Drawing.Size(580, 318),
            TabIndex = 0
        };

        Label textLabel = new Label() { Left = 15, Top = 20, Width = 60, Text = "Диапазон:" };

        TextBox textBox = new TextBox()
        {
            Left = 82,
            Top = 20,
            Width = 217,
            Anchor = AnchorStyles.Right | AnchorStyles.Top | AnchorStyles.Left
        };

        textBox.TextChanged += (sender, e) => { TextBox_TextChanged(sender, e); };
        textBox.KeyPress += (sender, e) => { TextBox_KeyPress(sender, e); };

        Button confirmation = new Button()
        {
            Text = "ОК",
            Left = 148,
            Top = 60,
            Width = 72,
            Height = 20,
            DialogResult = DialogResult.OK,
            Anchor = AnchorStyles.Right | AnchorStyles.Bottom
        };

        confirmation.Click += (sender, e) => { prompt.Close(); };

        Button cancel = new Button()
        {
            Text = "Отмена",
            Left = 228,
            Top = 60,
            Width = 72,
            Height = 20,
            DialogResult = DialogResult.OK,
            Anchor = AnchorStyles.Right | AnchorStyles.Bottom
        };

        confirmation.Click += (sender, e) => { prompt.Close(); };

        prompt.Controls.Add(confirmation);
        prompt.Controls.Add(cancel);
        prompt.Controls.Add(textLabel);
        prompt.Controls.Add(textBox);
        prompt.AcceptButton = confirmation;

        return prompt.ShowDialog() == DialogResult.OK ? textBox.Text : "";
    }

    private static void TextBox_KeyPress(object sender, KeyPressEventArgs e)
    {
        if (_numsAndSeprsRegex.IsMatch(e.ToString()))
        {
            _backup = ((TextBox)sender).Text;
        }
        else
        {
            e.Handled = true;
        }
    }

    private static void TextBox_TextChanged(object sender, EventArgs e)
    {
        Regex valRegex = new Regex(@"^\d{0,6}\s?(?:-\s?\d{0,6})?(?:,\s?\d{0,6}\s?(?:-\s?\d{0,6})?)*$");//new Regex(@"^\d{0,3}(?:-\d{0,3})?(?:,\d{0,3}(?:-\d{0,3})?)*$");
        Regex emptyRegex = new Regex(@"\s*");

        if (!valRegex.IsMatch(((TextBox)sender).Text) || !emptyRegex.IsMatch(((TextBox)sender).Text))
        {
            ((TextBox)sender).Text = _backup;
            ((TextBox)sender).SelectionStart = ((TextBox)sender).Text.Length;
        }
    }
}

