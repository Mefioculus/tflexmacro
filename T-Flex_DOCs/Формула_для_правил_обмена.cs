using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using TFlex.DOCs.Model.Macros;
using TFlex.DOCs.Model.Macros.ObjectModel;

public class Macro : MacroProvider
{
    public Macro(MacroContext context)
        : base(context)
    {
    }

    public override void Run()
    {

    }

    public string УдалитьТочки(string valueString)
    {
        return valueString.Replace(".", String.Empty);
    }

    public string ВставитьТочки(string valueString)
    {
        int num;
        bool isNum = int.TryParse(valueString, out num);
        if (!isNum && valueString.IndexOf('-') > 0 && (valueString.Substring(0, valueString.IndexOf('-')).Length == 6 || valueString.Substring(0, valueString.IndexOf('-')).Length == 7))
        {
            isNum = int.TryParse(valueString.Substring(0, valueString.IndexOf('-')), out num);
            valueString = valueString.Insert(3, ".");
        }

        if (valueString.StartsWith("УЯИС") || valueString.StartsWith("ШЖИФ") || valueString.StartsWith("УЖИЯ"))
        {
            // Изменяем значение
            valueString = valueString.Insert(4, ".").Insert(11, ".");
        }

        if (valueString.StartsWith("8А"))
        {
            // Изменяем значение
            valueString = valueString.Insert(3, ".").Insert(7, ".");
        }

        if (isNum && (valueString.Length == 6 || valueString.Length == 7))
        {
            valueString = valueString.Insert(3, ".");
        }
        return valueString;
    }
}
