using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using TFlex.DOCs.Model.Macros;
using TFlex.DOCs.Model.Macros.ObjectModel;
using System.Diagnostics;

public class Macro : MacroProvider
{
    public Macro(MacroContext context)
        : base(context)
    {
    }

    public override void Run()
    {
        System.Diagnostics.Process p = new System.Diagnostics.Process();
        
        string fullpath =  System.Reflection.Assembly.GetEntryAssembly().Location.ToString();
    	int pos = fullpath.LastIndexOf('\\') + 1;
    	string path = fullpath.Substring(0, pos);
        
        p.StartInfo.FileName = path + "MSProjectImportWizard.exe";
        p.Start();
    }
}
