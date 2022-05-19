using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using TFlex.DOCs.Model.Macros;
using TFlex.DOCs.Model.Macros.ObjectModel;
using TFlex.Technology;
using TFlex.DOCs.Model;
using TFlex.DOCs.Model.References;
using TFlex.Technology.References;
using TFlex.Technology.UI.Dialogs;
using TFlex.DOCs.Model.Classes;
using TFlex.DOCs.Model.References.Links;
using TFlex.DOCs.UI;
using TFlex.Technology.UI.Commands;

public class Macro : MacroProvider
{
    public Macro(MacroContext context)
        : base(context)
    {
    }

    public override void Run()
    {
        DSEConnector.Deploy(((TFlex.DOCs.UI.Objects.Managers.UIMacroContext)Context).OwnerWindow, Context.ReferenceObject);
    }
}
