using System;
using System.Linq;
using TFlex.DOCs.Model.Macros;
using TFlex.DOCs.Model.References;
using TFlex.DOCs.Model.References.Links;

public class Macro : MacroProvider
{
	private readonly Guid FilesLinkGuid = new Guid("9eda1479-c00d-4d28-bd8e-0d592c873303");
	
    public Macro(MacroContext context)
        : base(context)
    {
    }

    public void СоздатьОбъект()
    {
        ReferenceObject obj = Context.ReferenceObject;
        if (obj == null || obj.Prototype != null)
              return;

        string name = obj.Class.Attributes.GetValue<string>("InitialName");
        if(name != null && name.Length > 0)
        {
            Parameter["Наименование"] = name;
        }
    } 
}
