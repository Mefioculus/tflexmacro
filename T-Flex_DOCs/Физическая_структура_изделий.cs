using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using TFlex.DOCs.Model.Macros;
using TFlex.DOCs.Model.Macros.ObjectModel;
using TFlex.DOCs.References.PhysicalStructure;
using TFlex.DOCs.Model.References.Nomenclature;

public class Macro : MacroProvider
{
    public Macro(MacroContext context)
        : base(context)
    {
    }

    public override void Run()
    {
    }
    
    public void СоздатьСтруктуру()
    {
    	var dse = Context.ReferenceObject as NomenclatureObject;
        if(dse == null)
        	return;

        var reference = new PhysicalStructureReference(dse.Reference.Connection);

        PhysicalStructureObject newPhysicalStructureObject = reference.CreateReferenceObject(reference.Classes.PhysicalStructure) as PhysicalStructureObject;
        newPhysicalStructureObject.CreateSaveSet(); 
        newPhysicalStructureObject.SetProduct(dse);
        
        CreateChildElements(newPhysicalStructureObject, dse, null);
       
        newPhysicalStructureObject.EndChanges();
        
        Context.OpenReferenceWindow(reference, rootObject: newPhysicalStructureObject);
    }

    private void CreateChildElements(PhysicalStructureReferenceObject parent, NomenclatureReferenceObject dse, NomenclatureHierarchyLink linkToParent)
    {
        StructureElementObject newStructureElementObject = parent.Reference.CreateReferenceObject(parent, parent.Reference.Classes.StructureElement) as StructureElementObject;
        newStructureElementObject.SetProductProperties(dse, linkToParent);
        parent.SaveSet.Add(newStructureElementObject);

        foreach (NomenclatureHierarchyLink lnk in dse.Children.GetHierarchyLinks())
        {
            CreateChildElements(newStructureElementObject, lnk.ChildObject as NomenclatureReferenceObject, lnk);
        }
    }
}
