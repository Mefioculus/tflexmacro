using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using TFlex.DOCs.Model;
using TFlex.DOCs.Model.References;
using TFlex.DOCs.Model.Macros;
using TFlex.DOCs.Model.Macros.ObjectModel;
using TFlex.DOCs.Model.Search;

public class Macro : MacroProvider
{
	public ReferenceInfo connectionsReferenceInfo;
    public Reference connectionsReference;
	
    public Macro(MacroContext context)
        : base(context)
    {
    	//System.Diagnostics.Debugger.Launch();
        //System.Diagnostics.Debugger.Break();
    }

    public override void Run()
    {
        connectionsReferenceInfo = Context.Connection.ReferenceCatalog.Find(new Guid("9dd79ab1-2f40-41f3-bebc-0a8b4f6c249e"));				// справочник "Подключения"
    	connectionsReference = connectionsReferenceInfo.CreateReference();
    	connectionsReference.LoadSettings.AddMasterGroupParameters();
    	
    	ReferenceInfo unitsReferenceInfo = Context.Connection.ReferenceCatalog.Find(new Guid("01c51d4c-e07d-4f31-9346-5697399a09fb"));		// справочник Единицы измерения"
    	Reference unitsReference = unitsReferenceInfo.CreateReference();
    	unitsReference.LoadSettings.AddMasterGroupParameters();
    	
    	var selectedObjects = Context.GetSelectedObjects();
    	int counter = 0;
    	List<ReferenceObject> lst = new List<ReferenceObject>();
    	foreach (var connection in selectedObjects)
    	{
    		Filter filter = new Filter(unitsReferenceInfo);
    		filter.Terms.AddTerm(unitsReference.ParameterGroup[new Guid("f908d814-2a34-4dc2-bd7b-13ef41a5ec24")],
            ComparisonOperator.Equal, (int)connection[new Guid("19d31f8c-06d2-402b-85ee-bda3f5111e8c")].Value);
    		
    		var foundedUnits = unitsReference.Find(filter);
    		if (!foundedUnits.Any())
    			continue;
    		
    		var foundedUnit = foundedUnits.First();
    		
    		connection.BeginChanges();
    		connection[new Guid("d485a313-6228-4bbf-b40e-b29e82adbb68")].Value = foundedUnit[new Guid("57877a1f-7763-4583-952c-bbfd92ed2055")].Value.ToString();
    		lst.Add(connection);
    		if (counter == 1000)
    		{
    			try
            	{
            		Reference.EndChanges(lst);
            	}
            	catch
            	{
            		;
            	}
            	lst = new List<ReferenceObject>();
           		counter = 0;
    		}
    		counter++;
    	}
    	
    	if (counter > 0 && lst.Any())
    	{
    		try
        	{
        		Reference.EndChanges(lst);
        	}
        	catch
        	{
        		;
        	}
        	
        	lst = new List<ReferenceObject>();
    	}
    	
    }
}
