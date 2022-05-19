using System;
using System.Linq;
using TFlex.DOCs.Model.Macros;
using TFlex.DOCs.Model.References.Remarks;
using TFlex.DOCs.Model.References;

public class MacroTemplate : MacroProvider
{
    public MacroTemplate(MacroContext context)
        : base(context)
    {
    }

    // Устанавливает статус "Принято"
    public void УстановитьСтатусПринято()
    {
        ChangeStatusForSelectedRemarks(RemarkStatus.Accepted, false);
    }

    // Устанавливает статус "Выполнено"
    public void УстановитьСтатусВыполнено()
    {
        ChangeStatusForSelectedRemarks(RemarkStatus.Completed, true);
    }

    // Устанавливает статус "Отменено"
    public void УстановитьСтатусОтменено()
    {
        ChangeStatusForSelectedRemarks(RemarkStatus.Cancel, true);
    }

    // Устанавливает статус "Отклонено"
    public void УстановитьСтатусОтклонено()
    {
        ChangeStatusForSelectedRemarks(RemarkStatus.Rejected, true);
    }

    private void ChangeStatusForSelectedRemarks(RemarkStatus newStatus, bool setClosingAuthor)
    {
        var selectedObjects = Context.GetSelectedObjects()
            .OfType<RemarkReferenceObject>()
            .Where(so => so.CanEdit)
            .ToArray();

        if (selectedObjects.Length == 0)
            return;

        foreach (var selectedObject in selectedObjects)
        {
           ChangeStatusForRemark(newStatus, setClosingAuthor, selectedObject);
        }

        Context.RefreshReferenceObjects(selectedObjects);
    }
    
    private void ChangeStatusForRemark(RemarkStatus newStatus, bool setClosingAuthor, RemarkReferenceObject remark) 
    {
        remark.BeginChanges();
        
        try
        {
            remark.StatusType = newStatus;
            remark.ClosingDate.Value = DateTime.Now.Date;
            if (setClosingAuthor)
            	remark.ClosingAuthor.Value = ТекущийПользователь["Наименование"];

            remark.EndChanges();
        }
        catch
        {
        	if (remark.Changing)
        		remark.CancelChanges();
        	
        	throw;
        }
    }
    
    public void OпубликоватьЗамечание()
    {
    	var selectedRemarks = Context.GetSelectedObjects()
            .OfType<RemarkReferenceObject>()            
            .ToArray();

        if (selectedRemarks.Length == 0)
            return;
        
        foreach (var remark in selectedRemarks.Where(r => r is not RequestRemarkObject && r.Parent is RequestRemarkObject))
        {
        	var parent = remark.Parent.Parent;
        	OпубликоватьЗамечание(remark, parent);
        }
        
        foreach (var request in selectedRemarks.OfType<RequestRemarkObject>())
        {
        	var parent = request.Parent;
            ОпубликоватьВесьЗапрос(request, parent);
        }
        
        ОбновитьОкноСправочника();
    }
    
    private void OпубликоватьЗамечание(RemarkReferenceObject remark, ReferenceObject parent)
    {  
        if (!remark.CanEdit)
            return;
        
        var remarkCopy = remark.CreateCopy(remark.Class, parent); 
        remarkCopy.EndChanges();
        
        ChangeStatusForRemark(RemarkStatus.Completed, true, remark);
    }
    
    private void ОпубликоватьВесьЗапрос(RequestRemarkObject request, ReferenceObject parent)
    { 
    	foreach (var children in request.Children.OfType<RemarkReferenceObject>())
    	{
    		if (children is RequestRemarkObject childRequest)    			
    		{
    			ОпубликоватьВесьЗапрос(childRequest, parent);
    			continue;
    		}
    		
    		OпубликоватьЗамечание(children, parent);
    	}
    }
}
