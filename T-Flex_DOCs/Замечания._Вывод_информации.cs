using System.Collections.Generic;
using System.Linq;
using TFlex.DOCs.Common.Extensions;
using TFlex.DOCs.Model.Macros;
using TFlex.DOCs.Model.References.Remarks;

public class RemarkColumnsOutputMacro : MacroProvider
{
    public RemarkColumnsOutputMacro(MacroContext context)
        : base(context)
    {
    }

    /// <summary>
    /// Получить количество открытых замечаний
    /// </summary>
    /// <returns></returns>
    public int ПолучитьКоличествоОткрытыхЗамечаний()
    {
        var remarks = GetLinkedRemarks(true);
        var openedRemarks = remarks.Count(r => r.StatusType.OneOf(RemarkStatus.New, RemarkStatus.Accepted));

        return openedRemarks;
    }

    /// <summary>
    /// Получить количество замечаний
    /// </summary>
    /// <returns></returns>
    public int ПолучитьКоличествоЗамечаний()
    {
        var remarks = GetLinkedRemarks();
        return remarks.Count;
    }

    /// <summary>
    /// Проверить наличие открытых замечаний
    /// </summary>
    /// <returns></returns>
    public bool ПроверитьНаличиеОткрытыхЗамечаний()
    {
        var remarks = GetLinkedRemarks(true);
        var openedRemarksExistence = remarks.Any(r => r.StatusType.OneOf(RemarkStatus.New, RemarkStatus.Accepted));

        return openedRemarksExistence;
    }

    /// <summary>
    /// Проверить наличие замечаний
    /// </summary>
    /// <returns></returns>
    public bool ПроверитьНаличиеЗамечаний()
    {
        var remarks = GetLinkedRemarks();
        var anyStatusRemarksExistence = remarks.Any();
        return anyStatusRemarksExistence;
    }

    /// <summary>
    /// Получить замечания по объекту
    /// </summary>
    /// <param name="loadStatuses">Указывает, требуется ли загружать статусы вместе с замечаниями</param>
    /// <returns>Замечания</returns>
    private List<RemarkReferenceObject> GetLinkedRemarks(bool loadStatuses = false)
    {
        var currentObject = Context.ReferenceObject;
        if (currentObject is null)
            return new List<RemarkReferenceObject>();

        var remarksReference = new RemarksReference(Context.Connection);
        if (loadStatuses)
            remarksReference.LoadSettings.Add(RemarkReferenceObject.FieldKeys.Status);

        return remarksReference.FindRemarks(currentObject).SkipNulls().ToList();
    }
}
