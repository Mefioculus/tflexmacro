using System;
using TFlex.DOCs.Macro.NormativeReferenceInfo.Handlers.DocumentsNsi;
using TFlex.DOCs.Model.Macros;
using TFlex.DOCs.Model.References;


public class EventHandlerForDocumentNsiMacro : MacroProvider
{
    public EventHandlerForDocumentNsiMacro(MacroContext context)
        : base(context)
    {
    }

    /// <summary>
    /// GUID стадии "Утверждено"
    /// </summary>
    private static readonly Guid ApprovedGuid = new Guid("a5ea2e1c-d441-42fd-8f92-49840351d6c1");

    public void ИзменениеСтадииОбъекта()
    {
        var args = Context.ObjectChangedArgs as ObjectStageChangingEventArgs;

        if (args != null && args.NewStage.Guid == ApprovedGuid)
        {
            new GeneratorNewNsiReferenceCommand(this).СоздатьСправочник();
        }
    }
}

