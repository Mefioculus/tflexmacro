using System;
using TFlex.DOCs.Model.Macros;

using TFlex.DOCs.Model;
using TFlex.DOCs.Model.References;

public class Macro : MacroProvider
{
    public Macro(MacroContext context)
        : base(context)
    {
    }

    public override void Run()
    {
      
    }
    public void Создание()
    {
            Context.ReferenceObject.Links.OneToOne[new Guid("a04fcffd-0d5f-458c-8ea1-7dc960cf19fd")].SetLinkedObject(ClientView.Current.GetUser());
    }
}
