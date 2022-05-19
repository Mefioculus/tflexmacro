using System;
using TFlex.DOCs.Model.Macros;
using TFlex.DOCs.Model.Macros.ObjectModel;
using TFlex.DOCs.Model.References;

public class Macro : MacroProvider
{
    public Macro(MacroContext context)
        : base(context)
    {
    }

    public void Сохранение()
    {

         ReferenceObject obj = Context.ReferenceObject;
         if(obj == null)
           return;
        ReferenceObject company = obj.Links.ToOne[new Guid("05360587-0dc9-4667-aa0c-22c78e6eb6b6")].LinkedObject;
        if(company == null)
           throw new Exception("Не задана организация");
       //Копируем название организации в название поставщика
       obj.ParameterValues[new Guid("190763c7-c09f-4d3b-a8eb-37da04f262fe")].Value = company.ParameterValues[new Guid("e434c150-94e4-4cb6-a1de-9fc8b6408380")].GetString();
    }
}
