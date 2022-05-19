using System;
using TFlex.DOCs.Model;
using TFlex.DOCs.Model.Macros;
using TFlex.DOCs.Model.References;
using TFlex.DOCs.Model.Classes;

public class Macro : MacroProvider
{
    public Macro(MacroContext context)
        : base(context)
    {
    }

    public override void Run()
    {
      
    }

    public void Сохранение()
    {
        ReferenceObject obj = Context.ReferenceObject;
        string address = "";

        string index = Параметр["Почтовый индекс"].ToString();
        if(index.Length > 0)
             address += index + ", ";

        ReferenceObject city = obj.Links.ToOne[new Guid("e869437d-636d-40f5-bc84-fbebfb8f0899")].LinkedObject;
        if(city != null)
        {
             string CityName = city.ParameterValues[new Guid("615379d1-8663-4617-85d2-268ecbdc85dc")].GetString();
             if(CityName.Length > 0)
             {
                  address += "г. " + CityName + ", ";
             }
        }
        
        string adr = Параметр["Адрес"].ToString();
        if(adr.Length > 0)
             address += adr;
       
        Параметр["Полный адрес"] = address;

        ReferenceObject Поставщик = obj.Links.ToOne[new Guid("05360587-0dc9-4667-aa0c-22c78e6eb6b6")].LinkedObject;
        if(Поставщик != null)
       {
            String НазваниеОрганизации = obj.ParameterValues[new Guid("e434c150-94e4-4cb6-a1de-9fc8b6408380")].GetString();
            Поставщик.BeginChanges();
            Поставщик.ParameterValues[new Guid("190763c7-c09f-4d3b-a8eb-37da04f262fe")].Value = НазваниеОрганизации;
            Поставщик.EndChanges();
       }
    }

    public void СоздатьПараметрыПоставщика()
    {
        ReferenceObject obj = Context.ReferenceObject;
        ReferenceObject Поставщик = obj.Links.ToOne[new Guid("05360587-0dc9-4667-aa0c-22c78e6eb6b6")].LinkedObject;
        if(Поставщик != null)
            return;

       String НазваниеОрганизации = obj.ParameterValues[new Guid("e434c150-94e4-4cb6-a1de-9fc8b6408380")].GetString();
       if(НазваниеОрганизации.Length == 0)
       {
            System.Windows.Forms.MessageBox.Show("Не задано название организации");
            return;
       }

        ReferenceInfo info = ReferenceCatalog.FindReference(new Guid("ad3f4e41-959f-41ea-8b07-b90384119491"));
        if(info == null)
           return;
        Reference Поставщики = info.CreateReference();
        if (Поставщики == null)
           return;
       Поставщик = Поставщики.CreateReferenceObject();
       //Поставщик.Links.ToOne[new Guid("05360587-0dc9-4667-aa0c-22c78e6eb6b6")].SetLinkedObject(obj);
       Поставщик.ParameterValues[new Guid("190763c7-c09f-4d3b-a8eb-37da04f262fe")].Value = НазваниеОрганизации;
       Поставщик.EndChanges();
       obj.Links.ToOne[new Guid("05360587-0dc9-4667-aa0c-22c78e6eb6b6")].SetLinkedObject(Поставщик);
    }

    public void УдалитьПараметрыПоставщика()
    {
        ReferenceObject obj = Context.ReferenceObject;
        ReferenceObject Поставщик = obj.Links.ToOne[new Guid("05360587-0dc9-4667-aa0c-22c78e6eb6b6")].LinkedObject;
        if(Поставщик == null)
            return;
        obj.Links.ToOne[new Guid("05360587-0dc9-4667-aa0c-22c78e6eb6b6")].SetLinkedObject(null);
        //Поставщик.BeginChanges();
        Поставщик.Delete();
        //Поставщик.EndChanges();
    }

}


