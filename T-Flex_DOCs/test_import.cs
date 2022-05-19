using System;
using System.Collections.Generic;
using TFlex.DOCs.Model;
using TFlex.DOCs.Model.Macros; // Для работы с макроязыком
using TFlex.DOCs.Model.References; // Для работы со справочниками
using TFlex.DOCs.Model.Structure; // Для использования класса ParameterInfo
using TFlex.DOCs.Model.Search;


public class Macro : MacroProvider
{
    public Macro(MacroContext context)
        : base(context)
    {
     //   System.Diagnostics.Debugger.Launch();
     //   System.Diagnostics.Debugger.Break();
    }

    public override void Run()
    {
    runCompareTest();
    }
    
  private static class Tab
    {
        public static class SPEC
        {
            public static readonly string name_refer = "SPEC";
            public static readonly string nameTip = "SPEC";
            public static readonly Guid guid = new Guid("c587b002-3be2-46de-861b-c551ed92c4c1");
            public static readonly Guid ROW_NUMBER = new Guid("b6c7f1e2-d94c-4fde-966b-937a9979813f");
            public static readonly Guid Role_Список_номенклатуры_SPEC=new Guid("fba9104c-14ac-4c6a-a1eb-9823f717a1df");

        
        }


        public static class NORM
        {
           
            public static readonly string name_refer = "NORM";
            public static readonly Guid guid = new Guid("7e59b31a-eab5-4da5-938f-812703046345");
            public static readonly Guid ROW_NUMBER = new Guid("46a081c5-d993-4c60-af71-162bb11c1927");
       
        }

        public static class KLASM
        {
           
            public static readonly string name_refer = "KLASM";
            public static readonly Guid guid = new Guid("e253835e-260e-45f3-8ae8-600a7a8907a2");
            public static readonly Guid ROW_NUMBER = new Guid("9ec4fb11-0be1-4602-931a-18b479d9ac94");
            public static readonly Guid Role_Выгрузка_ведомости_материалов_в_FoxPro_2=new Guid("88c9749b-5f0a-408d-bf62-989ca641672a");
        
        }
     };
    
       public void runCompareTest()
    {


        var rowNum = new List<int> { 1, 2, 3, 4, 5, 6 };
        var klasmobj = GetAnalog(rowNum, Tab.KLASM.guid, Tab.KLASM.ROW_NUMBER);
        var specobj = GetAnalog(rowNum, Tab.SPEC.guid, Tab.SPEC.ROW_NUMBER);

        //ОбменДанными.ИмпортироватьОбъекты("Список номенклатуры SPEC", specobj);
        ОбменДанными.ИмпортироватьОбъекты("fba9104c-14ac-4c6a-a1eb-9823f717a1df", specobj);
        ОбменДанными.ИмпортироватьОбъекты("88c9749b-5f0a-408d-bf62-989ca641672a", klasmobj);
        ОбменДанными.ИмпортироватьОбъекты(Tab.SPEC.Role_Список_номенклатуры_SPEC.ToString(), specobj);
        ОбменДанными.ИмпортироватьОбъекты(Tab.KLASM.Role_Выгрузка_ведомости_материалов_в_FoxPro_2.ToString(), klasmobj);

    }
       
       
       
    public List<ReferenceObject> GetAnalog(List<int> row_number, Guid guidref, Guid parametr)

    {
        ReferenceInfo info = Context.Connection.ReferenceCatalog.Find(guidref);
        Reference reference = info.CreateReference();
        ParameterInfo parameterInfo = reference.ParameterGroup[parametr];
        List<ReferenceObject> result = reference.Find(parameterInfo, ComparisonOperator.IsOneOf, row_number);
        Console.WriteLine("GetAnalog");
        return result;
    }
       
}
