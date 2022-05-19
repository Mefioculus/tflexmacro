/*
TFlex.DOCs.UI.Client.dll
*/

using System;
using TFlex.DOCs.Model.Macros;
using TFlex.DOCs.Model.References;

public class AsuNsiARMStatic : MacroProvider
{
    public static Guid NameReference = Guid.Empty;

    private readonly string ИсточникиДанных = "7e7be674-84c8-45d2-bed4-f58376cddd50";
    private readonly string ИсточникиДанныхПараметрАналог = "bd2f577a-874f-4376-9e2e-e4b06e9f3123";

    public AsuNsiARMStatic(MacroContext context)
        : base(context)
    {
        /*if (Вопрос("Хотите запустить в режиме отладки?"))
        {
            System.Diagnostics.Debugger.Launch();
            System.Diagnostics.Debugger.Break();
        }*/
    }

    public override void Run()
    {

    }
    public string ПолучитьСправочник()
    {
        return ГлобальныйПараметр["Справочник"];
    }

    public Guid GetReference()
    {
        return NameReference;
    }

    public void ИзменитьГлобальныйПараметр()
    {
        if (Вопрос("Хотите запустить в режиме отладки?"))
        {
            System.Diagnostics.Debugger.Launch();
            System.Diagnostics.Debugger.Break();
        }
        var диалог = СоздатьДиалогВвода("Выбор");
        диалог.ДобавитьСтроковое("Справочник");
        if (!диалог.Показать())
            return;

        ГлобальныйПараметр["Справочник"] = диалог["Справочник"];
        //Сообщение("", NameReference.ToString());
        
        ОбновитьЭлементыУправления("item13");
    }
}

