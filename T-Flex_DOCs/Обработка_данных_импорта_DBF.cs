/*
TFlex.DOCs.SynchronizerReference.dll
*/

using TFlex.DOCs.Model.Macros;
using TFlex.DOCs.Synchronization.Database;
using TFlex.DOCs.Synchronization.Macros;

public class Macro_KZ_BDF : MacroProvider
{
    public Macro_KZ_BDF(MacroContext context) : base(context) { }

    public void СоставИзделий_Начало()
    {
        var context = Context as ExchangeDataMacroContext;

        if (context == null)
            return;

        var request = context.Settings.ExternalData as DatabaseRequest;

        if (request == null)
            return;

        // Выбираем от какой строки читать и сколько строк читаем
        /*
        var диалог = СоздатьДиалогВвода("Выберите от какой до какой строки импортировать данные");
        диалог.ДобавитьТекст("Если значения = 0, то будут импортироваться все данные", 1);
        диалог.ДобавитьЦелое("От строки");
        диалог.ДобавитьЦелое("Количество импортируемых строк");

        if (!диалог.Показать())
            Отменить();

        int offset = диалог["От строки"];
        int top = диалог["Количество импортируемых строк"];
        */
       
       
        int offset = 0;
        int top = 0;
        var option = request.GetOrAddQueryOption();
        option.Offset = offset;
        option.Top = top;

        // Сколько строк в пакете данных
        request.PacketSize = 1000; // По умолчанию 25000
    }

}
