/* Дополнительные ссылки
TFlex.DOCs.Model.Office.dll
*/

using TFlex.DOCs.Model.Macros;
using TFlex.DOCs.Model.References.RecordControlCards;


public class CreateRccTheSameTypeMacro : MacroProvider
{
    public CreateRccTheSameTypeMacro(MacroContext context)
        : base(context)
    {
    }

    public override void Run()
    {
        var rccObject = Context.ReferenceObject as RCCReferenceObject;
        if (rccObject == null)
            return;

        ЗакрытьДиалог(true, false);

        var referenceGuid = RCCReference.RCCReferenceGuid.ToString();
        var registrationJournalGuid = RCCReference.RCCRegistrationJouranLink.ToString();

        var classObjectGuid = ТекущийОбъект.Тип.Guid.ToString();
        var newObject = СоздатьОбъект(referenceGuid, classObjectGuid);

        var registrationJournal = ТекущийОбъект.СвязанныйОбъект[registrationJournalGuid];

        if (registrationJournal != null)
            newObject.Подключить(registrationJournalGuid, registrationJournal);

        ПоказатьДиалогСвойств(newObject);
    }
}

