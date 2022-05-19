/* Дополнительные ссылки
TFlex.DOCs.Model.Office.dll
*/

using TFlex.DOCs.Model.Macros;
using TFlex.DOCs.Model.References.RecordControlCards;


public class CreateRccWithTypeSelectionMacro : MacroProvider
{
    public CreateRccWithTypeSelectionMacro(MacroContext context)
        : base(context)
    {
    }

    public override void Run()
    {
        var rccObject = Context.ReferenceObject as RCCReferenceObject;
        if (rccObject == null)
            return;

        var referenceGuid = RCCReference.RCCReferenceGuid.ToString();
        var registrationJournalGuid = RCCReference.RCCRegistrationJouranLink.ToString();

        var classObjectAccessor = ПоказатьДиалогВыбораТипа(referenceGuid);
        if (classObjectAccessor == null)
            return;        

        ЗакрытьДиалог(true, false);

        var classObjectGuid = classObjectAccessor.Guid.ToString();
        var newObject = СоздатьОбъект(referenceGuid, classObjectGuid);

        var registrationJournal = ТекущийОбъект.СвязанныйОбъект[registrationJournalGuid];

        if (registrationJournal != null)
            newObject.Подключить(registrationJournalGuid, registrationJournal);

        ПоказатьДиалогСвойств(newObject);
    }
}

