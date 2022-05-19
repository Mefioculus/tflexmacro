/*
TFlex.DOCs.SynchronizerReference.dll
*/

using System;
using TFlex.DOCs.Model.Macros;
using TFlex.DOCs.References.ReferenceSyncRules;
using TFlex.DOCs.Synchronization.Objects;
using MDMTexts = TFlex.DOCs.Synchronization.Resources.Strings.Texts;

public class Macro : MacroProvider
{
    public Macro(MacroContext context)
        : base(context)
    {
    }

    private void ExecuteCommand(string caption, Action<DatasetSyncRuleReferenceObject> action)
    {
        var dataSetSyncRule = Context.ReferenceObject as DatasetSyncRuleReferenceObject;

        if (dataSetSyncRule == null)
            return;

        var waitingDialog = Context.CreateWaitingDialog();
        waitingDialog.Show(caption, false);

        action(dataSetSyncRule);
    }

    public void СоздатьСправочникРеестра()
    {
        ExecuteCommand(MDMTexts.CreateRegisterReference, RegisterReferenceHelper.CreateRegisterReference);
    }

    public void ДобавитьПараметрВнешнегоИдентификатораВСправочникРеестра()
    {
        ExecuteCommand(MDMTexts.CreateExternalParameterToRegister, RegisterReferenceHelper.AddExternalIdentificatorParameterToRegister);
    }

    public void ДобавитьПараметрСравненияВерсийВСправочникРеестра()
    {
        ExecuteCommand(MDMTexts.CreateCompareParameterToRegister, RegisterReferenceHelper.AddCompareVersionParameterToRegister);
    }

    public void ДобавитьСвязьНаСправочникРеестра()
    {
        ExecuteCommand(MDMTexts.CreateLinkToRegister, RegisterReferenceHelper.AddLinkToRegister);
    }

}
