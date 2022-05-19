/*
TFlex.DOCs.UI.Client.dll
TFlex.DOCs.SynchronizerReference.dll
*/

using System;
using System.Threading;
using TFlex.DOCs.Client;
using TFlex.DOCs.Client.ViewModels;
using TFlex.DOCs.Client.ViewModels.Base;
using TFlex.DOCs.Client.ViewModels.DataExchange.Dialogs;
using TFlex.DOCs.Client.ViewModels.DataExchange.References.MasterDataBindings;
using TFlex.DOCs.Model.Macros;
using TFlex.DOCs.References.MasterDataBindings;
using TFlex.DOCs.Synchronization.SyncData;
using TFlex.DOCs.Synchronization.SyncData.Engine;

public class Macro70 : MacroProvider
{
    public Macro70(MacroContext context)
        : base(context)
    {
    }

    Guid правило_обмена_Guid = new Guid("d29447c1-6bfa-4236-b7dc-f0aa69c1b1c3");

    public void Import()
    {
        var masterDataBinding = new MasterDataBindingsReference(Context.Connection).Find(правило_обмена_Guid) as MasterDataBindingReferenceObject;

        // Настройки
        var settings = CreateSettings(null, masterDataBinding, ExchangeDataMode.Import);

        switch (masterDataBinding.ExternalDataPlatform)
        {
            case ExternalDataPlatform.Xml:
                {
                    var selectViewModel = new RunImportFileViewModel(settings, masterDataBinding.Dispatcher);

                    if (!ApplicationManager.OpenDialog(selectViewModel, ((UIMacroContext)Context).OwnerViewModel))
                        return;
                }
                break;
            case ExternalDataPlatform.Database:
                {
                    // %%TODO Сделать диалог

                    //var databaseSynchronizer = masterDataBinding as DatabaseExchangeDataRuleReferenceObject;
                    //var callback = new UISyncDataCallback(databaseSynchronizer);
                    //var parameters = databaseSynchronizer.FillConnectionParameters(callback);

                    //if (parameters is null)
                    //    return false;

                    //var request = new DatabaseRequest(parameters);

                    //settings.ExternalData = request;
                }
                break;
            case ExternalDataPlatform.Conversion:
                {
                    if (!RunExportFileViewModel.Open(settings, masterDataBinding.Dispatcher, null, CancellationToken.None).Result)
                        return ;
                }
                break;
            default:
                throw new NotSupportedException(masterDataBinding.ExternalDataPlatform.ToString());
        }

        var results = Run(masterDataBinding, settings);
        Сообщение("", "Импорт завершен");
    }

    public void Export()
    {
        var masterDataBinding = new MasterDataBindingsReference(Context.Connection).Find(правило_обмена_Guid) as MasterDataBindingReferenceObject;

        // Настройки
        var settings = CreateSettings(null, masterDataBinding, ExchangeDataMode.Export);

        switch (masterDataBinding.ExternalDataPlatform)
        {
            case ExternalDataPlatform.Xml:
                {
                    var selectViewModel = new RunExportFileViewModel(settings, masterDataBinding.Dispatcher);
                    selectViewModel.Initialize(CancellationToken.None).Wait();


                    selectViewModel.IsReferenceMode = true;
                    selectViewModel.RootObject = Context.ReferenceObject;


                    if (!ApplicationManager.OpenDialog(selectViewModel, ((UIMacroContext)Context).OwnerViewModel))
                        return;
                }
                break;
#if DEBUG
            case ExternalDataPlatform.Database:
                {
                    // %%TODO Сделать диалог
                }
                break;
#endif
            default:
                throw new NotSupportedException(masterDataBinding.ExternalDataPlatform.ToString());
        }

        // Report(Texts.DataExport);
        var results = Run(masterDataBinding, settings);
        Сообщение("", "Экспорт завершен");
    }

    protected DataExchangeSettings CreateSettings(ISupportSelection viewModel, MasterDataBindingReferenceObject masterDataBinding,
        ExchangeDataMode exchangeDataMode)
    {
        var settings = masterDataBinding.CreateDefaultSettings();
        settings.ExchangeDataMode = exchangeDataMode;
        var dispatcher = masterDataBinding.Dispatcher;

        if (dispatcher != null)
            settings.Callback = new UISynchronizerCallback(dispatcher, settings, viewModel as ISupportProgressIndicator);

        return settings;
    }

    protected DataExchangeResults Run(MasterDataBindingReferenceObject masterDataBinding, DataExchangeSettings settings)
    {
        var validationResult = masterDataBinding.Validate(settings);

        if (validationResult.HasErrors)
        {
            ApplicationManager.ShowMessage(validationResult.GetErrorMessage(), ((UIMacroContext)Context).OwnerViewModel);
            return null;
        }

        return masterDataBinding.Run(settings);
    }

}
