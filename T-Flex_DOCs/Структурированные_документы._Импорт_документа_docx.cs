/*
PresentationFramework.dll
PresentationCore.dll
System.Xaml.dll
WindowsBase.dll
TFlex.DOCs.UI.Client.dll
TFlex.DOCs.StructuredDocumentImport.Word.dll
TFlex.DOCs.StructuredDocumentImport.ViewModel.dll
*/

using TFlex.DOCs.Model.Macros;
using TFlex.DOCs.StructuredDocumentImport.Word;
using TFlex.DOCs.StructuredDocumentImport.ViewModel;
using TFlex.DOCs.StructuredDocumentImport.ViewModel.ViewModels;
using TFlex.DOCs.StructuredDocumentImport.ViewModel.Dialogs;

namespace Macros
{
    public class CallUtilityMacro : MacroProvider
    {
        public CallUtilityMacro(MacroContext context)
            : base(context)
        {
        }

        public override void Run()
        {
            var dialogService = new DialogService();
            var mainViewModel = new MainWindowViewModel(
                new DocumentStructExtractor(),
                new DocumentStructUploader(Context),
                dialogService,
                new DocumentObfuscator());
            
            Context.RunOnUIThread(() => { Start(mainViewModel, dialogService); });
        }

        public ButtonValidator GetButtonValidator()
        {
            bool visible = !Context.Reference.IsSlave;
            return new ButtonValidator()
            {
                Enable = true,
                Visible = visible
            };
        }
        
        private void Start(MainWindowViewModel mainViewModel, IDialogService dialogService)
        {
            var window = new MainWindow(mainViewModel, dialogService);
            window.ShowDialog();
        }
    }
}
