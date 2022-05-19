using System;
using TFlex.DOCs.Model.Macros;
using TFlex.DOCs.Model.References.Reporting;
#if WPFCLIENT
using DevExpress.Xpf.Printing;
using System.Windows;
#else
using DevExpress.XtraPrinting;
#endif

public class Macro : MacroProvider
{
    public Macro(MacroContext context)
        : base(context)
    {
    }

    public void НазначитьНомер()
    {
        if (Параметр["Номер договора"] != "")
            return;

        ГлобальныйПараметр["Текущий номер договора"] += 1;
        Параметр["Наименование"] = "Договор №" + ГлобальныйПараметр["Текущий номер договора"];
        Параметр["Номер договора"] = ГлобальныйПараметр["Текущий номер договора"];
    }

    public void Создание()
    {
        Параметр["Дата договора"] = DateTime.Now;
    }

    public void СоздатьФайл()
    {
        var report = Context.ReferenceObject.GetObject(Guids.LinkAgreementGenerator) as Report;
        if (report is null)
            Ошибка("Не задан прототип договора");

        ReportGenerationContext reportContext = new ReportGenerationContext(Context.ReferenceObject, null);
        var reportIsDx = CheckDxReport(report);
        reportContext.OpenFile = !reportIsDx;
        reportContext.OverwriteReportFile = true;

        report.Generate(reportContext);

        if (!reportIsDx)
            return;

#if WPFCLIENT
        Context.RunOnUIThread(() =>
        {
            var control = new DocumentPreviewControl();
            var window = new Window
            {
                Title = reportContext.ReportFileName ?? System.IO.Path.GetFileNameWithoutExtension(reportContext.ReportFilePath),
                Content = control
            };

            window.Show();
            control.OpenDocument(reportContext.ReportFilePath);
        });
#else
        using var printingSystem = new PrintingSystem();
        printingSystem.LoadDocument(reportContext.ReportFilePath);
        printingSystem.PreviewFormEx.ShowDialog();
#endif
    }

    private bool CheckDxReport(Report report)
        => String.Equals(report.TemplateFile.Class.Extension, "repx", StringComparison.OrdinalIgnoreCase);

    private static class Guids
    {
        internal static readonly Guid LinkAgreementGenerator = new Guid("e703811f-f34d-44c1-8ece-3a3eabdcc449");
    }
}
