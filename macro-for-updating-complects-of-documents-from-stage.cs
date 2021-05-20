using System;
using System.Linq;
using System.IO;
using System.Collections;
using System.Collections.Generic;
using TFlex.DOCs.Model.Desktop; // Для применения изменений в файлах
using TFlex.DOCs.Model.Macros;
using TFlex.DOCs.Model.Macros.ObjectModel;
using TFlex.DOCs.Model.References; // для работы с справочниками и объектами справочников
using TFlex.DOCs.Model.References.Files; // Для работы с файловым справочником
using TFlex.DOCs.Model.References.Reporting; // Для формирования отчетов
using TFlex.DOCs.Model.FilePreview.CADExchange; // Для объединения grb файлов
using TFlex.DOCs.Model.FilePreview.CADService;  // Для экспорта grb файлов в другой формат
using TFlex.Model.Technology.References.SetOfDocuments; // Для работы с комплектами документов 

/*
Для работы макроса дополнительно следует указать две ссылки на библиотеки
TFlex.Model.Technology.dll
TFlex.Reporting.dll
*/

public class Macro : MacroProvider {
    public Macro(MacroContext context) : base(context) {
    }
    
    #region Guids

    public static class Guids {

        public static class References {
            public static Guid ТехнологическиеПроцессы = new Guid("353a49ac-569e-477c-8d65-e2c49e25dfeb");
            public static Guid ФайловыйСправочник = new Guid("a0fcd27d-e0f2-4c5a-bba3-8a452508e6b3");
        }
        public static class Classes {
            public static Guid БазовыйТехнологическийПроцесс = new Guid("cb230668-2a6a-4e10-988b-c6936fd06b06");
            public static Guid ТехнологическийКомплект = new Guid("dc1cf2a0-6c01-400d-9a42-9642b7496404");
        }
        public static class Links {
            public static Guid ТПДокументация = new Guid("cc38caed-f747-45ce-9fbf-771566841796");
            public static Guid ФайлОтчета = new Guid("6b18c3fc-7cd1-4ece-a526-cacad8101f09");
            public static Guid ФайлПодлинника = new Guid("148a64ed-3906-4da9-95fc-14bb018669f2");
        }

        public static class Folders {
            public static Guid ДиректорияДляДебагФайла = new Guid("61c60c06-71bd-4aeb-b67d-ef42d8ed04a7");
        }
    }

    #endregion

    #region Properties
    // В старой версии не поддерживается присовоение свойству начального значения
    //private string message { get; set; } = string.Empty;
    private string message { get; set; }
    private ReferenceObject currentTP { get; set; }
    private TechnologicalDocument document { get; set; }

    private FileObject currentReportFile { get; set; }
    //private int currentPercent { get; set; } = 0;
    //private int currentPercentLimit { get; set; } = 0;
    private int currentPercent { get; set; }
    private int currentPercentLimit { get; set; }

    #endregion Properties

    #region Точка входа в макрос

    public void ОбновлениеКомплектаДокументовИзПереходаБП(Объекты объектыИзБП) {
        // Точка входа для запуска макроса из перехода БП
        List<ReferenceObject>listOfTPs = new List<ReferenceObject>();
        foreach (Объект вложение in объектыИзБП) {
            RerefenceObject attachment = (ReferenceObject)вложение;

            if (attachment.Class.IsInherit(Guids.Classes.БазовыйТехнологическийПроцесс)) {
                listOfTPs.Add(attachment);
            }
        }
        ОбновлениеКомплектаДокументов(listOfTPs);
    }

    public void ОбновлениеКомплектаДокументовИзСтадииБП() {
        // Точка входа для запуска макроса из стадии БП
        List<ReferenceObject>listOfTPs = GetTPFromBP();
        ОбновлениеКомплектаДокументов(listOfTPs);
    }

    private void ОбновлениеКомплектаДокументов(List<ReferenceObject>listOfTPs) {
        ДиалогОжидания.Показать("Обновление комплекта документов", false);
        
        if ((listOfTPs == null) || (listOfTPs.Count == 0)) {
            Сообщение("Ошибка", "В бизнес-процессе не обнаружено ни одного прикрепленного технологического процесса");
            return;
        }
        //Сообщение("Информация", message);

        foreach (ReferenceObject tp in listOfTPs) {
             
            this.currentTP = tp;
            ProgressReset();
            ProgressSetLimit(15);
            ProgressShow();
            UpdateSetOfDocuments(tp);
        }
    }

    #endregion Точка входа в макрос

    #region Получение объектов из бизнес-процесса для вызова из стадии

    private List<ReferenctObject> GetTPFromBP() {
        // Метод для получения объектов из бизнес-процесса
        List<ReferenceObject> listOfTPs = new List<ReferenceObject>();

        ProcessReferenctObject process = null;

        ActiveActionReferenceObject activeAction = null;

        EventContext eventContext = Context as EventContext;

        if (eventContext != null) {
            var data = eventContext.Data as StateContextData;
            process = data.Process;
            activeAction = data.ActiveAction;
        }
        else {
            // Возможно придется расписать данный блок, так как предполагаю, что код выше будет отрабатывать не во всех случаях
            Сообщение("Ошибка", "Не удалось получить контекст события");
            return listOfTPs();
        }

        ActiveActionData activeActionData = activeAction.GetData<ActiveActionData>();

        foreach (ReferenceObject attachment in activeActionData.GetReferenceObject().ToList()) {
            if (attachment.Class.IsInherit(Guids.Classes.БазовыйТехнологическийПроцесс)) {
                listOfTPs.Add(attachment);
            }
        }

        return listOfTPs;
    }

    #endregion Получение объектов для работы

    #region UpdateSetOfDocument
    // Получаем список технологических комплектов, которые подключены к технологическому процессу
    private void UpdateSetOfDocuments(ReferenceObject tp) {
        var rootDocuments = tp.GetObjects(Guids.Links.ТПДокументация);
        if (!rootDocuments.Any()) {
            Сообщение("Ошибка", "К данному технологическому процессу не подключено никаких документов");
            return;
        }
        ProgressShow();

        var setOfDocuments = rootDocuments.FirstOrDefault(document =>
                document.Class.IsInherit(Guids.Classes.ТехнологическийКомплект)) as TechnologicalSet;

        ProgressShow();
        // Получаем папку, в которой будет располагаться документ
        FolderObject folder = setOfDocuments.Folder;
        if (folder == null) {
            Сообщение("Ошибка", "Директория, в которой будет располагаться комплект документов на найдена");
            return;
        }
        ProgressShow();

        List<string> files = new List<string>();

        setOfDocuments.Children.Reload();
        ProgressShow();

        // Получаем пути к файлам 
        ProgressSetLimit(50);
        ProgressShow(16);
        foreach (TechnologicalDocument document in setOfDocuments.Children.OfType<TechnologicalDocument>()) {
            // Запускаем обновление технологических документов
            FileObject documentFile = UpdateTechnologicalDocument(document);

            if (documentFile == null) {
                continue;
            }

            try {
                documentFile.GetHeadRevision();
                files.Add(documentFile.LocalPath);
            }
            catch {
                // Обработка возникающих ошибок
                Сообщение("Ошибка", "Возникла ошибка во время получения локального пути к файлам технологических документов");
            }
        }

        ProgressSetLimit(80);
        ProgressShow(51);

        if (!files.Any()) {
            Сообщение("Информация", "Не было обнаружено ни одного файла отчета для обновления комплекта документов");
            return;
        }
        ProgressShow();

        // Формирование коплекта документов на основании сформированных файлов отчетов.

        // Проверяем, есть ли уже сформированный файл комплекта
        if (setOfDocuments.File != null) {
            //Сообщение("Информация", "Комплект документов содержит файл оригинала");
            //Сообщение("Информация", string.Format("Путь, по которому располагается файл комплекта документов - '{0}'", ((FileObject)setOfDocuments.File).LocalPath));
            string pathToOriginalFile = ((FileObject)setOfDocuments.File).LocalPath;

            // Берем на редактирование сущестующий файл комплекта и обновляем в него новый комплект
            setOfDocuments.File.CheckOut(false);
            setOfDocuments.File.BeginChanges();
            
            CombineFilesProvider provider = new CombineFilesProvider(files, pathToOriginalFile) { IsEmbedded = true };
            provider.Execute(Context.Connection);

            setOfDocuments.File.EndChanges();
            Desktop.CheckIn(setOfDocuments.File, string.Format("Обновление комплекта технологических документов для ТП '{0}'", this.currentTP.ToString()), false);
        }
        else {
            Сообщение("Информация", "Комплект документов не содержит файла оригинала");
        }

        ProgressSetLimit(99);
        ProgressShow(81);
        ConvertGRBToPDF(setOfDocuments);

        ProgressShow(100);
    }

    #endregion

    #region ConvertGRBToPDF

    private void ConvertGRBToPDF(TechnologicalSet setOfDocuments) {

        ProgressShow();
        // Получаем файл, с которого будем формировать подлинник
        FileObject originalFile = setOfDocuments.File as FileObject;
        originalFile.GetHeadRevision();

        FileObject realFile = setOfDocuments.GetObject(Guids.Links.ФайлПодлинника) as FileObject;
        ProgressShow();

        // Узнаем, есть ли у комплекта документов файл подлинника.
        if (realFile != null) {
            // Берем существующий файл подлинника на изменение
            realFile.GetHeadRevision();
            realFile.CheckOut(false);
            realFile.BeginChanges();
            // Перезаписываем файл подлинника обновленными данными
            CreatePDFFile(originalFile.LocalPath, realFile.LocalPath);
            // Применяем изменения к файлу подлинника
            realFile.EndChanges();
            Desktop.CheckIn(realFile, string.Format("Обновление подлинника для '{0}'", setOfDocuments.ToString()), false);
        }
        else {
            // Если файла подлинника еще не существовало, тогда его можно создать.
            FolderObject folder = setOfDocuments.Folder as FolderObject;

            if (folder != null) {
                string pathToRealFile = Path.Combine(Path.GetTempPath(), string.Format("{0}.pdf", setOfDocuments.ToString()));
                CreatePDFFile(originalFile.LocalPath, pathToRealFile);

                FileReference fileReference = new FileReference(Context.Connection);
                realFile = fileReference.AddFile(pathToRealFile, folder);

                // Подключаем объект по связи к комплекту документов
                setOfDocuments.SetLinkedObject(Guids.Links.ФайлПодлинника, realFile);
                
                // Применяем изменения в системе
                Desktop.CheckIn(realFile, string.Format("Обновление подлинника для комплекта документов '{0}'", setOfDocuments.ToString()), false);
                // Удаляем временный файл
                File.Delete(pathToRealFile);


            }
            else {
                Сообщение("Ошибка", "В комплекте документов не указана папка для сохранения комплекта документов и подлинника");
            }
        }
        ProgressShow();
    }

    private void CreatePDFFile(string pathToOriginal, string pathToSave) {
        CadDocumentProvider provider = CadDocumentProvider.Connect(Context.Connection, ".grb");
        try {
            using (var document = provider.OpenDocument(pathToOriginal, true)) {
                ProgressShow();
                ExportContext exportContext = new ExportContext(pathToSave);
                PageInfo[] pages = document.GetPagesInfo();

                for (int i = 0; i < pages.Count(); i++) {
                    exportContext.Pages.Add(i);
                }
                ProgressShow();
                document.Export(exportContext);
                document.Close(false);
                ProgressShow();
            }
        }
        catch (Exception e) {
            Сообщение("Ошибка", string.Format("При формировании подлинника возникла ошибка: \n\n{0}", e.Message));
        }
    }

    #endregion

    #region UdpateTechnologicalDocument

    private FileObject UpdateTechnologicalDocument(TechnologicalDocument document) {
        ProgressShow();
        // Метод для обновления отчетов, привязанных к технологическому документу
        this.document = document;
        //Сообщение("Входящий документ", document.ToString());

        // Проверяем, если ли подключенный к документу файл отчета
        Report reportObject = document.ReportObject;
        ProgressShow();
        if (reportObject != null) {
            //Сообщение("Информация", string.Format("Для технологического процесса '{0}' был найден отчет '{1}'", document.ToString(), reportObject.ToString()));
        }
        else {
            Сообщение("Информация", string.Format("Для технологического процесса '{0}' не было найдена отчета", document.ToString()));
            return null;
        }
        ProgressShow();

        FileObject fileOfReport = GenerateReport(reportObject, this.currentTP);
        ProgressShow();

        if (fileOfReport != null)
            Desktop.CheckIn(fileOfReport, string.Format("Создание отчета '{0}' для технологического процесса '{1}'", document.ToString(), this.currentTP.ToString()), false);
        ProgressShow();

        return fileOfReport;

    }

    #endregion

    #region GenerateReport

    private FileObject GenerateReport (Report report, ReferenceObject tp) {
        
        // Получаем данные об уже сгенерированном отчете для того, чтобы его перезаписать
        this.currentReportFile = document.GetObject(Guids.Links.ФайлОтчета) as FileObject;
        if (this.currentReportFile == null)
            return null;
        ProgressShow();

        // Получаем контекст генерации отчета
        ReportGenerationContext reportContext = new ReportGenerationContext (tp, null);

        reportContext.OpenFile = false;
        reportContext.OverwriteReportFile = true;
        ProgressShow();

        FolderObject folderForSaving = GetFolderObjectOfOldReport(this.document);
        if (folderForSaving != null)
            reportContext.DefaultFolderObject = folderForSaving;

        string nameOfReportFile = GetNameOfOldReport(this.document);
        if (nameOfReportFile != null)
            reportContext.ReportFileName = nameOfReportFile;
        ProgressShow();


        report.Generate(reportContext);
        return reportContext.ReportFileObject as FileObject;
    }

    private FolderObject GetFolderObjectOfOldReport(TechnologicalDocument document) {
        if (this.currentReportFile != null)
            return this.currentReportFile.Parent as FolderObject;
        else
            return null;
    }

    private string GetNameOfOldReport(TechnologicalDocument document) {
        if (this.currentReportFile != null)
            return this.currentReportFile.ToString();
        else
            return null;
    }




    #endregion

    #region Calculate Percentage
    
    private void ProgressShow(int percent = -1) {
        if (percent == -1)
            percent = this.currentPercent;

        ДиалогОжидания.СледующийШаг(string.Format("Обновление комплекта ТД для ТП '{0}' - {1:000} %", this.currentTP.ToString(), percent));

        if (percent < this.currentPercentLimit)
            this.currentPercent = ++percent;
    }

    private void ProgressReset() {
        this.currentPercent = 0;
        this.currentPercentLimit = 0;
    }

    private void ProgressSetLimit(int limit) {
        this.currentPercentLimit = limit;
    }

    #endregion

}
