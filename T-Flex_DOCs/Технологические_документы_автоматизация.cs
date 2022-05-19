using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using TFlex.DOCs.Model.References;
using TFlex.DOCs.Model.References.Files;
using TFlex.DOCs.Model.Macros;
using TFlex.DOCs.Model.Macros.ObjectModel;
using TFlex.Reporting.Technology.Macros;
using TFlex.Reporting.CAD.MacroGenerator.ObjectModel;
using TFlex.DOCs.Model.References.Reporting;
using TFlex.DOCs.Model.References.Nomenclature;
using TFlex.DOCs.Model.FilePreview.CADService;
using TFlex.DOCs.UI.Objects.References.Files.Commands;
using Newtonsoft.Json;
using TFlex.DOCs.Client;
using TFlex.DOCs.Client.ViewModels;
using TFlex.DOCs.Client.ViewModels.References;
using ReferenceObjectWithLink = TFlex.DOCs.Model.References.Reporting.ReferenceObjectWithLink;

public class Macro : MacroProvider
{
    public Macro(MacroContext context)
        : base(context)
    {
    }

    public override void Run()
    {
    
    }
    
    public void СоздатьИнвКарточку()
    {
    	Объект документ = ТекущийОбъект;
    	if (документ == null)
    		return;
    	if (документ.СвязанныйОбъект["Карточка учёта"] != null)
    		return;
    	
    	string ТипКарточки = "Карточка учёта технологических документов"; 
    	if (!документ.Тип.ПорожденОт("Технологический документ"))
    		ТипКарточки = "Карточка учёта конструкторских документов";
    	
    	Объект карточка = СоздатьОбъект("Инвентарная книга", ТипКарточки);
    	карточка.СвязанныйОбъект["Документ"] = документ;
    	карточка["Наименование документа"] = документ["Наименование"];
    	карточка["Обозначение документа"] = документ["Обозначение"];
    	
    	карточка.Сохранить();
    	ПоказатьДиалогСвойств(карточка);
    }
    
    public void СменаСтадииХранениеДокументаИФайлы()
    {
    	Объект документ = ТекущийОбъект;
    	документ.ИзменитьСтадию("Хранение", true);
    	
    	foreach (var файл in документ.СвязанныеОбъекты["Файлы"])
    		файл.ИзменитьСтадию("Хранение", true);
    }
    
    public void СменаСтадииКорректировкаДокументаИФайлы()
    {
    	Объект документ = ТекущийОбъект;
    	документ.ИзменитьСтадию("Корректировка", true);
    	
    	foreach (var файл in документ.СвязанныеОбъекты["Файлы"])
    		файл.ИзменитьСтадию("Корректировка", true);
    }
    
    public void ФормированиеГППМаршруты()
    {
    	ФормированиеДокументаИОтчета("Маршруты");
    }
    
    public void ФормированиеГППОснастка()
    {
    	ФормированиеДокументаИОтчета("Оснастка");
    }
    
    public void ФормированиеГППВМ()
    {
    	ФормированиеДокументаИОтчета("ВМ");
    }
    
    // Добавлено АЭМ: Начало
    public void ФормированиеВедомостиНормРасходаДрагоценныхМеталлов()
    {
        ФормированиеДокументаИОтчета("ВедомостьДГ");
    }

    public void ФормированиеСводныхНормРасходаДрагоценныхМеталлов()
    {
        ФормированиеДокументаИОтчета("СводныеНормыДГ");
    }
    // Добавлено АЭМ: Конец
    
    private void ФормированиеДокументаИОтчета(string type)
    {
    	Объект ном = ТекущийОбъект;
    	ПользовательскийДиалог диалог = ПолучитьПользовательскийДиалог("Создание технологического документа");
    	диалог["Обозначение"] = "";
    	диалог["[Технологический документ].[С изделия №]"] = 0;
    	диалог["[Технологический документ].[По изделие №]"] = 0;
    	foreach (var абонент in диалог.СвязанныеОбъекты["Абоненты для рассылки"])
    		абонент.Удалить();
    	
    	Объект РодительскаяПапка = null;
    	
    	if(type == "Маршруты")
    	{
    		диалог["Наименование"] = string.Format("ГПП (маршруты) по выпуску изделия {0} - {1}", ном["Обозначение"], ном["Наименование"]);
    		диалог["Вид документа"] = 1;
    		РодительскаяПапка = НайтиОбъект("Документы", "[Наименование] = 'ГПП (МАРШРУТЫ)'");
    	}
    	else if (type == "Оснастка")
    	{
    		диалог["Наименование"] = string.Format("ГПП (оснастка) по выпуску изделия {0} - {1}", ном["Обозначение"], ном["Наименование"]);
    		диалог["Вид документа"] = 2;
    		РодительскаяПапка = НайтиОбъект("Документы", "[Наименование] = 'ГПП (ОСНАСТКА)'");
    	}
        else if (type == "ВМ")
    	{
    		диалог["Наименование"] = string.Format("Ведомость материалов для выпуска изделия {0} - {1}", ном["Обозначение"], ном["Наименование"]);
    		диалог["Вид документа"] = 3;
    		РодительскаяПапка = НайтиОбъект("Документы", "[Наименование] = 'ВЕДОМОСТЬ МАТЕРИАЛОВ'");
    	}
        else if (type == "ВедомостьДГ")
        {
            диалог["Наименование"] = string.Format("Ведомость норм расхода драгоценных металлов для выпуска изделия {0} - {1}", ном["Обозначение"], ном["Наименование"]);
            диалог["Вид документа"] = 4;
            РодительскаяПапка = НайтиОбъект("Документы", "[Наименование] = 'ВЕДОМОСТЬ НОРМ РАСХОДА ДРАГОЦЕННЫХ МЕТАЛЛОВ'");
        }
        else if (type == "СводныеНормыДГ")
        {
            диалог["Наименование"] = string.Format("Сводные нормы расхода драгоценных металлов для выпуска изделия {0} - {1}", ном["Обозначение"], ном["Наименование"]);
            диалог["Вид документа"] = 5;
            РодительскаяПапка = НайтиОбъект("Документы", "[Наименование] = 'СВОДНЫЕ НОРМЫ РАСХОДА ДРАГОЦЕННЫХ МЕТАЛЛОВ'");
        }
        else
        {
            Сообщение("Ошибка во время формирования документа", String.Format("Документа с типом '{0}' не существует. Формирование документа будет прекращено", type));
            return;
        }
    	
    	// Список объектов "Входящие изделия"
    	foreach (var izd in диалог.СвязанныеОбъекты["Входящие изделия"])
    		izd.Удалить();
    	
    	ReferenceObject nom = (ReferenceObject)ном;
    		
    	foreach (var con in nom.Children.RecursiveLoadHierarchyLinks())
    	{
    		if (con.ChildObject.Class.IsInherit(new Guid("7fa98498-c39c-44fc-bcaa-699b387f7f46")) || 		// Изделие
    		    con.ChildObject.Class.IsInherit(new Guid("1cee5551-3a68-45de-9f33-2b4afdbf4a5c")))			// Сборочная единица
    		{
    			Объект ЭСИ = Объект.CreateInstance(con.ChildObject, Context);
    			
    			Объект изделие = диалог.СоздатьОбъектСписка("Входящие изделия", "Изделие");
    			изделие["Наименование"] = string.Format("{0} - {1}", ЭСИ["Обозначение"], ЭСИ["Наименование"]);
    			изделие.СвязанныйОбъект["e2e0ef2c-300a-4fc3-9766-269d361bf11c"] = ЭСИ;
    			
    			изделие.Сохранить();
    		}
    	}
    	
    	if (диалог.ПоказатьДиалог())
    	{
    		Объект документ = НайтиОбъект("Документы", string.Format("[Наименование] = '{0}' И [Обозначение] = '{1}'", диалог["Наименование"], диалог["Обозначение"]));
    		if (документ != null)
    		{
    			документ.Изменить();
    			
        		документ["Дата документа"] = диалог["Дата документа"];
            	документ["С изделия №"] = диалог["[Технологический документ].[С изделия №]"];
            	документ["По изделие №"] = диалог["[Технологический документ].[По изделие №]"];
            	документ["Вид документа"] = диалог["[Технологический документ].[Вид документа]"];
        		
        		foreach (var абонент in диалог.СвязанныеОбъекты["Абоненты для рассылки"])
        		{
        			if (документ.СвязанныеОбъекты["Абоненты для рассылки"].Any(t => t["Наименование"].ToString() == абонент["Наименование"].ToString()))
        				continue;
        			
        			Объект абонентДок = документ.СоздатьОбъектСписка("Абоненты для рассылки", "Абонент");
        			абонентДок["Наименование"] = абонент["Наименование"];
        			абонентДок["Количество экземпляров"] = абонент["Количество экземпляров"];
        			
        			абонентДок.Сохранить();
        		}
        	    документ.Сохранить();
        	}
    		else
    		{
    			документ = СоздатьОбъект("Документы", "Технологический документ", РодительскаяПапка);
        		документ["Наименование"] = диалог["Наименование"];
        		документ["Обозначение"] = диалог["Обозначение"];
        		документ["Дата документа"] = диалог["Дата документа"];
        		документ["С изделия №"] = диалог["[Технологический документ].[С изделия №]"];
        		документ["По изделие №"] = диалог["[Технологический документ].[По изделие №]"];
        		документ["Вид документа"] = диалог["[Технологический документ].[Вид документа]"];
        		
        		foreach (var абонент in диалог.СвязанныеОбъекты["Абоненты для рассылки"])
        		{
        			Объект абонентДок = документ.СоздатьОбъектСписка("Абоненты для рассылки", "Абонент");
        			абонентДок["Наименование"] = абонент["Наименование"];
        			абонентДок["Количество экземпляров"] = абонент["Количество экземпляров"];
        			
        			абонентДок.Сохранить();
        		}
        		
        		foreach (var входящееИзделие in диалог.СвязанныеОбъекты["Входящие изделия"])
        		{
        			Объект вхИзделие = документ.СоздатьОбъектСписка("Входящие изделия", "Изделие");
        			вхИзделие["Наименование"] = входящееИзделие["Наименование"];
        			bool flag = (bool)входящееИзделие["Учитывать в документе"];
        			вхИзделие["Есть ГПП"] = !flag;
        			вхИзделие.СвязанныйОбъект["ЭСИ"] = входящееИзделие.СвязанныйОбъект["ЭСИ"];
        			
        			вхИзделие.Сохранить();
        		}
        		
        		документ.Подключить("Связанные документы", ном.СвязанныйОбъект["Связанный объект"]);
        		документ.Сохранить();
            }
    		
    		List<int> IDsAddToReport = new List<int>();
    			
    		IDsAddToReport.AddRange(диалог.СвязанныеОбъекты["Входящие изделия"].
    		                        //Where(t => (bool)t["Учитывать в документе"] == false).
    		                        Where(t => (bool)t["Учитывать в документе"] != false).
    		                        Select(t => (int)t.СвязанныйОбъект["e2e0ef2c-300a-4fc3-9766-269d361bf11c"]["Id"]));
    		Data data = new Data();
    		data.IDsAddToReport = IDsAddToReport;
    		data.From = (int)диалог["[Технологический документ].[С изделия №]"];
    		data.To = (int)диалог["[Технологический документ].[По изделие №]"];
    		data.Denotation = диалог["[Технологический документ].[Обозначение]"];
    		
            // Берем данные с конфигуратора
            var nomenclatureObject = Context.ReferenceObject as NomenclatureObject;
            var reference = new NomenclatureReference(Context.Connection);
            reference.DigitalStructureContext.Date = nomenclatureObject.Reference.DigitalStructureContext.Date;
            reference.DigitalStructureContext.ApplyDate = nomenclatureObject.Reference.DigitalStructureContext.ApplyDate;
            reference.DigitalStructureContext.ActiveStructures = nomenclatureObject.Reference.DigitalStructureContext.ActiveStructures;
            reference.DigitalStructureContext.ApplyDesignContext = nomenclatureObject.Reference.DigitalStructureContext.ApplyDesignContext;
            reference.DigitalStructureContext.ApplyCategoriesFilter = nomenclatureObject.Reference.DigitalStructureContext.ApplyCategoriesFilter;
            reference.DigitalStructureContext.ApplyProductFilter = nomenclatureObject.Reference.DigitalStructureContext.ApplyProductFilter;
            reference.DigitalStructureContext.DesignContext = nomenclatureObject.Reference.DigitalStructureContext.DesignContext;
            reference.DigitalStructureContext.Product = nomenclatureObject.Reference.DigitalStructureContext.Product;

            reference.Refresh();
            var cur = reference.Find(nomenclatureObject.SystemFields.Id) as NomenclatureObject;
    		
    		string serialized = JsonConvert.SerializeObject(data);
    		
    		ДиалогОжидания.Показать("Пожалуйста, подождите", true);
    		
    		Context.RunOnUIThread(() => 
            {
    		    Clipboard.SetText(serialized);	
    		    var mvm = ApplicationManager.MainViewModel as MainViewModel;
                var content = mvm.MdiContainer.SelectedItem.Content;
                var refVM = content as ReferenceExplorerViewModel;
                refVM?.TreeViewModel?.ReloadData();
                refVM?.GridViewModel?.ReloadData();
            });
    		
    		//СформироватьОтчет(type, ном, документ); 
            СформироватьОтчет2(type, cur, документ);
            
    		ДиалогОжидания.Скрыть();
    		
    		/*var mvm = ApplicationManager.MainViewModel as MainViewModel;
            var content = mvm.MdiContainer.SelectedItem.Content;
            var refVM = content as ReferenceExplorerViewModel;
            refVM?.TreeViewModel?.ReloadData();
            refVM?.GridViewModel?.ReloadData();*/
    	}
    	else
    		return;
    }
    
    private void СформироватьОтчет(string type, Объект ном, Объект документ)
    {
    	Объект отчет = null;
    	if(type == "Маршруты")
    	{
    		отчет = НайтиОбъект("Отчёты", "[Наименование] = 'ГПП маршруты'");
    	}
    	else if (type == "Оснастка")
    	{
    		отчет = НайтиОбъект("Отчёты", "[Наименование] = 'ГПП оснастка'");
    	}
        else if (type == "ВМ")
    	{
    		отчет = НайтиОбъект("Отчёты", "[Наименование] = 'ГПП Ведомость материалов'");
    	}
        else if (type == "ВедомостьДГ")
        {
            отчет = НайтиОбъект("Отчёты", "[Наименование] = '(Гуков) Ведомость норм расхода драгоценных металлов'");
        }
        else if (type == "СводныеНормыДГ")
        {
            отчет = НайтиОбъект("Отчёты", "[Наименование] = '(Гуков) Сводные нормы расхода драгоценных металлов'");
        }
        else
        {
            Сообщение("Ошибка во время формирования отчета", String.Format("Документа с типом '{0}' не существует. Формирование документа будет прекращено", type));
            return;
        }
    	if (отчет == null) {
            Сообщение("Ошибка во время поиска отчета", "Файл отчета не был найден, работа макроса будет прекращена");
        	return;
    	}
    	
    	ReportGenerationContext reportContext = new ReportGenerationContext((ReferenceObject)ном, null);
    	reportContext.OpenFile = false;
        ReferenceObject rep = (ReferenceObject)отчет;
        Report report = (Report)отчет;
        report.Generate(reportContext);
        var generatedFile = reportContext.ReportFileObject as FileObject;
        Объект файл = Объект.CreateInstance(generatedFile, Context);
        файл.Изменить();
        файл.Подключить("Документы", документ);
        
        файл.Сохранить();
        // Открываем файл
        
        /*generatedFile.GetHeadRevision();
        string filePath = generatedFile.LocalPath;
        
        CadDocumentProvider provider = CadDocumentProvider.Connect(Context.Connection, ".grb");
        using (CadDocument document = provider.OpenDocument(filePath, false))      // Открываем Cad-файл в режиме редактирования
    	{
    		document.Open(provider);
    		
    		ImportFilesUIHelper helper = new ImportFilesUIHelper(generatedFile.Reference, generatedFile.Parent, false);     // Открываем приложение
            helper.Edit(generatedFile);
    	}*/
    }
    // С учтетом контекста генератора
    private void СформироватьОтчет2(string type, NomenclatureObject cur, Объект документ)
    {
        Объект отчет = null;
        if (type == "Маршруты")
        {
            отчет = НайтиОбъект("Отчёты", "[Наименование] = 'ГПП маршруты'");
        }
        else if (type == "Оснастка")
        {
            отчет = НайтиОбъект("Отчёты", "[Наименование] = 'ГПП оснастка'");
        }
        else if (type == "ВМ")
        {
            отчет = НайтиОбъект("Отчёты", "[Наименование] = 'ГПП Ведомость материалов'");
        }
        else if (type == "ВедомостьДГ")
        {
            отчет = НайтиОбъект("Отчёты", "[Наименование] = '(Гуков) Ведомость норм расхода драгоценных металлов'");
        }
        else if (type == "СводныеНормыДГ")
        {
            отчет = НайтиОбъект("Отчёты", "[Наименование] = '(Гуков) Сводные нормы расхода драгоценных металлов'");
        }
        else
        {
            Сообщение("Ошибка во время формирования отчета", String.Format("Документа с типом '{0}' не существует. Формирование документа будет прекращено", type));
            return;
        }
        
        if (отчет == null) 
        {
            Сообщение("Ошибка во время поиска отчета", "Файл отчета не был найден, работа макроса будет прекращена");
            return;
        }

        List<ReferenceObject> findedObjects = new List<ReferenceObject>() { cur };
        IEnumerable<ReferenceObjectWithLink> referenceObjectWithLinks = findedObjects.Select(obj => new ReferenceObjectWithLink(obj));

        ReportGenerationContext reportContext = ReportGenerationContextFactory.CreateContext(referenceObjectWithLinks);
        reportContext.OpenFile = false;
        ReferenceObject rep = (ReferenceObject)отчет;
        Report report = (Report)отчет;
        report.Generate(reportContext);
        var generatedFile = reportContext.ReportFileObject as FileObject;
        Объект файл = Объект.CreateInstance(generatedFile, Context);
        файл.Изменить();
        файл.Подключить("Документы", документ);

        файл.Сохранить();
        // Открываем файл

        /*generatedFile.GetHeadRevision();
        string filePath = generatedFile.LocalPath;
        
        CadDocumentProvider provider = CadDocumentProvider.Connect(Context.Connection, ".grb");
        using (CadDocument document = provider.OpenDocument(filePath, false))      // Открываем Cad-файл в режиме редактирования
        {
            document.Open(provider);
            
            ImportFilesUIHelper helper = new ImportFilesUIHelper(generatedFile.Reference, generatedFile.Parent, false);     // Открываем приложение
            helper.Edit(generatedFile);
        }*/
    }

    public class Data
    {
    	public List<int> IDsAddToReport { get; set; }
    	public int? From { get; set; }
    	public int? To { get; set; }
    	public string Denotation { get; set; }
    	
    	public Data()
    	{
    		IDsAddToReport = new List<int>();
    	}
    }
}
