using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using TFlex.DOCs.Model.Macros;
using TFlex.DOCs.Model.Macros.ObjectModel;
using TFlex.DOCs.Model.References.Files;
using TFlex.DOCs.Model.References;
using TFlex.DOCs.Model.FilePreview.CADService;

using TFlex.DOCs.UI.Objects.References.Files.Commands;

public class Macro : MacroProvider
{
    public Macro(MacroContext context)
        : base(context)
    {
    }
    
    public override void Run()
    {
    	
    }

    public void СоздатьНовыйЭскиз()
    {
    	var прототип = НайтиПрототип("Файлы", "[Наименование] = '2D чертёж.grb'"); // Находим прототип.
        if (прототип == null)
            return;

        var папка = НайтиОбъект("Файлы", "Наименование", "Эскизы технологии"); // Находим папку для сохранения файла.
        if (папка == null)
            return;

        var скопированныеОбъекты = СкопироватьОбъект(прототип, папка);
        Объект ТП = null;
        if (!ТекущийОбъект.РодительскийОбъект.Тип.ПорожденОт("Технологический процесс"))
        	if (ТекущийОбъект.РодительскийОбъект.РодительскийОбъект.Тип.ПорожденОт("Технологический процесс"))
        	    ТП = ТекущийОбъект.РодительскийОбъект.РодительскийОбъект;
        else if (ТекущийОбъект.РодительскийОбъект.Тип.ПорожденОт("Технологический процесс"))
        	ТП = ТекущийОбъект.РодительскийОбъект;
        if (ТП == null)
        	return;
        
        var новыйФайл = скопированныеОбъекты.ПолучитьКопию(прототип); // созданный файл
        if (ТП.Тип.Имя == "Изготовление")
        	новыйФайл["Наименование"] = ТП["Наименование"] + " - " + ТП["Обозначение"] + " - " + ТекущийОбъект["Номер"] + ".grb";
        else
            новыйФайл["Наименование"] = ТП["Наименование"] + " - " + ТП["Обозначение ТП"] + " - " + ТекущийОбъект["Номер"] + ".grb";
        скопированныеОбъекты.Сохранить(); // сохраняем
        
    	ТекущийОбъект.Изменить();
    	ТекущийОбъект.СвязанныйОбъект["Чертеж детали"] = новыйФайл;
    	ТекущийОбъект.Сохранить();

    	FileObject file = (ReferenceObject)новыйФайл as FileObject;
    	file.GetHeadRevision();
    	string filePath = file.LocalPath;
    	
    	CadDocumentProvider provider = CadDocumentProvider.Connect(Context.Connection, ".grb");
    	using (CadDocument document = provider.OpenDocument(filePath, false))      // Открываем Cad-файл в режиме редактирования
    	{
    		document.Open(provider);
    		
    		ImportFilesUIHelper helper = new ImportFilesUIHelper(file.Reference, file.Parent, false);     // Открываем приложение
            helper.Edit(file);
    	}
    }
    
    public void СоздатьНаОсновеЧертежа()
    {
        var папка = НайтиОбъект("Файлы", "Наименование", "Эскизы технологии"); // Находим папку для сохранения файла.
        if (папка == null)
            return;
        
        Объект ТП = null;
        if (!ТекущийОбъект.РодительскийОбъект.Тип.ПорожденОт("Технологический процесс"))
        	if (ТекущийОбъект.РодительскийОбъект.РодительскийОбъект.Тип.ПорожденОт("Технологический процесс"))
        	    ТП = ТекущийОбъект.РодительскийОбъект.РодительскийОбъект;
        else if (ТекущийОбъект.РодительскийОбъект.Тип.ПорожденОт("Технологический процесс"))
        	ТП = ТекущийОбъект.РодительскийОбъект;
        if (ТП == null)
        	return;
        
        var прототип = ТП.СвязанныйОбъект["Чертеж детали"];
        if (прототип == null)
            return;
        
        var скопированныеОбъекты = СкопироватьОбъект(прототип, папка);
        
        var новыйФайл = скопированныеОбъекты.ПолучитьКопию(прототип); // созданный файл
        if (ТП.Тип.Имя == "Изготовление")
        	новыйФайл["Наименование"] = ТП["Наименование"] + " - " + ТП["Обозначение"] + " - " + ТекущийОбъект["Номер"] + ".grb";
        else
            новыйФайл["Наименование"] = ТП["Наименование"] + " - " + ТП["Обозначение ТП"] + " - " + ТекущийОбъект["Номер"] + ".grb";
        скопированныеОбъекты.Сохранить(); // сохраняем
        
    	ТекущийОбъект.Изменить();
    	ТекущийОбъект.СвязанныйОбъект["Чертеж детали"] = новыйФайл;
    	ТекущийОбъект.Сохранить();

    	FileObject file = (ReferenceObject)новыйФайл as FileObject;
    	file.GetHeadRevision();
    	string filePath = file.LocalPath;
    	
    	CadDocumentProvider provider = CadDocumentProvider.Connect(Context.Connection, ".grb");
    	using (CadDocument document = provider.OpenDocument(filePath, false))      // Открываем Cad-файл в режиме редактирования
    	{
    		document.Open(provider);
    		
    		ImportFilesUIHelper helper = new ImportFilesUIHelper(file.Reference, file.Parent, false);     // Открываем приложение
            helper.Edit(file);
    	}
    }
    
    public void СоздатьНаОсновеПредыдущей()
    {
        var папка = НайтиОбъект("Файлы", "Наименование", "Эскизы технологии"); // Находим папку для сохранения файла.
        if (папка == null)
            return;
        
        Объект ТП = null;
        if (ТекущийОбъект.РодительскийОбъект.Тип != "Технологический процесс")
        	if (ТекущийОбъект.РодительскийОбъект.РодительскийОбъект.Тип == "Технологический процесс")
        	    ТП = ТекущийОбъект.РодительскийОбъект.РодительскийОбъект;
        else if (ТекущийОбъект.РодительскийОбъект.Тип == "Технологический процесс")
        	ТП = ТекущийОбъект.РодительскийОбъект;
        
        int num = 0;
        Объект прототип = null;
        if (ТекущийОбъект.РодительскийОбъект.Тип == "Цехопереход")
        {
        	//num = (int)ТекущийОбъект.РодительскийОбъект["Номер"];
        	num = (int)ТекущийОбъект["№"];
        	if (num > 1)
            {
            	Объект операция = ТекущийОбъект.РодительскийОбъект.ДочерниеОбъекты.FirstOrDefault(t => (int)t["№"] == num - 1);
            	прототип = операция.СвязанныйОбъект["Чертеж детали"];
                if (прототип == null)
                    return;
            }
        }
        /*else
        {
        	num = (int)ТекущийОбъект["Номер"];
        	if (num > 1)
            {
            	Объект операция = ТП.ДочерниеОбъекты.FirstOrDefault(t => (int)t["Номер"] == num - 1);
            	прототип = операция.СвязанныйОбъект["Чертеж детали"];
                if (прототип == null)
                    return;
            }
        }*/
        
        var скопированныеОбъекты = СкопироватьОбъект(прототип, папка);
        
        var новыйФайл = скопированныеОбъекты.ПолучитьКопию(прототип); // созданный файл
        новыйФайл["Наименование"] = ТП["Наименование"] + " - " + ТП["Обозначение ТП"] + " - " + ТекущийОбъект["Номер"] + ".grb";
        скопированныеОбъекты.Сохранить(); // сохраняем
        
    	ТекущийОбъект.Изменить();
    	ТекущийОбъект.СвязанныйОбъект["Чертеж детали"] = новыйФайл;
    	ТекущийОбъект.Сохранить();

    	FileObject file = (ReferenceObject)новыйФайл as FileObject;
    	file.GetHeadRevision();
    	string filePath = file.LocalPath;
    	
    	CadDocumentProvider provider = CadDocumentProvider.Connect(Context.Connection, ".grb");
    	using (CadDocument document = provider.OpenDocument(filePath, false))      // Открываем Cad-файл в режиме редактирования
    	{
    		document.Open(provider);
    		
    		ImportFilesUIHelper helper = new ImportFilesUIHelper(file.Reference, file.Parent, false);     // Открываем приложение
            helper.Edit(file);
    	}
    }
}
