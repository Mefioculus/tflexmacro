using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using TFlex.DOCs.Model.Macros;
using TFlex.DOCs.Model.Macros.ObjectModel;

using TFlex.DOCs.Model.References.Files;

public class Macro : MacroProvider
{
    public Macro(MacroContext context)
        : base(context)
    {
    }

    public override void Run()
    {
    
    }
    
    
    public void Запустить_модуль ()
    {
    	Объект приложение = ТекущийОбъект;
    	string модуль = приложение.Параметр["Исполняемый модуль"];
    	Объект папка = приложение.СвязанныйОбъект["Папка приложения"];
    	Запустить_файл (папка, модуль);
    }
    
    public void Открыть_руководство ()
    {
    	Объект приложение = ТекущийОбъект;
    	string руководство = приложение.Параметр["Руководство пользователя"];
    	Объект папка = приложение.СвязанныйОбъект["Папка приложения"];
    	string relativePath = Определить_путь (папка, руководство);
    	Открыть_файл (relativePath);
    }
    
    
    public void Запустить_файл (Объект папка, string имя_файла)
    {
    	string relativePath = Определить_путь (папка, имя_файла);
    	if (relativePath != string.Empty)
    	{
    		Обновить_содержимое_папки (папка);
    		Открыть_файл (relativePath);
    	}
    	else
    		Ошибка ("Файл приложения не найден.");
    }
    
    
    public string Определить_путь (Объект папка, string имя_файла)
    {
    	string relativePath = string.Empty;
    	if (папка != null)
    	{
    		string path = папка.Параметр["Относительный путь"];
    		Объект файл = НайтиОбъект("Файлы", "[Относительный путь] начинается с '" + path + "' И [Наименование] = '" + имя_файла + "'");
    		if (файл != null)
    			relativePath = файл.Параметр["Относительный путь"];
    	}
    	return relativePath;
    }
    
    
    public void Обновить_содержимое_папки (Объект папка)
    {
    	if (папка != null)
    	{
    		string path = папка.Параметр["Относительный путь"];
    		FileReference reference = new FileReference();
    		FileReferenceObject file = reference.FindByRelativePath(path);
    		if (file != null)
    			file.GetHeadRevision();
    	}
    }
    
    
    public void Открыть_файл (string relativePath)
    {
    	FileReference reference = new FileReference();
    	FileReferenceObject file = reference.FindByRelativePath(relativePath);
    	if (file != null)
    	{
	    	file.GetHeadRevision();
	    	string path = file.LocalPath;
	    	System.Diagnostics.Process.Start(path);
	    }
    	else
    		Ошибка ("Файл по относительному пути '" + relativePath + "' не найден.");
    }
}
