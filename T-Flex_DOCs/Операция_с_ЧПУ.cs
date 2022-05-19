/* Ссылки
TFlex.Model.Technology.dll*/

using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using TFlex.DOCs.Model.Macros;
using TFlex.DOCs.Model.Macros.ObjectModel;
using TFlex.DOCs.Model.Desktop;
using TFlex.DOCs.Model.References;
using TFlex.DOCs.Model.References.Files;
using TFlex.Model.Technology.Macros.ObjectModel;

using TFlex.DOCs.Model.FilePreview.CADService;

public class Macro : MacroProvider
{
    public Macro(MacroContext context)
        : base(context)
    {
    }

    public override void Run()
    {
    }
    
  
    
  
    
    public string Заполнить_поле (int поле)
    {
    	string значение = "";
    	if (поле < 1 || поле > 4)
    		return значение;
    	
    	Объект оборудование = ТекущийОбъект.СвязанныеОбъекты["Оборудование"].FirstOrDefault();
    	if (оборудование == null && поле < 4)
    		return значение;
    	
    	switch (поле)
    	{
    		case 1:   //Наименование и марка станка
    			значение = оборудование.Параметр["Наименование"];
    			break;
    		case 2:   //Наименование и марка стойки ЧПУ
    			значение = оборудование.Параметр["Стойка ЧПУ"];
    			break;
    		case 3:   //Файл постпроцессора
    			Объект постпроцессор = оборудование.СвязанныйОбъект["Постпроцессор"];
    			if (постпроцессор != null)
    				значение = постпроцессор.Параметр["Наименование"];
    			break;
    	}
    	
    	значение = значение.Trim();
    	return значение;
    }
    
    private Объект Создать_файл (string folderPath, string наименование, string данные)
    {
    	FileReference fileReference = new FileReference();
    	string computer_path;
    	
    	//Задание целевой папки для импорта файлов
    	FolderObject parentFolder = fileReference.FindByRelativePath(folderPath) as FolderObject;
    	Directory.CreateDirectory(parentFolder.LocalPath);
    	
    	//Проверка наличия файла
    	FileObject fileObject = null;
    	Объект файл = НайтиОбъект("Файлы", "[Относительный путь]", folderPath + "\\" + наименование);
    	if (файл != null)
    	{
    		if (!Вопрос("Файл с указанным именем уже существует. Заменить?"))
    			return null;
    		
    		fileObject = (FileObject)файл;
    		bool редактируется = fileObject.IsCheckedOut;
    		
    		if (!редактируется)
    			fileObject.CheckOut(false);
    		
    		fileObject.GetHeadRevision();
    		
    		//Перезапись файла
    		computer_path = fileObject.LocalPath;
    		using(StreamWriter streamWriter = new StreamWriter(computer_path)) 
            {
                streamWriter.Write(данные);
                streamWriter.Close();
            }
    		
    		//Сохранение файла
    		if (!редактируется)
    			Desktop.CheckIn((DesktopObject)fileObject, "Генерация обменного файла", false);
    	}
    	else
    	{
    		//Формирование файла
    		computer_path = parentFolder.LocalPath + "\\" + наименование;
            if (File.Exists(computer_path))
                File.Delete(computer_path);
            File.AppendAllText(computer_path, данные);
    		
        	//Импорт файла
            fileObject = fileReference.AddFile(computer_path, parentFolder);
            Desktop.CheckIn((DesktopObject)fileObject, "Генерация обменного файла", false);
        }
    	
    	if (fileObject == null)
    		return null;
        return Объект.CreateInstance((ReferenceObject)fileObject, Context);
    }
    
    public void Создать_файл_обработки ()
    {
    	Операция операция = (Операция)ТекущийОбъект;
    	Объект файл = операция.СвязанныйОбъект["Обработка и траектории"];
    	if (файл != null)
    		return;
    	
    	//Поиск файла-источника
    	Объект прототип = НайтиПрототип("Файлы", "[Относительный путь] = 'Прототипы\\Прототип ЧПУ v14.grb'");
    	if (прототип == null)
    	{
    		Сообщение("Создание файла обработки", "Не удалось найти прототип для создания файла обработки и траектории.");
    		return;
    	}
    	
    	//Определение папки
    	string folderPath = ГлобальныйПараметр["Путь ЧПУ"];
    	Объект папка = НайтиОбъект("Файлы", "[Относительный путь]", folderPath);
    	if (папка == null)
    	{
    		Сообщение("Создание файла обработки", 
    		    "Не удалось найти папку, в которую должен быть скопирован файл. Проверьте правильность задания глобального параметра.");
    		return;
    	}
    	
    	//Задание файла обработки и траектории
    	string наименование = Наименование_файла(операция) + ".grb";
    	Объект существующий_файл = НайтиОбъект("Файлы", "[Относительный путь]", folderPath + "\\" + наименование);
    	if (существующий_файл != null)
    	{
    		if (Вопрос("В справочнике 'Файлы' уже существует файл обработки и траектории '" + наименование + "'. Подключить?"))
    			операция.СвязанныйОбъект["Обработка и траектории"] = существующий_файл;
    	}
    	else
    	{
    		var объекты = СкопироватьОбъект(прототип, папка, false, false);
    		Объект новый_файл = объекты.ПолучитьКопию(прототип);
    		новый_файл.Параметр["Наименование"] = наименование;
    		
    		CadDocumentProvider provider = CadDocumentProvider.Connect(Context.Connection, ".grb");
    		FileObject file = (ReferenceObject)новый_файл as FileObject;
    		file.GetHeadRevision();
    	    string filePath = file.LocalPath;
    		using (CadDocument document = provider.OpenDocument(filePath, false))      // Открываем Cad-файл в режиме редактирования
        	{
        		var variables = document.GetVariables();
        		var variable = variables.FirstOrDefault(cadVar => cadVar.Name == "$DOCs_Tools");
        		if (variable == null)
                {
                    document.Close(false);
                    return;
                }
        		if (операция.СвязанныйОбъект["Обменный файл с инструментом"] != null)
        		{
        			FileObject file_tool = (ReferenceObject)операция.СвязанныйОбъект["Обменный файл с инструментом"] as FileObject;
        			file_tool.GetHeadRevision();
        			string file_toolPath = file_tool.LocalPath;
        			variable.Value = file_toolPath;
        		}
        		        		
        		variable = variables.FirstOrDefault(cadVar => cadVar.Name == "$DOCs_MachineTools");     //модель оснастки
        		if (variable == null)
                {
                    document.Close(false);
                    return;
                }
        		

    			
        		variable = variables.FirstOrDefault(cadVar => cadVar.Name == "$DOCs_Part");
        		if (variable == null)
                {
                    document.Close(false);
                    return;
                }
        		if (операция.СвязанныйОбъект["Деталь"] != null)
        			if (операция.СвязанныйОбъект["Деталь"].СвязанныйОбъект["Связанный объект"].СвязанныеОбъекты["Файлы"].Count > 0)
        			{
        			    Объект деталь_файл = операция.СвязанныйОбъект["Деталь"].СвязанныйОбъект["Связанный объект"].СвязанныеОбъекты["Файлы"].FirstOrDefault(t => t["Наименование"].ToString().Contains(".grb"));
        				if (деталь_файл != null)
                		{
                			FileObject file_part = (ReferenceObject)деталь_файл as FileObject;
                			file_part.GetHeadRevision();
                			string file_partPath = file_part.LocalPath;
                			variable.Value = file_partPath;
                		}
        			}
        		
        		variable = variables.FirstOrDefault(cadVar => cadVar.Name == "$DOCs_Blank");
        		if (variable == null)
                {
                    document.Close(false);
                    return;
                }
        		if (операция.СвязанныйОбъект["Заготовка"] != null)
        			if (операция.СвязанныйОбъект["Заготовка"].СвязанныйОбъект["Связанный объект"].СвязанныеОбъекты["Файлы"].Count > 0)
        			{
        			    Объект заготовка_файл = операция.СвязанныйОбъект["Заготовка"].СвязанныйОбъект["Связанный объект"].СвязанныеОбъекты["Файлы"].FirstOrDefault(t => t["Наименование"].ToString().Contains(".grb"));
        				if (заготовка_файл != null)
                		{
                			FileObject file_blank = (ReferenceObject)заготовка_файл as FileObject;
                			file_blank.GetHeadRevision();
                			string file_blankPath = file_blank.LocalPath;
                			variable.Value = file_blankPath;
                		}
        			}
        		
        		variable = variables.FirstOrDefault(cadVar => cadVar.Name == "$DOCs_Postprocessor");
        		if (variable == null)
                {
                    document.Close(false);
                    return;
                }
        		variable.Value = ПутьПостпроцессора(операция);
        		
        		variables.Save();
        		// Закрываем документ, сохраняя его. 
                document.Close(true);
        	}
    		
        	объекты.Сохранить();
        	Desktop.CheckIn((DesktopObject)новый_файл, "Генерация файла обработки и траектории", false);
        	операция.СвязанныйОбъект["Обработка и траектории"] = новый_файл;
    	}
    	
    	//Проверка наличия программы
    	файл = операция.СвязанныйОбъект["Файл УП"];
    	if (файл != null)
    		return;
    	
    	//Задание программы для ЧПУ
    	наименование = Наименование_файла(операция) + ".nc";
    	существующий_файл = НайтиОбъект("Файлы", "[Относительный путь]", folderPath + "\\" + наименование);
    	if (существующий_файл != null)
    	{
    		if (Вопрос("В справочнике 'Файлы' уже существует файл программы ЧПУ '" + наименование + "'. Подключить?"))
    			операция.СвязанныйОбъект["Программа ЧПУ"] = существующий_файл;
    	}
    	else
    	{
    		try
    		{
        	   Объект новый_файл = СоздатьОбъект("Файлы", "Файл \"NC\"", папка);
        	   новый_файл.Параметр["Наименование"] = наименование;
        	   новый_файл.Сохранить();
        	   Desktop.CheckIn((DesktopObject)новый_файл, "Генерация файла программы ЧПУ", false);
        	   операция.СвязанныйОбъект["Файл УП"] = новый_файл;
        	}
    		catch(Exception)
    		{
    		}
    	}
    }
    
    private string Наименование_файла (Операция операция)
    {
    	ТехнологическийПроцесс техпроцесс = операция.ТехнологическийПроцесс;
    	string наименование = техпроцесс.Обозначение + "^" + техпроцесс.Наименование 
    		+ "^" + операция.Номер + " (" + операция["ID"] + ")";
    	return наименование;
    }
    
    private string ПутьПостпроцессора(Операция операция)
    {
    	Объект СТО = операция.СвязанныеОбъекты["Оборудование"].FirstOrDefault();
        string res = "";
        
        if (СТО.СвязанныйОбъект["Постпроцессор"] != null)
        {
            FileObject file_postprocessor = (ReferenceObject)СТО.СвязанныйОбъект["Постпроцессор"] as FileObject;
            file_postprocessor.GetHeadRevision();
            res = file_postprocessor.LocalPath;
        }
    
        return res;
    }
}
