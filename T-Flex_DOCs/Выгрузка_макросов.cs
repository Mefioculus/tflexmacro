using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using TFlex.DOCs.Model;
using TFlex.DOCs.Model.Macros;
using TFlex.DOCs.Model.References;
using TFlex.DOCs.Model.Macros.ObjectModel;

public class Macro : MacroProvider
{
    

    public string SavePath = "";

    public Macro(MacroContext context)
        : base(context)
    {
#if DEBUG
        System.Diagnostics.Debugger.Launch();
        System.Diagnostics.Debugger.Break();
#endif
    }


    /* Данный макрос нужен для того, чтобы экспортировать текст всех макросов
    в текстовые файлы */
    public override void Run()
    {
        SaveAllMacros();
    }


    public void Выгрузка_макросов()
    {
        // Инициализируем осноные переменные макроса
        string server_name = Context.Connection.InstanceName;
        string Наименование = "";
        string КодПрограммы = "";
        string ПутьКФайлу = "";
        string Путь = $@"C:\macros_tflex\{server_name}\";

        // Проверка наличия пути в системе
        if (Directory.Exists(Путь) == false)
        {
            DirectoryInfo di = Directory.CreateDirectory(Путь);
            Сообщение("Информация.", "Папка для сохранения экспортированных данных не была найдена и была создана по пути " + Путь);
        }

        // Для начала получаем доступ ко всем объектам справочника макросов
        Объекты макросы = НайтиОбъекты("Макросы", "[Наименование] != '0'");

        // Далее, для каждого объекта справочника получаем наименование и код программы
        foreach (var макрос in макросы)
        {
            // Получаем параметры макроса для записи в файлы
            Наименование = макрос.Параметр["Наименование"].ToString();
            Наименование = Наименование.Replace(" ", "_");
            if (Наименование.Contains('"'))
                Наименование = Наименование.Replace("\"", "");
            КодПрограммы = макрос.Параметр["Текст программы"];
            ПутьКФайлу = Путь + Наименование + ".cs";



            try
            {
                using (StreamWriter sw = new StreamWriter(ПутьКФайлу))
                {
                    sw.WriteLine(КодПрограммы);
                }

            }
            catch (Exception e)
            {
                Сообщение("", e.Message);
            }
        }

        Сообщение("", "Экспорт макросов завершен");
    }


    /// <summary>
    /// Сохраняет выбранные макросы.
    /// </summary>
    public void SaveSelectMacrosList()
    {
        GetPath();
        List<ReferenceObject> currentObjects = Context.GetSelectedObjects().ToList();
        SaveMacros(currentObjects);
    }

    public void SaveSelectMacros()
    {
        GetPath();
        ReferenceObject currentObject = Context.ReferenceObject;
        SaveMacros(currentObject);
        Сообщение("", "Экспорт макросов завершен");
    }

    /// <summary>
    /// Сохраняет все макросы из справочника макросы
    /// </summary>
    public void SaveAllMacros()
    {
        // Гуид справочника макросы
        Guid GuidMacrosRef = new Guid("3e6df4d0-b1d8-4375-978c-4da676604cca");
        GetPath();
        var allmacros = GetMacros(GuidMacrosRef);
        SaveMacros(allmacros);
    }

    public List<ReferenceObject> GetMacros(Guid guidref)
    {
        ReferenceInfo info = Context.Connection.ReferenceCatalog.Find(guidref);
        Reference reference = info.CreateReference();
        List<ReferenceObject> result = reference.Objects.ToList();
        return result;
    }

    public void GetPath()
    {
       string server_name = Context.Connection.InstanceName;
       SavePath = $@"C:\macros_tflex\{server_name}\";
       if (Directory.Exists(SavePath) == false)
        {
            DirectoryInfo di = Directory.CreateDirectory(SavePath);
            Сообщение("Информация.", "Папка для сохранения экспортированных данных не была найдена и была создана по пути " + SavePath);
        }
    }

    public void SaveMacros(List<ReferenceObject> currentObjects)
    {
        foreach (var macro in currentObjects)
        {
            SaveMacros(macro);
        }
        Сообщение("", "Экспорт макросов завершен");
    }

    public void SaveMacros(ReferenceObject currentObject)
    {
        string name_macros = "";
        string text_macros = "";
        string FullPath = "";

        name_macros = currentObject.GetObjectValue("Наименование").ToString();
        text_macros = currentObject.GetObjectValue("Текст программы").ToString();

        name_macros = name_macros.Replace(" ", "_");
        if (name_macros.Contains('"'))
            name_macros = name_macros.Replace("\"", "");


        FullPath = SavePath + name_macros + ".cs";


        try
        {
            using (StreamWriter sw = new StreamWriter(FullPath))
            {
                sw.WriteLine(text_macros);
            }

        }
        catch (Exception e)
        {
            Сообщение("", e.Message);
        }
        
    }


}
