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
    public Macro(MacroContext context)
        : base(context)
    {
    }


    /* Данный макрос нужен для того, чтобы экспортировать текст всех макросов
    в текстовые файлы */
    public override void Run()
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
            Наименование = макрос.Параметр["Наименование"];
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
}
