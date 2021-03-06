using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using TFlex.DOCs.Model.Macros;
using TFlex.DOCs.Model.Macros.ObjectModel;
//------------------------------------
using System.Text.RegularExpressions;

public class Macro : MacroProvider
{
    public Macro(MacroContext context)
        : base(context)
    {
    }
    static List<string> имя = new List<string>(); //перечень имен параметров
    static List<string> list = new List<string>(); //сортированный перечень имен параметров
    static List<int> формат = new List<int>();
    static List<string> формулы = new List<string>();
    static List<double> min = new List<double>();
    static List<double> max = new List<double>();
    static List<Объекты> списки = new List<Объекты>();
    static string[] функции = new string[] { " pi ", " e ", " sqrt(", " sin(", " cos(", " tg(", " ctg(", " sh(", " ch(", " th(", " log(", " exp(", " abs(" };
    static string[] шифр = new string[] { "?0?", "?1?", "?2?", "?3?", "?4?", "?5?", "?6?", "?7?", "?8?", "?9?", "?10?", "?11?", "?12?" };
    static string заголовок = "<!DOCTYPE html PUBLIC \"-//W3C//DTD XHTML 1.0 Transitional//EN\" \"http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd\"> " +
                "<html xmlns=\"http://www.w3.org/1999/xhtml\"> " +
                    "<head>" +
                        "<meta http-equiv=\"Content-Type\" content=\"text/html; charset=utf-8\" /><title> " +
                        "</title>" +
                        "<style type=\"text/css\"> " +
                            ".cs2654AE3A{text-align:left;text-indent:0pt;margin:0pt 0pt 0pt 0pt}" +
                            ".csFD04CD6A{color:#FF0000;background-color:transparent;font-family:Calibri;font-size:11pt;font-weight:normal;font-style:normal;}" +  // красный
                            ".csF8E8676A{color:#0000FF;background-color:transparent;font-family:Calibri;font-size:11pt;font-weight:normal;font-style:normal;}" +  // синий
                            ".csF96BE9E{color:#008000;background-color:transparent;font-family:Calibri;font-size:11pt;font-weight:normal;font-style:normal;}" +  // зеленый
                            ".csC8F6D76{color:#000000;background-color:transparent;font-family:Calibri;font-size:11pt;font-weight:normal;font-style:normal;}" +  // черный
                        "</style>" +
                    "</head>" +
                    "<body>";
    static string абзац = "<p class=\"cs2654AE3A\">";
    static string красный = "<span class=\"csFD04CD6A\">";
    static string синий = "<span class=\"csF8E8676A\">";
    static string зеленый = "<span class=\"csF96BE9E\">";
    static string черный = "<span class=\"csC8F6D76\">";
    static string сменацвета = "</span>";
    static string конецабзаца = "</p>";
    static string подвал = "</body></html>";

    public override void Run()
    {

    }

    public void ВычислитьЯчейку()
    {
        //Получить имя колонки для выделенной ячейки
        string name = ВыполнитьМакрос("АРМ.Взять объект ячейку", "ПолучитьКолонку", "Справочник");
        //записи в которых нужно вычислить колонку по выделенной ячейки
        Объекты выделены = ВыполнитьМакрос("АРМ.Взять объект ячейку", "ПолучитьВыбраныеОбъекты", "Справочник");
        Объект параметры = ТекущийОбъект.СвязанныйОбъект["Параметры документа НСИ"];
        Объект вычисляемый = параметры.СвязанныеОбъекты["Описание параметров НСИ"].FirstOrDefault(p => p["Наименование"] == name);
        if (вычисляемый["Тип параметра"] > 2)
            Ошибка("Выделенный параметр " + name + " имеет формат не поддерживающий вычисления по формуле! Используйте формулы для строковых, целых и действительных чисел!");
        Объекты наборпараметров = параметры.СвязанныеОбъекты["Описание параметров НСИ"];
        string формула = вычисляемый["Формула"];
        if (формула == "")
            Ошибка("Для данного параметра формула не задана!");
        foreach (Объект параметр in наборпараметров)
            имя.Add(параметр["Наименование"]);

        //------------------  Составить список имен переменных и отсортировать по длине имени --------------------------
        list = имя.OrderByDescending(item => item.Length).ToList(); //Отсортировать по длине по убыванию
        
        string[] words = формула.Split(new char[] { '+' });
        int i = 0;
        if (вычисляемый["Тип параметра"] != 0)
            формула = Зашифровать(формула); //Получение зашифрованной формулы для математических выражений
        else
        {
        foreach (var word in words)
            {
                words[i] = word.Trim();
                i += 1;                
            }
        }
        
        foreach (Объект запись in выделены)
        {
            //string выражение = формула;
            foreach (string txt in list)
            {
            	//Сообщение("", "txt = " + txt + "\n" + "list = " + list);
            	i = 0;
                foreach (var word in words)
                    {
                	   //Сообщение("word", word);
                	   if (word.IndexOf("\"") == -1)
                	   {
                	   	   words[i] = word.Replace(txt, запись[txt]);
                	   	   //Сообщение("А мы зашли и возможно заменили", words[i]);
                	   }
                	   i += 1;
                    }
                //выражение = выражение.Replace(txt, запись[txt]);
            }
            string выражениеце = String.Join("", words);
            string выражение = Regex.Replace(выражениеце,"\"", "");

            //результат по одной записи
            запись.Изменить();
            if (вычисляемый["Тип параметра"] == 0)
                запись[name] = выражение;
            else
            {
                выражение = Расшифровать(выражение);
                запись[name] = Результат(выражение);
            }
            запись.Сохранить();
        }

    }

    public void ВычислитьЯчейку_АРМ()
    {
        //Получить имя колонки для выделенной ячейки
        string name = ВыполнитьМакрос("АРМ.Взять объект ячейку", "ПолучитьКолонку", "Записи");
        //записи в которых нужно вычислить колонку по выделенной ячейке
        Объекты выделены = ВыполнитьМакрос("АРМ.Взять объект ячейку", "ПолучитьВыбраныеОбъекты", "Записи");
        //Объект параметры = ТекущийОбъект.СвязанныйОбъект["Параметры документа НСИ"];
        var справочник = выделены.FirstOrDefault().Справочник;
        Объект документ = НайтиОбъект("Документы НСИ", "Guid сгенерированного справочника", справочник.УникальныйИдентификатор);
        Объект параметры = документ.СвязанныйОбъект["Параметры документа НСИ"];

        Объект вычисляемый = параметры.СвязанныеОбъекты["Описание параметров НСИ"].FirstOrDefault(p => p["Наименование"] == name);
        if (вычисляемый["Тип параметра"] > 2)
            Ошибка("Выделенный параметр " + name + " имеет формат не поддерживающий вычисления по формуле! Используйте формулы для строковых, целых и действительных чисел!");
        Объекты наборпараметров = параметры.СвязанныеОбъекты["Описание параметров НСИ"];
        string формула = вычисляемый["Формула"];
        
        if (формула == "")
            Ошибка("Для данного параметра формула не задана!");
        foreach (Объект параметр in наборпараметров)
            имя.Add(параметр["Наименование"]);
        
        //------------------  Составить список имен переменных и отсортировать по длине имени --------------------------
        list = имя.OrderByDescending(item => item.Length).ToList(); //Отсортировать по длине по убыванию
        list.ForEach(Console.WriteLine);
        string[] words = формула.Split(new char[] { '+' });
        int i = 0;
        if (вычисляемый["Тип параметра"] != 0)
            формула = Зашифровать(формула); //Получение зашифрованной формулы для математических выражений
        else
        {
        foreach (var word in words)
            {
                words[i] = word.Trim();
                i += 1;                
            }
        }
        foreach (Объект запись in выделены)
        {
            //string выражение = формула;
            foreach (string txt in list)
            {
            	//Сообщение("", "txt = " + txt + "\n" + "list = " + list);
            	i = 0;
                foreach (var word in words)
                    {
                	   //Сообщение("word", word);
                	   if (word.IndexOf("\"") == -1)
                	   {
                	   	   words[i] = word.Replace(txt, запись[txt]);
                	   	   //Сообщение("А мы зашли и возможно заменили", words[i]);
                	   }
                	   i += 1;
                    }
                //выражение = выражение.Replace(txt, запись[txt]);
            }
            string выражениеце = String.Join("", words);
            string выражение = Regex.Replace(выражениеце,"\"", "");
            //Сообщение("Итог", выражение);
            //результат по одной записи
            запись.Изменить();
            if (вычисляемый["Тип параметра"] == 0)
                запись[name] = выражение;
            else
            {
                выражение = Расшифровать(выражение);
                запись[name] = Результат(выражение);
            }
            запись.Сохранить();
        }

    }

    private string Зашифровать(string формула)
    {
        //------------------  Составить список имен переменных и отсортировать по длине имени --------------------------
        //list = имя.OrderByDescending(item => item.Length).ToList(); //Отсортировать по длине по убыванию
        //------------------------------------------------------------------------------------------------------------    	
        ///зашифровать имена функций и констант для математических выражений

        for (int i = 0; i < 13; i++)
            формула = формула.Replace(функции[i], шифр[i]);  //Шифрованная формула
        return формула;
    }

    private string Расшифровать(string формула)
    {
        //Проверить на наличие символов алфавита в шифрованной формуле
        if (Regex.IsMatch(формула, "[a-zA-Zа-яА-я]"))
            Ошибка("Ошибка! Имеются нераспознаваемые имена переменных в выражении:  " + формула);
        ///расшифровать имена функций и констант
        for (int i = 0; i < 13; i++)
            формула = формула.Replace(шифр[i], функции[i]);
        return формула;
    }

    public double Результат(string выражение)
    {
        var mathParser = new MathParser.MathParser();
        double result = MathParser.MathParser.Вычислить(выражение);
        return result;
    }

    public void Справка()
    {
        Сообщение("Правила написания формул: ",
              "Правила написания строковых формул: " +
              "\n 1. Строковая формула пишется в виде маски состоящей из статического текста и имён переменных" +
              "\n 2. Имена переменных при вычислении формулы будут заменены на их значения" +
              "\n 3. Последовательность символов в статическом тексте на должна совпадать с именами переменных" +
              "\n 4. Для наименования переменных рекомендуется использовать латиницу (особенно для коротких имен из нескольких символов)" +
              "\n  " +
              "\n Правила написания математических выражений" +
              "\n 1. Имена переменных и констант (pi, e) должны быть отделены пробелами от символов" +
              "\n 2. Имена переменных не должны совпадать с названием констант (pi, e)" +
              "\n 3. Между именем функции и открывающейся скобкой не должно быть пробела, например,  sin(..." +
              "\n 4. В выражении не должен использоваться вопросительный знак (?)" +
              "\n 5. Перечень поддерживаемых функций:" +
              "\n    sqrt()  - Корень квадратный" +
              "\n    sin()   - Синус" +
              "\n    cos()   - Косинус" +
              "\n    tg()    - Тангенс" +
              "\n    ctg()   - Котангенс" +
              "\n    sh()    - Гиперболический синус" +
              "\n    ch()    - Гиперболический косинус" +
              "\n    th()    - Гиперболический тангенс" +
              "\n    log()   - Логарифм" +
              "\n    exp()   - Экспонента" +
              "\n    abs()   - Модуль" +
              "\n 6. Поддерживаемые символы в выражениях:" +
              "\n    +; -; *; /; ^ степень; (; ); √ корень квадратный "
             );
    }

    public void АРМ_ПроверкаЗначений()
    {
        //записи в которых нужно проверить значения
        //Объекты выделены = ВыполнитьМакрос("АРМ.Взять объект ячейку", "ПолучитьВыбраныеОбъекты", "Записи");
        Объекты записи = ВыполнитьМакрос("АРМ.Взять объект ячейку", "ПолучитьВыбраныеОбъекты", "Записи");
        //----------------------------Поправить Нужно реализовать получение врем енного справочника  

        Объект донор = записи.FirstOrDefault();
        //Объект временный = донор.Справочник;
        Guid справочник = донор.Справочник.УникальныйИдентификатор;
        //--------------------------    	
        Объект документНСИ = НайтиОбъект("Документы НСИ", "Guid сгенерированного справочника", справочник); //Найти документ НСИ, если можно, из контекста рабочей страницы
        if (документНСИ == null)
            return;
        Объект списокПараметров = документНСИ.СвязанныйОбъект["Параметры документа НСИ"];
        if (списокПараметров == null)
            return;
        Объекты параметры = списокПараметров.СвязанныеОбъекты["Описание параметров НСИ"];
        if (параметры.Count == 0)
            return;

        string формула = "";

        foreach (Объект параметр in параметры) //Включить все правила в списки
        {
            //записать имя параметра
            имя.Add(параметр["Наименование"]);
            //записать тип параметра
            формат.Add(параметр["Тип параметра"]);
            //Записать зашифрованные формулы для параметров
            формула = параметр["Формула"];
            if (формула != "") //Если есть формула, то зашифровать формулу и в список
            {
                if (параметр["Тип параметра"] == 1 || параметр["Тип параметра"] == 2)
                    формула = Зашифровать(формула); //Получение зашифрованной формулы для математических выражений
            }
            else
                формула = " ";
            формулы.Add(формула);
            // Записать диапазоны для параметров
            if (параметр["Использовать диапазон (Min)"] == true)
                min.Add(параметр["Диапазон значений (Min)"]);
            else
                min.Add(9999999);

            if (параметр["Использовать диапазон (Max)"] == true)
                max.Add(параметр["Диапазон значений (Max)"]);
            else
                max.Add(9999999);

            // Записать списки значений параметров
            if (параметр.СвязанныеОбъекты["Список значений"].Count() != 0)
                списки.Add(параметр.СвязанныеОбъекты["Список значений"]);
            else
                списки.Add(null);
        }
        // Проверка параметров у всех выделенных записей	
        int j = 0;
        foreach (Объект запись in записи)
        {
            string строка = ""; //абзац + черный + запись["ID"].ToString() ;

            int i = 0;
            foreach (Объект параметр in параметры) //Проверить все параметры 
            {
                строка = строка + абзац + черный + имя[i] + " Формула-";

                //проверить по формуле
                if (формулы[i] != " ")
                {
                    формула = формулы[i];
                    foreach (string txt in list)
                        формула = формула.Replace(txt, запись[txt]);

                    if (формат[i] != 0)
                    {
                        формула = Расшифровать(формула);
                        формула = Результат(формула).ToString();
                    }
                    if (запись[имя[i]] != формула)
                        строка = строка + сменацвета + красный + "(Вычислено != Задано)->(" + формула + " != " + запись[имя[i]] + ");" + сменацвета + конецабзаца + абзац + черный + " Диапазон-";
                    else
                        строка = строка + "ОК;" + сменацвета + конецабзаца + абзац + черный + "Диапазон-";
                }
                else
                    строка = строка + "Нет;" + сменацвета + конецабзаца + абзац + черный + "Диапазон-";

                //Проверить диапазон
                if (min[i] == 9999999 && max[i] == 9999999)
                    строка = строка + "Нет;" + сменацвета + конецабзаца + абзац + черный + "Значение-";
                else if (min[i] == 9999999 && min[i] > запись[имя[i]])
                    строка = строка + сменацвета + красный + " вне диапазона;" + сменацвета + черный + " Значение-";
                else if (max[i] == 9999999 && max[i] < запись[имя[i]])
                    строка = строка + сменацвета + красный + " вне диапазона;" + сменацвета + черный + " Значение-";
                else
                    строка = строка + сменацвета + зеленый + " ОК;" + сменацвета + конецабзаца + абзац + черный + "Значение-";
                //Сообщение ("Цикл ", запись.ToString());  
                // if (списки[j] == null)
                // Сообщение("списки", j + " - " ); //+ списки[j].ToString());
                //Проверить по списку значений

                if (списки[j] == null)
                    строка = строка + " допустимые значения не заданы;" + сменацвета + конецабзаца;
                else
                {
                    Объект совпал = списки[i].FirstOrDefault(p => p["Значение"] == запись[имя[i]]);
                    if (совпал == null)
                        строка = строка + сменацвета + красный + " не соответсвует списку допустимых значений;" + сменацвета + конецабзаца;
                    else
                        строка = строка + сменацвета + зеленый + " соответсвует списку допустимых значений;" + сменацвета + конецабзаца;
                }

                i++;
            }

            запись.Изменить();
            запись["Анализ символов"] = заголовок + строка + подвал;
            запись.Сохранить();
            j++;
        }
    }
    //--------------------------------------

}
