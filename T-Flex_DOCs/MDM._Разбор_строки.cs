using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text.RegularExpressions;
using TFlex.DOCs.Model.Macros;
using TFlex.DOCs.Model.Macros.ObjectModel;

public class Macro : MacroProvider
{
    public Macro(MacroContext context)
        : base(context)
    {
       // if (Вопрос("Хотите запустить в режиме отладки?"))
        //{
        //    System.Diagnostics.Debugger.Launch();
          //  System.Diagnostics.Debugger.Break();
        //}
    }
    public override void Run() { }

    private string Депробел(string s)
    {
        var i = 0; var k = s.Length - 1;
        while ((i < k) && (s[i] == ' ')) i++;
        while ((i < k) && (s[k] == ' ')) k--;
        if (i == k) return "";
        var issp = false;
        var t = new char[k - i + 1];
        var acnt = 0;
        for (int u = i; u <= k; u++)
        {
            var ch = s[u];
            if (ch != ' ') { t[acnt++] = ch; issp = false; }
            else if (!issp) { t[acnt++] = ch; issp = true; }
        }
        Array.Resize(ref t, acnt);
        return new string(t);
    }

    private void АРМ_Извлечь_параметр(string строка, Объекты правила, Объект запись, string имя, int формат)
    {
        string temp = "";//ненужный обрезок строки, если мы режем "до"
        Объект исхПараметр = правила.FirstOrDefault(obj => !string.IsNullOrEmpty(obj["Исходный параметр"]));
        if (исхПараметр != null)//если есть параметр Исходный, режем данные из столбца с таким названием
            строка = запись[исхПараметр["Исходный параметр"]];

        foreach (Объект правило in правила)
        {
            int caseSwitch = правило["Правило"];//переменная для switch с Правилом
            bool del_with_substr = правило["Удалить с подстрокой"];
            string substring = "";//подстрока, относительно которой будет производиться обрезка строки, == шаблону по смыслу, если 3 кейс
            int amt = -1;//номер по счету "Опорной" подстроки
            amt = правило["По счету"];

            if (правило["Подстрока"] == "Пробел")
                substring = " ";
            else
                substring = (правило["Подстрока"]);
            if (substring == "")
                Ошибка("Пустая подстрока.");

            int n = строка.IndexOf(substring);
            List<int> list_of_index = new List<int>();//{ n }; это список всех индексов вхождения подстроки
            if (n != -1 && (caseSwitch == 1 || caseSwitch == 2))
            {
                while (n != -1)
                {
                    list_of_index.Add(n);
                    n = строка.IndexOf(substring, n + substring.Length);
                }
            }
            else if (n == -1 && (caseSwitch == 1 || caseSwitch == 2))
            {
                Ошибка("Указанный символ \"" + substring + "\" не найден.");
            }

            if (temp != "")
            {
                int amount = new Regex(substring).Matches(temp).Count;//кол-во подстрок, которые удалились при обрезке "до"
                amt -= amount;
            }
            if ((caseSwitch == 1 || caseSwitch == 2) && (amt <= 0 || amt > list_of_index.Count))
                Ошибка("Неверно указано правило \"По счету\".");

            Объект список, значение;
            switch (caseSwitch)
            {
                case 1://Обрезать ДО подстроки
                    if (del_with_substr)//если обрезаем с подстрокой
                    {
                        temp = строка.Substring(0, list_of_index[amt - 1] + substring.Length);
                        строка = строка.Substring(list_of_index[amt - 1] + substring.Length);
                    }
                    else
                    {
                        temp = строка.Substring(0, list_of_index[amt - 1]);
                        строка = строка.Substring(list_of_index[amt - 1]);
                    }
                    break;
                case 2://Обрезать ПОСЛЕ подстроки
                    if (del_with_substr)//если обрезаем с подстрокой
                        строка = строка.Substring(0, list_of_index[amt - 1]);
                    else
                        строка = строка.Substring(0, list_of_index[amt - 1] + substring.Length);
                    break;
                case 3://Наименование покрытия (определить по параметру)
                    string значение_шаблона = запись[substring];//запись[шаблон]
                    if (значение_шаблона == "")
                        Сообщение("Предупреждение", "Убедитесь в правильности порядка параметров.");
                    else
                    {
                        список = НайтиОбъект("Списки допустимых значений", "Наименование", "Покрытие");

                        значение = список.СвязанныеОбъекты["Допустимые значения"].FirstOrDefault(Об => Об["Значение"] == значение_шаблона);
                        строка = значение["Наименование"];
                    }
                    break;
                case 4://По подстроке
                    if (n != -1)
                        строка = правило["Результат"];
                    break;
                default:
                    break;
            }
        }
        try
        {
            запись.Изменить();
            switch (формат)
            {
                case 24://0
                    запись[имя] = строка;
                    break;
                case 3://1
                    int resint = int.Parse(строка);
                    запись[имя] = resint;
                    break;
                case 4://2
                    double resdouble = double.Parse(строка);
                    запись[имя] = resdouble;
                    break;
                case 15://3
                    DateTime restime = DateTime.Parse(строка);
                    запись[имя] = restime;
                    break;
                case 5://4
                    bool resbool = bool.Parse(строка);
                    запись[имя] = resbool;
                    break;
            }
            запись.Сохранить();
        }
        catch (System.FormatException)
        {
            Ошибка("Несоответствие данных параметру.");
        }
    }

    public void АРМ_Извлечь_параметры()
    {
        Объекты записи = ВыполнитьМакрос("ea0d8d3c-395b-48d5-9d42-dab98371522c", "ПолучитьВыбраныеОбъекты", "Записи");//массив столбцов
        var справочник = записи.FirstOrDefault().Справочник;
        Объект документ = НайтиОбъект("Документы НСИ", "Guid сгенерированного справочника", справочник.УникальныйИдентификатор);
        Объект параметры = документ.СвязанныйОбъект["Параметры документа НСИ"];
        int формат;
        foreach (Объект запись in записи)
        {
            string имя = ВыполнитьМакрос("ea0d8d3c-395b-48d5-9d42-dab98371522c", "ПолучитьКолонку", "Записи");
            string строка = запись["Сводное наименование"];
            строка = Депробел(строка);//удаляем лишние пробелы
            if (имя == "Сводное наименование")
            {
                Объекты имена_параметров = параметры.СвязанныеОбъекты["Описание параметров НСИ"];
                foreach (Объект имя_параметра in имена_параметров)
                {
                    //Объекты правила = имя_параметра.СвязанныеОбъекты["Список правил разбора строки"];
                    Объекты правила = имя_параметра.СвязанныеОбъекты["Правила разбора строки"];
                    формат = имя_параметра["Тип параметра"];
                    имя = имя_параметра["8a2be7ac-8953-4330-b6e1-cb7a8fbaf5c1"];//[наименование]
                    if (правила.Count != 0)
                        АРМ_Извлечь_параметр(строка, правила, запись, имя, формат);
                }
            }
            else
            {
                Объект параметр = параметры.СвязанныеОбъекты["Описание параметров НСИ"].FirstOrDefault(об => об["Наименование"] == имя);

                if (параметр == null)
                    Ошибка("Параметр \"" + имя + "\" не найден.");

                Объекты правила = параметр.СвязанныеОбъекты["Правила разбора строки"];
                формат = параметр["Тип параметра"];
                АРМ_Извлечь_параметр(строка, правила, запись, имя, формат);
            }
        }
        ОбновитьЭлементыУправления("Записи");
    }
}
