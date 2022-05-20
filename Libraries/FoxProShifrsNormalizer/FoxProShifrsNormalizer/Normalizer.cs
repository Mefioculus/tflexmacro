using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FoxProShifrsNormalizer
{
    public class Normalizer
    {

        public string prefix = "МК-";
        public bool setprefix = false;
        public string NormalizeShifrsFromFox(string shifr)
        {
            // Метод для преобразования обозначений, полученных из Fox
            // Список исключений, в которых не нужно производить нормализации
            if (shifr.ToLower().Contains("северо"))
                return shifr;

            // Создаем переменную для того, чтобы можно было использовать метод int.TryParse()
            long shifrInt;

            // Узнаем, цифра ли содержится в значении
            bool isNum = long.TryParse(shifr, out shifrInt);

            // Если входные данные не являются числовым значением и содержат в себе символ "-", и часть строки до тире состоит из шести или из семи символов
            if (!isNum && shifr.IndexOf('-') > 0 && (shifr.Substring(0, shifr.IndexOf('-')).Length == 6 || shifr.Substring(0, shifr.IndexOf('-')).Length == 7))
            {
                isNum = long.TryParse(shifr.Substring(0, shifr.IndexOf('-')), out shifrInt);
                shifr = shifr.Insert(3, ".");
            }

            // Случай, когда по определению обозначение не состоит из цифровых значений
            if ((shifr.StartsWith("УЯИС") || shifr.StartsWith("ШЖИФ") || shifr.StartsWith("УЖИЯ")) && shifr.Length > 11)
            {
                shifr = shifr.Insert(4, ".").Insert(11, ".");
            }

            if (shifr.StartsWith("3905") && shifr.Length > 4)
            {

                shifr = shifr.Insert(4, ".");

            }

            if (shifr.StartsWith("8А") && (shifr.Length > 7))
            {
                shifr = shifr.Insert(3, ".").Insert(7, ".");
            }

            // В данном случае рассматривается вариант, когда значение полностью состоит из цифр
            if (isNum && (shifr.Length == 6 || shifr.Length == 7))
            {
                shifr = shifr.Insert(3, ".");
            }

            if (setprefix)
            {
                shifr = shifr.Insert(0, prefix);
            }

            return shifr;
        }

        public string NormalizeShifrsToFox(string shifr)
        {
            // Метод для преобразования обозначений для передачи в Fox
            if (!setprefix && shifr.StartsWith(prefix))
            {
                shifr = shifr.Substring(prefix.Length, shifr.Length - prefix.Length);
            }

            return shifr.Replace(".", "");
        }
    }
}
