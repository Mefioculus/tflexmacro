using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection.Emit;
//using System.Windows.Forms;
using TFlex.DOCs.Model.Macros;
using TFlex.DOCs.Model.Macros.ObjectModel;
using TFlex.DOCs.Model.References;

public class Macro : MacroProvider
{
    public Macro(MacroContext context)
        : base(context)
    {
        //if (Вопрос("Хотите запустить в режиме отладки?"))
        //{
          // System.Diagnostics.Debugger.Launch();
          // System.Diagnostics.Debugger.Break();
        //}
    }

    public override void Run()
    {
        //ТекущийОбъект["Наименование"];
        //ВыполнитьМакрос("1df791b4-66d1-4382-9812-170ed6ae6220");
        //DateTime fileTime = DateTime.Parse(ТекущийОбъект["Дата последнего изменения"].ToString());

        if (ТекущийОбъект["Стадия"] == "Утверждено")
        {
            DateTime startTime = DateTime.Now;

            //ВыполнитьМакрос("92e6443b-c661-4853-af7c-93e9a01e500a");
            var цехопереходы = ВыбранныеОбъекты.SelectMany(тп => тп.ДочерниеОбъекты).Where(цп => цп.Тип.ПорожденОт("Цехопереход")).ToArray();
            ОбменДанными.ИмпортироватьОбъекты("Выгрузка маршрутов в FoxPro", цехопереходы);
            ОбменДанными.ИмпортироватьОбъекты("Выгрузка ведомости материалов в FoxPro 2", ВыбранныеОбъекты);

            string обозначение = ТекущийОбъект["Обозначение"].ToString().Replace(".", "");
            string фильтр = String.Format("[SHIFR] = '{0}", обозначение);
            Объекты marchpOUT = НайтиОбъекты("MARCHP OUT", фильтр);
            var newMarchpOUT = (Объекты)marchpOUT.Where(obj => DateTime.Parse(obj["Дата последнего изменения"].ToString()) > startTime).ToList();

            Объекты normOUT = НайтиОбъекты("NORM OUT", фильтр);
            var newNormOUT = (Объекты)normOUT.Where(obj => DateTime.Parse(obj["Дата последнего изменения"].ToString()) > startTime).ToList();
            ///////
            //string результат = string.Empty;

            //foreach (var item in newMarchpOUT)
            //{
            //    результат += item.ToString() + Environment.NewLine;
            //}
            //foreach (var item1 in newNormOUT)
            //{
            //    результат += item1.ToString() + Environment.NewLine;
            //}
            
            //Сообщение("Объекты, сформированные во втором пункте", результат);
            ///////
            ОбменДанными.ЭкспортироватьОбъекты("Выгрузка маршрутов в базу данных ", newMarchpOUT);
            ОбменДанными.ЭкспортироватьОбъекты("Выгрузка ведомости материалов в базу данных", newMarchpOUT);
            ОбменДанными.ЭкспортироватьОбъекты("Выгрузка маршрутов в базу данных ", newNormOUT);
            ОбменДанными.ЭкспортироватьОбъекты("Выгрузка ведомости материалов в базу данных", newNormOUT);
        }
    }

    public void ByTimer()
    {
        DateTime startTime = DateTime.Now.AddHours(-1);

        Объекты объекты = НайтиОбъекты("SPEC_OUT", "[Id] > 0");
        Объекты заПоследнийЧас = (Объекты)объекты.Where(obj => DateTime.Parse(obj["Дата последнего изменения"].ToString()) > startTime);

        ОбменДанными.ИмпортироватьОбъекты("Выгрузка в базу данных спецификации", заПоследнийЧас);
    }
}

