/* Ссылки
TFlex.Model.Technology.dll
*/

using TFlex.DOCs.Model.Macros;
using TFlex.DOCs.Model.Macros.ObjectModel;
using TFlex.Model.Technology.Macros.ObjectModel;

namespace TechnologicalMaps
{
    public class TechnologicalMapsName : MacroProvider
    {
        public TechnologicalMapsName(MacroContext context)
            : base(context)
        {
        }

        public override void Run()
        {
        }

        //ВыполнитьМакрос("Технологические карты (Наименование)", "НаименованиеОтчёта")

        public string НаименованиеОтчёта()
        {
            Объект объект = ТекущийОбъект;

            //Техпроцесс
            if (объект.Тип.ПорожденОт("Базовый технологический процесс"))
            {
                ТехнологическийПроцесс техпроцесс = (ТехнологическийПроцесс)объект;
                return ОтчётНаТехпроцесс(техпроцесс);
            }

            //Операция
            if (объект.Тип.ПорожденОт("Операция") || объект.Тип.ПорожденОт("Технологическая операция"))
            {
                Операция операция = (Операция)объект;
                return ОтчётНаОперацию(операция);
            }

            //Объект электронной структуры
            if (объект.Тип.ПорожденОт(TFlex.DOCs.Model.References.Nomenclature.NomenclatureTypes.Keys.Object))
                return объект.ToString();

            return "[не задано]";
        }

        private string ОтчётНаТехпроцесс(ТехнологическийПроцесс техпроцесс)
        {
            string name = техпроцесс.Наименование + " " + техпроцесс.Обозначение;
            return name;
        }

        private string ОтчётНаОперацию(Операция операция)
        {
            string name = операция.ТехнологическийПроцесс.Наименование + " №" + операция.Номер + " " + операция.Наименование;
            return name;
        }
    }
}
