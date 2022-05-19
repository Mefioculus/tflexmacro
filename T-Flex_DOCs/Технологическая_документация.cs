using System;
using TFlex.DOCs.Model.Macros;
using TFlex.DOCs.Model.Macros.ObjectModel;
using TFlex.DOCs.Model.References;
using TFlex.Model.Technology.Macros;
using TFlex.Model.Technology.Macros.ObjectModel;
using TFlex.Model.Technology.References.SetOfDocuments;

namespace PDM
{
    public class Технологическая_документация : TechnologyMacroProvider
    {
        private int _companyCodeLength;
        private int _docCodeLength;

        public Технологическая_документация(MacroContext context) : base(context)
        {
            _companyCodeLength = ГлобальныйПараметр["Использовать код предприятия"] == true ? ГлобальныйПараметр["Код предприятия"].ToString().Length + 1 : 0;
            _docCodeLength = 11 + _companyCodeLength;
        }

        public override void Run()
        {
        }

        public void НазначитьОбозначение()
        {
            ДиалогВводаТехнологии диалог = СоздатьДиалогТехнологии("Обозначение технологического документа");
            диалог.Изменить();

            string code = ТекущийОбъект.Параметр["Обозначение"];
            if (code.Length == _docCodeLength)
            {
                диалог["Код предприятия"] = ГлобальныйПараметр["Код предприятия"];
                диалог["Код вида документации"] = code.Substring(_companyCodeLength, 2);
                диалог["Код вида технологического процесса"] = code.Substring(_companyCodeLength + 2, 1);
                диалог["Код вида обработки"] = code.Substring(_companyCodeLength + 3, 2);
                диалог["Номер"] = code.Substring(_companyCodeLength + 6, 5);
                диалог["Обозначение документа"] = ТекущийОбъект.Параметр["Обозначение"];
            }
            else
            {
                диалог["Код предприятия"] = ГлобальныйПараметр["Код предприятия"];
                диалог["Код вида документации"] = "";
                диалог["Код вида технологического процесса"] = "";
                диалог["Код вида обработки"] = "";

                диалог["Вид документации"] = "";
                диалог["Вид технологического процесса"] = "";
                диалог["Вид обработки"] = "";

                диалог["Номер"] = "";
                диалог["Обозначение документа"] = "";
            }

            диалог["Наименование"] = ТекущийОбъект["Наименование"];
            диалог.Сохранить();

            if (диалог.ПоказатьДиалог())
            {
                ТекущийОбъект["Обозначение"] = диалог["Обозначение документа"];
                ТекущийОбъект["Наименование"] = диалог["Наименование"];
            }
        }

        public void Нумеровать()
        {
            string код = ВычислитьОбозначениеДокумента(ТекущийОбъект["Код вида документации"], ТекущийОбъект["Код вида технологического процесса"], ТекущийОбъект["Код вида обработки"]);

            if (String.IsNullOrEmpty(код))
                return;

            Объект запись = НайтиОбъект("Реестр обозначений технологической документации", "Код", код);

            int number = 1;
            if (запись == null)
            {
                запись = СоздатьОбъект("Реестр обозначений технологической документации", "Запись");
                запись["Код"] = код;
            }
            else
            {
                запись.Изменить();
                number = запись["Номер"] + 1;
            }

            запись["Номер"] = number;
            запись.Сохранить();
                            
            ТекущийОбъект["Номер"] = number.ToString("00000");
            ТекущийОбъект["Обозначение документа"] = код + "." + ТекущийОбъект["Номер"];
            return;
        }

        public void НаследоватьНаименованиеДокумента()
        {
            ТекущийОбъект.Параметр["Наименование"] = ТекущийОбъект.Параметр["Вид документации"];
        }

        private string ВычислитьОбозначениеДокумента(string кодВидаДокументации, string кодВидаТП, string кодПередела, string номер = null)
        {
            string result = String.Empty;

            if (String.IsNullOrEmpty(кодВидаДокументации) || String.IsNullOrEmpty(кодВидаТП) || String.IsNullOrEmpty(кодПередела))
                return result;

            if (ГлобальныйПараметр["Использовать код предприятия"] == true)
                result = ГлобальныйПараметр["Код предприятия"] + ".";

            result = result + кодВидаДокументации + кодВидаТП + кодПередела;
            if (!String.IsNullOrEmpty(номер))
                result = result + "." + номер;

            return result;
        }

        public void ЗавершениеИзмененияСвязиОбъекта()
        {
            if (Context.ChangedLink.LinkGroup.Guid == SetOfDocumentsReferenceObject.RelationKeys.DocSetToReportLink)
                ВычислитьНаименованиеДокумента();
        }

        private void ВычислитьНаименованиеДокумента()
        {
            if ((ReferenceObject)CurrentObject is Document document)
            {
                if (String.IsNullOrEmpty(document.Name) && document.ReportObject != null)
                    document.Name.Value = document.ReportObject.Name;
            }
        }
    }
}
