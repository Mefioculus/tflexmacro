using System;
using System.Linq;
using TFlex.DOCs.Model.Macros;
using TFlex.DOCs.Model.Macros.ObjectModel;
using TFlex.DOCs.Model.References.BomSections;
using TFlex.DOCs.Model.References.Nomenclature;
using TFlex.DOCs.Model.Search;
using TFlex.DOCs.Model.Structure;
using System.Collections.Generic;
using TFlex.DOCs.Model.References;

namespace PDM_DenotationFNN
{
    public class Macro : MacroProvider
    {
        private const string ObjectFieldSeparator = " - ";

        #region Guids

        private static class Guids
        {
            public static class Parameters
            {
                public static readonly Guid Объект = new Guid("27f5916b-1911-44c8-8511-a4f6dee948d9");
            }

            public static class Classes
            {
                public static readonly Guid Изменение = new Guid("f40ea698-bfaa-4143-9534-6276ddec0955");
                public static readonly Guid ТехнологическоеИзменение = new Guid("a8c61089-dc59-43c8-9296-f2629fec2c4e");
            }

            public static class Links
            {
                public static readonly Guid ИзмененияАктуальныйВариант = new Guid("d962545d-bccb-40b5-9986-257b57032f6e");
                public static readonly Guid ИзмененияАктуальныйВариантТП = new Guid("254c4753-4b42-454e-84cc-f5abc82b2448");

                public static readonly Guid ИзмененияЦелеваяРевизия = new Guid("7898c148-6434-494a-bb27-f19f31a5baa2");
            }
        }

        #endregion

        public Macro(MacroContext context)
            : base(context)
        {
        }

        public void СформироватьОбозначение()
        {
            ПользовательскийДиалог диалог = ПолучитьПользовательскийДиалог("Формирование обозначения по классификатору");
            диалог.Изменить();
            if (string.IsNullOrEmpty(диалог["Наименование"]))
                диалог["Наименование"] = "Формирование обозначения по классификатору";

            string кодОрганизации = ГлобальныйПараметр["Код предприятия"];
            if (!String.IsNullOrEmpty(кодОрганизации))
                диалог["Код организации"] = кодОрганизации;

            ПоказатьПорядковыйНомер(диалог);

            диалог.Сохранить();
            if (!диалог.ПоказатьДиалог())
                return;

            string обозначение = СформироватьОбозначение(диалог["Код организации"], диалог["Код классификатора"]);
            // получить связанный документ номенклатурного объекта
            Объект номенклатурныйДокумент = ТекущийОбъект.СвязанныйОбъект["Связанный объект"];
            номенклатурныйДокумент["Обозначение"] = обозначение;
            if (String.IsNullOrEmpty(номенклатурныйДокумент["Код ФНН"]))
                номенклатурныйДокумент["Код ФНН"] = обозначение;
        }

        public void ЗавершениеИзмененияПараметраДиалога()
        {
            if (ИзмененныйПараметр == "Код организации" || ИзмененныйПараметр == "Код классификатора")
            {
                ПоказатьПорядковыйНомер(ТекущийОбъект);
            }
        }

        public void Заполнить_Обозначение_Из_Кода_ФНН()
        {
            if (ТекущийОбъект.Тип.ПорожденОт("Стандартное изделие") || ТекущийОбъект.Тип.ПорожденОт("Прочее изделие") || ТекущийОбъект.Тип.ПорожденОт("Материал"))
                return;

            string кодФНН = ТекущийОбъект["Код ФНН"];
            string обозначение = ТекущийОбъект["Обозначение"];

            if (String.IsNullOrEmpty(обозначение))
            {
                if (String.IsNullOrEmpty(кодФНН))
                {
                    int номер = ГлобальныйПараметр["Счётчик кода ФНН"];
                    номер += 1;
                    ГлобальныйПараметр["Счётчик кода ФНН"] = номер;

                    кодФНН = String.Format("ВН.{0}", номер);
                    ТекущийОбъект["Код ФНН"] = кодФНН;
                }

                ТекущийОбъект["Обозначение"] = кодФНН;
            }
            else
            {
                if (String.IsNullOrEmpty(кодФНН))
                    ТекущийОбъект["Код ФНН"] = обозначение;
            }
        }

        public string КолонкаНоменклатура()
        {
            string обозначение = ТекущийОбъект["Обозначение"];
            string наименование = ТекущийОбъект["Наименование"];
            string вариант = ТекущийОбъект["Название варианта"];
            return String.Concat(обозначение, !String.IsNullOrEmpty(обозначение) ? " - " : String.Empty, наименование, !String.IsNullOrEmpty(вариант) ? String.Format(" # [{0}]", вариант) : String.Empty);
        }

        public void ЗаполнитьПараметрыОбъекта()
        {
            var pdmBaseObject = Context.ReferenceObject as NomenclatureReferenceObject;
            if (pdmBaseObject is null)
                return;

            if (pdmBaseObject.Class.IsMaterialObject)
                Заполнить_Обозначение_Из_Кода_ФНН();

            var objectParameter = pdmBaseObject.ParameterValues.GetParameter(Guids.Parameters.Объект);
            if (objectParameter is null)
                return;

            string denotation = pdmBaseObject is NomenclatureObject pdmObject ? pdmObject.Denotation.Value : String.Empty;
            string name = pdmBaseObject.Name.Value;

            string objectValue = String.Concat(
                denotation,
                !String.IsNullOrEmpty(denotation) ? ObjectFieldSeparator : String.Empty,
                name);

            if (pdmBaseObject.Reference.ParameterGroup.SupportsRevisions && !pdmBaseObject.Class.IsFolder)
            {
                string revisionName = pdmBaseObject.SystemFields.RevisionName;
                if (!String.IsNullOrEmpty(revisionName))
                {
                    objectValue = String.Concat(
                        objectValue,
                        ObjectFieldSeparator,
                        revisionName);
                }
            }

            objectParameter.Value = objectValue;
        }

        public string КолонкаНоменклатураВСтруктуре()
        {
            string обозначение = ТекущийОбъект["[Изделие для структуры]->[Обозначение]"];
            string наименование = ТекущийОбъект["[Изделие для структуры]->[Наименование]"];
            string вариант = ТекущийОбъект["[Изделие для структуры]->[Название варианта]"];
            return String.Concat(обозначение, !String.IsNullOrEmpty(обозначение) ? " - " : String.Empty, наименование, !String.IsNullOrEmpty(вариант) ? String.Format(" [{0}]", вариант) : String.Empty);
        }

        public string КолонкаСпецификация()
        {
            if (ТекущееПодключение == null)
                return "Отсутствует применяемость";

            bool входитВСпецификацию = ТекущееПодключение["Входит в спецификацию"];
            if (!входитВСпецификацию)
                return "Не входит в спецификацию";

            string bomName = ТекущееПодключение["Раздел спецификации"];
            if (String.IsNullOrEmpty(bomName))
                return "Раздел спецификации не задан";

            string[] sections = bomName.Split('\\', '/');
            if (sections.Length > 0)
            {
                var referenceInfo = Context.Connection.ReferenceCatalog.Find(SpecialReference.BomSections);
                var bomSectionsReference = referenceInfo.CreateReference() as BomSectionsReference;
                if (bomSectionsReference == null)
                    return bomName;

                var filter = new Filter(bomSectionsReference.ParameterGroup);
                BomSectionsReferenceObject bomSection = null;
                foreach (string раздел in sections)
                {
                    filter.Terms.Clear();
                    filter.Terms.AddTerm(bomSectionsReference.ParameterGroup[BomSectionsReferenceObject.FieldKeys.Name], ComparisonOperator.Equal, раздел);
                    if (bomSection == null)
                        bomSection = bomSectionsReference.Find(filter).FirstOrDefault() as BomSectionsReferenceObject;
                    else
                        bomSection = bomSectionsReference.Find(filter, 1, bomSection).FirstOrDefault() as BomSectionsReferenceObject;

                    if (bomSection == null)
                        break;
                }

                if (bomSection != null)
                    return String.Format("{0} {1}", bomSection.Code.ToString(), bomSection.Name);
            }

            return bomName;
        }

        public string КолонкаДокументВИзменении()
        {
            string обозначение = String.Empty;
            string наименование = String.Empty;

            if (ТекущийОбъект.Тип.ПорожденОт(Guids.Classes.Изменение))
            {
                var referenceObject = (ReferenceObject)ТекущийОбъект;
                referenceObject.TryGetObject(Guids.Links.ИзмененияЦелеваяРевизия, out var revision);
                if (revision != null)
                    return revision.ToString();

                if (referenceObject.TryGetObject(Guids.Links.ИзмененияАктуальныйВариант, out var document))
                {
                    if (document is NomenclatureObject nomenclatureObject)
                    {
                        обозначение = nomenclatureObject.Denotation;
                        наименование = nomenclatureObject.Name;
                    }
                }
            }
            else if (ТекущийОбъект.Тип.ПорожденОт(Guids.Classes.ТехнологическоеИзменение))
            {
                var документ = ТекущийОбъект.СвязанныйОбъект[Guids.Links.ИзмененияАктуальныйВариантТП.ToString()];
                if (документ != null)
                {
                    обозначение = документ["Обозначение ТП"];
                    наименование = документ["Наименование"];
                }
            }

            return String.Concat(обозначение, !String.IsNullOrEmpty(обозначение) ? " - " : String.Empty, наименование);
        }

        public ICollection<int> ПолучитьИдентификаторыИспользованияВСтруктурах()
        {
            var pdmBaseObject = Context.ReferenceObject as NomenclatureReferenceObject;
            if (pdmBaseObject is null)
                return Array.Empty<int>();

            var pdmReference = pdmBaseObject.Reference as NomenclatureReference;
            if (pdmReference is null)
                return Array.Empty<int>();

            return pdmReference.ConfigurationSettings.EditableStructures.Select(x => x.Id).ToArray();

        }

        private void ПоказатьПорядковыйНомер(Объект диалог)
        {
            string код = ПолучитьКод(диалог["Код организации"], диалог["Код классификатора"]);
            int порядковыйНомер = ПолучитьПорядковыйНомер(код, false);
            диалог["Порядковый номер"] = порядковыйНомер.ToString("D3");
        }

        private string СформироватьОбозначение(string кодПредприятия, string кодКлассификатора)
        {
            string код = ПолучитьКод(кодПредприятия, кодКлассификатора);
            int порядковыйНомер = ПолучитьПорядковыйНомер(код, true);

            // формирование строки обозначения
            return String.Format("{0}.{1}", код, порядковыйНомер.ToString("D3"));
        }

        private int ПолучитьПорядковыйНомер(string код, bool изменитьСчётчик)
        {
            int порядковыйНомер = 1;

            Объект счётчик = НайтиСчётчик(код);
            if (счётчик != null)
                порядковыйНомер = счётчик["Номер"] + 1;

            if (изменитьСчётчик)
            {
                if (счётчик == null)
                {
                    счётчик = СоздатьОбъект("3595d4b3-8272-4229-83b0-1b0281685800", "7872dc95-a1c9-4dff-baf8-f7d84a2d13a5");
                    счётчик["Код"] = код;
                }
                else
                {
                    счётчик.Изменить();
                }

                счётчик["Номер"] = порядковыйНомер;
                счётчик.Сохранить();
            }

            return порядковыйНомер;
        }

        private Объект НайтиСчётчик(string код)
        {
            string фильтр = String.Format("[Код] = '{0}'", код);
            // справочник "Реестр обозначений ЕСКД"
            return НайтиОбъект("3595d4b3-8272-4229-83b0-1b0281685800", фильтр);
        }

        private string ПолучитьКод(string кодПредприятия, string кодКлассификатора)
        {
            return String.Format("{0}.{1}", кодПредприятия, кодКлассификатора);
        }

        public List<int> ПолучитьСписокИзменений()
        {
            var idList = new List<int>();

            GetRevisionChangelist(ТекущийОбъект, ref idList);

            return idList;
        }

        private void GetRevisionChangelist(Объект currentObject, ref List<int> idList)
        {
            if (currentObject is null)
                return;

            string фильтр = $"[7898c148-6434-494a-bb27-f19f31a5baa2].[ID] = {currentObject.Id}";
            var вводящиеИзменения = НайтиОбъекты("Изменения", фильтр);

            foreach (var вводящееИзменение in вводящиеИзменения)
            {
                if (idList.Contains(вводящееИзменение.Id))
                    continue;

                idList.Add(вводящееИзменение.Id);

                var исходныеРевизии = вводящееИзменение.СвязанныеОбъекты["Исходные ревизии"];

                foreach (var исходнаяРевизия in исходныеРевизии)
                    GetRevisionChangelist(исходнаяРевизия.СвязанныйОбъект["Ревизия"], ref idList);
            }
        }
    }
}
