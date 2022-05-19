using System;
using System.Linq;
using TFlex.DOCs.Model;
using TFlex.DOCs.Model.Macros;
using TFlex.DOCs.Model.Macros.ObjectModel;
using TFlex.DOCs.Model.References;
using TFlex.DOCs.Model.References.Catalogs;
using TFlex.DOCs.Model.Search;
using TFlex.DOCs.Model.Search.Path;
using TFlex.DOCs.Model.Structure;

namespace CatalogCreator
{
    public class Macro : MacroProvider
    {
        private string _catalogReferenceParameterName = string.Empty;
        private ParameterInfo _currentReferenceParameter;

        private bool _showEmbeddedObjects = false;
        private string _skippedSymbols = string.Empty;

        public Macro(MacroContext context)
            : base(context)
        {
        }

        public void ПоказатьПользовательскийДиалогСозданияКаталога()
        {
            // ТекущийОбъект - объект справочника

            ПользовательскийДиалог диалог = ПолучитьПользовательскийДиалог("Создание каталога");
            диалог.Изменить();
            // Гуид текущего справочника
            диалог["Текущий справочник"] = ТекущийОбъект.Справочник.УникальныйИдентификатор;
            string наименованиеКаталога = ТекущийОбъект["Наименование"];
            диалог["Наименование"] = наименованиеКаталога;
            диалог["Идентификатор текущего объекта"] = ТекущийОбъект["ID"];
            диалог.Заголовок = string.Format("Создание каталога '{0}'", наименованиеКаталога);
            диалог.Сохранить();

            диалог.ПоказатьДиалог();
        }

        public void СоздатьКаталог()
        {
            var connection = Context.Connection;
            if (connection == null)
                Ошибка("connection");

            // ТекущийОбъект - пользовательский диалог "Создание каталога"
            Объект диалог = ТекущийОбъект;

            _catalogReferenceParameterName = диалог["Параметр каталога"].ToString().Split('.').Last();

            string pathString = диалог["Параметр текущего справочника"].ToString();

            ReferencePath referencePath = null;
            if (ReferencePath.TryParse(pathString, Context.Connection, out referencePath))
            {
                ParameterPathItem lastPathItem = referencePath.Last() as ParameterPathItem;
                _currentReferenceParameter = lastPathItem.Parameter;
            }

            _showEmbeddedObjects = диалог["Показывать вложенные"];

            int пропуститьСимволов = диалог["Пропустить"];
            for (int i = 0; i < пропуститьСимволов; i++)
            {
                _skippedSymbols = string.Format("{0}_", _skippedSymbols);
            }

            Guid исходныйСправочник = диалог["Текущий справочник"];
            Guid справочникДляКаталога = диалог["Справочник для каталога"];
            string наименованиеКаталога = диалог["Наименование"];
            int идентификаторВыбранногоОбъекта = диалог["Идентификатор текущего объекта"];

            if (СоздатьКаталог(исходныйСправочник, справочникДляКаталога, наименованиеКаталога, идентификаторВыбранногоОбъекта))
                Сообщение("Создание каталога", string.Format("Создание каталога '{0}' с папками поиска успешно завершено.", наименованиеКаталога));
        }

        private bool СоздатьКаталог(Guid исходныйСправочник, Guid справочникДляКаталога, string наименованиеКаталога, int идентификаторВыбранногоОбъекта)
        {
            Объект объектДляКаталога = НайтиОбъект(исходныйСправочник.ToString(), "ID", идентификаторВыбранногоОбъекта);
            if (объектДляКаталога == null)
            {
                Сообщение("Создание каталога", "Не найден исходный объект для создания каталога.");
                return false;
            }

            ReferenceInfo catalogReferenceInfo = Context.Connection.ReferenceCatalog.Find(справочникДляКаталога);
            if (catalogReferenceInfo == null)
            {
                Сообщение("Создание каталога", string.Format("Не найден справочник '{0}' для создания каталога.", справочникДляКаталога));
                return false;
            }

            // Менеджер каталогов справочника
            CatalogManager catalogManager = new CatalogManager(catalogReferenceInfo);
            if (catalogManager.Any(c => c.Name == наименованиеКаталога))
            {
                Сообщение("Создание каталога", string.Format("Каталог '{0}' в справочнике '{1}' уже существует.", наименованиеКаталога, catalogReferenceInfo.Name));
                return false;
            }

            CreateCatalog(catalogManager, (ReferenceObject)объектДляКаталога, catalogReferenceInfo, наименованиеКаталога);

            return true;
        }

        private void CreateCatalog(CatalogManager catalogManager, ReferenceObject referenceObject, ReferenceInfo catalogReferenceInfo, string catalogName)
        {
            // Создаём новый каталог
            var catalog = new Catalog(catalogManager);
            catalog.Name = catalogName;
            catalog.Save();

            SearchFolder searchFolder = new SearchFolder(catalog);
            FillSearchFolderParameters(searchFolder, referenceObject, catalogReferenceInfo.Description);
            searchFolder.Save();

            RecursiveSearchFolderCreation(searchFolder, referenceObject, catalogReferenceInfo.Description);
        }

        private void RecursiveSearchFolderCreation(SearchFolder parentFolder, ReferenceObject referenceObject, ParameterGroup parameterGroup)
        {
            foreach (ReferenceObject child in referenceObject.Children)
            {
                SearchFolder childSearchFolder = new SearchFolder(parentFolder);
                FillSearchFolderParameters(childSearchFolder, child, parameterGroup);
                childSearchFolder.Save();

                RecursiveSearchFolderCreation(childSearchFolder, child, parameterGroup);
            }
        }

        /// <summary>
        /// Установка параметров папки поиска
        /// </summary>
        private void FillSearchFolderParameters(SearchFolder searchFolder, ReferenceObject referenceObject, ParameterGroup parameterGroup)
        {
            // наименование папки поиска
            searchFolder.Name = referenceObject.ToString();

            // Отображать вложенные объекты, если не будет дочерних папок поиска
            if (_showEmbeddedObjects)
                searchFolder.ShowObjects = !referenceObject.Children.Any();

            // Значение указанного параметра
            string parameterValue = referenceObject[_currentReferenceParameter].ToString();
            int prefixIndex = parameterValue.IndexOf("_");
            // формирование фильтра для папки поиска
            string filterString = string.Format("{0} соответствует маске '{1}{2}{3}'", _catalogReferenceParameterName, _skippedSymbols, parameterValue, prefixIndex < 0 ? "%" : string.Empty);

            Filter filter = null;
            if (Filter.TryParse(filterString, parameterGroup, out filter))
                searchFolder.SetFilter(filter);
        }
    }
}

