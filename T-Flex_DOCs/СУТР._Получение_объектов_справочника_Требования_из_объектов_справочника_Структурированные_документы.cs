using System;
using System.Collections.Generic;
using TFlex.DOCs.Model.Classes;
using TFlex.DOCs.Model.Macros;
using TFlex.DOCs.Model.References;

namespace Macroses
{
    /// <summary>
    /// Макрос служит для создания элементов справочника Требования на основе уже созданных объектов справочника Структурированные документы
    /// </summary>
    public class RequirementsFromStructuredDocumentsMacros : MacroProvider
    {
        // Guid справочника "Требования"
        private static readonly Guid RequirementsReferenceGuid = new Guid("48c51985-0f22-4315-a965-7b49888f4098");

        // Guid справочника "Структурированные документы"
        private static readonly Guid StructuredDocumentsReferenceGuid =
            new Guid("2b610882-d4a3-41b3-859d-8545191f8671");

        // Guid типа "Спецификация требований" справочника "Требования". Это корневой объект
        private static readonly Guid RequirementsSpecificationClass = new Guid("3707fb29-42c9-4a44-b10c-51ff809ffc64");

        // Guid типа "Требование" справочника "Требования"
        private static readonly Guid RequirementClass = new Guid("979d4e8f-5253-441d-ba45-745385e9a8a7");

        // Guid типа "Заголовок" справочника "Требования"
        private static readonly Guid RequirementsHeaderClass = new Guid("eb51c585-ce8c-4def-a359-489b653da7ba");

        // Guid типа "Заголовок" справочника "Cтруктурированные документы"
        private static readonly Guid StructuredDocumentsHeaderClass = new Guid("cb3602e4-5b36-4494-a49e-a46780cf5351");

        // Guid типа "Фрагмент" справочника "Cтруктурированные документы"
        private static readonly Guid StructuredDocumentsFragmentClass =
            new Guid("182b2876-98c4-488c-945e-10897b179c92");

        // Guid типа "Вспомогательный текст" справочника "Требования"
        private static readonly Guid RequirementsHelperTextClass = new Guid("2adbbb8d-c617-44c2-8027-62ec6a341f88");

        // Guid типа "Вспомогательный текст" справочника "Cтруктурированные документы"
        private static readonly Guid StructuredDocumentsHelperTextClass = new Guid("c50cbf8b-c32d-4c05-ac24-393fd0ca3447");

        // Guid параметра "Наименование, текст" справочника "Требования"
        private static readonly Guid NameTextRequirementParameterGuid =
            new Guid("49d6731b-9a72-4817-a23a-15bc919752d5");

        // Guid параметра "Наименование, текст" справочника "Структурированные документы"
        private static readonly Guid NameTextStructuredDocumentsParameterGuid =
            new Guid("ae937205-1945-4832-b841-00f75a306b8c");

        // Guid связи "Структурированные документы" справочника "Требования"
        private static readonly Guid StructuredDocumentsLinkGuid = new Guid("63b766a5-7daa-4b54-a3ee-124e7a25c4f7");

        // Guid параметра "Форматированный текст" справочника "Требования"
        private static readonly Guid RequirementsRichTextParameterGuid =
            new Guid("eb330278-a3ce-4982-89b1-1362527d30f0");

        // Guid параметра "Форматированный текст" справочника "Структурированные документы"
        private static readonly Guid StructuredDocumentsRichTextParameterGuid =
            new Guid("533cbfa4-a165-4c43-9ff2-fe8d648d1e8c");

        private IDictionary<Guid, Guid> _typesMappingDictionary;

        public RequirementsFromStructuredDocumentsMacros(MacroContext context) : base(context)
        {
            _typesMappingDictionary = new Dictionary<Guid, Guid>()
            {
                {StructuredDocumentsHeaderClass, RequirementsHeaderClass}, //соответствие заголовок - заголовок
                {StructuredDocumentsFragmentClass, RequirementClass}, // соответствие фрагмент - требование
                {StructuredDocumentsHelperTextClass, RequirementsHelperTextClass}, // соответствие вспомогательный текст - вспомогательный текст
            };
        }

        public override void Run()
        {
            Start();
        }

        private void Start()
        {
            if (Context.ReferenceObject == null)
                return;

            if (Context.Reference.ParameterGroup.Guid != RequirementsReferenceGuid)
            {
                Сообщение("Внимание!", "Макрос запущен не на справочнике \"Требования\"!");
                return;
            }

            var contextReferenceObject = Context.ReferenceObject;

            // если ни один тип из справочника Требования, выходим
            if (contextReferenceObject.Class.Guid != RequirementsSpecificationClass
                && contextReferenceObject.Class.Guid != RequirementClass
                && contextReferenceObject.Class.Guid != RequirementsHeaderClass)
            {
                Сообщение("Внимание!",
                    "Указанный тип объекта не соответствует типу Спецификации требований, требования или заголовок");
                return;
            }

            // выбрать объекты справочника Структурированные документы
            var selectedStructuredDocuments = SelectDocumentObjects();
            if (selectedStructuredDocuments.Count == 0)
                return;

            ДиалогОжидания.Показать("Пожалуйста подождите - идет операция получения исходных данных", false);

            FillReferenceRequirements(contextReferenceObject, selectedStructuredDocuments);

            ДиалогОжидания.Скрыть();
            ОбновитьОкноСправочника();
        }

        private void FillReferenceRequirements(ReferenceObject requirementObject,
            List<ReferenceObject> selectedDocuments)
        {
            if (requirementObject is null)
                return;

            foreach (var documentObject in selectedDocuments)
            {
                StartFill(documentObject, requirementObject);
            }
        }

        private void StartFill(ReferenceObject documentObject, ReferenceObject requirementObject)
        {
            Guid classGuidFromRequirements = GetClassGuidFromRequirements(documentObject.Class.Guid);
            var childRequirementClass = Context.Reference.Classes.Find(classGuidFromRequirements);
            var newObj = NewRequirementObject(requirementObject, documentObject, childRequirementClass);
            FillReqRecursive(documentObject, newObj);
        }

        private void FillReqRecursive(ReferenceObject documentObject, ReferenceObject parentReqReferenceObject)
        {
            if (parentReqReferenceObject.Changing)
                parentReqReferenceObject.EndChanges();

            foreach (var docObject in documentObject.Children)
            {
                StartFill(docObject, parentReqReferenceObject);
            }
        }

        private ReferenceObject NewRequirementObject(ReferenceObject parentObject,
            ReferenceObject structuredDocumentObject,
            ClassObject requirementClass)
        {
            var newObj = Context.Reference.CreateReferenceObject(parentObject, requirementClass);

            newObj.ParameterValues[NameTextRequirementParameterGuid].Value =
            	structuredDocumentObject.ParameterValues[NameTextStructuredDocumentsParameterGuid].GetString();

            newObj.ParameterValues[RequirementsRichTextParameterGuid].Value = 
            	structuredDocumentObject.ParameterValues[StructuredDocumentsRichTextParameterGuid].GetString();

            newObj.AddLinkedObject(StructuredDocumentsLinkGuid, structuredDocumentObject);
            return newObj;
        }

        private List<ReferenceObject> SelectDocumentObjects()
        {
            var list = new List<ReferenceObject>();

            var dialog = СоздатьДиалогВыбораОбъектов(StructuredDocumentsReferenceGuid.ToString());
            dialog.Заголовок = "Выберите объект из справочника Структурированные документы для наполнения требований";
            dialog.МножественныйВыбор = true;

            if (!dialog.Показать())
                return list;

            if (dialog.ВыбранныеОбъекты == null || dialog.ВыбранныеОбъекты.Count == 0)
                return list;

            var selectedObjects = dialog.ВыбранныеОбъекты;
            ReferenceObject previousSelectedDocumentParentReferenceObject = null;
            foreach (var selectedObject in selectedObjects)
            {
                if (previousSelectedDocumentParentReferenceObject == null)
                    previousSelectedDocumentParentReferenceObject = (ReferenceObject)selectedObject.ParentObject;
                var parentReferenceObject = (ReferenceObject)selectedObject.ParentObject;

                if (parentReferenceObject != previousSelectedDocumentParentReferenceObject)
                    throw new InvalidOperationException("Выбранные узлы должны принадлежать одному родителю!");

                list.Add((ReferenceObject)selectedObject);
            }

            return list;
        }

        private Guid GetClassGuidFromRequirements(Guid classGuidFromStructuredDocuments)
        {
            if (!_typesMappingDictionary.TryGetValue(classGuidFromStructuredDocuments, out var result))
                throw new NotSupportedException(
                    "Выбранный тип в справочнике Структурированные документы не имеет соответствий ни с каким типом из справочника Требования");

            return result;
        }
    }
}
