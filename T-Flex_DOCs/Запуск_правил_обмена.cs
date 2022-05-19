using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using TFlex.DOCs.Model.Macros;
using TFlex.DOCs.Model.Macros.ObjectModel;
using TFlex.DOCs.Model.References;
using TFlex.DOCs.Model.References.Nomenclature;
using TFlex.DOCs.Model.Search;

public class MacroExchangeRules : MacroProvider
{
    public MacroExchangeRules(MacroContext context)
        : base(context)
    {
        if (Context.Connection.ClientView.HostName == "MOSINS")
            if (Вопрос("Хотите запустить в режиме отладки?"))
            {
                System.Diagnostics.Debugger.Launch();
                System.Diagnostics.Debugger.Break();
            }
    }

    public override void Run()
    {
        var currentObejct = Context.ReferenceObject;
        if (currentObejct == null)
            return;

        var childrenObjectList = GetAllChindrenNomenclature(currentObejct);
        childrenObjectList.Add(currentObejct);

        var tpObjectsList = GetAllTPObjectList(childrenObjectList);
        var materialObjects = GetAllTPMaterialObjectList(tpObjectsList);
        var perehodsObjects = GetAllPerehodList(tpObjectsList);
        var riggingList = GetAllRiggingList(perehodsObjects);
        var allTechDocuments = GetAllTechDocumentList(childrenObjectList);

        StringBuilder result = new StringBuilder();

        ЗапуститьПравилоОбменаВедомостиМатериалов();

        if (materialObjects.Count > 0)
            result.AppendLine(RunDataExchandeExport("Выгрузка материалов в базу данных (без аналогов)", materialObjects));

        if (perehodsObjects.Count > 0)
            result.AppendLine(RunDataExchandeExport("Выгрузка маршрутов в базу данных (без аналогов)", perehodsObjects));

        if (riggingList.Count > 0)
            result.AppendLine(RunDataExchandeExport("Выгрузка в базу данных оснастки (без аналогов)", riggingList));

        if (allTechDocuments.Count > 0)
            result.AppendLine(RunDataExchandeExport("Выгрузка ГПП в базу данных (без аналогов)", allTechDocuments));



        if (String.IsNullOrWhiteSpace(result.ToString()))
            Сообщение("Выполнено", "Ошибок не обнаружено");
        else
            Ошибка(result.ToString());
    }

    public void ЗапуститьПравилоОбменаВедомостиМатериалов()
    {
        var result = ОбменДанными.Экспортировать("Выгрузка ведомости материалов в FoxPro 2", показыватьДиалог: false);
        if (result.HasErrors)
            return;

        var normOutObjects = FindNormOutObjects();
        if (normOutObjects.Count == 0)
            return;

        RunDataExchandeExport("Выгрузка ведомости материалов в базу данных", normOutObjects);

        var saveSet = new List<ReferenceObject>();
        foreach (var referenceObject in normOutObjects)
        {
            try
            {
                referenceObject.BeginChanges(false);
                referenceObject[Guids.IsImportParameter].Value = false;
                saveSet.Add(referenceObject);
            }
            catch (Exception) { }
        }

        if (saveSet.Count > 0)
            Reference.EndChanges(saveSet);
    }

    public void ОчиститьВсеПризнакиОбъектов()
    {
        var reference = Context.Connection.ReferenceCatalog.Find(Guids.NormOutReference)?.CreateReference();
        if (reference == null)
            return;

        reference.LoadSettings.Add(Guids.IsImportParameter);
        var objects = reference.Objects;
        var saveSet = new List<ReferenceObject>();
        foreach (var referenceObject in objects)
        {
            try
            {
                referenceObject.BeginChanges(false);
                referenceObject[Guids.IsImportParameter].Value = false;
            }
            catch (Exception) { }
        }

        if (saveSet.Count > 0)
            Reference.EndChanges(saveSet);
    }

    private List<ReferenceObject> FindNormOutObjects()
    {
        var result = new List<ReferenceObject>();
        var reference = Context.Connection.ReferenceCatalog.Find(Guids.NormOutReference)?.CreateReference();
        if (reference == null)
            return result;

        var parameterInfo = reference.ParameterGroup.OneToOneParameters.Find(Guids.IsImportParameter);
        if (parameterInfo == null)
            return result;

        reference.LoadSettings.Add(parameterInfo);
        var filter = new Filter(reference.ParameterGroup);
        filter.Terms.AddTerm(parameterInfo, ComparisonOperator.Equal, true);
        return reference.Find(filter);
    }

    private List<ReferenceObject> GetAllTPMaterialObjectList(List<ReferenceObject> allTpObjects)
    {
        var result = new List<ReferenceObject>();
        foreach (var tpObject in allTpObjects)
        {
            var materialListObjectList = tpObject.GetObjects(Guids.MaterialBlankListObjects).FindAll(m => m.Class.Guid == Guids.MaterialClass);
            foreach (var materialListObject in materialListObjectList)
            {
                var linkedMaterial = materialListObject.GetObject(Guids.MaterianBlankLinkMaterial);
                if (linkedMaterial != null)
                    result.Add(linkedMaterial);
            }
        }

        return result.Distinct().ToList();
    }

    private List<ReferenceObject> GetAllTechDocumentList(List<ReferenceObject> referenceObjectList)
    {
        var result = new List<ReferenceObject>();

        foreach (var referenceObjcet in referenceObjectList)
        {
            if (!(referenceObjcet is NomenclatureObject nomObject))
                continue;

            var linkedObjcet = nomObject.LinkedObject;
            if (linkedObjcet.Reference.Name != "Документы")
                continue;

            result.AddRange(GetLinkedTPDocuments(linkedObjcet));
        }

        return result.Distinct().ToList();
    }

    /// <summary>
    /// Получает все связанные технологические документы 
    /// </summary>
    /// <param name="documentObject"></param>
    /// <returns></returns>
    private List<ReferenceObject> GetLinkedTPDocuments(ReferenceObject documentObject)
    {
        var result = new List<ReferenceObject>();
        var linkedDocuments = documentObject.GetObjects(Guids.LinkedDocumentsLink).FindAll(d => d.Class.Guid == Guids.TechDocumentClass);
        foreach (var linkedDocument in linkedDocuments)
        {
            var viewDocument = linkedDocument[Guids.ViewDocumentParameter].GetInt32();
            if (viewDocument == 1 || viewDocument == 2)
                result.Add(linkedDocument);
        }

        return result;
    }
    //=> documentObjcet.GetObjects(Guids.LinkedDocumentsLink).FindAll(d => 
    //(d.Class.Guid == Guids.TechDocumentClass) && (d[Guids.ViewDocumentParameter].GetInt32()))

    private string RunDataExchandeExport(string name, List<ReferenceObject> objects)
        => ОбменДанными.ЭкспортироватьОбъекты(name, Объекты.CreateInstance(objects, Context), false).ErrorMessage;

    private string RunDataExchandeExport(string name) => ОбменДанными.Экспортировать(name, показыватьДиалог: false).ErrorMessage;

    /// <summary>
    /// Получает всю оснастку у объектов переходов
    /// </summary>
    /// <param name="perehods"></param>
    /// <returns></returns>
    private List<ReferenceObject> GetAllRiggingList(List<ReferenceObject> perehods)
    {
        var result = new List<ReferenceObject>();
        perehods.ForEach(ro => result.AddRange(GetRiggingList(ro)));
        return result.Distinct().ToList();
    }

    /// <summary>
    /// Получает все техпроцессы у номенклатурных единиц
    /// </summary>
    /// <param name="nomObjectList"></param>
    /// <returns></returns>
    private List<ReferenceObject> GetAllTPObjectList(List<ReferenceObject> nomObjectList)
    {
        var result = new List<ReferenceObject>();
        try
        {
            nomObjectList.ForEach(ro => result.AddRange(ro.GetObjects(Guids.TPLink)));
        }
        catch (Exception) { }

        return result;
    }

    /// <summary>
    /// Получает переходу у всех техпроцессов
    /// </summary>
    /// <param name="tpObjectList"></param>
    /// <returns></returns>
    private List<ReferenceObject> GetAllPerehodList(List<ReferenceObject> tpObjectList)
    {
        var result = new List<ReferenceObject>();
        tpObjectList.ForEach(ro => result.AddRange(GetPerehodInTP(ro)));
        return result;
    }

    /// <summary>
    /// Получает все переходы у номенклатурного объекта
    /// </summary>
    /// <param name="parentNomObject"></param>
    /// <returns></returns>
    private List<ReferenceObject> GetAllPerehodList(ReferenceObject parentNomObject)
    {
        var allChildrenObjectList = GetAllChindrenNomenclature(parentNomObject);
        allChildrenObjectList.Add(parentNomObject);
        var allTpObjects = GetAllTPObjectList(allChildrenObjectList);
        if (allTpObjects.Count == 0)
            Ошибка("Не найдены техпроцессы");
        var perehods = GetAllPerehodList(allTpObjects);
        if (perehods.Count == 0)
            Ошибка("Не найдены переходы");

        return perehods;
    }

    /// <summary>
    /// Получает все дочерние объекты номенклатуры
    /// </summary>
    /// <param name="nomObject"></param>
    /// <returns></returns>
    private List<ReferenceObject> GetAllChindrenNomenclature(ReferenceObject nomObject) => nomObject.Children.RecursiveLoad();

    /// <summary>
    /// Получает все переходы у Техпроцесса
    /// </summary>
    /// <param name="tpObject"></param>
    /// <returns></returns>
    private List<ReferenceObject> GetPerehodInTP(ReferenceObject tpObject) => tpObject.Children.Where(ob => ob.Class.Guid == Guids.PerehodClass).ToList();

    /// <summary>
    /// Получает связанные объекты оснащения
    /// </summary>
    /// <param name="perehodObject"></param>
    /// <returns></returns>
    private List<ReferenceObject> GetRiggingList(ReferenceObject perehodObject) => perehodObject.GetObjects(Guids.RiggingLink);

    private static class Guids
    {
        public static Guid[] ChandeClassGuids = new Guid[]
        {
            new Guid("08309a17-4bee-47a5-b3c7-57a1850d55ea"), new Guid("1cee5551-3a68-45de-9f33-2b4afdbf4a5c")
        };

        public static Guid NormOutReference = new Guid("991875e9-7376-4eb0-884f-63dee51a6d66");

        public static Guid IsImportParameter = new Guid("97d8d865-be6d-458a-a5ea-c10dd7685fa6");

        /// <summary>
        /// Документы тип Технологический документ параметр Вид документа 
        /// </summary>
        public static Guid ViewDocumentParameter = new Guid("0a50c800-161a-4f6a-8738-e712693a02b3");

        /// <summary>
        /// Документы связь связанные документы
        /// </summary>
        public static Guid LinkedDocumentsLink = new Guid("b840c7cc-bc01-48db-84a0-3706b7aba745");

        /// <summary>
        /// ТП - Изготовление список объектов Материал\заготовка
        /// </summary>
        public static Guid MaterialBlankListObjects = new Guid("8f505469-c1b4-4caf-a7c1-3bc7ea8c2bbe");

        /// <summary>
        /// ТП - Изготовление список объектов Материал\заготовка связь материал
        /// </summary>
        public static Guid MaterianBlankLinkMaterial = new Guid("0ea298d6-094e-4286-a08a-ed86ca042424");

        /// <summary>
        /// ТП - Изготовление список объектов Материал\заготовка тип Материал
        /// </summary>
        public static Guid MaterialClass = new Guid("35ee20c9-d771-4f90-8573-0505c2a7e398");

        /// <summary>
        /// Цехопереход связь Каталог оснащения
        /// </summary>
        public static Guid RiggingLink = new Guid("c7af468f-95dd-4835-a562-c0f96e170e4e");

        /// <summary>
        /// ТП тип цехопереход
        /// </summary>
        public static Guid PerehodClass = new Guid("459ae48b-165b-44fd-8b3e-890298f2c3d7");

        /// <summary>
        /// Связь Технология
        /// </summary>
        public static Guid TPLink = new Guid("ba824125-2d20-4b50-b14f-0e5bfe9b4db4");

        /// <summary>
        /// Документы тип Технологический документ
        /// </summary>
        public static Guid TechDocumentClass = new Guid("9745b167-0e66-43c2-91c0-899a0149b19e");

        /// <summary>
        /// Деталь
        /// </summary>
        public static Guid DetailClass = new Guid("08309a17-4bee-47a5-b3c7-57a1850d55ea");

        /// <summary>
        /// Сборочная единица
        /// </summary>
        public static Guid AssemblyUnitClass = new Guid("1cee5551-3a68-45de-9f33-2b4afdbf4a5c");
    }
}
