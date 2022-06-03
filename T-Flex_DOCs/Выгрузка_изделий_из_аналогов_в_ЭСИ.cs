using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using TFlex.DOCs.Model.Macros;
using TFlex.DOCs.Model.References;
using TFlex.DOCs.Model.Macros.ObjectModel;
using TFlex.DOCs.Model;
using TFlex.DOCs.Model.Structure;
using TFlex.DOCs.Model.Search;

public class Macro : MacroProvider
{
    private Reference СписокНоменклатурыСправочник { get; set; }
    private Reference ПодключенияСправочник { get; set; }
    private Reference ЭсиСправочник { get; set; }
    private Reference ДокументыСправочник { get; set; }
    private Reference ЭлектронныеКомпонентыСправочник { get; set; }
    private Reference МатериалыСправочник { get; set; }
    private string ДиректорияДляЛогов { get; set; }

    public Macro(MacroContext context)
        : base(context) {

        // Получаем экземпляры справочников для работы
        СписокНоменклатурыСправочник = Context.Connection.ReferenceCatalog.Find(Guids.References.СписокНоменклатурыFoxPro).CreateReference();
        ПодключенияСправочник = Context.Connection.ReferenceCatalog.Find(Guids.References.Подключения).CreateReference();
        ЭсиСправочник = Context.Connection.ReferenceCatalog.Find(Guids.References.ЭСИ).CreateReference();
        ЭлектронныеКомпонентыСправочник = Context.Connection.ReferenceCatalog.Find(Guids.References.ЭлектронныеКомпоненты).CreateReference();
        МатериалыСправочник = Context.Connection.ReferenceCatalog.Find(Guids.References.Материалы).CreateReference();
        
        // Создаем директорию для ведения логов
        ДиректорияДляЛогов = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "Логи выгрузки из аналогов в ЭСИ");

        if (!Directory.Exists(ДиректорияДляЛогов))
            Directory.CreateDirectory(ДиректорияДляЛогов);

    }

    private static class Guids {
        public static class References {
            public static Guid СписокНоменклатурыFoxPro = new Guid("c9d26b3c-b318-4160-90ae-b9d4dd7565b6");
            public static Guid Подключения = new Guid("9dd79ab1-2f40-41f3-bebc-0a8b4f6c249e");
            public static Guid ЭСИ = new Guid("853d0f07-9632-42dd-bc7a-d91eae4b8e83");
            public static Guid ЭлектронныеКомпоненты = new Guid("2ac850d9-5c70-45c2-9897-517ab571b213");
            public static Guid Материалы = new Guid("c5e7ae00-90f2-49e9-a16c-f51ed087752a");
        }

        public static class Parameters {
            // Параметры справочника "Список номерклатуры FoxPro"
            public static Guid НоменклатураОбозначение = new Guid("1bbb1d78-6e30-40b4-acde-4fc844477200");

            // Параметры справчоника "Подключения"
            public static Guid ПодключенияСборка = new Guid("4a3cb1ca-6a4c-4dce-8c25-c5c3bd13a807");
            public static Guid ПодключенияКомплектующая = new Guid("7d1ac031-8c7f-49b5-84b8-c5bafa3918c2");

            // Параметры справочника "ЭСИ"
            
            // Параметры справочника "Документы"
            
            // Параметры справочника "Электронные компоненты"
            
            // Параметры справочника "Материалы"
            
        }

        public static class Links {
            public static Guid СвязьСпискаНоменклатурыНаЭСИ = new Guid("ec9b1e06-d8d5-4480-8a5c-4e6329993eac");
        }
    }

    public override void Run() {
        // Производим загрузку всей необходимой информации
        Dictionary<string, ReferenceObject> номенклатура = GetNomenclature();
        Dictionary<string, List<ReferenceObject>> подключения = GetLinks();

        // Запрашиваем у пользователя перечень изделий, по которым нужно произвести выгрузку
        List<string> изделияДляВыгрузки = GetShifrsFromUserToImport(номенклатура);

        // Определяем позиции справочника "Список номенклатуры FoxPro", которые необходимо обрабатывать во время выгрузки
        HashSet<ReferenceObject> номенклатураДляСоздания = GetNomenclatureToProcess(номенклатура, подключения, изделияДляВыгрузки);

        // Производим поиск и (при необходимости) создание объектов в ЭСИ и смежных справочниках
        List<ReferenceObject> созданныеДСЕ = FindOrCreateNomenclatureObjects(номенклатураДляСоздания);

        // Производим соединение созданных ДСЕ в иерархию при помощи подключений
        ConnectCreatedObjects(созданныеДСЕ, подключения);

        Message("Информация", "Работа макроса завершена");
    }


    /// <summary>
    /// Получает все объекты справочника
    /// </summary>
    public ReferenceObjectCollection GeAlltRefObj(Guid guidref)
    {
        ReferenceInfo info = Context.Connection.ReferenceCatalog.Find(guidref);
        Reference reference = info.CreateReference();
        var result = reference.Objects;
        return result;
    }

    /// <summary>
    /// Функция возвращает словарь с изделиями
    /// </summary>
    private Dictionary<string, ReferenceObject> GetNomenclature()
    {
        var ListNum = GeAlltRefObj(Guids.References.СписокНоменклатурыFoxPro);
        var dictListNum = ListNum.ToDictionary(objref => (objref[Guids.Parameters.НомерклатураОбозначение].Value.ToString()));
        return dictListNum;
    }

    /// <summary>
    /// Функция возвращает словарь с подключениями
    /// </summary>
    private Dictionary<string, List<ReferenceObject>> GetLinks()
    {
           var RefConnectNum = GeAlltRefObj(Guids.References.Подключения);

        Dictionary<string,List<ReferenceObject>> dict = new Dictionary<string,List<ReferenceObject>>(300000);
        foreach (var item in RefConnectNum)
        {
            string shifr = (String)item[Guids.Parameters.ПодключенияСборка].Value;
            if (dict.ContainsKey(shifr))
            {
                    dict[shifr].Add(item);
            }
            else
            {
                dict.Add(shifr, new List<ReferenceObject>() { item });
            }
        }
        return dict;
    }

     /// <summary>
    /// Получает объекты справочника по условию если parametr содердит строку str
    /// </summary>
    public List<ReferenceObject> GeFiltertRefObj(String str, Guid guidref, Guid parametr)
    {
        ReferenceInfo info = Context.Connection.ReferenceCatalog.Find(guidref);
        Reference reference = info.CreateReference();
        ParameterInfo parameterInfo = reference.ParameterGroup[parametr];
        List<ReferenceObject> result = reference.Find(parameterInfo, ComparisonOperator.Equal, str);
        return result;
    }

    /// <summary>
    /// Функция запрашивает у пользователя, какие изделия необходимо выгрузить.
    /// Если введенное пользователем изделие отсутствует среди всех изделий, должно выдаваться об этом сообщение
    /// Так же должна быть возможность снова ввести данные.
    /// </summary>
    private List<string> GetShifrsFromUserToImport(Dictionary<string, ReferenceObject> nomenclature)
    {

         List<string> result = new List<string>();
       string shifr = "УЯИС.731353.037";
       var filter =  GeFiltertRefObj(shifr, Guids.References.СписокНоменклатурыFoxPro, Guids.Parameters.НомерклатураОбозначение);
        if (filter != null)
        {
            result = (filter.Select(objref => (objref[Guids.Parameters.НомерклатураОбозначение].Value.ToString()))).ToList();            
        }

        return result;
    }

    private HashSet<ReferenceObject> GetNomenclatureToProcess(Dictionary<string, ReferenceObject> nomenclature, Dictionary<string, List<ReferenceObject>> links, List<string> shifrs) {
        // Для каждого шифра создаем объект, реализующий интерфейс ITree, получаем входящие объекты, добавляем их в HashSet (для исключения дубликатов)
        // В конце пишем лог, в котором записываем информацию о сгенерированном дереве, количестве входящих объектов и их структуре
        HashSet<ReferenceObject> result = new HashSet<ReferenceObject>();
        foreach (string shifr in shifrs) {
            NomenclatureTree tree = new NomenclatureTree(nomenclature, links, shifr);
            result.UnionWith(tree.AllReferenceObjects);
        }
        return result;
    }

    private List<ReferenceObject> FindOrCreateNomenclatureObjects(HashSet<ReferenceObject> nomenclature) {
        // Функция принимает записи справочника "Список номенклатуры FoxPro" для создания объектов с справочнике ЭСИ и смежных справочников
        // Функция возвращает найденные или созданные записи справочника ЭСИ
        //
        // Необходимо реализовать:
        // - Поиск объектов в справочнике ЭСИ
        // - Создание объектов, если они не были найдены
        // - Подключение созданных или найденных объектов к соответствующим записям справочника "Список номенклатуры FoxPro"
        // - Возврат всех найденных/созданных объектов для последующей с ними работы
        // - Вывод лога о всех произведенных действиях
        List<ReferenceObject> result = new List<ReferenceObject>();

        // ВАЖНО: при проведении поиска нужно проверять на то, что найденный объект единственный. Если он не единственный, тогда нужно выдать ошибку для принятия решения по поводу обработки данного случая
        
        foreach (ReferenceObject nom in nomenclature) {
            // Пробуем получить объект по связи на справочник ЭСИ. Если получилось, добавляем его в result, дальнейший код не выполняем
            //
            // Пробуем найти объект в справочнике ЭСИ. Если получилось, подключаем его к объекту nom по связи, добавляем в result, дальнейший код не выполняем
            //
            // Определяем тип объекта
            TypeOfObject type = DefineTypeOfObject(nom);
            // На основе полученных данных о типе производим поиск в смежных справочниках. Если найти объект получилось, создаем номенклатурный объект на его основе, добавляем в result, подключаем в nom,
            // дальнейший код не выполняем
            //
            // Если найти объект не получилось, в завосомости от типа объекта создаем объект в соответствующем справочнике, создаем ЭСИ, добавляем в result, подключаем в nom
        }

        return result;
    }

    private void ConnectCreatedObjects(List<ReferenceObject> createdObjects, Dictionary<string, List<ReferenceObject>> links) {
        // Функция принимает созданные номенклатурный объекты, а так же объекты справочника "Подключения"
        // 
        // Необходимо реализовать:
        // - Анализ полученных объектов (у них уже могут быть связи, так как там будут и найденные позиции)
        // - Создание/Корректировка/Удаление связей при необходимости
        // - Вывод лога о всех произведенных действиях
    }

    private TypeOfObject DefineTypeOfObject(ReferenceObject nomenclature) {
        return TypeOfObject.НеОпределено;
    }

    // Интерфейсы
    public interface ITree {
        // Название изделия
        string NameProduct { get; } // Шифр корневой ноды дерева
        INode RootObject { get; } // Корневая нода дерева
        HashSet<ReferenceObject> AllReferenceObjects { get; set; } // Уникальные объекты дерева

        // Создание сообщение для лога с деревом изделия
        string GenerateLog();
    }

    public interface INode {
        string Name { get; }
        int Level { get; }
        ITree Tree { get; }
        INode Parent { get; }
        List<INode> Children { get; }
        ReferenceObject NomenclatureObject { get; }
    }

    // Перечисления

    private enum TypeOfObject {
        НеОпределено,
        Изделие,
        СборочнаяЕдиница,
        СтандартноеИзделие,
        ПрочееИзделие,
        Деталь,
        ЭлектронныйКомпонент,
        Материал,
        Другое,
    }

    // Классы
    private class NomenclatureTree : ITree {
        public string NameProduct { get; private set; }
        public INode RootObject { get; private set; }
        public HashSet<ReferenceObject> AllReferenceObjects { get; set; }

        public NomenclatureTree(Dictionary<string, ReferenceObject> nomenclature, Dictionary<string, List<ReferenceObject>> links, string shifr) {
            this.NameProduct = shifr;
            this.RootObject = new NomenclatureNode(this, null, nomenclature, links, shifr);
        }

        public string GenerateLog() {
            return string.Empty;
        }
    }

    private class NomenclatureNode : INode {
        public string Name { get; private set; }
        public int Level { get; private set; }
        public ITree Tree { get; private set; }
        public INode Parent { get; private set; }
        public List<INode> Children { get; private set; }
        public ReferenceObject NomenclatureObject { get; private set; }

        public NomenclatureNode (ITree tree, INode parent, Dictionary<string, ReferenceObject> nomenclature, Dictionary<string, List<ReferenceObject>> links, string shifr) {
            this.Tree = tree;
            this.Parent = parent;

            this.Level = parent == null ? 1 : parent.Level + 1;

            this.NomenclatureObject = nomenclature.ContainsKey(shifr) ?
                nomenclature[shifr] :
                throw new Exception($"Во время создания дерева изделия '{tree.NameProduct}' возникла ошибка. Обозначение '{shifr}' отсутствует в справочнике 'Список номенклатуры FoxPro'");

            // Подключаем номенклатурный объект в список всех объектов дерева
            this.Tree.AllReferenceObjects.Add(this.NomenclatureObject);

            // Получаем название объекта
            this.Name = (string)this.NomenclatureObject[Guids.Parameters.НоменклатураОбозначение].Value;

            // Рекурсивно получаем потомков
            this.Children = new List<INode>();
            foreach (string childShifr in links[shifr].Select(link => (string)link[Guids.Parameters.ПодключенияСборка].Value)) {
                this.Children.Add(new NomenclatureNode(this.Tree, this, nomenclature, links, childShifr));
            }
        }
    }

}
