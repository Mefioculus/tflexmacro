using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using TFlex.DOCs.Model;
using TFlex.DOCs.Model.Structure;
using TFlex.DOCs.Model.Desktop;
using TFlex.DOCs.Model.Classes;
using TFlex.DOCs.Model.Macros;
using TFlex.DOCs.Model.Macros.ObjectModel;
using TFlex.DOCs.Model.References;
using TFlex.DOCs.Model.References.Documents;
using TFlex.DOCs.Model.References.Nomenclature;

public class Macro : MacroProvider
{
    public Macro(MacroContext context)
        : base(context)
    {
    }

    #region поля и свойства основного класса Macro
    private Random random;
    private NomenclatureHierarchyLink hLink;


    private static class Guids
    {
        public static Guid ЭлектроннаяСтруктураИзделий = new Guid("853d0f07-9632-42dd-bc7a-d91eae4b8e83");

        public static Guid ТипСборочнаяЕдиница = new Guid("dd2cb8e8-48fa-4241-8cab-aac3d83034a7");
        public static Guid ТипДеталь = new Guid("7c41c277-41f1-44d9-bf0e-056d930cbb14");
        public static Guid ТипИзделие = new Guid("11d2fb6f-baa7-401c-bbd9-7f3222f5c5e8");

        public static Guid ПапкаОбучение = new Guid("22f1e188-9f37-43c4-82d6-6d21ed5f16e2");
        public static Guid ПараметрНаименование = new Guid("45e0d244-55f3-4091-869c-fcf0bb643765");
        public static Guid ПараметрОбозначение = new Guid("ae35e329-15b4-4281-ad2b-0e8659ad2bfb");

        public static Guid СвязьНаТипыСтруктур = new Guid("77726357-b0eb-4cea-afa5-182e21eb6373");
        public static Guid КоличествоПодключение = new Guid("3f5fc6c8-d1bf-4c3d-b7ff-f3e636603818");

        public static Guid LogicalObjectGuid = new Guid("49c7b3ec-fa35-4bb1-92a5-01d4d3a40d16");

    }

    private DocumentReference documentReference;
    private NomenclatureReference nomenclatureReference;

    public NomenclatureReference NomenclatureRef
    {
        get
        {
            if (nomenclatureReference == null)
                nomenclatureReference = new NomenclatureReference(Context.Connection);
            
            return nomenclatureReference;
        }
    }

    public DocumentReference DocumentRef
    {
        get
        {
            if (documentReference == null)
                documentReference = new DocumentReference(Context.Connection);
            
            return documentReference;
        }
    }
    #endregion


    #region СозданиеТестовойСборки - метод для генерации случайного дерева состава
    public void СозданиеТестовойСборки()
    {
        // Метод для воссоздания тестовой сборки для последующей обработки данных
        // Проверка на то, что справочники, необходимые для работы были найдены
        if ((DocumentRef == null) || (NomenclatureRef == null))
        {
            Message("Ошибка", "Не удалось получить справочник документы для создания нового объекта");
            return;
        }
        

        // Тестовые сборка будут размещаться в папке для обучения, так что получаем эту папку для размещения данных
        ReferenceObject folderForLearning = NomenclatureRef.Find(Guids.ПапкаОбучение);
        if (folderForLearning == null)
        {
            Message("Информация", "Папка для обучения не была найдена");
            return;
        }

        // Инициализируем класс для рандомизирования результатов
        random = new Random();
        // Создаем изделие
        ReferenceObject изделие = CreateNomenclatureObject("Тестовое изделие", GenerateRandomString("ИЗД"), Guids.ТипИзделие);
        Desktop.CheckIn(изделие, "Создание тестового изделия", false);

        // Подключаем изделие к папке назвначения
        LinkObjects(folderForLearning, изделие);

        // Код для случайного составления структуры
        GenerateTree(изделие, 3);

        Message("Информация", "Создание тестового дерева изделия завершено успешно");
        

    }
    #endregion

    #region LinkObjects - метод для создания связи между двумя объектами. Количество усанавливается случайным образом
    private void LinkObjects(ReferenceObject parent, ReferenceObject child)
    {
        hLink = null;
        // Подключаем дочерние изделие к родителю
        hLink = parent.CreateChildLink(child) as NomenclatureHierarchyLink;

        if (hLink == null)
        {
            Message("Информация", string.Format(
                        "Не удалось подключить {0} к {1}",
                        child[Guids.ПараметрОбозначение],
                        parent[Guids.ПараметрОбозначение]));
            return;
        }

        if (random.Next(2) == 0)
            hLink.Amount.Value = 1;
        else
            hLink.Amount.Value = random.Next(10);

        hLink.EndChanges();
    }
    #endregion

    #region GenerateTree - рекурсивный метод для генерации состава изделия
    private void GenerateTree (ReferenceObject parent, int limit)
    {
        if (limit > 0)
        {
            // Добавляем рандомное количество деталей и сборок через рекурсивный вызов этой функции
            // Создаем дочерние объекты
            int rand;
            for (int i = 0; i <= limit; i++)
            {
                rand = random.Next(3);
                if (rand == 0)
                {
                    ReferenceObject сборка = CreateNomenclatureObject("Тестовая сборка", GenerateRandomString("СБ"), Guids.ТипСборочнаяЕдиница);
                    Desktop.CheckIn(сборка, string.Format("Создание тестовой сборки в '{0}'", parent.ToString()), false);
                    LinkObjects(parent, сборка);
                    if (limit != 1)
                        limit = limit - random.Next(1, limit + 1);
                    else
                        limit = 0;
                    GenerateTree(сборка, limit);

                }
                else
                {
                    ReferenceObject деталь = CreateNomenclatureObject("Тестовая деталь", GenerateRandomString("ДЕТ"), Guids.ТипДеталь);
                    Desktop.CheckIn(деталь, string.Format("Создание тестовой детали в '{0}'", parent.ToString()), false);
                    LinkObjects(parent, деталь);
                }
            }
        }
        else
        {
            // Добавляем одну деталь
            ReferenceObject деталь = CreateNomenclatureObject("Тестовая деталь", GenerateRandomString("ДЕТ"), Guids.ТипДеталь);

            // Подключаем изделие к родителю
            LinkObjects(parent, деталь);
        }
    }
    #endregion

    private string GenerateRandomString(string str)
    {
        return string.Format("{0}{1}_{2}", str, random.Next(99999999).ToString(), DateTime.Now.ToString("dd.MM.yy"));
    }

    #region CreateNomenclatureObject - метод для создания номенклатурного объекта по названию и типу объекта
    private ReferenceObject CreateNomenclatureObject(string name, string denotation, Guid typeOfObject)
    {
        ClassObject type = DocumentRef.Classes.Find(typeOfObject);
        EngineeringDocumentObject newDocument 
            = DocumentRef.CreateReferenceObject(type) as EngineeringDocumentObject;

        if (newDocument == null)
        {
            Message("Ошибка", "Не удалось создать новый документ в справочнике документы");
            return null;
        }

        // Присваиваем документу обозначение и наименование
        newDocument.Name = name;
        newDocument.Denotation = denotation;

        newDocument.EndChanges();

        // Подключаем документ к номенклатуре

        return NomenclatureRef.CreateNomenclatureObject(newDocument);
    }
    #endregion

    #region УдалениеТестовойСборки - метод для удаления изделия и всех входящих в него деталей
    public void УдалениеТестовойСборки()
    {
        ReferenceObject изделие = GetObjectOnDelete();
        DeleteTreeRecursive(изделие);
        DeleteFromDocumentsAndNomenclature(изделие);
        Message("Информация", "Удаление тестового изделия произошло успешно");
    }

    private ReferenceObject GetObjectOnDelete()
    {
        ReferenceObject изделие = Context.ReferenceObject;

        if (изделие == null)
            return null;

        if (!Question(string.Format("Удалить изделие '{0}' и все входящие в него элементы?", изделие.ToString())))
            изделие = null;

        return изделие;
    }

    private void DeleteTreeRecursive(ReferenceObject изделие)
    {
        // Для начала нужно проверить, есть ли дочерние подключения у этого объекта
        if (изделие == null)
        {
            Message("Ошибка", "Объект не был найден");
            return;
        }

        var children = изделие.Children.GetHierarchyLinks();

        if (children != null)
        {
            foreach (ComplexHierarchyLink hLink in children)
            {
                ReferenceObject child = hLink.ChildObject;
                DeleteTreeRecursive(child);
                DeleteFromDocumentsAndNomenclature(child);
            }
        }
    }

    private void DeleteFromDocumentsAndNomenclature(ReferenceObject изделие)
    {
        //Начинаем удаление объекта
        изделие.CheckOut(true);
        Desktop.CheckIn(изделие, string.Format("Удаление изделия '{0}'", изделие.ToString()), false);
    }

    #endregion

    #region ПерепривязываниеНовыхРевизийКДругомуИзделию
    public void ПерепривязываниеНовыхРевизийКДругомуИзделию()
    {
        ReferenceObject from = GetReferenceObjectFrom();
        ReferenceObject to = GetReferenceObjectTo();
        string message = string.Empty;

        if (to == null)
        {
            message = "Не выбрано изделие для привязывания новых ревизий состава '{0}'";
            Message("Ошибка", string.Format(message, from.ToString()));
            return;
        }

        message = "Произвести создание новых ревизий для состава '{0}' и привязывания их к '{1}'";

        if (!Question(string.Format(message, from.ToString(), to.ToString())))
            return;

        List<ReferenceObject> listOfStructures = GetFromUserListOfStructures();

        
        RecursiveCreateRevision(from, to, listOfStructures);
        Message("Информация", "Создание ревизий и подключение их к новому изделию успешно завершено");
    }
    #endregion

    #region Методы для получения объекта, с которого будут браться данный и объекта, в который данные будут копироваться
    private ReferenceObject GetReferenceObjectFrom()
    {
        return Context.ReferenceObject;
    }

    private ReferenceObject GetReferenceObjectTo()
    {
        ДиалогВыбораОбъектов диалог = СоздатьДиалогВыбораОбъектов("Электронная структура изделий");
        диалог.Заголовок = "Выбор изделия для копирования состава";
        диалог.МножественныйВыбор = false;

        if (диалог.Показать())
            return (ReferenceObject)диалог.ВыбранныеОбъекты.FirstOrDefault();
        else
            return null;
    }
    #endregion


    #region RecursiveCreateRevision - основной метод, в котором будет происходить создание копии структуры с другими ревизиями
    private void RecursiveCreateRevision(ReferenceObject from, ReferenceObject to, List<ReferenceObject> listOfStructures)
    {
        string message = string.Empty;

        // Получаем все подключеные к изделию объекты
        List<ComplexHierarchyLink> hLinks = from.Children.GetHierarchyLinks();
        if (hLinks != null)
        {
            foreach (ComplexHierarchyLink hLink in hLinks)
            {
                // Получаем дочерний объект и создаем для него новую ревизию
                ReferenceObject childFrom = hLink.ChildObject;
                ReferenceObject newRevisionChild = childFrom.CreateRevision();

                // Удаляем связь новой ревизии с старым объектом
                ClearLinks(childFrom, newRevisionChild);

                // Привязываем новую ревизию к новому объекту
                var newLink = newRevisionChild.CreateParentLink(to);

                // Устанавливаем корректное значение количества в подключении
                newLink[Guids.КоличествоПодключение].Value = hLink[Guids.КоличествоПодключение].Value;
                newLink.EndChanges();

                
                // Присваиваем подключению ревизии те структуры, к которым она будет относиться
                ChangeStructures(newLink, listOfStructures);
                
                message = "Создание новой ревизии для '{0}' и подключение ревизии к '{1}'";
                Desktop.CheckIn(newRevisionChild, string.Format(message, newRevisionChild.ToString(), to.ToString()), false);

                // Вызываем данный метод рекурсивно для погружения глубже
                RecursiveCreateRevision(childFrom, newRevisionChild, listOfStructures);
            }
        }
        else
            return;
    }
    #endregion

    #region ClearLinks - Метод для удаления связи между двумя объектами
    private void ClearLinks(ReferenceObject parent, ReferenceObject child)
    {
        // Для начала получаем необходимое подключение
        ComplexHierarchyLink hLink = null;
        string message = string.Empty;

        hLink = parent.GetChildLink(child);
        if (hLink != null)
            if (!parent.DeleteLink(hLink))
            {
                message = "Не удалось отключить новую ревизию '{0}' от '{1}";
                Message("Ошибка", string.Format(message, child.ToString(), parent.ToString()));
            }
            
        else
        {
            message = "Не удалось получить подключение '{0} к '{1}'";
            Message("Ошибка", string.Format(message, child.ToString(), parent.ToString()));
        }
    }
    #endregion


    #region GetFromUserListOfStructures - метод для получения от пользователя списка структур для подлюкчения к ним новой ревизии
    private List<ReferenceObject> GetFromUserListOfStructures()
    {
        List<ReferenceObject> listOfStructures = new List<ReferenceObject>();
        ДиалогВыбораОбъектов диалог = СоздатьДиалогВыбораОбъектов("Типы структур изделий");
        диалог.Заголовок = "Выбор структур для новых ревизий";
        диалог.МножественныйВыбор = true;

        if (!диалог.Показать())
            return null;

        foreach (Объект structure in диалог.ВыбранныеОбъекты)
            listOfStructures.Add((ReferenceObject)structure);

        return listOfStructures;
    }
    #endregion


    #region ChangeStructures - метод для изменения структуры, к которой будет относиться новая ревизия
    private void ChangeStructures(ComplexHierarchyLink hLink, List<ReferenceObject> listOfStructures)
    {
        if (listOfStructures == null)
            return;
        hLink.BeginChanges();
        // Удаляем старые связи на структуры
        hLink.ClearLinks(Guids.СвязьНаТипыСтруктур);
        
        // Создаем новые связи на структуры
        foreach (ReferenceObject structure in listOfStructures)
            hLink.AddLinkedObject(Guids.СвязьНаТипыСтруктур, structure);
        hLink.EndChanges();
    }
    #endregion

    #region СозданиеТестовогоИзделия
    public void СозданиеТестовогоИзделия()
    {
        ReferenceObject currentObject = Context.ReferenceObject;
        random = new Random();

        ReferenceObject newObject = CreateNomenclatureObject ("Тестовое изделие", GenerateRandomString("ИЗД"), Guids.ТипИзделие);
        Desktop.CheckIn(newObject, "Создание тестового изделия", false);
        
        if (currentObject != null)
            // Подключаем созданный номенклатурный объект к текущему объекту
            LinkObjects(currentObject, newObject);
        else
        {
            Message("Ошибка", "Не указан текущий объект для подключения тестового изделия");
            return;
        }

        Message("Информация", string.Format("Тестовое изделие '{0}' создано в '{1}'", newObject.ToString(), currentObject.ToString()));
    }
    #endregion

    #region ПолученияСпискаРевизий
    public void ПолучениеСпискаРевизий()
    {
        // Получаем текущий объект и создаем для него дополнительные ревизии
        
        ReferenceObject currentObject = Context.ReferenceObject;

        // Пробуем найти все ревизии данного объекта в базе данных
        Guid guidOfObject = currentObject.SystemFields.LogicalObjectGuid;

        // Производим поиск объектов в справочнике
        string message = string.Empty;

        // Получаем ParameterInfo по guid
        ParameterInfo logicalObjectGuid = NomenclatureRef.ParameterGroup[Guids.LogicalObjectGuid];
        // Ищем объект в справочнике по параметру
        List<ReferenceObject> listOfRevisions = NomenclatureRef.Find(logicalObjectGuid, guidOfObject.ToString());

        foreach (ReferenceObject revision in listOfRevisions)
        {
            message += string.Format("{0}\n", revision.ToString());
        }
        Message("Информация", message);
    }
    #endregion
        
}
