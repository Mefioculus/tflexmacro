using System;
using System.IO;
using System.Linq;
using System.Collections;
using System.Collections.Generic;
using System.Windows.Forms;

using TFlex.DOCs.Model;
using TFlex.DOCs.Common.Encryption; // Для осуществления подключения к серверу
using TFlex.DOCs.Model.Macros;
using TFlex.DOCs.Model.Macros.ObjectModel;
using TFlex.DOCs.Model.Structure;
using TFlex.DOCs.Model.Search;
using TFlex.DOCs.Model.Classes;
using TFlex.DOCs.Model.References;
using TFlex.DOCs.Model.References.Files;

/*
 * Для работы макроса так же потребуется подключение в качестве ссылки библиотеки TFlex.DOCs.Common.dll
*/

public class Macro : MacroProvider {
    public Macro(MacroContext context)
        : base (context) {
        }

    public override void Run() {
        int startPadding = 0;
        int deltaPadding = 5;

        StructureDataBase structure = new StructureDataBase(Context.Connection, "Рабочая база", startPadding, deltaPadding);
        Message($"Количество объектов {structure.FrendlyName}", structure.AllGuids.Count);
        //Message($"Электронная структура объектов базы: {structure.FrendlyName}", structure.GetInfoAboutItem("853d0f07-9632-42dd-bc7a-d91eae4b8e83"));

        StructureDataBase otherStructure = GetStructureFromOtherServer(@"Gukovry", "TFLEX-DOCS:21324", "Макет", startPadding, deltaPadding);
        if (otherStructure != null) {
            Message($"Количество объектов {otherStructure.FrendlyName}", otherStructure.AllGuids.Count);
            //Message($"Электронная структура объектов базы: {otherStructure.FrendlyName}", otherStructure.GetInfoAboutItem("853d0f07-9632-42dd-bc7a-d91eae4b8e83"));
        }
        
    }

    private StructureDataBase GetStructureFromOtherServer(string userName, string serverAddress, string dataBaseFrendlyName, int padding, int delta) {

        StructureDataBase resultStructure = null;

        using (ServerConnection serverConnection = ServerConnection.Open(userName, "123", serverAddress)) {
            if (!serverConnection.IsConnected)
                Message("Ошибка", $"Не удалось подключиться к серверу '{serverAddress}' под пользователем '{userName}'");
            else
                resultStructure = new StructureDataBase(serverConnection, dataBaseFrendlyName, padding, delta);
        }

        return resultStructure;
    }


    public abstract class BaseNode {
        // Основные параметры
        public Guid Guid { get; set; }
        public int Id { get; set; }
        public string Name { get; set; }

        // Структурные параметры
        public BaseNode Parent { get; set; }
        public StructureDataBase Root { get; set; }
        public Dictionary<Guid, BaseNode> ChildNodes { get; set; } = new Dictionary<Guid, BaseNode>();

        // Статус и отображение различия
        public CompareResult Status { get; set; }
        public TypeOfNode Type { get; set; } = TypeOfNode.неопределено;
        public List<string> Difference { get; set; } = new List<string>();
        public string StringDifference => Difference.Count == 0 ? string.Empty : string.Join("; ", this.Difference);
        public List<string> MissedNodes { get; set; } = new List<string>();

        // Декоративные параметры
        public int Padding { get; set; }
        public int Delta { get; set; }
        public string StringRepresentation => $"({this.Type.ToString()}) {this.Name} (ID: {this.Id.ToString()}; Guid: {this.Guid.ToString()})";

        public void CompareWith(BaseNode otherNode) {
            if (this.Guid != otherNode.Guid) {
                throw new Exception($"Невозможно сравнить объекты с разными Guid: {this.Guid.ToString()} => {otherNode.Guid.ToString()}");
            }

            if (this.Type != otherNode.Type) {
                throw new Exception($"Невозможно произвести сравнение объектов с разными типами: {this.Type} => {otherNode.Type}. ({this.Name})");
            }

            // Производим сравнение
            bool diffId = this.Id != otherNode.Id;
            bool diffName = this.Name != otherNode.Name;

            if (diffId || diffName) {
                if (diffId && diffName)
                    this.Status = CompareResult.ПолноеРазличие;
                else
                    this.Status = diffId ? CompareResult.РазличиеId : CompareResult.РазличиеName;
            }
            else {
                this.Status = CompareResult.ПолноеСоответствие;
            }

            if (diffId)
                this.Difference.Add($"Id: '{this.Id}' => '{otherNode.Id}'");
            if (diffName)
                this.Difference.Add($"Name: '{this.Name}' => '{otherNode.Name}'");

        }

        public abstract void Compare(BaseNode node);
        public abstract string ToString(string type);
    
    }

    public class StructureDataBase : BaseNode {
        // Поля, свойственные только корню
        public string FrendlyName { get; set; }
        public Dictionary<Guid, List<BaseNode>> AllNodes { get; private set; } = new Dictionary<Guid, List<BaseNode>>();
        public List<Guid> AllGuids { get; private set; } = new List<Guid>();

        public StructureDataBase(ServerConnection connection, string frendlyName, int padding, int delta) {
            this.Id = 0;
            this.Guid = new Guid("00000000-0000-0000-0000-000000000000");
            this.Parent = null;
            this.Root = null;
            this.Type = TypeOfNode.база;

            this.Name = connection.ServerName;
            this.FrendlyName = frendlyName;
            this.Padding = padding;
            this.Delta = delta;

            foreach (ReferenceInfo refInfo in connection.ReferenceCatalog.GetReferences()) {
                this.AllGuids.Add(refInfo.Guid);
                StructureReference reference = new StructureReference(refInfo, this, null, this.Padding, this.Delta);
                this.ChildNodes.Add(reference.Guid, reference);
                if (!this.AllNodes.ContainsKey(reference.Guid))
                    this.AllNodes.Add(reference.Guid, new List<BaseNode>() { reference });
                else
                    this.AllNodes[reference.Guid].Add(reference);
            }
        }

        /*
        public string GetInfoAboutItem(Guid guid, string type = "all") {
            if (this.AllNodes.ContainsKey(guid)) {
                return AllNodes[guid][0].ToString(type);
            }
            else return $"Не удалось найти Guid: {guid.ToString()}";
        }
        
        public string GetInfoAboutItem(string stringGuid, string type = "all") {
            return GetInfoAboutItem(new Guid(stringGuid), type);
        }

        */

        public override void Compare(BaseNode node) {
            StructureDataBase otherStructure = node as StructureDataBase;
            if (otherStructure == null)
                throw new Exception("Ошибка при сравнении объектов");
            // Сравниваем все справочники
            List<Guid> allReferenceGuidsInCurrentStructure = this.ChildNodes.Select(kvp => kvp.Key).ToList<Guid>();
            List<Guid> allReferenceGuidsInOtherStructure = otherStructure.ChildNodes.Select(kvp => kvp.Key).ToList<Guid>();

            var missingInCurrent = allReferenceGuidsInCurrentStructure.Except(allReferenceGuidsInOtherStructure);
            var missingInOther = allReferenceGuidsInOtherStructure.Except(allReferenceGuidsInCurrentStructure);
            var existInBoth = allReferenceGuidsInCurrentStructure.Intersect(allReferenceGuidsInOtherStructure);

            foreach (Guid guid in missingInCurrent) {
                this.MissedNodes.Add(otherStructure.ChildNodes[guid].StringRepresentation);
            }

            foreach (Guid guid in missingInOther) {
                this.ChildNodes[guid].Status = CompareResult.Новый;
            }

            // Запускаем сравнение справочников, которые есть и в первой и во второй структуре
            foreach (Guid guid in existInBoth) {
                this.ChildNodes[guid].Compare(otherStructure.ChildNodes[guid]);
            }
        }

        public void AddGuid(Guid guid) {
            this.AllGuids.Add(guid);
        }

        public void AddNode(BaseNode node) {
            if (!this.AllNodes.ContainsKey(node.Guid))
                this.AllNodes.Add(node.Guid, new List<BaseNode>() { node });
            else
                this.AllNodes[node.Guid].Add(node);
        }

        public override string ToString(string type) {
            string stringPadding = this.Padding == 0 ? string.Empty : new string(' ', this.Padding);
            switch (type) {
                case "all":
                    return $"Структура справочников сервера {this.Name}:\n\n{stringPadding}{string.Join("\n", this.ChildNodes.Select(kvp => $" {kvp.Value.ToString("all")}"))}";
                default:
                    return string.Empty;
            }
        }
    }

    public class StructureReference : BaseNode {

        public StructureReference(ReferenceInfo referenceInfo, StructureDataBase root, BaseNode parent, int padding, int delta) {
            this.Type = TypeOfNode.справочник;
            this.Guid = referenceInfo.Guid;
            this.Id = referenceInfo.Id;
            this.Name = referenceInfo.Name;
            this.Parent = parent;
            this.Root = root;
            this.Status = CompareResult.НеОбработано;
            this.Padding = padding;
            this.Delta = delta;

            foreach (ClassObject classObject in referenceInfo.Classes.AllClasses.Where(cl => cl is ClassObject)) {
                this.Root.AddGuid(classObject.Guid);
                StructureClass structureClass = new StructureClass(classObject, this.Root, this, this.Padding + this.Delta, this.Delta);
                this.ChildNodes.Add(structureClass.Guid, structureClass);
                this.Root.AddNode(structureClass);
            }
        }

        public override void Compare(BaseNode node) {

            // Производим сравнение данной ноды
            this.CompareWith(node);

            // Запускаем сравнение для всех входящих нод


            StructureReference otherReference = node as StructureReference;
            if (otherReference == null)
                throw new Exception("При сравнении объекта справочника возникла ошибка");

        }

        public override string ToString(string type) {
            string stringPadding = this.Padding == 0 ? string.Empty : new string(' ', this.Padding);
            switch (type) {
                case "all":
                    return $"{stringPadding}(справочник) {this.Name}\n{string.Join("\n", this.ChildNodes.Select(kvp => kvp.Value.ToString("all")))}";
                default:
                    return string.Empty;
            }
        }
    }

    public class StructureClass : BaseNode {

        public StructureClass(ClassObject classObject, StructureDataBase root, BaseNode parent, int padding, int delta) {
            this.Type = TypeOfNode.тип;
            this.Guid = classObject.Guid;
            this.Id = classObject.Id;
            this.Name = classObject.Name;
            this.Parent = parent;
            this.Root = root;
            this.Status = CompareResult.НеОбработано;
            this.Padding = padding;
            this.Delta = delta;

            // Производим поиск групп параметров справочника
            foreach (ParameterGroup group in classObject.GetAllGroups()) {
                this.Root.AddGuid(group.Guid);
                StructureParameterGroup parameterGroup = new StructureParameterGroup(group, this.Root, this, this.Padding + this.Delta, this.Delta);
                this.ChildNodes.Add(parameterGroup.Guid, parameterGroup);
                this.Root.AddNode(parameterGroup);
            }
        }

        public override void Compare(BaseNode node) {
            this.CompareWith(node);
        }

        public override string ToString(string type) {
            string stringPadding = this.Padding == 0 ? string.Empty : new string(' ', this.Padding);
            switch (type) {
                case "all":
                    return $"{stringPadding}(тип) {this.Name}\n{string.Join("\n", this.ChildNodes.Select(kvp => kvp.Value.ToString("all")))}";
                default:
                    return string.Empty;
            }
        }
        
    }

    public class StructureParameterGroup : BaseNode {

        public StructureParameterGroup(ParameterGroup group, StructureDataBase root, BaseNode parent, int padding, int delta) {
            this.Type = TypeOfNode.группа;
            this.Guid = group.Guid;
            this.Id = group.Id;
            this.Name = group.Name;
            this.Parent = parent;
            this.Root = root;
            this.Status = CompareResult.НеОбработано;
            this.Padding = padding;
            this.Delta = delta;

            // Получаем параметры из группы параметров
            foreach (ParameterInfo parameterInfo in group.Parameters) {
                this.Root.AddGuid(parameterInfo.Guid);
                StructureParameter parameter = new StructureParameter(parameterInfo, this.Root, this, this.Padding + this.Delta, this.Delta);
                this.ChildNodes.Add(parameter.Guid, parameter);
                this.Root.AddNode(parameter);
            }
        }

        public override void Compare(BaseNode node) {
            this.CompareWith(node);
        }

        public override string ToString(string type) {
            string stringPadding = this.Padding == 0 ? string.Empty : new string(' ', this.Padding);
            switch (type) {
                case "all":
                    return $"{stringPadding}(группа) {this.Name}\n{string.Join("\n", this.ChildNodes.Select(kvp => kvp.Value.ToString("all")))}";
                default:
                    return string.Empty;
            }
        }
    }

    public class StructureParameter : BaseNode {

        public StructureParameter(ParameterInfo parameterInfo, StructureDataBase root, BaseNode parent, int padding, int delta) {
            this.Type = TypeOfNode.параметр;
            this.Guid = parameterInfo.Guid;
            this.Id = parameterInfo.Id;
            this.Name = parameterInfo.Name;
            this.Root = root;
            this.Parent = parent;
            this.Status = CompareResult.НеОбработано;
            this.Padding = padding;
            this.Delta = delta;
        }

        public override void Compare(BaseNode node) {
            this.CompareWith(node);
        }

        public override string ToString(string type) {
            string stringPadding = this.Padding == 0 ? string.Empty : new string(' ', this.Padding);
            switch (type) {
                case "all":
                    return $"{stringPadding}(Параметр) {this.Name}";
                default:
                    return string.Empty;
            }
        }
    }

    public class StructureLink : BaseNode {

        public StructureLink(int id, Guid guid, string name, int padding, int delta) {
            this.Type = TypeOfNode.связь;
            this.Guid = guid;
            this.Id = id;
            this.Name = name;
            this.Padding = padding;
            this.Delta = delta;
        }

        public override void Compare(BaseNode node) {
            this.CompareWith(node);
        }

        public override string ToString(string type) {
            string stringPadding = this.Padding == 0 ? string.Empty : new string(' ', this.Padding);
            switch (type) {
                case "all":
                    return string.Empty;
                default:
                    return string.Empty;
            }
        }
    }

    public enum CompareResult {
        НеОбработано,
        ПолноеСоответствие,
        Отсутствует,
        Новый,
        РазличиеId,
        РазличиеName,
        ПолноеРазличие
    }

    public enum TypeOfNode {
        неопределено,
        база,
        справочник,
        тип,
        группа,
        параметр,
        связь
    }
}
