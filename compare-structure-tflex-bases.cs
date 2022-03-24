using System;
using System.Text;
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
        int indent = 5;

        StructureDataBase structure = new StructureDataBase(Context.Connection, "Рабочая база", indent);
        structure.ExcludeKey("id");
        //Message($"Количество объектов {structure.FrendlyName}", structure.AllGuids.Count);
        //Message($"Объекты {structure.FrendlyName}", structure.ToString("all"));
        //Message($"Электронная структура объектов базы: {structure.FrendlyName}", structure.GetInfoAboutItem("853d0f07-9632-42dd-bc7a-d91eae4b8e83"));

        StructureDataBase otherStructure = GetStructureFromOtherServer(@"Gukovry", "TFLEX-DOCS:21324", "Макет", indent);
        if (otherStructure != null) {
            //Message($"Количество объектов {otherStructure.FrendlyName}", otherStructure.AllGuids.Count);
            //Message($"Электронная структура объектов базы: {otherStructure.FrendlyName}", otherStructure.GetInfoAboutItem("853d0f07-9632-42dd-bc7a-d91eae4b8e83"));
        }

        // Производим сравнение структур
        structure.Compare(otherStructure);
        Message("Разница", structure.PrintDifferences());
        Message("Информация", "Работа макроса завершена");
        
    }

    private StructureDataBase GetStructureFromOtherServer(string userName, string serverAddress, string dataBaseFrendlyName, int indent) {

        StructureDataBase resultStructure = null;

        using (ServerConnection serverConnection = ServerConnection.Open(userName, "123", serverAddress)) {
            if (!serverConnection.IsConnected)
                Message("Ошибка", $"Не удалось подключиться к серверу '{serverAddress}' под пользователем '{userName}'");
            else
                resultStructure = new StructureDataBase(serverConnection, dataBaseFrendlyName, indent);
        }

        return resultStructure;
    }


    public abstract class BaseNode {
        // Основные параметры
        public Guid Guid { get; set; }

        // Структурные параметры
        public BaseNode Parent { get; set; }
        public StructureDataBase Root { get; set; }
        public Dictionary<Guid, BaseNode> ChildNodes { get; set; } = new Dictionary<Guid, BaseNode>();
        private Dictionary<string, string> Properties = new Dictionary<string, string>();

        // Статус и отображение различия
        public CompareResult Status { get; set; }
        public TypeOfNode Type { get; set; } = TypeOfNode.неопределено;
        public List<string> Difference { get; set; } = new List<string>();
        public string StringDifference => GetStringRepresentationOfDifference();
        public List<string> MissedNodes { get; set; } = new List<string>();
        public bool IsVisibleForDiffReport { get; set; } = false;

        // Декоративные параметры
        public int Indent { get; set; }
        public int Level { get; set; }
        public string IndentString { get; set; }
        public string StringRepresentation => $"({this.Type.ToString()}) {this["name"]} (ID: {this["id"]}; Guid: {this.Guid.ToString()})";

        public BaseNode(StructureDataBase root, BaseNode parent, int indent, TypeOfNode type) {
            this.Root = root != null ? root : (StructureDataBase)this;
            this.Parent = parent;
            this.Indent = this.Root != null ? this.Root.Indent : indent;
            this.Type = type;
            this.Level = this.Parent == null ? 0 : this.Parent.Level + 1;
            this.IndentString = GetIndentString();
        }

        public BaseNode(StructureDataBase root, BaseNode parent, TypeOfNode type) : this(root, parent, 5, type) {
        }

        public string this[string key] {
            get {
                try {
                    return this.Properties[key];
                }
                catch (Exception e) {
                    throw new Exception($"При попытке получения значения по ключу '{key}' возникла ошибка:\n{e.Message}");
                }
            }
            set {
                this.Properties[key] = value;
            }
        }

        public void CompareWith(BaseNode otherNode) {
            if (this.Guid != otherNode.Guid) {
                throw new Exception($"Невозможно сравнить объекты с разными Guid: {this.Guid.ToString()} => {otherNode.Guid.ToString()}");
            }

            if (this.Type != otherNode.Type) {
                throw new Exception($"Невозможно произвести сравнение объектов с разными типами: {this.Type} => {otherNode.Type}. ({this["name"]})");
            }

            foreach (string key in this.Properties.Keys) {
                if ((this[key] != otherNode[key]) && !this.Root.ExcludedKeys[this.Type].Contains(key))
                    this.Difference.Add($"[{key}]: {this[key]} => {otherNode[key]}");
            }

            if (this.Difference.Count == 0)
                this.Status = CompareResult.ПолноеСоответствие;
            else if (this.Difference.Count == this.Properties.Count)
                this.Status = CompareResult.ПолноеРазличие;
            else
                this.Status = CompareResult.ЧастичноеРазличие;
        }

        public void Compare(BaseNode otherNode) {
            this.CompareWith(otherNode);

            // Сравниваем все справочники
            List<Guid> allReferenceGuidsInCurrentStructure = this.ChildNodes.Select(kvp => kvp.Key).ToList<Guid>();
            List<Guid> allReferenceGuidsInOtherStructure = otherNode.ChildNodes.Select(kvp => kvp.Key).ToList<Guid>();

            var missingInOther = allReferenceGuidsInCurrentStructure.Except(allReferenceGuidsInOtherStructure);
            var missingInCurrent = allReferenceGuidsInOtherStructure.Except(allReferenceGuidsInCurrentStructure);
            var existInBoth = allReferenceGuidsInCurrentStructure.Intersect(allReferenceGuidsInOtherStructure);

            foreach (Guid guid in missingInCurrent) {
                this.MissedNodes.Add(otherNode.ChildNodes[guid].StringRepresentation);
            }

            foreach (Guid guid in missingInOther) {
                this.ChildNodes[guid].Status = CompareResult.Новый;
            }

            // Запускаем сравнение справочников, которые есть и в первой и во второй структуре
            foreach (Guid guid in existInBoth) {
                this.ChildNodes[guid].Compare(otherNode.ChildNodes[guid]);
            }

            if (this.HaveDifference())
                SetAllParentsToVisible();
        }

        private void SetAllParentsToVisible() {
            this.IsVisibleForDiffReport = true;
            BaseNode currentType = this.Parent;

            while (true) {
                if ((currentType == null) || (currentType.IsVisibleForDiffReport == true))
                    break;
                currentType.IsVisibleForDiffReport = true;
            }
        }

        private bool HaveDifference() {
            if (this.Difference.Count != 0)
                return true;
            if (this.MissedNodes.Count != 0)
                return true;
            if (this.Status == CompareResult.Новый)
                return true;
            return false;
        }

        public string GetIndentString() {
            return this.Level != 0 ? new string(' ', (int)(this.Level * this.Indent)) : string.Empty;
        }

        private string GetStringRepresentationOfDifference() {
            StringBuilder result = new StringBuilder();

            if (this.HaveDifference()) {
                result.AppendLine(this.StringRepresentation);
                if (this.Difference.Count != 0) {
                    result.AppendLine("Обнаруженные отличия:");
                    result.AppendLine(string.Join("; ", this.Difference));
                }
                if (this.MissedNodes.Count != 0) {
                    result.AppendLine("Отсутствующие объекты:");
                    result.AppendLine(string.Join("\n", this.MissedNodes));
                }
            }

            return result.ToString();
        }

        public string PrintDifferences() {
            return string.Empty;
        }

        public override string ToString() {
            return $"{this.IndentString}({this.Type.ToString()}) {this["name"]}\n{string.Join(string.Empty, this.ChildNodes.Select(kvp => kvp.Value.ToString()))}";
        }
    
    }

    public class StructureDataBase : BaseNode {
        // Поля, свойственные только корню
        public string FrendlyName { get; set; }
        public Dictionary<Guid, List<BaseNode>> AllNodes { get; private set; } = new Dictionary<Guid, List<BaseNode>>();
        public List<Guid> AllGuids { get; private set; } = new List<Guid>();

        // Словарь с исключениями
        public Dictionary<TypeOfNode, List<string>> ExcludedKeys { get; private set; } = new Dictionary<TypeOfNode, List<string>>();

        public StructureDataBase(ServerConnection connection, string frendlyName, int indent) : base(null, null, indent, TypeOfNode.база) {
            this.FrendlyName = frendlyName;

            this.Guid = new Guid("00000000-0000-0000-0000-000000000000");
            this["id"] = "0";
            this["name"] = connection.ServerName;
            this["frendlyName"] = this.FrendlyName;

            foreach (ReferenceInfo refInfo in connection.ReferenceCatalog.GetReferences()) {
                this.AllGuids.Add(refInfo.Guid);
                StructureReference reference = new StructureReference(refInfo, this, this);
                this.ChildNodes.Add(reference.Guid, reference);
                if (!this.AllNodes.ContainsKey(reference.Guid))
                    this.AllNodes.Add(reference.Guid, new List<BaseNode>() { reference });
                else
                    this.AllNodes[reference.Guid].Add(reference);
            }
        }

        public void ExcludeKey(string key) {
            foreach (string name in Enum.GetNames(typeof(TypeOfNode)).Skip(1).ToArray<string>()) {
                this.ExcludeKey(key, (TypeOfNode)Enum.Parse(typeof(TypeOfNode), name));
            }
        }

        public void ExcludeKey(string key, TypeOfNode type) {
            if (type == TypeOfNode.неопределено)
                throw new Exception("При добавлении исключения не поддерживается неопределенный тип объекта");

            if (this.ExcludedKeys.ContainsKey(type))
                this.ExcludedKeys[type].Add(key);
            else
                this.ExcludedKeys[type] = new List<string>() { key };
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

    }

    public class StructureReference : BaseNode {

        public StructureReference(ReferenceInfo referenceInfo, StructureDataBase root, BaseNode parent) : base(root, parent, TypeOfNode.справочник) {
            this.Guid = referenceInfo.Guid;
            this["id"] = referenceInfo.Id.ToString();
            this["name"] = referenceInfo.Name;
            this.Status = CompareResult.НеОбработано;

            foreach (ClassObject classObject in referenceInfo.Classes.AllClasses.Where(cl => cl is ClassObject)) {
                this.Root.AddGuid(classObject.Guid);
                StructureClass structureClass = new StructureClass(classObject, this.Root, this);
                this.ChildNodes.Add(structureClass.Guid, structureClass);
                this.Root.AddNode(structureClass);
            }
        }

    }

    public class StructureClass : BaseNode {

        public StructureClass(ClassObject classObject, StructureDataBase root, BaseNode parent) : base(root, parent, TypeOfNode.тип) {
            this.Guid = classObject.Guid;
            this["id"] = classObject.Id.ToString();
            this["name"] = classObject.Name;
            this.Status = CompareResult.НеОбработано;

            // Производим поиск групп параметров справочника
            foreach (ParameterGroup group in classObject.GetAllGroups()) {
                this.Root.AddGuid(group.Guid);
                StructureParameterGroup parameterGroup = new StructureParameterGroup(group, this.Root, this);
                this.ChildNodes.Add(parameterGroup.Guid, parameterGroup);
                this.Root.AddNode(parameterGroup);
            }
        }

    }

    public class StructureParameterGroup : BaseNode {

        public StructureParameterGroup(ParameterGroup group, StructureDataBase root, BaseNode parent) : base(root, parent, TypeOfNode.группа) {
            this.Guid = group.Guid;
            this["id"] = group.Id.ToString();
            this["name"] = group.Name;
            this.Status = CompareResult.НеОбработано;

            // Получаем параметры из группы параметров
            foreach (ParameterInfo parameterInfo in group.Parameters) {
                this.Root.AddGuid(parameterInfo.Guid);
                StructureParameter parameter = new StructureParameter(parameterInfo, this.Root, this);
                this.ChildNodes.Add(parameter.Guid, parameter);
                this.Root.AddNode(parameter);
            }
        }

    }

    public class StructureParameter : BaseNode {

        public StructureParameter(ParameterInfo parameterInfo, StructureDataBase root, BaseNode parent) : base(root, parent, TypeOfNode.параметр) {
            this.Guid = parameterInfo.Guid;
            this["id"] = parameterInfo.Id.ToString();
            this["name"] = parameterInfo.Name;
            this.Status = CompareResult.НеОбработано;
        }

    }

    public class StructureLink : BaseNode {

        public StructureLink(StructureDataBase root, BaseNode parent, int id, Guid guid, string name) : base(root, parent, TypeOfNode.связь) {
            this.Guid = guid;
            this["id"] = id.ToString();
            this["name"] = name;
        }

    }

    public enum CompareResult {
        НеОбработано,
        РазницаВПотомках,
        ПолноеСоответствие,
        ПолноеРазличие,
        ЧастичноеРазличие,
        Новый
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
