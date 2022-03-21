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

        StructureDataBase otherStructure = GetStructureFromOtherServer(@"Gukovry", "TFLEX-DOCS:21324", "Макет", startPadding, deltaPadding);
        if (otherStructure != null) {
            Message($"Количество объектов {otherStructure.FrendlyName}", otherStructure.AllGuids.Count);
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


    public interface IStructureNode {
        public Guid Guid { get; set; }
        public int ID { get; set; }
        public string Name { get; set; }
        public IStructureNode Parent { get; set; }
        public StructureDataBase Root { get; set; }
        public CompareResult Status { get; set; }
        public int Padding { get; set; }
        public int Delta { get; set; }

        public string ToString(string type);
    
    }

    public class StructureDataBase {
        public string Name { get; set; }
        public string FrendlyName { get; set; }
        public int Padding { get; set; }
        public int Delta { get; set; }
        public Dictionary<Guid, StructureReference> References { get; set; }
        public Dictionary<Guid, List<IStructureNode>> AllNodes { get; private set; }
        public List<Guid> AllGuids { get; private set; }

        public StructureDataBase(ServerConnection connection, string frendlyName, int padding, int delta) {
            this.Name = connection.ServerName;
            this.FrendlyName = frendlyName;
            this.Padding = padding;
            this.Delta = delta;

            this.AllNodes = new Dictionary<Guid, List<IStructureNode>>();
            this.AllGuids = new List<Guid>();

            this.References = new Dictionary<Guid, StructureReference>();
            foreach (ReferenceInfo refInfo in connection.ReferenceCatalog.GetReferences()) {
                this.AllGuids.Add(refInfo.Guid);
                StructureReference reference = new StructureReference(refInfo, this, null, this.Padding, this.Delta);
                this.References.Add(reference.Guid, reference);
                if (!this.AllNodes.ContainsKey(reference.Guid))
                    this.AllNodes.Add(reference.Guid, new List<IStructureNode>() { reference });
                else
                    this.AllNodes[reference.Guid].Add(reference);
            }
        }

        public void AddGuid(Guid guid) {
            this.AllGuids.Add(guid);
        }

        public void AddNode(IStructureNode node) {
            if (!this.AllNodes.ContainsKey(node.Guid))
                this.AllNodes.Add(node.Guid, new List<IStructureNode>() { node });
            else
                this.AllNodes[node.Guid].Add(node);
        }

        public string ToString(string type) {
            string stringPadding = this.Padding == 0 ? string.Empty : new string(' ', this.Padding);
            switch (type) {
                case "all":
                    return $"Структура справочников сервера {this.Name}:\n\n{stringPadding}{string.Join("\n", this.References.Select(kvp => $" {kvp.Value.ToString("all")}"))}";
                default:
                    return string.Empty;
            }
        }
    }

    public class StructureReference : IStructureNode {
        public Guid Guid { get; set; }
        public int ID { get; set; }
        public string Name { get; set; }
        public IStructureNode Parent { get; set; }
        public StructureDataBase Root { get; set; }
        public CompareResult Status { get; set; }
        public int Padding { get; set; }
        public int Delta { get; set; }
        
        public Dictionary<Guid, StructureClass> Classes { get; set; }

        public StructureReference(ReferenceInfo referenceInfo, StructureDataBase root, IStructureNode parent, int padding, int delta) {
            this.Guid = referenceInfo.Guid;
            this.ID = referenceInfo.Id;
            this.Name = referenceInfo.Name;
            this.Parent = parent;
            this.Root = root;
            this.Status = CompareResult.НеОбработано;
            this.Padding = padding;
            this.Delta = delta;

            this.Classes = new Dictionary<Guid, StructureClass>();
            foreach (ClassObject classObject in referenceInfo.Classes.AllClasses.Where(cl => cl is ClassObject)) {
                this.Root.AddGuid(classObject.Guid);
                StructureClass structureClass = new StructureClass(classObject, this.Root, this, this.Padding + this.Delta, this.Delta);
                this.Classes.Add(structureClass.Guid, structureClass);
                this.Root.AddNode(structureClass);
            }
        }

        public string ToString(string type) {
            string stringPadding = this.Padding == 0 ? string.Empty : new string(' ', this.Padding);
            switch (type) {
                case "all":
                    return $"{stringPadding}(справочник) {this.Name}\n{string.Join("\n", this.Classes.Select(kvp => kvp.Value.ToString("all")))}";
                default:
                    return string.Empty;
            }
        }
    }

    public class StructureClass : IStructureNode {
        public Guid Guid { get; set; }
        public int ID { get; set; }
        public string Name { get; set; }
        public IStructureNode Parent { get; set; }
        public StructureDataBase Root { get; set; }
        public CompareResult Status { get; set; }
        public int Padding { get; set; }
        public int Delta { get; set; }
        
        public Dictionary<Guid, StructureParameterGroup> ParameterGroups { get; set; }
        public Dictionary<Guid, StructureLink> Links { get; set; }

        public StructureClass(ClassObject classObject, StructureDataBase root, IStructureNode parent, int padding, int delta) {
            this.Guid = classObject.Guid;
            this.ID = classObject.Id;
            this.Name = classObject.Name;
            this.Parent = parent;
            this.Root = root;
            this.Status = CompareResult.НеОбработано;
            this.Padding = padding;
            this.Delta = delta;

            this.ParameterGroups = new Dictionary<Guid, StructureParameterGroup>();
            this.Links = new Dictionary<Guid, StructureLink>();

            // Производим поиск групп параметров справочника
            foreach (ParameterGroup group in classObject.GetAllGroups()) {
                this.Root.AddGuid(group.Guid);
                StructureParameterGroup parameterGroup = new StructureParameterGroup(group, this.Root, this, this.Padding + this.Delta, this.Delta);
                this.ParameterGroups.Add(parameterGroup.Guid, parameterGroup);
                this.Root.AddNode(parameterGroup);
            }
        }

        public string ToString(string type) {
            string stringPadding = this.Padding == 0 ? string.Empty : new string(' ', this.Padding);
            switch (type) {
                case "all":
                    return $"{stringPadding}(тип) {this.Name}\n{string.Join("\n", this.ParameterGroups.Select(kvp => kvp.Value.ToString("all")))}";
                default:
                    return string.Empty;
            }
        }
        
    }

    public class StructureParameterGroup : IStructureNode {
        public Guid Guid { get; set; }
        public int ID { get; set; }
        public string Name { get; set; }
        public IStructureNode Parent { get; set; }
        public StructureDataBase Root { get; set; }
        public CompareResult Status { get; set; }
        public int Padding { get; set; }
        public int Delta { get; set; }

        public Dictionary<Guid, StructureParameter> Parameters { get; set; }

        public StructureParameterGroup(ParameterGroup group, StructureDataBase root, IStructureNode parent, int padding, int delta) {
            this.Guid = group.Guid;
            this.ID = group.Id;
            this.Name = group.Name;
            this.Parent = parent;
            this.Root = root;
            this.Status = CompareResult.НеОбработано;
            this.Padding = padding;
            this.Delta = delta;

            this.Parameters = new Dictionary<Guid, StructureParameter>();

            // Получаем параметры из группы параметров
            foreach (ParameterInfo parameterInfo in group.Parameters) {
                this.Root.AddGuid(parameterInfo.Guid);
                StructureParameter parameter = new StructureParameter(parameterInfo, this.Root, this, this.Padding + this.Delta, this.Delta);
                this.Parameters.Add(parameter.Guid, parameter);
                this.Root.AddNode(parameter);
            }
        }

        public string ToString(string type) {
            string stringPadding = this.Padding == 0 ? string.Empty : new string(' ', this.Padding);
            switch (type) {
                case "all":
                    return $"{stringPadding}(группа) {this.Name}\n{string.Join("\n", this.Parameters.Select(kvp => kvp.Value.ToString("all")))}";
                default:
                    return string.Empty;
            }
        }
    }

    public class StructureParameter : IStructureNode {
        public Guid Guid { get; set; }
        public int ID { get; set; }
        public string Name { get; set; }
        public IStructureNode Parent { get; set; }
        public StructureDataBase Root { get; set; }
        public CompareResult Status { get; set; }
        public int Padding { get; set; }
        public int Delta { get; set; }

        public StructureParameter(ParameterInfo parameterInfo, StructureDataBase root, IStructureNode parent, int padding, int delta) {
            this.Guid = parameterInfo.Guid;
            this.ID = parameterInfo.Id;
            this.Name = parameterInfo.Name;
            this.Root = root;
            this.Parent = parent;
            this.Status = CompareResult.НеОбработано;
            this.Padding = padding;
            this.Delta = delta;
        }

        public string ToString(string type) {
            string stringPadding = this.Padding == 0 ? string.Empty : new string(' ', this.Padding);
            switch (type) {
                case "all":
                    return $"{stringPadding}(Параметр) {this.Name}";
                default:
                    return string.Empty;
            }
        }
    }

    public class StructureLink : IStructureNode {
        public Guid Guid { get; set; }
        public int ID { get; set; }
        public string Name { get; set; }
        public IStructureNode Parent { get; set; }
        public StructureDataBase Root { get; set; }
        public CompareResult Status { get; set; }
        public int Padding { get; set; }
        public int Delta { get; set; }

        public StructureLink(int id, Guid guid, string name, int padding, int delta) {
            this.Guid = guid;
            this.ID = id;
            this.Name = name;
            this.Padding = padding;
            this.Delta = delta;
        }

        public Dictionary<Guid, IStructureNode> GetNodes() {
            // TODO: Реализовать метод получения всех дочерних нод
            return null;
        }

        public string ToString(string type) {
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
        ОтличаетсяНаименование,
        ОтличаетсяID,
        ОтличаетсяНаименованиеИID
    }
}
