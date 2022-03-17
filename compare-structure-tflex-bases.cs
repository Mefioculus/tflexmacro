using System;
using System.IO;
using System.Linq;
using System.Collections;
using System.Collections.Generic;

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

        StructureDataBase structure = new StructureDataBase(Context.Connection);
        Message("Структура базы данных", structure.ToString(0, 5));
        StructureDataBase otherStructure = GetStructureFromOtherServer(@"Gukovry", "TFLEX-DOCS:21324");
        if (otherStructure != null) {
            Message("Структура базы данных", otherStructure.ToString(0, 5));
        }
        
    }

    private StructureDataBase GetStructureFromOtherServer(string userName, string serverAddress) {

        StructureDataBase resultStructure = null;

        using (ServerConnection serverConnection = ServerConnection.Open(userName, "123", serverAddress)) {
            if (!serverConnection.IsConnected)
                Message("Ошибка", $"Не удалось подключиться к серверу '{serverAddress}' под пользователем '{userName}'");
            else
                resultStructure = new StructureDataBase(serverConnection);
        }

        return resultStructure;
    }



    public class StructureDataBase {
        public string Name { get; set; }
        public Dictionary<Guid, StructureReference> References { get; set; }

        public StructureDataBase(ServerConnection connection) {
            this.Name = connection.ServerName;

            this.References = new Dictionary<Guid, StructureReference>();
            foreach (ReferenceInfo refInfo in connection.ReferenceCatalog.GetReferences()) {
                this.References.Add(refInfo.Guid, new StructureReference(refInfo));
            }
        }

        public string ToString(int currPadding, int deltaPadding) {
            string stringPadding = currPadding == 0 ? string.Empty : new string(' ', currPadding);
            return $"Структура справочников сервера {this.Name}:\n\n{stringPadding}{string.Join("\n", this.References.Select(kvp => $" {kvp.Value.ToString(currPadding, deltaPadding)}"))}";
        }
    }

    public class StructureReference {
        public Guid Guid { get; set; }
        public int ID { get; set; }
        public string Name { get; set; }
        
        public Dictionary<Guid, StructureClass> Classes { get; set; }

        public StructureReference(ReferenceInfo referenceInfo) {
            this.Guid = referenceInfo.Guid;
            this.ID = referenceInfo.Id;
            this.Name = referenceInfo.Name;

            this.Classes = new Dictionary<Guid, StructureClass>();
            foreach (ClassObject classObject in referenceInfo.Classes.AllClasses.Where(cl => cl is ClassObject)) {
                this.Classes.Add(classObject.Guid, new StructureClass(classObject));
            }
        }

        public string ToString(int currPadding, int deltaPadding) {
            string stringPadding = currPadding == 0 ? string.Empty : new string(' ', currPadding);
            return $"{stringPadding}(справочник) {this.Name}\n{string.Join("\n", this.Classes.Select(kvp => kvp.Value.ToString(currPadding + deltaPadding, deltaPadding)))}";
        }
    }

    public class StructureClass {
        public Guid Guid { get; set; }
        public int ID { get; set; }
        public string Name { get; set; }
        
        public Dictionary<Guid, StructureParameterGroups> ParameterGroups { get; set; }
        public Dictionary<Guid, StructureLink> Links { get; set; }

        public StructureClass(ClassObject classObject) {
            this.Guid = classObject.Guid;
            this.ID = classObject.Id;
            this.Name = classObject.Name;

            this.ParameterGroups = new Dictionary<Guid, StructureParameterGroups>();
            this.Links = new Dictionary<Guid, StructureLink>();

            // Производим поиск групп параметров справочника
            foreach (ParameterGroup group in classObject.GetAllGroups()) {
                this.ParameterGroups.Add(group.Guid, new StructureParameterGroups(group));
            }
        }

        public string ToString(int currPadding, int deltaPadding) {
            string stringPadding = currPadding == 0 ? string.Empty : new string(' ', currPadding);
            return $"{stringPadding}(тип) {this.Name}\n{string.Join("\n", this.ParameterGroups.Select(kvp => kvp.Value.ToString(currPadding + deltaPadding, deltaPadding)))}";
        }
        
    }

    public class StructureParameterGroups {
        public Guid Guid { get; set; }
        public int ID { get; set; }
        public string Name { get; set; }

        public Dictionary<Guid, StructureParameter> Parameters { get; set; }

        public StructureParameterGroups(ParameterGroup group) {
            this.Guid = group.Guid;
            this.ID = group.Id;
            this.Name = group.Name;
        }

        public string ToString(int currPadding, int deltaPadding) {
            string stringPadding = currPadding == 0 ? string.Empty : new string(' ', currPadding);
            return $"{stringPadding}(группа) {this.Name}";
        }
    }

    public class StructureParameter {
        public Guid Guid { get; set; }
        public int ID { get; set; }
        public string Name { get; set; }

        public StructureParameter(int id, Guid guid, string name) {
            this.Guid = guid;
            this.ID = id;
            this.Name = name;
        }
    }

    public class StructureLink {
        public Guid Guid { get; set; }
        public int ID { get; set; }
        public string Name { get; set; }

        public StructureLink(int id, Guid guid, string name) {
            this.Guid = guid;
            this.ID = id;
            this.Name = name;
        }
    }
}
