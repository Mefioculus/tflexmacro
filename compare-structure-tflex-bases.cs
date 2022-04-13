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
using TFlex.DOCs.Model.References.Events;

/*
 * Для работы макроса так же потребуется подключение в качестве ссылки библиотеки TFlex.DOCs.Common.dll
*/

public class Macro : MacroProvider {
    public Macro(MacroContext context)
        : base (context) {
        }

    public override void Run() {
        int indent = 5;

        // Названия полей
        string serverNameField = "Имя сервера";
        string frendlyNameField = "Короткое название базы";
        string userNameField = "Имя пользователя";
        string passwordField = "Пароль";
        string pathToSaveDiffField = "Путь для сохранения отчета";
        string directionOfCompareField = "Обратное сравнение";
        string excludedFromComparingFields = "Исключенные поля";

        // Переменные

        string serverName = "TFLEX-DOCS:21324";
        string userName = "Gukovry";
        string frendlyName = "Макет";
        string pathToFile = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "result.txt");
        string password = "123";

        // Перечень полей, по которым идет сравнение
        List<string> allFields = new List<string>() {
            "id",
            "name",
            "frendlyName",
            "countOfEvents",
            "eventNames",
            "countOfHandlers",
            "handlerNames",
            "lengthParam",
            "nullableParam",
            "typenameParam",
            "unitParam",
            "valuesParam",
            "typeLink",
            "countOfDialogPages"
        };

        InputDialog dialog = new InputDialog(this.Context, "Укажите параметры для подключения к базе данных");
        dialog.AddString(serverNameField, serverName);
        dialog.AddString(frendlyNameField, frendlyName);
        dialog.AddString(userNameField, userName);
        dialog.AddString(passwordField, password);
        dialog.AddString(pathToSaveDiffField, pathToFile);
        dialog.AddMultiselectFromList(excludedFromComparingFields, allFields);
        dialog.AddFlag(directionOfCompareField, false);

        if (dialog.Show()) {
            serverName = dialog[serverNameField];
            userName = dialog[userNameField];
            frendlyName = dialog[frendlyNameField];
            pathToFile = dialog[pathToSaveDiffField];
            password = dialog[passwordField];

            // Приступаем к чтению данных
            StDataBase structure = new StDataBase(Context.Connection, "Текущая база", indent);
            StDataBase otherStructure = GetStructureFromOtherServer(userName, serverName, frendlyName, indent);

            if (dialog[excludedFromComparingFields] != null) {
                foreach (object key in dialog[excludedFromComparingFields]) {
                    structure.ExcludeKey((string)key);
                    otherStructure.ExcludeKey((string)key);
                }
            }

            if (dialog[directionOfCompareField] == false) {
                structure.Compare(otherStructure);
                File.WriteAllText(pathToFile, structure.PrintDifferences());
            }
            else {
                otherStructure.Compare(structure);
                File.WriteAllText(pathToFile, otherStructure.PrintDifferences());
            }

            // Прозводим открытие файла в блокноте
            System.Diagnostics.Process notepad = new System.Diagnostics.Process();
            notepad.StartInfo.FileName = "notepad.exe";
            notepad.StartInfo.Arguments = pathToFile;
            notepad.Start();

        }
        else
            return;
    }

    private StDataBase GetStructureFromOtherServer(string userName, string serverAddress, string dataBaseFrendlyName, int indent) {

        StDataBase resultStructure = null;

        using (ServerConnection serverConnection = ServerConnection.Open(userName, "123", serverAddress)) {
            if (!serverConnection.IsConnected)
                Message("Ошибка", $"Не удалось подключиться к серверу '{serverAddress}' под пользователем '{userName}'");
            else
                resultStructure = new StDataBase(serverConnection, dataBaseFrendlyName, indent);
        }

        return resultStructure;
    }


    public abstract class StBaseNode {
        // Основные параметры
        public Guid Guid { get; set; }

        // Структурные параметры
        public StBaseNode Parent { get; set; }
        public StDataBase Root { get; set; }
        public Dictionary<Guid, StBaseNode> ChildNodes { get; set; } = new Dictionary<Guid, StBaseNode>();
        private Dictionary<string, string> Properties = new Dictionary<string, string>();
        public Dictionary<string, string>.KeyCollection Keys => this.Properties.Keys;

        // Статус и отображение различия
        public CompareResult Status { get; set; }
        public TypeOfNode Type { get; set; } = TypeOfNode.UDF;
        public List<string> Differences { get; set; } = new List<string>();
        public List<string> MissedNodes { get; set; } = new List<string>();

        public bool IsVisibleForDiffReport { get; set; } = false;

        // Декоративные параметры
        public int Indent { get; set; }
        public int Level { get; set; }
        public string IndentString { get; set; }
        public string StringRepresentation => $"({this.Type.ToString()}) {this["name"]} (ID: {this["id"]}; Guid: {this.Guid.ToString()})";

        public StBaseNode(StDataBase root, StBaseNode parent, int indent, TypeOfNode type) {
            this.Root = root != null ? root : (StDataBase)this;
            this.Parent = parent;
            //this.Indent = this.Root != null ? this.Root.Indent : indent;
            this.Indent = 5;
            this.Type = type;
            this.Level = this.Parent == null ? 0 : this.Parent.Level + 1;
            this.IndentString = GetIndentString();
            this.Status = CompareResult.NTP;
        }

        public StBaseNode(StDataBase root, StBaseNode parent, TypeOfNode type) : this(root, parent, 5, type) {
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

        public void CompareWith(StBaseNode otherNode) {
            if (this.Guid != otherNode.Guid) {
                throw new Exception($"Невозможно сравнить объекты с разными Guid: {this.Guid.ToString()} => {otherNode.Guid.ToString()}");
            }

            if (this.Type != otherNode.Type) {
                throw new Exception($"Невозможно произвести сравнение объектов с разными типами: {this.Type} => {otherNode.Type}. ({this["name"]})");
            }

            foreach (string key in this.Keys) {
                if ((this[key] != otherNode[key]) && (!this.Root.ExcludedKeys.ContainsKey(this.Type) || !this.Root.ExcludedKeys[this.Type].Contains(key)))
                    this.Differences.Add($"[{key}]: {this[key]} => {otherNode[key]}");
            }

            if (this.Differences.Count == 0)
                this.Status = CompareResult.EQL;
            else
                this.Status = CompareResult.DIF;
        }

        public void Compare(StBaseNode otherNode) {
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
                this.ChildNodes[guid].Status = CompareResult.NEW;
                this.ChildNodes[guid].SetAllChildNodesStatus(CompareResult.NEW);
                //this.ChildNodes[guid].SetAllChildNodesVisible(); // - Данную строку следует распомментировать, если нужно сделать видимыми все дочерние элементы нового объекта (которые тоже соответственно будут новыми)
                this.ChildNodes[guid].IsVisibleForDiffReport = true;
            }

            // Запускаем сравнение справочников, которые есть и в первой и во второй структуре
            foreach (Guid guid in existInBoth) {
                this.ChildNodes[guid].Compare(otherNode.ChildNodes[guid]);
            }

            if (this.HaveDifference())
                SetAllParentsToVisible();
        }

        private void SetAllChildNodesStatus(CompareResult status) {
            this.Status = status;
            foreach (StBaseNode node in this.ChildNodes.Select(kvp => kvp.Value)) {
                node.SetAllChildNodesStatus(status);
            }
        }

        private void SetAllChildNodesVisible() {
            this.IsVisibleForDiffReport = true;
            foreach (StBaseNode node in this.ChildNodes.Select(kvp => kvp.Value)) {
                node.SetAllChildNodesVisible();
            }
        }

        private void SetAllParentsToVisible() {
            this.IsVisibleForDiffReport = true;
            StBaseNode currentType = this.Parent;

            while (true) {
                if ((currentType == null) || (currentType.IsVisibleForDiffReport == true))
                    break;
                currentType.IsVisibleForDiffReport = true;
            }
        }

        private bool HaveDifference() {
            if (this.Differences.Count != 0)
                return true;
            if (this.MissedNodes.Count != 0)
                return true;
            if (this.Status == CompareResult.NEW)
                return true;
            return false;
        }

        public string GetIndentString() {
            return this.Level != 0 ? new string(' ', (int)(this.Level * this.Indent)) : string.Empty;
        }

        public string PrintDifferences() {
            string tree = $"{this.IndentString}({this.Type.ToString()})"; // Строковое представление входимости объектов и их тип
            string name = this["name"]; // Имя объекта
            string stat = this.Status.ToString(); // Статус объекта
            string diff = this.Differences.Count.ToString(); // Количество отличий
            string nw = this.ChildNodes.Where(kvp => kvp.Value.Status == CompareResult.NEW).Count().ToString(); // Количество новых
            string miss = this.MissedNodes.Count.ToString(); // Количество отсутствующих
            string childElements = string.Join(string.Empty, this.ChildNodes.Where(kvp => kvp.Value.IsVisibleForDiffReport).Select(kvp => kvp.Value.PrintDifferences())); // Информация по входяхим элементам

            // Получаем детальную информацию о позиции
            StringBuilder details = new StringBuilder();
            if ((this.Differences.Count != 0) || this.MissedNodes.Count != 0)
                details.AppendLine();
            foreach (string difference in this.Differences)
                details.AppendLine($"{"--DETAILS--  ", 26}- (diff){difference}");
            foreach (string missed in this.MissedNodes)
                details.AppendLine($"{"--DETAILS--  ", 26}- (miss){missed}");


            return $"{tree, -25} {name,-80} (stat: {stat, 3}, diff: {diff, 2}, new: {nw, 2}, miss: {miss, 2}){details}\n{childElements}";
        }

        public override string ToString() {
            return $"{this.IndentString}({this.Type.ToString()}) {this["name"]} ({this.Status})\n{string.Join(string.Empty, this.ChildNodes.Select(kvp => kvp.Value.ToString()))}";
        }
    
    }

    public class StDataBase : StBaseNode {
        // Поля, свойственные только корню
        public string FrendlyName { get; set; }
        public Dictionary<Guid, List<StBaseNode>> AllNodes { get; private set; } = new Dictionary<Guid, List<StBaseNode>>();
        public List<Guid> AllGuids { get; private set; } = new List<Guid>();

        // Словарь с исключениями
        public Dictionary<TypeOfNode, List<string>> ExcludedKeys { get; private set; } = new Dictionary<TypeOfNode, List<string>>();

        public StDataBase(ServerConnection connection, string frendlyName, int indent) : base(null, null, indent, TypeOfNode.SRV) {
            this.FrendlyName = frendlyName;

            this.Guid = new Guid("00000000-0000-0000-0000-000000000000");
            this["id"] = "0";
            this["name"] = connection.ServerName;
            this["frendlyName"] = this.FrendlyName;

            foreach (ReferenceInfo refInfo in connection.ReferenceCatalog.GetReferences()) {
                this.AllGuids.Add(refInfo.Guid);
                StReference reference = new StReference(refInfo, this, this);
                this.ChildNodes.Add(reference.Guid, reference);
                if (!this.AllNodes.ContainsKey(reference.Guid))
                    this.AllNodes.Add(reference.Guid, new List<StBaseNode>() { reference });
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
            if (type == TypeOfNode.UDF)
                throw new Exception("При добавлении исключения не поддерживается неопределенный тип объекта");

            if (this.ExcludedKeys.ContainsKey(type))
                this.ExcludedKeys[type].Add(key);
            else
                this.ExcludedKeys[type] = new List<string>() { key };
        }

        public StBaseNode GetInfoAboutReference(Guid guid) {
            if (this.AllNodes.ContainsKey(guid))
                return this.ChildNodes[guid];
            return null;
        }

        public void AddGuid(Guid guid) {
            this.AllGuids.Add(guid);
        }

        public void AddNode(StBaseNode node) {
            if (!this.AllNodes.ContainsKey(node.Guid))
                this.AllNodes.Add(node.Guid, new List<StBaseNode>() { node });
            else
                this.AllNodes[node.Guid].Add(node);
        }

    }

    public class StReference : StBaseNode {

        public StReference(ReferenceInfo referenceInfo, StDataBase root, StBaseNode parent) : base(root, parent, TypeOfNode.REF) {
            this.Guid = referenceInfo.Guid;
            this["id"] = referenceInfo.Id.ToString();
            this["name"] = referenceInfo.Name;

            foreach (ClassObject classObject in referenceInfo.Classes.AllClasses.Where(cl => cl is ClassObject)) {
                this.Root.AddGuid(classObject.Guid);
                StClass structureClass = new StClass(classObject, this.Root, this);
                this.ChildNodes.Add(structureClass.Guid, structureClass);
                this.Root.AddNode(structureClass);
            }
        }

    }

    public class StClass : StBaseNode {

        public StClass(ClassObject classObject, StDataBase root, StBaseNode parent) : base(root, parent, TypeOfNode.TYP) {
            this.Guid = classObject.Guid;
            this["id"] = classObject.Id.ToString();
            this["name"] = classObject.Name;
            this["countOfEvents"] = classObject.Events.GetUserEvents().Count.ToString();
            this["eventNames"] = string.Join("; ", classObject.Events.GetUserEvents().Select(ev => ev.Name).OrderBy(name => name));
            this["countOfHandlers"] = classObject.Events.Handlers.Count.ToString();
            this["handlerNames"] = string.Join("; ", classObject.Events.Handlers.Select(handler => handler.HandlerName).OrderBy(name => name));
            this["countOfDialogPages"] = classObject.Dialog != null ?
                classObject.Dialog.Groups.Select(group => group.Pages != null ? group.Pages.Count : 0).Sum().ToString() :
                "Отсутствует";


            // Производим поиск групп параметров справочника
            foreach (ParameterGroup group in classObject.GetAllGroups()) {
                this.Root.AddGuid(group.Guid);

                if (group.IsLinkGroup) {
                    StLink link = new StLink(group, this.Root, this);
                    this.ChildNodes.Add(link.Guid, link);
                    this.Root.AddNode(link);
                }
                else {
                    StParameterGroup parameterGroup = new StParameterGroup(group, this.Root, this);
                    this.ChildNodes.Add(parameterGroup.Guid, parameterGroup);
                    this.Root.AddNode(parameterGroup);
                }
            }

        }

    }

    public class StParameterGroup : StBaseNode {

        public StParameterGroup(ParameterGroup group, StDataBase root, StBaseNode parent) : base(root, parent, TypeOfNode.GRP) {
            this.Guid = group.Guid;
            this["id"] = group.Id.ToString();
            this["name"] = group.Name;

            // Получаем параметры из группы параметров
            foreach (ParameterInfo parameterInfo in group.Parameters) {
                this.Root.AddGuid(parameterInfo.Guid);
                StParameter parameter = new StParameter(parameterInfo, this.Root, this);
                this.ChildNodes.Add(parameter.Guid, parameter);
                this.Root.AddNode(parameter);
            }
        }

    }

    public class StParameter : StBaseNode {

        public StParameter(ParameterInfo parameterInfo, StDataBase root, StBaseNode parent) : base(root, parent, TypeOfNode.PRM) {
            this.Guid = parameterInfo.Guid;
            this["id"] = parameterInfo.Id.ToString();
            this["name"] = parameterInfo.Name;
            this["lengthParam"] = parameterInfo.Length.ToString();
            this["nullableParam"] = parameterInfo.Nullable.ToString();
            this["typenameParam"] = parameterInfo.TypeName != null ? parameterInfo.TypeName : "null";
            this["unitParam"] = parameterInfo.Unit != null ? parameterInfo.Unit.Name : "null";
            this["valuesParam"] = parameterInfo.ValueList != null ? string.Join("; ", parameterInfo.ValueList.Select(item => $"{item.Name} - {item.Value}")) : "null";
        }

    }

    public class StLink : StBaseNode {

        public StLink(ParameterGroup group, StDataBase root, StBaseNode parent) : base(root, parent, TypeOfNode.LNK) {
            this.Guid = group.Guid;
            this["id"] = group.Id.ToString();
            this["name"] = group.Name;
            this["typeLink"] = group.LinkType.ToString();
        }

    }

    public enum CompareResult {
        NTP, //Не обработано
        EQL, //Соответствует
        DIF, //Есть разница 
        NEW  //Новый (отсутствует в сравниваемом справочнике)
    }

    public enum TypeOfNode {
        UDF, //Не определено
        SRV, //Сервер
        REF, //Справочник
        TYP, //Тип
        GRP, //Группа
        PRM, //Параметр
        LNK  //Связь
    }
}
