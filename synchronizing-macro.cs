using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Windows.Forms;
using TFlex.DOCs.Model;
using TFlex.DOCs.Model.Macros;
using TFlex.DOCs.Model.References;



// Макрос для синхронизации макросов, которые находятся в справочнике "Макросы"

public class Macro : MacroProvider {

    public Macro (MacroContext context) : base(context)
    {
    }


    #region Fields and Properties

    private string nameOfLogFile = "sync-metadata.json";
    // Переменная, в которой будет храниться текущий лог
    private SyncMetaData currentSyncMetaData = new SyncMetaData();
    private Dictionary<string, MacroObject> dictOfMacros = new Dictionary<string, MacroObject>();
    private string serverName = string.Empty;
    private string pathToDirectory = string.Empty;
    
    // Списки макросов
    private List<MacroObject> onlyOnLocal = new List<MacroObject>();
    private List<MacroObject> onlyOnRemote = new List<MacroObject>();
    private List<MacroObject> differentOnBothBase = new List<MacroObject>();
    private List<MacroObject> sameOnBothBase = new List<MacroObject>();

    #region Guids

    private static class Guids {
        public static class References {
            public static Guid MacroReference = new Guid("3e6df4d0-b1d8-4375-978c-4da676604cca");
        }

        public static class Properties {
            public static Guid NameOfMacro = new Guid("8334853d-8b04-4716-b3e2-19bc0a360384");
            public static Guid CodeOfMacro = new Guid("3d654359-8567-49b3-8060-516f5f2f2ad2");
        }
    }

    #endregion Guids


    #endregion Fields and Properties

    #region Entry Points

    #region Run entry

    public override void Run() {
    // Основная точка входа
    // Данный метод будет проводить синхронизацию в общем порядке.
    // Остальные методы в точках входа будут выполнять более специфические задачи
        Testing(); 
    }

    #endregion Run entry

    #region Testing entry

    public void Testing() {
        // Для начала получаем список объектов из справочника
        serverName = Context.Connection.ServerName;

        GetAllMacrosFromRemoteBase();
        GetAllMacrosFromLocalBase();

        // Сортируем все полученные макросы для последующего анализа

        foreach (KeyValuePair<string, MacroObject> kvp in dictOfMacros) {
            if (kvp.Value.PerformAnalize()) {
                SortingMacroToLists(kvp.Value);
            }
            else
                Error(string.Format("Во время обработки {0} возникла ошибка, обратитесь к администратору для ее устранения\n{1}", kvp.Value.NameMacro, kvp.Value.ToString()));
        }

        // Выводим получившиеся списки макросов

        // Список макросов, над которыми не нужно производить никаких действий
        string message = "Список макросов, которые не требуют синхронизации:\n\n";

        foreach (MacroObject macroObject in sameOnBothBase) {
            message += string.Format("- {0}\n", macroObject.ToString("{nam}"));
        }

        Message("Информация", message);

        // Список макросов, которых нет на локальной машине
        message = "Список макросов, которых нет на локальной базе:\n\n";

        foreach (MacroObject macroObject in onlyOnRemote) {
            message += string.Format("- {0}\n", macroObject.ToString("{nam}"));
        }

        Message("Информация", message);

        // Список макросов, которых нет на удаленном сервере
        message = "Список макросов, которых нет на удаленной базе:\n\n";

        foreach (MacroObject macroObject in onlyOnLocal) {
            message += string.Format("- {0}\n", macroObject.ToString("{nam}"));
        }

        Message("Информация", message);

        // Список макросов, которые есть и там и там, но у них есть разница в коде
        message = "Список макросов, по синхронизации которых требуется принять решение:\n\n";

        foreach (MacroObject macroObject in differentOnBothBase) {
            message += string.Format("- {0}\n", macroObject.ToString("{nam}"));
        }

        Message("Информация", message);

        Message("Информация", "Работа макроса завершена");
    }

    #endregion Testing entry


    #endregion Entry Points

    #region Service methods
    
    #region Method SelectDirectory
    // Метод выбора директории для синхронизации при помощи диалога Windows.Forms
    private void SelectDirectory() {
        FolderBrowserDialog dialog = new FolderBrowserDialog();
        dialog.Description =
            "Выберите директорию для проведения синхронизации макросов";
        // Позволить пользователю создавать новую директорию из диалогового окна
        dialog.ShowNewFolderButton = true;
        // Данное свойство не принимает строки в качестве значения и требует специального класса SpecialFolder.
        // Для того, чтобы покрыть все возможнные варианты, была выбрана директория "Мой компьютер", так как она позволяет
        // получить доступ ко всему, что находится на компьютере
        dialog.RootFolder = Environment.SpecialFolder.MyComputer;

        if (dialog.ShowDialog() == DialogResult.OK)
            pathToDirectory = dialog.SelectedPath;
    }
    #endregion Method SelectDirectory

    //TODO Предусмотреть возможность использованимя стандартной директории для синхронизации, которую можно прописать в глобальных переменных (или зашить в макрос)

    #region Method GetAllMacrosFromRemoteBase
    // Метод для получения всех макросов, которые есть в системном справочнике Макросы
    private bool GetAllMacrosFromRemoteBase() {
        Reference macroReference = Context.Connection.ReferenceCatalog.Find(Guids.References.MacroReference).CreateReference();
        macroReference.Objects.Load();

        
        foreach (ReferenceObject macroFromRef in macroReference.Objects) {
            //TODO Добавить проверку на то, что это макрос, а не блок схема

            string name = macroFromRef[Guids.Properties.NameOfMacro].ToString();
            string code = macroFromRef[Guids.Properties.CodeOfMacro].ToString();
            DateTime modification = macroFromRef.SystemFields.EditDate;

            if (!dictOfMacros.ContainsKey(name)) {
                // Данного объекта еще не было добавлено
                MacroObject currentMacroObject = new MacroObject();
                // Производим первичное наполнение объекта
                currentMacroObject.NameMacro = name;
                currentMacroObject.CodeFromRemoteBase = code;
                currentMacroObject.DateOfRemoteMacro = modification;
                currentMacroObject.Location = LocationOfMacro.Remote;

                dictOfMacros[name] = currentMacroObject;
            }
            else {
                // Данный объект уже есть в базе, следовательно нужно обновить информацию о нем
                // Данный случай не должен срабатывать при чтении всех макросов с удаленной базы.
                // Если данная ветка сработала, значит у двух макросов одинаковое название,
                // что на данный момент не предусмотрено работой данного макроса
                string messageTemplate = "Во время чтения макросов из базы '{0}' было обнаружено несколько макросов с одинаковым наименованием - '{1}'";
                messageTemplate += "\nМакрос прекратит свою работу, так как его логика не предусматривает таких случаев.";
                messageTemplate += "\nЕсли вы увидели данное сообщение, свяжитесь с разработчиком макроса для того, чтобы макрос был модифицирован для обработки подобных случаев";
                Message("Внимание", string.Format(messageTemplate,
                                                    serverName,
                                                    name));

                // Прекращаем работу макроса
                return false;
            }
            
        }
        // Сообщаем вызвавшему об успешном завершении работы макроса
        return true;
    }
    #endregion Method GetAllMacrosFromRemoteBase

    #region Method GetAllMacrosFromLocalBase
    // Метод для получения всех макросов, которые находятся в синхронизуемой директории на локальной машине
    // Данный метод должен запускаться после того, как произойдет чтение из удаленной базы
    private bool GetAllMacrosFromLocalBase() {
        // Получаем путь к директории для синхронизации макросов
        SelectDirectory();

        if (pathToDirectory != string.Empty) {
            // Производим поиск макросов, расположенных в локальной директории
            foreach (string file in Directory.GetFiles(pathToDirectory, "*.cs")) {

                // Получаем всю необходимую информацию из файла
                string name = Path.GetFileNameWithoutExtension(file);
                string code = File.ReadAllText(file);
                DateTime modification = File.GetLastWriteTime(file);

                // Производим создание нового MacroObject или обновление уже существующего
                if (!dictOfMacros.ContainsKey(name)) {
                    // Случай, когда данного макроса не было на удаленной базе
                    MacroObject currentMacroObject = new MacroObject();
                    currentMacroObject.NameMacro = name;
                    currentMacroObject.CodeFromLocalBase = code;
                    currentMacroObject.DateOfLocalMacro = modification;
                    currentMacroObject.Location = LocationOfMacro.Local;

                    dictOfMacros[name] = currentMacroObject;
                }
                else {
                    // Случай, когда данный макрос уже существовал на удаленной базе
                    dictOfMacros[name].CodeFromLocalBase = code;
                    dictOfMacros[name].DateOfLocalMacro = modification;
                    dictOfMacros[name].Location = LocationOfMacro.RemoteAndLocal;
                }
            }
        }
        else
            return false;

        return true;
    }
    #endregion Method GetAllMacrosFromLocalBase

    private void SortingMacroToLists(MacroObject macroObject) {
        // TODO Реализовать метод сортировки
        if (macroObject.Action == SyncAction.LeaveUntouched) {
            sameOnBothBase.Add(macroObject);
            return;
        }
        if (macroObject.Action == SyncAction.CreateLocalOrDeleteRemote) {
            onlyOnRemote.Add(macroObject);
            return;
        }
        if (macroObject.Action == SyncAction.CreateRemoteOrDeleteLocal) {
            onlyOnLocal.Add(macroObject);
            return;
        }

        // Все остальные случаи добавляются в список отличающихся макросов,
        // которые при этом присутствуют на обеих базах
        differentOnBothBase.Add(macroObject);
    }
    
    // TODO
    // Реализовать метод, который будет обновлять (создавать) макросы на основании файлов в папке синхронизации
    
    #endregion Service methods

    #region Service classes

    // Дата класс, предназначенный для хранения всей важной информации, связанной с макросами

    private enum SyncAction {
        Unknown, // Стартовое действие, которое установлено по умолчанию до произведения анализа макросов на локальной и удаленной машине
        Update, // Случай, когда не понятно, какой макрос новее
        UpdateLocal, // Случай, когда макросы различаются и на удаленной машине более свежая версия
        CreateLocalOrDeleteRemote, // Случай, когда на локальной машине отсутствует макрос, который присутствует на удаленной машине
        UpdateRemote, // Случай, когда макросы различаются  и на локальной машине более свежая версия 
        CreateRemoteOrDeleteLocal, // Случай, когда на удаленной машине отсутствует макрос, который присутствует на локальной машине
        LeaveUntouched // Случай, когда макросы присутствуют на обоих базах и в их коде нет никаких отличий
    }

    private enum LocationOfMacro {
        Unknown, // Стартовое значение свойства, которое установлено по умолчанию до проведедния поиска макросов на удаленной базе и на локальной машине
        Remote, // Макрос есть только на удаленной машине
        Local, // Макрос есть только на локальной машине
        RemoteAndLocal // Макрос есть на удаленной и на локальной машине
    }

    private class MacroObject {
        public string NameMacro { get; set; }
        public string CodeFromLocalBase { get; set; }
        public string CodeFromRemoteBase { get; set; }
        public DateTime DateOfLocalMacro { get; set; }
        public DateTime DateOfRemoteMacro { get; set; }
        //public Guid GuidOfMacro { get; set; }   TODO определиться с тем, нужно ли это свойство
        public SyncAction Action { get; set; } = SyncAction.Unknown;
        public LocationOfMacro Location { get; set; } = LocationOfMacro.Unknown;

        //TODO Реализовать метод анализа макро объекта (который будет выбирать действие, которое нужно совершить над макросом
        public bool PerformAnalize() {
            // TODO Расставить ветки таким образом, чтобы с самого начала выполнялись самые распространенные случаи для увеличения быстродействия системы
            // Случай, когда на локальном компьютере есть макрос, которого нет на удаленной базе
            if (this.Location == LocationOfMacro.Local) {
                this.Action = SyncAction.CreateRemoteOrDeleteLocal;
                return true;
            }
            // Случай, когда на удаленной базе есть макрос, которого нет на локальной машине
            if (this.Location == LocationOfMacro.Remote) {
                this.Action = SyncAction.CreateLocalOrDeleteRemote;
                return true;
            }
            // Если макрос есть а в локальной директории и на удаленной базе, необхомо проверить, есть ли разница в коде макросов
            if (this.Location == LocationOfMacro.RemoteAndLocal) {
                if (this.CodeFromRemoteBase == this.CodeFromLocalBase) {
                    this.Action = SyncAction.LeaveUntouched;
                    return true;
                }
                else {
                    // Самый основной случай, когда будет разница между локальным репозиторием
                    int dateDiff = this.DateOfRemoteMacro.CompareTo(this.DateOfLocalMacro);
                    if (dateDiff > 0) {
                        // Случай, когда макрос на удаленной машине свежее макроса на локальной машине
                        this.Action = SyncAction.UpdateLocal;
                        return true;
                    }
                    else if (dateDiff < 0) {
                        // Случай, когда макрос на локальной машине свежее макроса на удаленной машине
                        this.Action = SyncAction.UpdateRemote;
                        return true;
                    }
                    else {
                        // Редкий случай, когда дата макроса на локальной машине совпадает до секунды с датой
                        // на удаленной машине (скорее всего такой вариант просто невозможен)
                        this.Action = SyncAction.Update;
                        return true;
                    }
                }
            }
            // Если код дошел до этой ветки, следовательно объект имеет необработанный тип LocationOfMacro,
            // чего не должно происходить. В этом случае мы возвращаем false, чтобы обозначить, что возникла
            // ошибка во время обработки
            return false;
        }

        public override string ToString() {
            return string.Format("{0}\nAction: {1}; Location: {2};",
                                    this.NameMacro,
                                    this.Action.ToString(),
                                    this.Location.ToString());
        }

        public string ToString(string template) {
            // Метод, который возвращает строковое отображение объекта в зависимости от переданного значения
            // Он распознает несколько ключевых слов внутри строки
            // {act}, {loc}, {nam}, {loccod}, {remcode}, {locdat}, {remdat}
            
            template = template.Replace("{nam}", "{0}");
            template = template.Replace("{act}", "{1}");
            template = template.Replace("{loc}", "{2}");
            template = template.Replace("{loccod}", "{3}");
            template = template.Replace("{remcod}", "{4}");
            template = template.Replace("{locdat}", "{5}");
            template = template.Replace("{remdat}", "{6}");
            
            return string.Format(template, this.NameMacro,
                                            this.Action.ToString(),
                                            this.Location.ToString(),
                                            this.CodeFromLocalBase,
                                            this.CodeFromRemoteBase,
                                            this.DateOfLocalMacro.ToString("dd.MM.yyyy"),
                                            this.DateOfRemoteMacro.ToString("dd.MM.yyyy"));
        }

        //TODO Реализовать метод, который будет производить действия над данный макросом в соответствии с тем, какое дайствие в нем выбрано
    }

    private class SyncMetaData {
        // Данный класс должен хранить данные о всех макросах, которые были перенесены путем синхронизации.
        // Так же он должен хранить данные о дате последней модификации макроса из справочника и дате, когда этот макрос был
        // в последний раз изменен на локальном компьютере
        // ДАнный класс должена хратить историю синхронизаций (не всех запусков, а именно историю переносов информации
        
        public string ServerName { get; private set; }
        public DateTime Date { get; private set; }

        public SyncMetaData () {
        }
    }

    #endregion Service classes
}
