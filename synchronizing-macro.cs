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
    private string serverName = Context.Connection.ServerName;
    private string pathToDirectory = string.Empty;

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
        Message("Информация", "Работа макроса завершена");
    }

    #endregion Testing entry


    #endregion Entry Points

    #region Service methods
    
    #region Method SelectDirectory
    // Метод выбора директории для синхронизации при помощи диалога Windows.Forms
    private SelectDirectory() {
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

    #region Method GetAllMacrosFromRemoteBase
    // Метод для получения всех макросов, которые есть в системном справочнике Макросы
    private bool GetAllMacrosFromRemoteBase() {
        Reference macroReference = Context.Connection.ReferenceCatalog.Find(Guids.References.MacroReference).CreateReference;
        macroReference.Objects.Load();

        
        foreach (ReferenceObject macroFromRef in macroReference) {
            //TODO Добавить проверку на то, что это макрос, а не блок схема

            string name = macroFromRef[Guids.Properties.NameOfMacro];
            string code = macroFromRef[Guids.Properties.CodeOfMacro];
            DateTime modification = macroFromRef.SystemFields.LastEdit;

            if (!dictOfMacros.ContainsKey(nameOfMacro)) {
                // Данного объекта еще не было добавлено
                MacroObject currentMacroObject = new MacroObject();
                // Производим первичное наполнение объекта
                currentMacroObject.NameMacro = name;
                currentMacroObject.CodeFromRemoteBase = code;
                currentMacroObject.DateOfRemoteMacro = modification;
                currentMacroObject.Location = LocationOfMacro.Remote;

                dictOfMacros[nameOfMacro] = currentMacroObject;
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
                                                    nameOfMacro));

                // Прекращаем работу макроса
                return false;
            }
            
            // Сообщаем вызвавшему об успешном завершении работы макроса
            return true;
        }
    }
    #endregion Method GetAllMacrosFromRemoteBase

    #region Method GetAllMacrosFromLocalBase
    // Метод для получения всех макросов, которые находятся в синхронизуемой директории на локальной машине
    // Данный метод должен запускаться после того, как произойдет чтение из удаленной базы
    private bool GetAllMacrosFromLocalBase(string pathToDirectory) {
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
    
    // TODO
    // Реализовать метод, который будет обновлять (создавать) макросы на основании файлов в папке синхронизации
    
    #endregion Service methods

    #region Service classes

    // Дата класс, предназначенный для хранения всей важной информации, связанной с макросами

    private enum SyncAction {
        UpdateFile,
        CreateFile,
        UpdateMacro,
        CreateMacro,
        DeleteMacro,
        DeleteFile,
        None
    }

    private enum LocationOfMacro {
        Remote,
        Local,
        RemoteAndLocal,
        Unknown
    }

    private class MacroObject {
        public string NameMacro { get; set; }
        public string CodeFromLocalBase { get; set; }
        public string CodeFromRemoteBase { get; set; }
        public DateTime DateOfLocalMacro { get; set; }
        public DateTime DateOfRemoteMacro { get; set; }
        //public Guid GuidOfMacro { get; set; }   TODO определиться с тем, нужно ли это свойство
        public SyncAction Action { get; set; } = SyncAction.None;
        public LocationOfMacro Location { get; set; } = LocationOfMacro.Unknown;

        //TODO Реализовать метод анализа макро объекта (который будет выбирать действие, которое нужно совершить над макросом

        //TODO Реализовать метод вывода объекта в виде текста;

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
            this.ServerName = Context.Connection.ServerName;
        }
    }

    #endregion Service classes
}
