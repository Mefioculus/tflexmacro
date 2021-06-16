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

    private string nameOfLogFile = "last_sync.json";
    // Переменная, в которой будет храниться текущий лог
    private SyncronizationLog currentSyncLog = new SyncronizationLog();

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

    public override void Run() {
    // Основная точка входа
    // Данный метод будет проводить синхронизацию в общем порядке.
    // Остальные методы в точках входа будут выполнять более специфические задачи
    
        string pathToDirectory = SelectDirectory();
        if (pathToDirectory == string.Empty) {
            return;
        }

        // Получаем файл, в котором будет храниться информация об предыдущей синхронизации
        string pathToLogFile = Path.Combine(pathToDirectory, nameOfLogFile);

        //TODO
        //Реализовать проверку сервера при синхронизации
        //(Если предыдущая синхроназация проводилась с другого сервера, должно выскочить предупреждение)

        //TODO
        //Основаная последовательность действий при синхронизации
        //- Получаем директорию от пользователя
        //- В данной директории производим поиск json файла, который будет хранит в себе данные
        //о прошлой синхронизации
        //-  Если данный файл был в наличии, производим его чтение, если его нет, считаем, что загрузка проиозводится впервые
        //- Произвести чтение всех данных из справочника с макросами
        //
        //ЕСЛИ ЗАГРУЗКА ПРОИЗВОДИТСЯ ВПЕРВЫЕ
        //- Проверям данную директорию на наличие файлов cs
        //- Если файлов нет, производим создание файлов в соответствии с тем, что получилось прочитать из справочника
        //- Если файлы есть, нужно произвести сравнение, создать отсутствующие файлы, спросить про перезапись присутствующих
        //и спросить про удаление лишних
        //
        //ЕСЛИ ЗАГРУЗКА ПРОИЗВОДИТСЯ НЕ В ПЕРВЫЙ РАЗ
        //- Произвести чтение файла истории
        //- Произвести сравнение файла истории с списком макросов, которые были прочтены из справочника
        //- Произвести сравнение файла истории с списокм макросов, которые были прочтены из справчоника
        //- Спросить что делать с отсутствующими на локальном компьютере файлами (следует ли их удалять в справочнике, или же из следует снова воссоздать в директории
        //- Спосить, что делать с отсутствующими в справочнике макросами (следует ли их добавть, следует ли удалить их из локальной директории, или же следует ничего не предпринимать)
        //- Спрость, что делать с измененными файлами, которые имеют более свежую дажу модификации на локальной машине
        //- Спросить, что делать с измененными файлами, которые имеют более свежую дату модификации в справочнике
        
        //- В конце вывести справочное сообщение с произведенными изменениями
        //- Все изменения в процессе фиксировать в файле истории
        //- По завершению синхронизации произвести запись файла истории для последующих интеграций


        // Производим чтение всех макросов, находящихся в справочнике
        Dictionary<string, MacrosObject> macroFromMacroReference = GetAllMacrosFromMacroReference();

        // Производим поиск и чтение файла истории
        if (File.Exists(pathToLogFile)) {
            // Данная ветка предполагает, что синхронизация ранее производилась
        }
        else {
            // Данная ветка предполгатает, что синхнонизация производится в первый раз
        }

        // Производим чтение всех макросов из директории
        Dictionary<string, MacrosObject> macroFromLocalMachine = GetAllMacrosFromLocalMachine(pathToDirectory);


        // Производим анализ полученных данных

    }

    public void MethodForTestPurpose() {
        // Проверка запроса директории
        string directory = SelectDirectory();
        Message("Проверка выбора директории", string.Format("Выбранная директория для проведения синхронизации - '{0}'", directory));

        // Проверка запроса данных из справочника


        // Проверка запроса данных с локальной директории
    }


    #endregion Entry Points

    #region Service methods
    // Метод выбора директории для синхронизации при помощи диалога Windows.Forms
    private string SelectDirectory() {
        FolderBrowserDialog dialog = new FolderBrowserDialog();
        dialog.Description =
            "Выберите директорию для проведения синхронизации макросов";
        // Позволить пользователю создавать новую директорию из диалогового окна
        dialog.ShowNewFolderButton = true;
        // Устанавливаем стартовый путь для диалога на папку "Мои документы"
        dialog.RootFolder = Environment.SpecialFolder.Personal;

        if (dialog.ShowDialog() == DialogResult.OK) {
            return dialog.SelectedPath;
        }
        
        return string.Empty;
    }


    // Метод для получения всех макросов, которые есть в системном справочнике Макросы
    private Dictionary<string, MacrosObject> GetAllMacrosFromMacroReference() {

        currentSyncLog.ServerName = Context.Connection.ServerName;

        Dictionary<string, MacrosObject> result = new Dictionary<string, MacrosObject>();

        Reference macroReference = Context.Connection.ReferenceCatalog.Find(Guids.References.MacroReference).CreateReference();
        // Производим загрузку объектов справочника на первом уровне иерархии (так как справочник плоский,
        // то это загрузка всех объектов
        macroReference.Objects.Load();

        foreach (ReferenceObject macroObj in macroReference.Objects) {
            // Проходим через все объекты справочника и собираем данные
            MacrosObject docsMacro = new MacrosObject();
            docsMacro.Name = (string)macroObj[Guids.Properties.NameOfMacro];
            docsMacro.Code = (string)macroObj[Guids.Properties.CodeOfMacro];
            docsMacro.LastModificationDate = macroObj.SystemFields.EditDate;
            docsMacro.GuidOfMacro = macroObj.SystemFields.Guid;

            result[docsMacro.Name] = docsMacro;
        }

        return result;
    }

    // Метод для получения всех макросов, которые находятся в синхронизуемой директории на локальной машине
    private Dictionary<string, MacrosObject> GetAllMacrosFromLocalMachine(string pathToDirectory) {
        Dictionary<string, MacrosObject> result = new Dictionary<string, MacrosObject>();

        // Получаем пути ко всем файлам с расширением *.cs
        string[] files = Directory.GetFiles(pathToDirectory, "*.cs");

        if (files.Length > 0) {
            foreach (string file in files) {
                MacrosObject localMacro = new MacrosObject();
                localMacro.Name = Path.GetFileNameWithoutExtension(file);
                localMacro.Code = File.ReadAllText(file);
                //TODO
                //Реализовать получение даты не из параметров файла, а из специального файла, который
                //будет генерироваться во время первой синхранизации и будет хранить параметры синхронизации
                localMacro.LastModificationDate = File.GetLastWriteTime(file);
                result[localMacro.Name] = localMacro;
            }
        }

        return result;
    }
    
    // TODO
    // Реализовать метод сравнения макросов
    
    #endregion Service methods

    #region Service classes

    // Дата класс, предназначенный для хранения всей важной информации, связанной с макросами
    private class MacrosObject {
        public string Name { get; set; }
        public string Code { get; set; }
        // string AuthorOfLastModification { get; private set; }
        public DateTime LastModificationDate { get; set; }
        public Guid GuidOfMacro { get; set; }
    }

    // TODO
    // Реализовать класс, который будет хранить информацию о последней синхронизации
    private class SyncronizationLog {
        // Данный класс должен хранить данные о всех макросах, которые были перенесены путем синхронизации.
        // Так же он должен хранить данные о дате последней модификации макроса из справочника и дате, когда этот макрос был
        // в последний раз изменен на локальном компьютере
        // ДАнный класс должена хратить историю синхронизаций (не всех запусков, а именно историю переносов информации
        
        public string ServerName { get; set; }
    }

    #endregion Service classes
}
