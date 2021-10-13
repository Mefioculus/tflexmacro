using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using TFlex.DOCs.Model.Macros; // Для работы с макроязыком
using TFlex.DOCs.Model.Macros.ObjectModel; // Для работы с макроязыком
using TFlex.DOCs.Model.References; // Для работы со справочниками
using TFlex.DOCs.Model.References.Nomenclature; // Для получения объекта справочника документы
using TFlex.DOCs.Model.Structure; // Для использования класса ParameterInfo

public class Macro : MacroProvider {

    #region Fields and Properties of class Macro

    private static Reference documentReference;
    private static Reference nomenclatureReference;

    #endregion Fields and Properties of class Macro

    #region Constructor

    public Macro(MacroContext context)
        : base(context)
    {
        // Получаем объекты справочников документов и электронной структуры изделий
        documentReference = Context.Connection.ReferenceCatalog.Find(Guids.References.Документы).CreateReference();
        nomenclatureReference = Context.Connection.ReferenceCatalog.Find(Guids.References.Номенклатура).CreateReference();

        // Проверка на то, что справочники были успешно загружены
        if (documentReference == null || nomenclatureReference == null)
            throw new Exception("При выполнении макроса возникла ошибка. Не удалось подключить доступ к справочникам 'Документы' и 'Электронная структура изделия'");
    }

    #endregion Constructor

    #region Guids class

    private static class Guids {
        public static class References {
            public static Guid Документы = new Guid("ac46ca13-6649-4bbb-87d5-e7d570783f26");
            public static Guid Номенклатура = new Guid("853d0f07-9632-42dd-bc7a-d91eae4b8e83");
        }

        public static class Parameters {
            public static class Номенклатура {
                public static Guid Обозначение = new Guid("ae35e329-15b4-4281-ad2b-0e8659ad2bfb");
                public static Guid Наименование = new Guid("45e0d244-55f3-4091-869c-fcf0bb643765");
            }
        }

        public static class Links {
        }
    }

    #endregion Guids class


    public override void Run() {
        // Попробуем найти объект по его обозначению
        string shifr = "0000-00-1";
        string name = string.Empty;

        // Производим поиск данного объекта
        var searchedObject = new ObjectInTflex(shifr, name);

        Message("Найденный номенклатурный объект", string.Format(
                    "{0}\n{1}",
                    searchedObject.NomenclatureObject.ToString(),
                    searchedObject.NomenclatureObject.Guid.ToString()
                    ));

        Message("Найденный объект справочника документы", string.Format(
                    "{0}\n{1}",
                    searchedObject.DocumentObject.ToString(),
                    searchedObject.DocumentObject.Guid.ToString()
                    ));

        //System.Diagnostics.Debugger.Launch();
        //System.Diagnostics.Debugger.Break();

        
    }

    #region GetInfo(ReferenceObject)
    private string GetInfo(ReferenceObject refObj) {
        string result = string.Empty;
        result += string.Format("Информация собирается для объекта '{0}'\n", refObj.ToString());

        LoadSettings settings = new LoadSettings(refObj.Reference.ParameterGroup);
        refObj.Links.Fill(settings);
        // Получение информации по связям объекта
        // Получаем подключения ко многим
        
        // Получение связанного объекта
        
        // Получаем связи объекта
        // Для начала получаем менеджер по работе с связями объекта к одному
        var managerToOneLinks = refObj.Links.ToOne;
        var managerToManyLinks = refObj.Links.ToMany;
        var managerToOneComplexHLinks = refObj.Links.ToOneToComplexHierarchy;
        var managerToManyComplexHLinks = refObj.Links.ToManyToComplexHierarchy;

        result += "Список подключений к одному:\n";
        foreach (var link in managerToOneLinks) {
            result += string.Format("- {0}\n", link.ToString());
        }
        result += "\n";

        result += "Список подключений ко многим:\n";
        foreach (var link in managerToManyLinks) {
            result += string.Format("- {0}\n", link.ToString());
        }
        result += "\n";

        result += "Список подключений к одному (сложная иерархия):\n";
        foreach (var link in managerToOneComplexHLinks) {
            result += string.Format("- {0}\n", link.ToString());
        }
        result += "\n";

        result += "Список подключений ко многим (сложная иерархия):\n";
        foreach (var link in managerToManyComplexHLinks) {
            result += string.Format("- {0}\n", link.ToString());
        }
        result += "\n";

        // Пробуем получить связь на справочник документы
        return result;
    }
    #endregion GetInfo(ReferenceObject)

    #region ЗагрузитьИзмененияСтруктурыИзделияИзFox()
    // Метод для применения изменений
    public bool ЗагрузитьИзмененияСтруктурыИзделияИзFox(
            Dictionary<int, Dictionary<string, string>> oldData,
            Dictionary<int, Dictionary<string, string>> newData) {

        // oldData содержит данные из справочника с предыдущей выгрузкой, отфильтрованные
        // только по тем записям, по которым произошли изменения
        // newData содержит данные из справочника с новой выгрузкой, отфильтрованные только по тем записям,
        // которые измелись относительно предыдущей выгрузки
        
        // Производим анализ входных данных
        // Получаем список всех ключей, по которым нужно произвести итерацию
        List<int> rowNumbers = GetAllRowNumbers(oldData, newData);

        // Рассортируем информацию по назначению
        List<int> IDsOfChangedData = new List<int>();
        List<int> IDsOfDeletedData = new List<int>();
        List<int> IDsOfAddedData = new List<int>();

        foreach (int rowNumber in rowNumbers) {
            // Если запись содержится в двух словарях, значит в ней есть изменения, которые необходимо произвести
            if ((oldData.ContainsKey(rowNumber)) && (newData.ContainsKey(rowNumber)))
                IDsOfChangedData.Add(rowNumber);
            // Если запись отсутствует в словаре с новой информацией, значит она была удалена в Fox
            else if (!newData.ContainsKey(rowNumber))
                IDsOfDeletedData.Add(rowNumber);
            // Другой вариант - запись отсутствует в старом словаре, следовательно это новая запись и ее необходимо добавить
            else 
                IDsOfAddedData.Add(rowNumber);
        }

        // Производим изменение текущих записей
        ApplyChangesToTflex(IDsOfChangedData, oldData, newData);

        // Производим добавление новых записей
        ApplyAddingsToTflex(IDsOfAddedData, newData);

        // Производим удаление записей
        ApplyDeletingToTflex(IDsOfDeletedData, oldData);



        // Произвести проверку того, что oldData соответствует записи в T-Flex

        // Далее, необходимо произвести необходимые изменения


        // Возвращаем истину, если в процесса загрузки информации не возникло ошибок
        return true;
    }
    #endregion ЗагрузитьИзмененияСтруктурыИзделияИзFox()

    
    #region Sevrice methods
    // Дополнительные методы, которые требуются для работы метода ЗагрузитьИзмененияСтруктурыИзделияИзFox()

    private List<int> GetAllRowNumbers(
                Dictionary<int, Dictionary<string, string>> oldData,
                Dictionary<int, Dictionary<string, string>> newData) {
        // Реализовать код, который получает номера всех записей (из первого и второго словаря
        return new List<int>();
    }

    private string ПреобразоватьОбозначениеFoxTflex(string value) {
        // TODO Найти макрос, который производит подстановку данных и перенести функциональность сюда
        return string.Empty;
    }

    private string ПреобразоватьОбозначениеTFlexFox(string value) {
        return value.Replace(".", string.Empty);
    }

    private void ApplyChangesToTflex(List<int> IDs, Dictionary<int, Dictionary<string, string>> oldData, Dictionary<int, Dictionary<string, string>> newData) {
        // Метод для применения изменений
        // TODO Реализовать код для применения изменений к структуре
        foreach (int ID in IDs) {
            if (oldData[ID]["SHIFR"] != newData[ID]["SHIFR"]) {
                // В этом случае мы в старом изделии убираем это подключение, в новом изделии его создаем

                // Проверяем, соответствует ли наименование новому найденному объекту. Если не соответствует, производим правку

                // Приводим параметры подключения (POS, PRIM, IZD) в соответствие с тем, что хранится в newData
            }
            else {
                // Данный случай подразумевает, что мы остаемся в том же подключении, но правим какие-то другие изменения
                if (oldData[ID]["NAIM"] != newData[ID]["NAIM"]) {
                    // Корректируем название документа и объекта электронной структуры изделия
                }

                // Приводим остальные параметры подключения (POS, PRIM, IZD) в соответствие с тем, что хранится в newData
            }
        }
        
        
    }

    private void ApplyAddingsToTflex(List<int> IDs, Dictionary<int, Dictionary<string, string>> newData) {
        // Метод для применения создания новых записей
        // TODO Реализовать код для добавления новой записи в структруру
        foreach (int ID in IDs) {
            // Производим проверку наличия данной записи в базе T-Flex

            // Если данная запись есть в базе, получаем к ней доступ и создаем новое подключение
            
            // Если данной записи в базе нет, создаем новые объекты в справочнике документы и ЭСИ
            
            // Производим подключение объекта в структуру
        }
    }

    private void ApplyDeletingToTflex(List<int> IDs, Dictionary<int, Dictionary<string, string>> oldData) {
        // Метод для применения удаления записей
        foreach (int ID in IDs) {
            // TODO Реализловать код для удаления записи из структуры
            // Производим поиск данной записи в базе.
            // Если запись была найдена, находим необходимое подключение и удаляем его.

            // Если не получилось найти запись или не получилось найти необходимое подключение, пишем сообщение в лог

        }
    }
    #endregion Sevrice methods

    #region serviceClasses

    #region ObjectInTFlex
    private class ObjectInTflex {
        public ReferenceObject DocumentObject;
        public ReferenceObject NomenclatureObject;
        // Свойство, которое показывает, успешно ли был произведен поиск объектов
        public bool IsDataComplete => ((this.DocumentObject != null) & (this.NomenclatureObject != null));

        public ObjectInTflex(string shifr, string name) {
            // Производим поиск по справочнику документов
            // TODO Добавить анализ найденных данных.
            // Узнать, что получилось добавить, а чего не получилось
            // TODO Реазизовать интерфейс для создания нового подключения через этот служебный класс
            this.NomenclatureObject = FindInNomenclatureReference(shifr, name);
            this.DocumentObject = FindInDocumentReference(shifr, name);
        }
        
        #region FindInNomenclatureReference()
        
        // Метод для поиска объекта в справочнике ЭСИ
        private ReferenceObject FindInNomenclatureReference(string shifr, string name) {
            // Производим поиск объекта в справочнике "ЭСИ" по обозначению
            ParameterInfo shifrParam = nomenclatureReference.ParameterGroup[Guids.Parameters.Номенклатура.Обозначение];
            List<ReferenceObject> searchedObject = nomenclatureReference.Find(shifrParam, shifr);

            if (searchedObject.Count == 1) {
                return searchedObject[0];
            }
            else if (searchedObject.Count == 0) {
                // Случай, когда объект найти не получилось
                
                // Необходимо сообщить об этом пользователю

                // Как варинат, необходимо создать объект с таким обозначением и наименованем
            }
            else {
                // Случай, когда было найдено несколько объектов
                // Об этом нужно сообщить пользователю
            }

            return null;
        }

        #endregion FindInNomenclatureReference()

        #region FindInDocumentReference();

        // Метод для поиска объекта в справочнике документов
        private ReferenceObject FindInDocumentReference(string designation, string name) {
            // Для начала пробуем получить данный объект по связи с номенклатурного объекта
            if (this.NomenclatureObject != null) {
                NomenclatureObject nomObj = this.NomenclatureObject as NomenclatureObject;

                if (nomObj.HasLinkedObject)
                    return nomObj.LinkedObject;
                else {
                    // Пытаемся найти нужный объект и подключить его

                    // Или же создаем объект и подключаем его
                }

            }
            return null;
        }

        #endregion FindInDocumentReference();
    }
    #endregion ObjectInTFlex

    #region Enums
    #region TypeMessage

    private enum TypeMessage {
        Info,
        Error,
        Stat
    }

    #endregion TypeMessage

    #region TypeProcess

    private enum TypeProcess {
        Add,
        Delete,
        Modify
    }

    #endregion TypeProcess
    #endregion Enums
    #region class Logger
    private static class Logger {
        
        #region Fields and Properties

        public static string ResultMessage { get; private set; }
        public static string PathToDirectory { get; set; } // Путь к директории, где будет храниться файл с логами
        public static string NameOfFile { get; private set; } // Название файла логов

        // Сообщение, которое будет содержать информацию о статистике
        public static List<string> AddingProcessStats { get; set; } = new List<string>();
        public static List<string> DeletingProcessStats { get; set; } = new List<string>();
        public static List<string> ModifyingProcessStats { get; set; } = new List<string>();

        // Сообщения с ошибкой
        public static List<string> AddingProcessErrors { get; private set; } = new List<string>();
        public static List<string> ModifyingProcessErrors { get; private set; } = new List<string>();
        public static List<string> DeletingProcessErrors { get; private set; } = new List<string>();

        // Информационные сообщения
        public static List<string> AddingProcessInfo { get; private set; } = new List<string>();
        public static List<string> ModifyingProcessInfo { get; private set; } = new List<string>();
        public static List<string> DeletingProcessInfo { get; private set; } = new List<string>();

        // Парамеры Логгера
        public static bool PrintInfo { get; set; } = false;
        public static bool PrintStat { get; set; } = false;

        #endregion Fields and Properties

        #region AddMessage()

        public static void AddMessage(string textOfMessage, TypeProcess typeOfProcess, TypeMessage typeOfMessage = TypeMessage.Info) {
            
            // Если выставлен флаг не писать стат сообщение, ничего не добавляем
            if ((typeOfMessage == TypeMessage.Stat) & (!PrintStat))
                return;
            // Если выставлен флаг не писать инфо сообщение, ничего не добавляем
            if ((typeOfMessage == TypeMessage.Info) & (!PrintInfo))
                return;

            // Метод для добавления нового сообщения
            if (typeOfMessage == TypeMessage.Info) {

                if (typeOfProcess == TypeProcess.Add)
                    AddingProcessInfo.Add(textOfMessage); // Информационное сообщение о добавлении
                else if (typeOfProcess == TypeProcess.Delete)
                    DeletingProcessInfo.Add(textOfMessage); // Информационное сообщение о удалении
                else if (typeOfProcess == TypeProcess.Modify)
                    ModifyingProcessInfo.Add(textOfMessage); // Информационное сообщение о изменении
            }
            else if (typeOfMessage == TypeMessage.Error) {
                if (typeOfProcess == TypeProcess.Add)
                    AddingProcessErrors.Add(textOfMessage); // Ошибка при добавлении
                else if (typeOfProcess == TypeProcess.Delete)
                    DeletingProcessErrors.Add(textOfMessage); // Ошибка при удалении
                else if (typeOfProcess == TypeProcess.Modify)
                    ModifyingProcessErrors.Add(textOfMessage); // Ошибка при изменении
            }
            else if (typeOfMessage == TypeMessage.Stat) {
                if (typeOfProcess == TypeProcess.Add)
                    AddingProcessStats.Add(textOfMessage); // Сообщения с статистикой, которая будет писаться в заглавии сообщения
                if (typeOfProcess == TypeProcess.Delete)
                    DeletingProcessStats.Add(textOfMessage); // Сообщения с статистикой, которая будет писаться в заглавии сообщения
                if (typeOfProcess == TypeProcess.Modify)
                    ModifyingProcessStats.Add(textOfMessage); // Сообщения с статистикой, которая будет писаться в заглавии сообщения
            }
            else {
                // Обработка случаев, когда пользователь ввел неожиданное сочетание параметров
                throw new Exception(string.Format("Выбрано сочетание полей ({0}, {1}), для которого не сделана обработка", typeOfProcess.ToString(), typeOfMessage.ToString()));
            }
        }

        #endregion AddMessage()

        #region GenerateMessage()
        
        public static void GenerateMessage() {
            // Формирование сообщений для процесса добавления нового ДСЕ в состав изделия
            //TODO Реализовать генерацию сообщения 
            if (AddingProcessStats.Count != 0) {
            }

            if (AddingProcessInfo.Count != 0) {
            }

            if (AddingProcessErrors.Count != 0) {
            }
            
            // Формирование сообщений для процесса изменения ДСЕ в составе изделия
            if (ModifyingProcessStats.Count != 0) {
            }

            if (ModifyingProcessInfo.Count != 0) {
            }

            if (ModifyingProcessErrors.Count != 0) {
            }

            // Формирование сообщений для процесса удаления ДСЕ из состава изделия
            if (DeletingProcessStats.Count != 0) {
            }

            if (DeletingProcessInfo.Count != 0) {
            }

            if (DeletingProcessErrors.Count != 0) {
            }

            // Запускаем запись в файл
            WriteLog();
        }

        #endregion GenerateMessage()

        #region WriteLog()

        private static void WriteLog() {
            NameOfFile = string.Format("[{0}] Импорт структуры изделий из FoxPro в T-Flex", DateTime.Now.ToString("dd.MM.yyyy HH:mm"));
            string pathToFile = Path.Combine(PathToDirectory, NameOfFile);

            File.WriteAllText(pathToFile, ResultMessage);
        }

        #endregion WriteLog()
        
    }
    #endregion class Logger


    #endregion serviceClasses
}


