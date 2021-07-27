#region Рекурсивное получение списка изделий для случайной ДСЕ
// Метод, который будет рекурсивно просматривать родителей дсе для получения списка всех изделий
private Объекты ПолучитьРодительскиеИзделияДляДСЕ (Объект дсе) {
    Объекты списокИзделий = new Объекты();

    Объекты родители = дсе.РодительскиеОбъекты;

    foreach (Объект родитель in родители) {
        // Проверяем, не является ли данный родитель изделием
        if ((родитель.РодительскиеОбъекты.Count == 0) || (СодержитПапку(родитель))) {
            списокИзделий.Add(родитель);
            // Данная ветвь будет срабатывать в том случае, если данное две является изделием (концом ветки)
        }
        else {
            // Данная ветвь будет срабатываеть, если данное изделие является промежуточным звеном
            foreach (Объект изделие in ПолучитьРодительскиеИзделияДляДСЕ(родитель)) {
                if (!списокИзделий.Contains(изделие)) {
                    списокИзделий.Add(изделие);
                }
            }

        }

    }

    return списокИзделий;
}

private bool СодержитПапку(Объект дсе) {
    foreach (Объект родитель in дсе.РодительскиеОбъекты) {
        if (родитель.Тип == "Папка") {
            return true;
        }
    }
    return false;
}
#endregion Рекурсивное получение списка изделий для случайной ДСЕ

#region Получение списка цехопереходов, которые наследуют свои доступы


// Для работы данного кода потребуются так же следующие пространства имен:
//using TFlex.DOCs.Model.References;
//using TFlex.DOCs.Model.Access;
//using TFlex.DOCs.Model.Classes;
// А так же возможно ссылки на следующие библиотеки
//Cs.UI.Objects.dll
//TFlex.DOCs.UI.Common.dll
//TFlex.DOCs.UI.Types.dll
//TFlex.DOCs.Common.dll


Reference techReference = Context.Connection.ReferenceCatalog.Find(new Guid("353a49ac-569e-477c-8d65-e2c49e25dfeb")).CreateReference();

// Получаем список технологический процессов
techReference.Objects.Load();

string message = "Список цехопереходов с явными правами:";

foreach (ReferenceObject process in techReference.Objects) {
    // Получаем список цехопереходов
    foreach (ReferenceObject perehod in process.Children) {
        AccessManager accessManager = AccessManager.GetReferenceObjectAccess(perehod);
        if (!accessManager.IsInherit) {
            message += string.Format("\nТП (id {0}): {1}; ЦЗ (id {2}): {3}", process.SystemFields.Id.ToString(), process.ToString(), perehod.SystemFields.Id.ToString(), perehod.ToString());
        }
    }
}

Message("Информация", message);

#endregion Получение списка цехопереходов, которые наследуют свои доступы

#region Получение списка переменных и списка свойств из документа CAD

// Для того, чтобы данный макрос работал, потребуется подключить следующие пространства имен
//using TFlex.DOCs.Model.FilePreview.CADService; // Данное пространство имен требуется для классов, связанных с документом CAD

// Макрос для получения всей информации из CAD документа, которую только можно получить при помощи
// встроенного в DOCs функционала

string pathToCadFile = "C:\\\\testDocument.grb";

CadDocumentProvider provider = CadDocumentProvider.Connect(Context.Connection, ".grb");

// Открываем документ в режиме чтения
using (CadDocument document = provider.OpenDocument(pathToCadFile, true)) {
    // Проверка, был ли открыт документ
    if (document == null) {
        string fileName = Path.GetFileName(pathToCadFile);
        Message("Ошибка экспорта", "Файл '{0}' не может быть открыт", fileName);
        return;
    }

    // Получаем список переменных документа
    VariableCollection variables = document.GetVariables();

    string message = "Список переменных CAD документа:";
    // Отображаем переменные
    foreach (Variable cadVar in variables) {
        message += string.Format("\nName: '{0}'; Value: '{1}'", cadVar.Name, cadVar.Value.ToString());
    }
    Message("Информация", message);

    // Получаем список фрагментов документа
    FragmentCollection fragments = document.GetFragments3D();

    foreach (var fragment in fragments) {
        message = string.Format("Список свойств CAD фрагмента '{0}':", fragment.ToString());
        var properties = fragment.GetProperties();
        message = "Список свойств CAD документа:";
        foreach (var property in properties) {
            message += string.Format("\n{0}", property.ToString());
        }
        Message("Информация", message);
    }
}
#endregion Получение списка переменных и списка свойств из документа CAD

#region Разбор содержимого подключения

public void ПолучитьВсеПодключенияОбъекта() {
    // Метод для получения всех подключений объекта
    Объект сборка = ТекущийОбъект;
    Подключения подкл1 = сборка.ДочерниеПодключения;

    string message = string.Join("\n", подкл1.Select(подключение => подключение.ДочернийОбъкт.ToString()));

    Message("", message);

    foreach (Подключение подключение in подкл1) {
        Объекты применяемости = подключение.СвязанныеОбъекты["Статусы и применяемость"];
        message = string.Join("\n", применяемости.Select(prim => prim.ToString()));
        Message("", message);
    }
    
}

#endregion Разбор содержимого подключения

#region Получение данных из таблицы FoxPro 

// Данный код требует для работы подключения пространства имен System.Reflections;
public void ПолучитьДанныеИзDBF() {
    string pathToDbfFile = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "spec.dbf");
    string pathToDbfDataReader = @"D:\Библиотеки dotnet\dbfdatareader.0.7.0\lib\net48\DbfDataReader.dll";

    Assembly dbfDataReader = Assembly.LoadFrom(pathToDbfDataReader);

    Type dataReaderType = dbfDataReader.GetType("DbfDataReader.DbfDataReader");
    if (dataReaderType == null) {
        Message("Ошибка", "Не получилось извлечь тип");
        return;
    }

    // Получаем необходимые методы данного класса
    MethodInfo read = dataReaderType.GetMethod("Read");
    MethodInfo getString = dataReaderType.GetMethod("GetString");

    object obj = Activator.CreateInstance(dataReaderType, new object[] {pathToDbfFile});
    

    List<string> result = new List<string>();
    while ((bool)read.Invoke(obj, new object[] {})) {
        result.Add((string)getString.Invoke(obj, new object[] {1}));
    }

    string message = string.Join("\n", result);

    Message("", message);
    Message("", "Работа макроса завершена");

    
    /*
    string pathToFile = "path/to/file.dbf";
    using (DbfDataReader dataReader = new DbfDataReader(pathToFile)) {
        while (dataReader.Read()) {
            var value = dataReader.GetString(0);
        }
    }
    */
}




#endregion Получение данных из таблицы FoxPro

