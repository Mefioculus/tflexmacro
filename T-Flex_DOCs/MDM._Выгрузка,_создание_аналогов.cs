/*
TFlex.DOCs.UI.Client.dll
*/

using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using TFlex.DOCs.Client.ViewModels.Base;
using TFlex.DOCs.Client.ViewModels.Layout;
using TFlex.DOCs.Model;
using TFlex.DOCs.Model.Access;
using TFlex.DOCs.Model.Classes;
using TFlex.DOCs.Model.Macros;
using TFlex.DOCs.Model.Macros.ObjectModel.Layout;
using TFlex.DOCs.Model.Notification;
using TFlex.DOCs.Model.References;
using TFlex.DOCs.Model.References.Events;
using TFlex.DOCs.Model.References.Users;
using TFlex.DOCs.Model.Search;
using TFlex.DOCs.Model.Stages;
using TFlex.DOCs.Model.Structure;
using TFlex.DOCs.Model.Structure.Builders;

public class AsuNsiLoad : MacroProvider
{
    public AsuNsiLoad(MacroContext context)
        : base(context)
    {
    }

    private int _createdReferenceId = 0;

    private StringBuilder _errorStringBuilder = new StringBuilder();
    /// <summary>
    /// Гуид объекта ГИП Администратор НСИ
    /// </summary>
    private Guid _MDMAdminObjectGuid = new Guid("671ef0fd-ce8a-4d21-9763-f5d5bb35f3b4");
    private Guid _autorAccessGuid = new Guid("00ee8f37-9a89-4efc-9997-adf0dc19927e");

    /// <summary>
    /// Описание списка значений базового параметра "Состояние"
    /// </summary>
    private static Dictionary<int, string> _parametersDictionary = new Dictionary<int, string>()
    {
        {0, "Добавлено"},
        {1, "Очистка данных"},
        {2, "Установлен эталон"},
        {3, "Обработано"},
    };

    public override void Run()
    {
    }

    /// <summary>
    /// Подключение настроек очистки в аналоговом справочнике.
    /// Запускается мультиселектом из справочника аналога.
    /// </summary>
    public void ConnectionSettingsClearInAnalogReference()
    {
        var contextSettings = CreateContextSettings();

        var analogReferenceObjects = GetFilteredSelectedObjects(contextSettings);

        foreach (var analogReferenceObject in analogReferenceObjects)
        {
            ReferenceObject settingsReferenceObject;

            var settingsReferenceObjects = GetFilteredSettingClearReferenceObjects(analogReferenceObject, contextSettings);
            //settingsClearDataReference.Find(filter);
            if (settingsReferenceObjects.Count > 0)
            {
                if (settingsReferenceObjects.Count > 1)
                {
                    //Записываем в лог если было найдено больше одной записи.
                    _errorStringBuilder.Append(string.Format("Объект аналога с ID:{0} , найдено несколько настроек очистки{1}",
                            analogReferenceObject.SystemFields.Id, Environment.NewLine));
                }

                //Если нашлась одна запись, то подключаем запись настроек к аналогу
                settingsReferenceObject = settingsReferenceObjects.First();
            }
            else
            {
                settingsReferenceObject = CreateSettingClearDataReferenceObject(analogReferenceObject, contextSettings);
            }

            SetAnalogParameter(analogReferenceObject, contextSettings, settingsReferenceObject.Guid);
        }

        if (_errorStringBuilder.Length > 0)
        {
            Сообщение("Предупреждение",
                string.Format("Во время обработки возникли ошибки: {0}{1}",
                   Environment.NewLine, _errorStringBuilder));
        }
    }

    private void SetAnalogParameter(ReferenceObject analogReferenceObject, ContextSettings contextSettings,
        Guid settingReferenceObjectGuid)
    {
        if (!TryBeginChanges(analogReferenceObject))
        {
            _errorStringBuilder.Append(
                string.Format("Объект аналога с ID:{0} , заблокирован другим пользователем{1}",
                    analogReferenceObject.SystemFields.Id, Environment.NewLine));
            return;
        }

        //По дальнейшей договоренности тут будет подключение объект по связи.
        analogReferenceObject[contextSettings.AnalogSettingClearData].Value = settingReferenceObjectGuid;
        analogReferenceObject.EndChanges();
    }

    /// <summary>
    /// Метод запускает процесс загрузки данных из Excel файла.
    /// Используется на рабочей странице "АРМ нормализатора"
    /// </summary>
    public void LoadAnalogReferenceFromExcel()
    {
        ReferenceObject referenceObject = GetFocusedObjectFromWorkingPage(Settings.WorkingPageItemName);
        if (referenceObject == null)
            Ошибка(String.Format("Не найден выбранный объект на элементе управления: {0}", Settings.WorkingPageItemName));

        Guid referenceGuid = referenceObject[Settings.DataSourceParameterReferenceAnalog].GetGuid();
        if (!DeleteOldAnalogReference(referenceGuid))
            return;

        //Добавляем обработчик на событие создания справочника
        Context.Connection.EventWatcher.WatchReferenceCatalogChanged(null, OnReferenceCatalogChanges);

        StartImportProcessReference();

        if (_createdReferenceId == 0)
            return;

        ReferenceBuilder referenceBuilder = GetReferenceBulder(_createdReferenceId);
        var parameterGroup = referenceBuilder.ParameterGroup;
        if (referenceBuilder == null)
            return;

        if (referenceBuilder.IsActive)
            referenceBuilder.Deactivate();

        SetAccessReference(referenceBuilder);

        MoveToAnalogReferenceFolder(referenceBuilder);
        ChangeImportedParemeters(parameterGroup);

        CreateBaseParameters(parameterGroup);

        var baseClassObject = referenceBuilder.ParameterGroup.Classes.AllClasses.First();

        SetClassObjectName(baseClassObject, Settings.ClassObjectName);

        //CreateUserEventButton(baseClassObject, Settings.NSIEventName, Settings.MacroNsiGuid, Settings.MacroNsiMetodConnectionSettingsName);
        //CreateUserEventButton(baseClassObject, Settings.MDMEventName, Settings.MacroMDMSeachGuid);
        //CreateUserEventButton(baseClassObject, Settings.ConnectTypeNSIEventName, Settings.MacroMDMSetTypeNSIGuid);

        SetMDMStageScheme(referenceBuilder);

        referenceBuilder.Save();
        referenceBuilder.Activate();

        SetDefaulMDMStageAllObjects(referenceBuilder);

        SetAnalogReferenceObjectPаrameter(referenceObject, referenceBuilder.ParameterGroup.Guid);

        if (_errorStringBuilder.Length != 0)
        {
            Сообщение("Предупреждение", String.Format("Внимание, во время обработки возникли ошибки:{0}{1}",
                Environment.NewLine, _errorStringBuilder));
        }
        else
        {
            Сообщение("Выполнено", String.Format("Справочник: {0} был загружен", referenceBuilder.Name));
        }
    }


    private void SetAccessReference(ReferenceBuilder referenceBuilder)
    {
        var userReference = new UserReference(Context.Connection);
        var MDMAdmin = userReference.Find(_MDMAdminObjectGuid) as UserReferenceObject;
        if (MDMAdmin == null)
        {
            _errorStringBuilder.Append(String.Format("Не найдена роль 'Администратор НСИ'{0}", Environment.NewLine));
            return;
        }
        AccessGroup autorAccess = AccessGroup.Find(Context.Connection, _autorAccessGuid);
        if (autorAccess == null)
        {
            _errorStringBuilder.Append(String.Format("Не найден авторский доступ для справочника{0}", Environment.NewLine));
            return;
        }

        AccessManager accesseManager = AccessManager.GetReferenceAccess(referenceBuilder.ReferenceInfo);
        accesseManager.SetAccess(MDMAdmin, autorAccess);
        accesseManager.Save();
    }

    private void SetDefaulMDMStageAllObjects(ReferenceBuilder referenceBuilder)
    {
        var reference = referenceBuilder.ReferenceInfo.CreateReference();
        var defaultStage = referenceBuilder.DefaultStage.Stage;
        var referenceObjects = reference.Objects;
        defaultStage.Set(referenceObjects);
    }

    /// <summary>
    /// Задает название типа справочника
    /// </summary>
    /// <param name="classObject">Тип справочника</param>
    /// <param name="name">Наименование типа</param>
    private void SetClassObjectName(ClassObject classObject, string name)
    {
        ClassObjectBuilder classObjectBuilder = new ClassObjectBuilder(classObject);
        classObjectBuilder.Name = name;
        classObjectBuilder.Save();
    }

    private void SetMDMStageScheme(ReferenceBuilder referenceBuilder)
    {
        var mdmScheme = Scheme.GetSchemes().FirstOrDefault(x => x.Guid == Settings.MDMAnalogStageScheme);
        if (mdmScheme == null)
        {
            _errorStringBuilder.Append(String.Format("Не найдена схема для подключения к справочнику"));
            return;
        }

        var addedStage = mdmScheme.Stages.FirstOrDefault(s => s.Guid == Settings.AddedStageGuid);

        SetStageScheme(referenceBuilder, mdmScheme, addedStage);
    }

    private void SetStageScheme(ReferenceBuilder referenceBuilder, Scheme scheme, SchemeStage schemeStage = null)
    {
        referenceBuilder.SupportsStages = true;
        referenceBuilder.Scheme = scheme;

        if (schemeStage != null)
            referenceBuilder.DefaultStage = schemeStage;
    }

    private void MoveToAnalogReferenceFolder(ReferenceBuilder referenceBuilder)
    {
        var foundAnalogCatalogFolder = Context.Connection.ReferenceCatalog.FindFolderByName(Settings.AnalogCatalogFolderGuid);
        if (foundAnalogCatalogFolder == null)
            return;

        referenceBuilder.MoveToAnotherFolder(foundAnalogCatalogFolder);
    }

    private static bool TryBeginChangesReferenceObject(ReferenceObject referenceObject)
    {
        if (referenceObject.CanEdit)
        {
            referenceObject.BeginChanges();
            return true;
        }

        return referenceObject.Changing;
    }

    private static bool TryBeginChanges(ReferenceObject referenceObject)
    {
        if (referenceObject.IsCheckedOut)
        {
            if (referenceObject.IsCheckedOutByCurrentUser)
            {
                return TryBeginChangesReferenceObject(referenceObject);
            }
        }
        else if (referenceObject.CanCheckOut)
        {
            return TryBeginChangesReferenceObject(referenceObject);
        }

        return TryBeginChangesReferenceObject(referenceObject);
    }

    private ContextSettings CreateContextSettings()
    {
        var parameters = Context.Reference.ParameterGroup.OneToOneParameters;

        var settingsClearDataReference = FindReference(Settings.SettingsClearData);
        if (settingsClearDataReference == null)
            throw new NullReferenceException("Не найден справочник 'Настройки очистки данных'");

        var documentNSI = FindReference(Settings.DocumentNSI);
        if (documentNSI == null)
            throw new NullReferenceException("Не найден справочник 'Документы НСИ'");

        var settingsCleadDataClassObject = settingsClearDataReference.Classes.AllClasses.First();

        var contextSettings = new ContextSettings
        {
            AnalogTypeNSI = GetParameterGuidByName(parameters, "Тип НСИ"),
            AnalogDocNSI = GetParameterGuidByName(parameters, "Документ НСИ"),
            AnalogSettingClearData = GetParameterGuidByName(parameters, "Настройка очистки данных"),
            AnalogParameters = GetNotSystemsParameterInfoList(Context.Reference.ParameterGroup),
            AnalogReferenceGuid = Context.Reference.ParameterGroup.Guid,
            SettingsClearDataReference = settingsClearDataReference,
            SettingsClearDataReferenceClassObject = settingsCleadDataClassObject,
            DocumentNSI = documentNSI

        };

        return contextSettings;
    }

    private static List<ParameterInfo> GetNotSystemsParameterInfoList(ParameterGroup parameterGroup)
    {
        return parameterGroup.Parameters.Where(par => !par.IsSystem).ToList();
    }

    /// <summary>
    /// Возвращает отфильтрованные объекты справочника "Настройки очистки данных"
    /// </summary>
    /// <param name="analogReferenceObject"></param>
    /// <returns></returns>
    private List<ReferenceObject> GetFilteredSettingClearReferenceObjects(ReferenceObject analogReferenceObject, ContextSettings contextSettings)
    {
        Guid docNSIGuid = analogReferenceObject[contextSettings.AnalogDocNSI].GetGuid();

        string filterString = string.Format("[{0}] = '{1}' И [{2}] = '{3}'",
            Settings.SettingsClearDataDocNSI,
            docNSIGuid,
            Settings.SettingsClearDataReferenceAnalog,
            contextSettings.AnalogReferenceGuid);


        if (!Filter.TryParse(filterString, contextSettings.SettingsClearDataReference.ParameterGroup, out Filter filter))
            return new List<ReferenceObject>();

        var referenceObjects = contextSettings.SettingsClearDataReference.Find(filter);

        return referenceObjects;
    }

    /// <summary>
    /// Большой метод будет создавать настройки очистки данных, с со списком значения соответствия параметров
    /// </summary>
    /// <param name="analogReferenceObject"></param>
    /// <returns></returns>
    private ReferenceObject CreateSettingClearDataReferenceObject(ReferenceObject analogReferenceObject, ContextSettings contextSettings)
    {
        var clearDataReferenceObject =
            contextSettings.SettingsClearDataReference.CreateReferenceObject(contextSettings.SettingsClearDataReferenceClassObject);

        clearDataReferenceObject[Settings.SettingsClearDataReferenceAnalog].Value = contextSettings.AnalogReferenceGuid;
        clearDataReferenceObject[Settings.SettingsClearDataDocNSI].Value = analogReferenceObject[contextSettings.AnalogDocNSI].GetGuid();
        clearDataReferenceObject[Settings.SettingsClearDataName].Value = "Настройка";

        //Заполнение связи
        //clearDataReferenceObject.SetLinkedObject(Settings.SettingsClearDataLinkDocNSI, analogReferenceObject.GetObject(contextSettings.));

        var objectsListMatchParameters =
            CreateObjectsListMatchParametersForSettingClearData(clearDataReferenceObject, analogReferenceObject, contextSettings);

        clearDataReferenceObject[Settings.SettingsClearDataEqualParameters].Value = objectsListMatchParameters.Item2;

        clearDataReferenceObject.EndChanges();

        return clearDataReferenceObject;
    }

    /// <summary>
    /// Метод создает список объектов при создании настройки очистки данных
    /// </summary>
    /// <param name="clearDataReferenceObject"></param>
    /// <param name="analogReferenceObject"></param>
    /// <param name="contextSettings"></param>
    /// <returns>Возвращает коллекцию с описанием, и значением указывающим, что созданный параметр одинаковый </returns>
    private Tuple<List<ReferenceObject>, bool> CreateObjectsListMatchParametersForSettingClearData(ReferenceObject clearDataReferenceObject,
        ReferenceObject analogReferenceObject, ContextSettings contextSettings)
    {
        var tempReferenceNotSystemParameters = FindTempReferenceNotSystemParameters(analogReferenceObject, contextSettings);
        if (!tempReferenceNotSystemParameters.Any())
            return null;

        bool allParametersIsEqual = true;

        List<ReferenceObject> createdMatchReferenceObject = new List<ReferenceObject>();

        foreach (var tempReferenceNotSystemParameter in tempReferenceNotSystemParameters)
        {

            ReferenceObject createMatchReferenceObject = CreateAndCheckMatchReferenceObject(clearDataReferenceObject, contextSettings,
                tempReferenceNotSystemParameter, out bool parametersIsEqual);

            if (!parametersIsEqual)
                allParametersIsEqual = false;

            createdMatchReferenceObject.Add(createMatchReferenceObject);
        }

        return new Tuple<List<ReferenceObject>, bool>(createdMatchReferenceObject, allParametersIsEqual);
    }

    private List<ParameterInfo> FindTempReferenceNotSystemParameters(ReferenceObject analogReferenceObject,
        ContextSettings contextSettings)
    {
        ReferenceObject documentNSI = FindDocumentNSIReferenceObject(analogReferenceObject, contextSettings);
        if (documentNSI == null)
            return null;

        //Временный справочник который указан в документе НСИ. Параметр имеет тип строка...
        string guidGenerateValue = documentNSI[Settings.DocumentNSIGuidGenerateReference].GetString();


        if (!Guid.TryParse(guidGenerateValue, out Guid documentNsiGenerateReference))
            return null;

        var tempReference = FindReference(documentNsiGenerateReference);
        if (tempReference == null)
            return null;

        return tempReference.ParameterGroup.Parameters.Where(par => !par.IsSystem).ToList();
    }

    /// <summary>
    /// Создает объект соответствия параметров 
    /// </summary>
    /// <param name="clearDataReferenceObject"></param>
    /// <param name="contextSettings"></param>
    /// <param name="tempReferenceNotSystemParameter"></param>
    /// <returns></returns>
    private static ReferenceObject CreateAndCheckMatchReferenceObject(ReferenceObject clearDataReferenceObject,
        ContextSettings contextSettings, ParameterInfo tempReferenceNotSystemParameter, out bool isParameterEqual)
    {
        //В любом случае будет создана запись в комментарии будет записан результат.
        var createMatchReferenceObject = clearDataReferenceObject.CreateListObject
            (Settings.SettingsClearDataMatchParameter, Settings.SettingsClearDataMatchParameterClassObject);

        createMatchReferenceObject[Settings.SettingsClearDataMatchParameterTemp].Value =
            tempReferenceNotSystemParameter.Guid;

        string comment = string.Empty;

        string name = string.Format("[{0}]", tempReferenceNotSystemParameter.Name);

        //Пытаемся найти соответствующий параметр в аналоге
        var findAnalogParameter = contextSettings.AnalogParameters.FirstOrDefault(par => par.Name == tempReferenceNotSystemParameter.Name);

        isParameterEqual = true;

        if (findAnalogParameter == null)
        {
            comment = "Ошибка: параметр аналога не найден";
            isParameterEqual = false;
        }
        else
        {
            //Если найден соответствующий по наименование параметр, то подключаем его независимо от типа
            createMatchReferenceObject[Settings.SettingsClearDataMatchParameterAnalog].Value = findAnalogParameter.Guid;

            name = string.Format("[0] - [0]", tempReferenceNotSystemParameter.Name);

            if (findAnalogParameter.Type != tempReferenceNotSystemParameter.Type)
            {
                isParameterEqual = false;
                comment = "Ошибка: типы параметров не соответствуют";
            }
        }

        //Вот тут у меня вопрос. Может быть как то упростить запись параметров наименования, обозначения. name comment
        createMatchReferenceObject[Settings.SettingsClearDataMatchParameterName].Value = name;
        createMatchReferenceObject[Settings.SettingsClearDataMatchParameterComment].Value = comment;

        createMatchReferenceObject.EndChanges();

        return createMatchReferenceObject;
    }

    private ReferenceObject FindDocumentNSIReferenceObject(ReferenceObject analogReferenceObject, ContextSettings contextSettings)
    {
        //Вся эта конструкция проверок, позволяет не обращаться к серверу, за повторным поиском документа НСИ
        Guid documentNsiGuid = analogReferenceObject[contextSettings.AnalogDocNSI].GetGuid();

        if (contextSettings.DocumentNSIReferenceObject != null &&
            contextSettings.DocumentNSIReferenceObject.Guid == documentNsiGuid)
        {
            return contextSettings.DocumentNSIReferenceObject;
        }

        var documentNSI = contextSettings.DocumentNSI.Find(documentNsiGuid);
        if (documentNSI == null)
        {
            _errorStringBuilder.Append(string.Format("Не найден документ НСИ с Guid = '{0}'",
                analogReferenceObject[contextSettings.AnalogDocNSI].GetGuid()));
            return null;
        }

        contextSettings.DocumentNSIReferenceObject = documentNSI;

        return documentNSI;
    }

    private Reference FindReference(Guid referenceGuid)
    {
        var parameterInfo = Context.Connection.ReferenceCatalog.Find(referenceGuid);
        if (parameterInfo == null)
            return null;

        return parameterInfo.CreateReference();
    }

    private List<ReferenceObject> GetFilteredSelectedObjects(ContextSettings contextSettings)
    {
        var filtredReferenceObjects = new List<ReferenceObject>();

        //Пока не ясно работает ли загрузка параметров
        //Context.Reference.LoadSettings.AddParameters(Settings.TypeNSI , docNSI,settingClearData);

        foreach (var selectedObject in Context.GetSelectedObjects())
        {
            //Нужно все проверить по параметрам. параметров много, поэтому вынесу в Foreach
            if (CheckParametersObject(contextSettings, selectedObject))
                filtredReferenceObjects.Add(selectedObject);
        }

        return filtredReferenceObjects;
    }

    private static bool CheckParametersObject(ContextSettings contextSettings, ReferenceObject selectedObject)
    {
        return selectedObject[contextSettings.AnalogTypeNSI].GetGuid() != Guid.Empty &&
               selectedObject[contextSettings.AnalogDocNSI].GetGuid() != Guid.Empty &&
               selectedObject[contextSettings.AnalogSettingClearData].GetGuid() == Guid.Empty;
    }

    /// <summary>
    /// Возвращает гуид параметра из указанной группы параметров
    /// </summary>
    /// <param name="parameters">Группа параметров для поиска</param>
    /// <param name="nameParameter">Имя параметра</param>
    /// <returns></returns>
    private static Guid GetParameterGuidByName(ParameterInfoCollection parameters, string nameParameter)
    {
        var parameter = parameters.FindByName(nameParameter);
        if (parameter == null)
            throw new NullReferenceException("В текущем справочнике нет параметра: " + nameParameter);

        return parameter.Guid;
    }

    private void SetAnalogReferenceObjectPаrameter(ReferenceObject referenceObject, Guid referenceGuid)
    {
        if (TryBeginChanges(referenceObject))
        {
            referenceObject[Settings.DataSourceParameterReferenceAnalog].Value = referenceGuid;
            referenceObject.EndChanges();

            return;
        }

        _errorStringBuilder.Append(
            string.Format("Ошибка при взятии в редактирование выбранного объекта с ID:{0}{1}" +
                          "Возможно объект заблокирован другим пользователем{1}",
                referenceObject.SystemFields.Id, Environment.NewLine));
    }

    /// <summary>
    /// Удаляет справочник аналог
    /// </summary>
    /// <returns>Возвращает true если справочник был удален</returns>
    private bool DeleteOldAnalogReference(Guid referenceGuid)
    {
        var analogReferenceInfo = Context.Connection.ReferenceCatalog.Find(referenceGuid);
        if (analogReferenceInfo == null)
            return true;

        if (!Вопрос(string.Format("Справочник аналога: {0}, будет удален.{1}Продолжить выполнение?",
            analogReferenceInfo.Name, Environment.NewLine)))
        {
            return false;
        }

        var deleteReferenceBuilder = new ReferenceBuilder(analogReferenceInfo);
        if (deleteReferenceBuilder.IsActive)
            deleteReferenceBuilder.Deactivate();

        ReferenceBuilder.Delete(analogReferenceInfo);

        return true;
    }

    /// <summary>
    /// Создает пользовательское событие, кнопку на основе группы параметров
    /// </summary>
    /// <param name="parameterGroup"></param>
    private void CreateUserEventButton(ClassObject classObject, string buttonName, Guid macrosGuid, string methodName = "")
    {
        var userEvent = new UserEvent(classObject.Events)
        {
            Name = buttonName,
            Button = new UserEventButtonInfo
            {
                Text = buttonName,
                Position = UserEventButtonInfo.PositionInMenu.AfterCommands,
                EditObject = false,
                ExecuteForEachObject = false
            }
        };
        userEvent.Save();

        AddMacroEventHandlerToEvent(classObject, userEvent, macrosGuid, methodName);
    }

    /// <summary>
    /// Добавляет обработчик к заданному событию
    /// </summary>
    /// <param name="classObject">Тип справочника</param>
    /// <param name="userEvent">Событие справочника</param>
    /// <param name="macrosGuid">Выполняемый макрос</param>
    /// <param name="methodName">Метод выполняемого макроса</param>
    private void AddMacroEventHandlerToEvent(ClassObject classObject, UserEvent userEvent, Guid macrosGuid, string methodName = "")
    {
        var macros = (TFlex.DOCs.Model.References.Macros.Macro)Context.Connection.References.Macros.Find(macrosGuid);
        if (macros == null)
            return;

        MacroEventHandler macroEventHandler = new MacroEventHandler(classObject, userEvent, macros)
        {
            EntryPoint = methodName
        };

        macroEventHandler.Save();
    }

    /// <summary>
    /// Меняет основные параметры после импорта
    /// </summary>
    /// <param name="referenceBuilder"></param>
    private void ChangeImportedParemeters(ParameterGroup parameterGroup)
    {
        var parameters = parameterGroup.Parameters;
        foreach (var parameterInfo in parameters)
        {
            if (parameterInfo.IsSystem)
                continue;

            ParameterInfoBuilder parameterBulder = new ParameterInfoBuilder(parameterInfo);
            parameterBulder.AllowEdit = false;
            parameterBulder.Save();
        }
    }

    /// <summary>
    /// Создает базовые параметры в отдельной группе параметров в справочнике Документ НСИ, Тип НСИ
    /// </summary>
    /// <param name="parameterGroup"></param>
    private void CreateBaseParameters(ParameterGroup parameterGroup)
    {
        //Создаем группу параметров, для визуального восприятия
        ParameterGroupBuilder parameterGroupBuilder =
            new ParameterGroupBuilder(ParameterGroupType.TableOneToOne,
                                      parameterGroup,
                                      parameterGroup.Classes.BaseClasses.FirstOrDefault())
            {
                Name = "Базовые параметры",
                Comment = "Создано через систему NSI"
            };

        parameterGroupBuilder.Save();

        foreach (var parameterDescription in GetBaseParameterForCreatingAnalogReference())
        {
            ParameterInfoBuilder parameterInfoBuilder = new ParameterInfoBuilder(parameterGroupBuilder.ParameterGroup)
            {
                Name = parameterDescription.Name,
                IsVisible = true,
                Type = parameterDescription.Type
            };

            if (!string.IsNullOrEmpty(parameterDescription.UserControl))
                parameterInfoBuilder.UserControl = parameterDescription.UserControl;

            if (parameterDescription.ParametersValueList != null)
            {
                parameterInfoBuilder.ContainsValueList = true;

                foreach (var keyValuePair in parameterDescription.ParametersValueList)
                {
                    var listValue = new ListValue(parameterInfoBuilder.ParameterValueList)
                    { Name = keyValuePair.Value, Value = keyValuePair.Key };

                    parameterInfoBuilder.ParameterValueList.Add(listValue);
                }
            }

            parameterInfoBuilder.Save();
        }
    }

    /// <summary>
    /// Возвращает конструктор справочника
    /// </summary>
    /// <param name="createdReferenceId">Идентификатор справочника</param>
    /// <returns></returns>
    private ReferenceBuilder GetReferenceBulder(int createdReferenceId)
    {
        Context.Connection.ReferenceCatalog.Load();
        ReferenceInfo referenceInfo = Context.Connection.ReferenceCatalog.Find(createdReferenceId);
        if (referenceInfo == null)
            return null;

        return new ReferenceBuilder(referenceInfo);
    }

    private void OnReferenceCatalogChanges(int referenceid, ChangeType changetype, int clientview)
    {
        if (_createdReferenceId == 0 && changetype == ChangeType.Added)
            _createdReferenceId = referenceid;
    }

    /// <summary>
    /// Возвращает выбранный объект справочника на рабочей странице
    /// </summary>
    /// <param name="itemName">Имя элемента управления</param>
    /// <returns>Возвращает выбранный объект</returns>
    public ReferenceObject GetFocusedObjectFromWorkingPage(string itemName)
    {
        var supportSelectionVM = GetSupportSelection(itemName);
        if (supportSelectionVM == null)
            return null;

        var focusedObject = supportSelectionVM.FocusedObject;

        if (!(focusedObject is TFlex.DOCs.Client.ViewModels.References.ReferenceObjectViewModel referenceObjectViewModel))
            return null;

        return referenceObjectViewModel.ReferenceObject;
    }

    private ISupportSelection GetSupportSelection(string itemName)
    {
        IWindow currentWindow = Context.GetCurrentWindow();

        ILayoutItem foundItem = currentWindow.FindItem(itemName);
        if (foundItem == null)
            return null;

        //Окно справочника
        if (foundItem is ReferenceWindowLayoutItemViewModel referenceWindowLayoutItemVM)
            return referenceWindowLayoutItemVM.InnerViewModel as ISupportSelection;

        return null;
    }

    /// <summary>
    /// Запустить процесс импорта внешних данных
    /// </summary>
    /// <returns></returns>
    private bool StartImportProcessReference()
    {
        int version = 0;
        bool is64Bit = false;

        if (!GetOfficeDescription(ref version, ref is64Bit))
            is64Bit = false;

        if (version == 0)
            is64Bit = false;

        var info = new ProcessStartInfo();

        string path = Environment.CurrentDirectory;

        string _fileName = Path.Combine(path,
            String.Format("TFlex.DOCs.ReferenceImport{0}.exe",
                is64Bit ? ".x64" : string.Empty));

        string connectionParametersData = Context.Connection.ConnectionParameters.Serialize();

        info.FileName = _fileName;

        info.Arguments = String.Format("\"{0}\" \"{1}\"",
            connectionParametersData,
            version);

        info.CreateNoWindow = true;
        info.WindowStyle = ProcessWindowStyle.Normal;
        info.UseShellExecute = false;

        if (!String.IsNullOrEmpty(_fileName))
        {
            if (File.Exists(_fileName))
            {
                var process = Process.Start(info);

                if (process != null)
                {
                    process.WaitForExit();

                    if (process.ExitCode != 0)
                        throw new Exception(_fileName);
                }
            }
            else
            {
                throw new FileNotFoundException(string.Format("Файл не найден : {0}", _fileName), _fileName);
            }
        }
        return true;
    }

    /// <summary>
    /// Версия и разрядность Microsoft Office
    /// </summary>
    /// <param name="version">Номер версии</param>
    /// <param name="is64Bit">В случае x64 true, в случае x32 false</param>
    /// <returns>Возвращает true, если Office установлен</returns>
    private bool GetOfficeDescription(ref int version, ref bool is64Bit)
    {
        bool isOS64Bit = Environment.Is64BitOperatingSystem;
        RegistryView registryView = isOS64Bit ? RegistryView.Registry64 : RegistryView.Registry32;

        is64Bit = isOS64Bit;
        version = 0;

        using (RegistryKey localMachineKey = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, registryView))
        using (RegistryKey officeKey = localMachineKey.OpenSubKey(@"Software\Microsoft\Office\Common"))
        {
            if (officeKey == null)
            {
                if (is64Bit)
                {
                    using (RegistryKey officeKey32 = localMachineKey.OpenSubKey(@"Software\Wow6432Node\Microsoft\Office\Common"))
                    {
                        if (officeKey32 == null)
                            return false;

                        version = (int)officeKey32.GetValue("LastAccessInstall", 0);
                        is64Bit = false;
                    }
                }
                else
                {
                    return false;
                }
            }
            else
            {
                version = (int)officeKey.GetValue("LastAccessInstall", 0);
            }
        }

        return true;
    }

    private class ParameterDescription
    {
        public ParameterDescription()
        {
            Type = ParameterType.UniqueIdentifier;
        }

        public string Name { get; set; }

        public string UserControl { get; set; }

        public ParameterType Type { get; set; }

        public Dictionary<int, string> ParametersValueList { get; set; }
    }

    private class ContextSettings
    {
        public Guid AnalogTypeNSI = Guid.Empty;
        public Guid AnalogDocNSI = Guid.Empty;
        public Guid AnalogSettingClearData = Guid.Empty;
        public Guid AnalogReferenceGuid = Guid.Empty;

        public List<ParameterInfo> AnalogParameters;

        public Reference SettingsClearDataReference;
        public ClassObject SettingsClearDataReferenceClassObject;

        public Reference DocumentNSI;
        /// <summary>
        /// Документ НСИ добавил, так как вероятно все будет запускаться для всех объектов с один документом НСИ.
        /// </summary>
        public ReferenceObject DocumentNSIReferenceObject;

    }

    /// <summary>
    /// Возвращает список из базовых параметров, с описанием
    /// </summary>
    /// <returns></returns>
    private static IEnumerable<ParameterDescription> GetBaseParameterForCreatingAnalogReference()
    {
        IEnumerable<ParameterDescription> parameters = new ParameterDescription[]
        {
            new ParameterDescription() {Name = "Тип НСИ", UserControl = "<RepositoryGuidEditItemXMLData><UserControlName>GuidParameterEdit</UserControlName><ReferenceGuid>00bf7ef0-6080-4edd-a548-95b44df465c4</ReferenceGuid></RepositoryGuidEditItemXMLData>"},
            new ParameterDescription() {Name = "Документ НСИ", UserControl = "<RepositoryGuidEditItemXMLData><UserControlName>GuidParameterEdit</UserControlName><ReferenceGuid>a169916f-fa02-417b-b52f-63de54b06a59</ReferenceGuid></RepositoryGuidEditItemXMLData>"},
            new ParameterDescription() {Name = "Настройка соответствия", UserControl = "<RepositoryGuidEditItemXMLData><UserControlName>GuidParameterEdit</UserControlName><ReferenceGuid>212a5ec8-3f36-4501-bb46-082d200ba05f</ReferenceGuid></RepositoryGuidEditItemXMLData>"},
            //new ParameterDescription() {Name = "Состояние", Type = ParameterType.Int, ParametersValueList = _parametersDictionary},
            new ParameterDescription() {Name = "Комментарий", Type = ParameterType.NText},
            new ParameterDescription() {Name = "Внешний идентификатор", Type = ParameterType.NText},
            new ParameterDescription() {Name = "Статус актуальности", Type = ParameterType.Int, ParametersValueList = new Dictionary<int, string>() {{0, "Актуален"}, {1, "Удален" }, {2, "Отменен" }}}
        };

        return parameters;
    }

    private static class Settings
    {
        public static string ClassObjectName = "Запись";

        public static string WorkingPageItemName = "ImportExcel";

        public static Guid MDMAnalogStageScheme = new Guid("2c70a48e-a182-4da7-b300-4ed1a7059516");
        public static Guid AddedStageGuid = new Guid("b253355a-29ea-4db3-9e5c-2ffe8c27a0ee");

        /// <summary>
        /// Аналоги прикладных систем
        /// </summary>
        public static string AnalogCatalogFolderGuid = "Аналоги прикладных систем";//new Guid("483673d8-c6f6-49d2-bc33-d75957495f0b");

        /// <summary>
        /// Гуид текущего макроса запускаемого в системе
        /// Необходим при создании справочника аналога
        /// </summary>
        public static readonly Guid MacroNsiGuid = new Guid("cfbdbe0a-7825-4673-b68b-316da1d6224b");

        /// <summary>
        /// Название метода из текущего макроса который запускает настройку очистки
        /// Необходим при создании справочника аналога
        /// </summary>
        public static readonly string MacroNsiMetodConnectionSettingsName = "ConnectionSettingsClearInAnalogReference";

        public static readonly string NSIEventName = "Подключить настройки соответствия";

        /// <summary>
        /// Макрос (MDM и НСИ) Макрос для операции нечеткого поиска в справочнике
        /// </summary>
        public static readonly Guid MacroMDMSeachGuid = new Guid("826b4adb-12bc-480a-b01c-0f9cf2841346");
        public static readonly string MacroMDMMetodRun = "Run";
        public static readonly string MDMEventName = "Поиск";

        /// <summary>
        /// Макрос "MDM. Аналоги. Выбор типа НСИ"
        /// </summary>
        public static readonly Guid MacroMDMSetTypeNSIGuid = new Guid("90b92bdf-6fe2-4e07-8508-d4b19ff3c46c");
        public static readonly string ConnectTypeNSIEventName = "Подключить ТИП НСИ";

        public static readonly Guid DataSourceParameterReferenceAnalog = new Guid("bd2f577a-874f-4376-9e2e-e4b06e9f3123");

        public static readonly Guid SettingsClearData = new Guid("7d26dc3e-b6a4-4a04-9ea0-34f65b57ddba");
        public static readonly Guid SettingsClearDataReferenceAnalog = new Guid("4ba85cec-97b3-4209-ac2b-72661eadddea");
        public static readonly Guid SettingsClearDataDocNSI = new Guid("585bb58d-583c-48aa-a3bf-991981453993");
        public static readonly Guid SettingsClearDataName = new Guid("85fa3139-29a2-4513-8987-60ff2b335e74");
        public static readonly Guid SettingsClearDataEqualParameters = new Guid("b7e60b7c-eab6-4f88-819a-82f6136e1173");

        //Пока не используется
        //public static readonly Guid SettingsClearDataLinkDocNSI = new Guid("723e232a-a144-4bd7-974f-778e8a61ca00");

        public static readonly Guid SettingsClearDataMatchParameter = new Guid("cbd16b27-e0fe-40a0-90a0-c4883283037c");
        public static readonly Guid SettingsClearDataMatchParameterAnalog = new Guid("f91a0706-4407-4306-ab60-f941a4aa8728");
        public static readonly Guid SettingsClearDataMatchParameterTemp = new Guid("248760c5-3620-4530-ad85-59bad92ed1c1");
        public static readonly Guid SettingsClearDataMatchParameterName = new Guid("91226aca-5ddb-4402-ba78-f72a7a2c0498");
        public static readonly Guid SettingsClearDataMatchParameterComment = new Guid("25367970-1f16-4304-8359-59a1cf7e2054");

        public static readonly Guid SettingsClearDataMatchParameterClassObject = new Guid("467bc90c-167f-46b2-90b8-f81db0aeeba0");

        public static readonly Guid DocumentNSI = new Guid("a169916f-fa02-417b-b52f-63de54b06a59");
        public static readonly Guid DocumentNSIGuidGenerateReference = new Guid("79d316f0-0f13-4ee9-9316-e41f27389333");
    }
}
