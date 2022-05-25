using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using DbfDataReader;
using System.Windows.Forms;
using TFlex.DOCs.Model;
using TFlex.DOCs.Model.Macros; // Для работы с макроязыком
using TFlex.DOCs.Model.Macros.ObjectModel; // Для работы с макроязыком
using TFlex.DOCs.Model.Classes; // Для работы с ClassObject (требуется для создания нового объекта заданного типа в справочнике)
using TFlex.DOCs.Model.References; // Для работы со справочниками
using TFlex.DOCs.Model.References.Nomenclature; // Для получения объекта справочника документы через объект ЭСИ
using TFlex.DOCs.Model.References.Documents; // Для получения объекта справочника ЭСИ через объект справочника Документы
using TFlex.DOCs.Model.Desktop; // Для удаления объектов, а так же применения изменений

using TFlex.DOCs.Model.Structure; // Для использования класса ParameterInfo
using TFlex.DOCs.Model.Search;
using System.Data.Common;
using System.Data;
// Для выполнения кода в разных потоках
using System.Threading;
using System.Threading.Tasks;
using System.Globalization;
using Npgsql;
using System.Diagnostics;

//using NpsglLib;



/*
Для работы данного макроса так же требуется подключение сторонних библиотек
- DbfDataReader.dll
- System.Data.Common.dll
- System.Data.dll
*/

//TODO Создать класс для хранения и обработки записей, которые необходимо грузить в T-Flex
//TODO Реализовать загрузку материалов
//TODO Для стандартных изделий и электронных компонентов сделать автоматическое заполнение поля ГОСТ (если оно не было заполнено ранее)
//TODO Реализовать механизм сохранения необработанных строк для того, чтобы можно было провести дозагрузку данных после исправления ошибок
//TODO - Реализовать сохранение необработанных записей в файл
//TODO - Реализовать метод чтения необработанных записей из файла
//TODO - Реализовать метод, который будет производить проверку на наличие такого файла и проводить дозагрузку данных
//TODO Переписать итоговое сообщение так, чтобы там не отображались нулевые строки, а так же добавить туда статус "Пересоздано"
//TODO Изменить логику работы свойств HasError, HasExcluded, переместить случай с НеТребуетсяСоздавать из ошибочного в исключенные
//TODO Исправить определение типа таким образом, чтобы при наличии объекта ЭСИ или Документа, тип брался оттуда (И проводилась соответствующая корректировка в списке номенклатуры)
//TODO Добавить в валидацию проверку типа в Списке номенклатуры и с объектах ЭСИ (и документах)
//TODO Предусмотреть вариант с тем, что может при определенном стечении обстоятельств может возникнуть потребность поправить разные подключения, но будет несколько раз поправлена одна и та же связь (это касается случая, когда между
//двумя ДСЕ будет существовать несколько подключений.

public class Macro : MacroProvider {

    #region Fields and Properties of class Macro

    private static ReferenceInfo documentReferenceInfo; // Информация о справочнике "Документы"
    private static ReferenceInfo nomenclatureReferenceInfo; // Информация о справочнике "ЭСИ"
    private static ReferenceInfo listOfNomenclatureReferenceInfo; // Информация о справочнике "Список номенклатуры"
    private static ReferenceInfo componentReferenceInfo; // Информация о справочнике "Электронные компоненты"
    private static ReferenceInfo specReferenceInfo; // Информация о справочике Spec

    private static Reference documentReference; // Справочник "Документы"
    private static Reference nomenclatureReference; // Справочник "ЭСИ"
    private static Reference listOfNomenclatureReference; // Справочник "Список номенклатуры"
    private static Reference componentReference; // Справочник "Электронные компоненты"
    private static Reference specReference; // Справочник "Spec"

    private static ParameterInfo shifrOfDocument; // Информация о параметре "Обозначение" в справочнике список "Документы"
    private static ParameterInfo shifrOfNomenclature; // Информация о параметре "Обозначение" в справочнике "ЭСИ"
    private static ParameterInfo shifrOfListNomenclature; // Информация о параметре "Обозначение" в справочнике "Список номенклатуры"
    private static ParameterInfo shifrOfComponent; // Информация о параметре "Обозначение" в справочнике "Список номенклатуры"
    private static ParameterInfo nameOfDocument; // Информация о параметре "Наименование" в справочнике список "Документы"
    private static ParameterInfo nameOfNomenclature; // Информация о параметре "Наименование" в справочнике "ЭСИ"
    private static ParameterInfo nameOfListNomenclature; // Информация о параметре "Наименование" в справочнике "Список номенклатуры"
    private static ParameterInfo nameOfComponent; // Информация о параметре "Наименование" в справочнике "Список номенклатуры"
                                                  //private static ParameterInfo shifrOfSpec; // Информация о параметре "SHIFR" в справочнике "Spec"
                                                  //private static ParameterInfo izdOfSpec; // Информация о параметер "IZD" в справочинке "Spec"

    //private static Dictionary<string, List<ReferenceObject>> allConnections;

    // Списки ключевых слов для определения того, является ли изделие с неопределенным типом стандартным изделием или же электронным компонентом
    private static string[] keyWordsElectricalComponents = {
        "рез-р",
        "к-р ",
        "диод",
        "тиристор",
        "вилка",
        "розетка",
        "транзистор",
        "предохранитель",
        "реле",
        "контактор",
        "микросхема",
        "стабилитрон",
        "триак",
        "переключатель",
        "трансформатор",
        "дроссель",
        "резонатор",
        "сердечник",
        "гнездо",
        "вставка плавкая",
        "клемма",
        "штеккер",
        "выключатель",
        "тумблер",
        "кнопка",
        "плата",
        "мост",
        "блок"
    };

    private static string[] keyWordsStandartParts = {
        "подшипник",
        "шарик",
        "ролик",
        "шайба",
        "гайка",
        "болт",
        "кольцо",
        "винт",
        "заклепка",
        "штифт",
        "шунт",
        "канистра",
        "хомут",
        "профиль",
        "прокладка",
        "гвоздь",
        "вентилятор",
        "бирка",
        "наконечник",
        "втулка",
        "ключ",
        "кронштейн",
        "пружина",
        "шуруп",
        "шпилька",
        "замок",
        "футорка",
        "ящик",
        "амортизатор",
        "ведро"
    };

    // Словари с соответствием типов цифровым значениям в поле "Тип объекта" справочника "Список номенклатуры FoxPro"
    private static Dictionary<int, TypeOfObject> dictIntToTypes;
    private static Dictionary<TypeOfObject, int> dictTypesToInt;
    private static Dictionary<string, TypeOfObject> dictStringToTypes;
    private static Dictionary<TypeOfObject, ClassObject> dictTypesToClassObject;
    public static Dictionary<int, string> analog_old_dict = new Dictionary<int, string>();
    public static Dictionary<int, string> dbf_old_dict = new Dictionary<int, string>();
    public static Dictionary<int, string> dbf_new_dict = new Dictionary<int, string>();
    public static Dictionary<int, Dictionary<string, string>> RowDbfChangeNew = new Dictionary<int, Dictionary<string, string>>();
    public static Dictionary<int, Dictionary<string, string>> RowDbfChangeOld = new Dictionary<int, Dictionary<string, string>>();
    public static Dictionary<int, Dictionary<string, string>> RowDbfDel = new Dictionary<int, Dictionary<string, string>>();
    public static Dictionary<int, Dictionary<string, string>> RowDbfDel_all = new Dictionary<int, Dictionary<string, string>>();
    public static Dictionary<int, Dictionary<string, string>> RowDbfAdd = new Dictionary<int, Dictionary<string, string>>();

    public Dictionary<int, string> RowDbfAddParent = new Dictionary<int, string>();
    public static List<ReferenceObject> liststruct;
    public static List<string> RowNumParent = new List<string>();

    public static DateTime datelog = DateTime.Now;
    //string log = $"{pathdir}\\{(datelog.ToShortDateString()).Replace(".","")}";
    // string pathdir = @"C:\AEMexport";
    public static string pathdir = $"C:\\AEMexport\\{(datelog.ToShortDateString()).Replace(".", "")}";
    public static string patholddbf = @"C:\AEMexport\olddbf";
    public static string pathnewdbf = @"C:\AEM";

    //static object locker = new object();

    #endregion Fields and Properties of class Macro

    #region Constructor

    public Macro(MacroContext context)
        : base(context)
    {
        // Получаем информацию о справочниках
        documentReferenceInfo = Context.Connection.ReferenceCatalog.Find(Guids.References.Документы);
        nomenclatureReferenceInfo = Context.Connection.ReferenceCatalog.Find(Guids.References.Номенклатура);
        listOfNomenclatureReferenceInfo = Context.Connection.ReferenceCatalog.Find(Guids.References.СписокНоменклатуры);
        componentReferenceInfo = Context.Connection.ReferenceCatalog.Find(Guids.References.ЭлектронныеКомпоненты);
        specReferenceInfo = Context.Connection.ReferenceCatalog.Find(Guids.References.Spec);

        // Получаем справочники "Документы", "ЭСИ", "Список номенклатуры FoxPro", "Электронные компоненты", "Spec"
        documentReference = documentReferenceInfo.CreateReference();
        nomenclatureReference = nomenclatureReferenceInfo.CreateReference();
        listOfNomenclatureReference = listOfNomenclatureReferenceInfo.CreateReference();
        componentReference = componentReferenceInfo.CreateReference();
        specReference = specReferenceInfo.CreateReference();

        // Обнуляем содержание словаря allConnections на случай, если там осталась информация после предыдущего запуска макроса
        //allConnections = null;


        // Проверка на то, что справочники были успешно загружены
        if (
            documentReference == null ||
            nomenclatureReference == null ||
            listOfNomenclatureReference == null ||
            componentReference == null ||
            specReference == null
                ) {

            // Определяем, какой из справочников не получилось инициализировать, чтобы указать это в ошибке
            string errorMessage = string.Empty;

            errorMessage += documentReference == null ? "- Справочнику 'Документы'\n" : string.Empty;
            errorMessage += nomenclatureReference == null ? "- Справочнику 'Электронная структура изделий'\n" : string.Empty;
            errorMessage += listOfNomenclatureReference == null ? "- Справочнику 'Список номенклатуры FoxPro'\n" : string.Empty;
            errorMessage += componentReference == null ? "- Справочнику 'Электронные компоненты'\n" : string.Empty;
            errorMessage += specReference == null ? "- Справочнику 'Spec'\n" : string.Empty;

            throw new Exception(string.Format("При выполнении макроса возникла ошибка. Не удалось получить доступ к:\n{0}", errorMessage));
        }

        // Получаем объекты ParameterInfo для поиска в справочниках по параметру
        shifrOfDocument = documentReference.ParameterGroup.FindRelation(Guids.Parameters.Документы.ДанныеДляСпецификации)[Guids.Parameters.Документы.Обозначение];
        shifrOfNomenclature = nomenclatureReference.ParameterGroup[Guids.Parameters.Номенклатура.Обозначение];
        shifrOfListNomenclature = listOfNomenclatureReference.ParameterGroup[Guids.Parameters.СписокНоменклатуры.Обозначение];
        shifrOfComponent = componentReference.ParameterGroup[Guids.Parameters.ЭлектронныеКомпоненты.Обозначение];
        nameOfDocument = documentReference.ParameterGroup[Guids.Parameters.Документы.Наименование];
        nameOfNomenclature = nomenclatureReference.ParameterGroup[Guids.Parameters.Номенклатура.Наименование];
        nameOfListNomenclature = listOfNomenclatureReference.ParameterGroup[Guids.Parameters.СписокНоменклатуры.Наименование];
        nameOfComponent = componentReference.ParameterGroup[Guids.Parameters.ЭлектронныеКомпоненты.Наименование];

        // Проверка на то, что все параметры, необходимые для поиска объектов, были получены
        if (shifrOfDocument == null || shifrOfNomenclature == null || shifrOfListNomenclature == null || shifrOfComponent == null)
            throw new Exception("При выполнении макроса возникла ошибка. " +
                    "Не удалось получить доступ к одному или нескольким описанием параметров " +
                    "(Обозначения в справочниках 'Документы', 'Электронная структура изделия', 'Список номенклатуры', 'Электронные комноненты')"
                    );

        // Заполняем справочники для трансформации типов в их цифровое представление
        dictIntToTypes = new Dictionary<int, TypeOfObject>();
        dictIntToTypes.Add(0, TypeOfObject.НеОпределено);
        dictIntToTypes.Add(1, TypeOfObject.СборочнаяЕдиница);
        dictIntToTypes.Add(2, TypeOfObject.СтандартноеИзделие);
        dictIntToTypes.Add(3, TypeOfObject.ПрочееИзделие);
        dictIntToTypes.Add(4, TypeOfObject.Изделие);
        dictIntToTypes.Add(5, TypeOfObject.Деталь);
        dictIntToTypes.Add(6, TypeOfObject.ЭлектронныйКомпонент);
        dictIntToTypes.Add(7, TypeOfObject.Материал);
        dictIntToTypes.Add(8, TypeOfObject.Другое);

        dictTypesToInt = dictIntToTypes.ToDictionary(kvp => kvp.Value, kvp => kvp.Key);

        dictStringToTypes = new Dictionary<string, TypeOfObject>();
        dictStringToTypes.Add("сборочная единица", TypeOfObject.СборочнаяЕдиница);
        dictStringToTypes.Add("деталь", TypeOfObject.Деталь);
        dictStringToTypes.Add("стандартное изделие", TypeOfObject.СтандартноеИзделие);
        dictStringToTypes.Add("прочее изделие", TypeOfObject.ПрочееИзделие);
        dictStringToTypes.Add("изделие", TypeOfObject.Изделие);
        dictStringToTypes.Add("электронный компонент", TypeOfObject.ЭлектронныйКомпонент);
        dictStringToTypes.Add("другое", TypeOfObject.Другое);

        dictTypesToClassObject = new Dictionary<TypeOfObject, ClassObject>();
        dictTypesToClassObject.Add(TypeOfObject.СборочнаяЕдиница, documentReference.Classes.Find(Guids.Types.Документы.СборочнаяЕдиница));
        dictTypesToClassObject.Add(TypeOfObject.СтандартноеИзделие, documentReference.Classes.Find(Guids.Types.Документы.СтандартноеИзделие));
        dictTypesToClassObject.Add(TypeOfObject.ПрочееИзделие, documentReference.Classes.Find(Guids.Types.Документы.ПрочееИзделие));
        dictTypesToClassObject.Add(TypeOfObject.Изделие, documentReference.Classes.Find(Guids.Types.Документы.Изделие));
        dictTypesToClassObject.Add(TypeOfObject.Деталь, documentReference.Classes.Find(Guids.Types.Документы.Деталь));
        dictTypesToClassObject.Add(TypeOfObject.ЭлектронныйКомпонент, componentReference.Classes.Find(Guids.Types.ЭлектронныеКомпоненты.ЭлектронныйКомпонент));
        dictTypesToClassObject.Add(TypeOfObject.Другое, documentReference.Classes.Find(Guids.Types.Документы.Другое));



        // /* 111
#if DEBUG
        System.Diagnostics.Debugger.Launch();
        System.Diagnostics.Debugger.Break();
#endif
        //   */
    }

    #endregion Constructor

    #region Guids class

    private static string Host = "tflex-test";//"localhost";
    private static string User = "analog_foxpro_read"; //"analog_foxpro_read";
    private static string DBname = "tflex-analog-import-dbf";
    private static string Password = "tdocs528"; //"postgres529"; //"tdocs528";
    private static string Port = "5449"; //"5433";

    private static class Guids {
        public static class References {
            public static Guid Документы = new Guid("ac46ca13-6649-4bbb-87d5-e7d570783f26");
            public static Guid Номенклатура = new Guid("853d0f07-9632-42dd-bc7a-d91eae4b8e83");
            public static Guid СписокНоменклатуры = new Guid("c9d26b3c-b318-4160-90ae-b9d4dd7565b6");
            public static Guid ЭлектронныеКомпоненты = new Guid("2ac850d9-5c70-45c2-9897-517ab571b213");
            public static Guid Spec = new Guid("c587b002-3be2-46de-861b-c551ed92c4c1");
        }

        public static class Parameters {
            public static class Номенклатура {
                public static Guid Обозначение = new Guid("ae35e329-15b4-4281-ad2b-0e8659ad2bfb");
                public static Guid Наименование = new Guid("45e0d244-55f3-4091-869c-fcf0bb643765");
            }

            public static class Документы {
                public static Guid Обозначение = new Guid("b8992281-a2c3-42dc-81ac-884f252bd062");
                public static Guid Наименование = new Guid("7e115f38-f446-40ce-8301-9b211e6ce5fd");
                public static Guid ДанныеДляСпецификации = new Guid("aa5a9c14-85b8-45a5-8fb9-be72286fb4db"); // Группа параметров
            }

            public static class СписокНоменклатуры {
                public static Guid Обозначение = new Guid("1bbb1d78-6e30-40b4-acde-4fc844477200");
                public static Guid Наименование = new Guid("c531e1a8-9c6e-4456-86aa-84e0826c7df7");
                public static Guid ТипНоменклатуры = new Guid("3c7a075f-0b53-4d68-8242-9f76ca7b2e97");
            }

            public static class ЭлектронныеКомпоненты {
                public static Guid Обозначение = new Guid("65e0e04a-1a6f-4d21-9eb4-dfe5a135ec3b");
                public static Guid Наименование = new Guid("01184891-8364-4a5c-bf05-2163e1f3d460");
            }

            public static class Spec {
                public static Guid Shifr = new Guid("2a855e3d-b00a-419f-bf6f-7f113c4d62a0");
                public static Guid Izd = new Guid("817eca21-1e7e-46e0-b80a-e61685bef5f7");
                public static Guid Naim = new Guid("cd16e2f4-a499-435c-b578-83c01bbee723");
            }
        }

        public static class Links {
            public static class СписокНоменклатуры {
                public static Guid Номенклатура = new Guid("ec9b1e06-d8d5-4480-8a5c-4e6329993eac");
                public static Guid Spec = new Guid("a356fe70-1a77-4e8a-9a90-38ba118431a8");
            }
        }

        public static class hLinks {
            public static Guid Количество = new Guid("3f5fc6c8-d1bf-4c3d-b7ff-f3e636603818");
            public static Guid Позиция = new Guid("ab34ef56-6c68-4e23-a532-dead399b2f2e");
        }

        public static class Types {
            public static class Номенклатура {
                public static Guid Деталь = new Guid("08309a17-4bee-47a5-b3c7-57a1850d55ea");
                public static Guid СборочнаяЕдиница = new Guid("1cee5551-3a68-45de-9f33-2b4afdbf4a5c");
                public static Guid Изделие = new Guid("7fa98498-c39c-44fc-bcaa-699b387f7f46");
                public static Guid СтандартноеИзделие = new Guid("87078af0-d5a1-433a-afba-0aaeab7271b5");
                public static Guid ПрочееИзделие = new Guid("f50df957-b532-480f-8777-f5cb00d541b5");
                public static Guid Другое = new Guid("c314a54e-8cea-453c-9160-ef6645584294");
                public static Guid ЭлектронныйКомпонент = new Guid("c314a54e-8cea-453c-9160-ef6645584294");
                public static Guid Материал = new Guid("f7f45e16-ceba-4d26-a9af-f099a2e2fca6");
            }

            public static class Документы {
                public static Guid Деталь = new Guid("7c41c277-41f1-44d9-bf0e-056d930cbb14");
                public static Guid СборочнаяЕдиница = new Guid("dd2cb8e8-48fa-4241-8cab-aac3d83034a7");
                public static Guid Изделие = new Guid("11d2fb6f-baa7-401c-bbd9-7f3222f5c5e8");
                public static Guid СтандартноеИзделие = new Guid("582dad76-1b07-4c4b-b97d-cc89b0149aa6");
                public static Guid ПрочееИзделие = new Guid("83e1ef55-0658-4e3e-afeb-d8fceee3c86d");
                public static Guid Другое = new Guid("89fcfb8f-5377-4481-b4d8-a465cb02f7a5");
            }

            public static class ЭлектронныеКомпоненты {
                public static Guid ЭлектронныйКомпонент = new Guid("a2a106b7-3bb8-4853-a16e-de441f73e499");
            }
        }

    }

    #endregion Guids class

    #region Entry points

    public override void Run() {
        //++

        var dataStart = DateTime.Now;
        //var t = GetTypeRef(new Guid("8324a9a7-51a9-4946-bde2-31fb458edd64"), new Guid("2911d8d9-b712-408a-9bae-c4b1dcc85f8d"));
        //var t2 = GetTypeRef();

        // runCompareTest();


        //////////////////////////////////
        CreateFolder();
        string runCopyDbf = "robocopy_net_test.bat";
        string runCopyOldDbf = "robocopy_old_copy.bat";
        CopyFoxProFiles(runCopyDbf);
        CopyFoxProFiles(runCopyOldDbf);

        //CopyFoxProFiles();


        //RunImpotDbftoAnalog();
        
        //runImportAnalogToSystemRef();

        runCompare(false);

        /////////////////////////////////

        //runCompareDebag();
        //runCompareDebag();
        //Connect.connectBase();

        /*        var precount = 0;
                var result_link_count = UpgradeLinkRoles();
                while (result_link_count > 0)
                {

                    if (precount == (result_link_count))
                    {
                        break;
                    }
                    precount = result_link_count;

                    result_link_count=UpgradeLinkRoles();
                }*/






        //UpgradeLinkRoles();
        var dataStop = DateTime.Now;
        Сообщение("ЗАГРУЗКА ЗАВЕРШИЕНА", String.Format("Загрузка завершена {0}", (dataStop - dataStart).ToString()));
    }



    /// <summary>
    /// Создание папки для dbf
    /// </summary>
    public void CreateFolder()
    {
        //+
        if (Directory.Exists(pathdir) == false)
        {
            Directory.CreateDirectory(pathdir);
        }
        if (Directory.Exists(patholddbf) == false)
        {
            Directory.CreateDirectory(patholddbf);
        }
        if (Directory.Exists(pathnewdbf) == false)
        {
            Directory.CreateDirectory(pathnewdbf);
        }
    }

    /// <summary>
    /// Проставление типа номенклатуры для выбраных объектов
    /// </summary>
    public void selectObj()
    {

        Объект item = ТекущийОбъект;
        //Объекты объекты = ВыбранныеОбъекты;
        //var refername = объект.Reference;
        //  foreach (var item in объекты)
        //  {
        var tip = 0;
        // string Shifr, string Naim
        if (item.Reference.ToString().Equals("SPEC"))
            tip = SetTipNomenclature(item["SHIFR"].ToString(), item["NAIM"].ToString());
        if (item.Reference.ToString().Equals("Список номенклатуры FoxPro"))
            tip = SetTipNomenclature(item["Обозначение"].ToString().Replace(".", ""), item["Наименование"].ToString().Replace(".", ""));

        if (tip != 0)
        {
            item.Изменить();
            item["Тип Номенклатуры"] = tip;
            item.Сохранить();
        }
        //  }
        //SetTipNomenclature(объект["Обозначение"].ToString(), объект["Наименование"].ToString());


    }

    public void CopyFoxProFiles(string pathToFileString)
    {
        string AEMDBF_DIR = @"C:\AEM";
        if (Directory.Exists(AEMDBF_DIR) == false)
        {
            Directory.CreateDirectory(AEMDBF_DIR);
        }

        string pathToDirectory = @"\\fs\share\Отдел АСУП\Служба PLM\Маркин А.А\test\";
        //string pathToFileString = "copy_net_test.bat";
        //string pathToFileString = "robocopy_net_test.bat";
        ProcessStartInfo infoStartProcess = new ProcessStartInfo();
        infoStartProcess.WorkingDirectory = pathToDirectory;
        infoStartProcess.FileName = pathToFileString;
        infoStartProcess.CreateNoWindow = false;
        infoStartProcess.WindowStyle = ProcessWindowStyle.Hidden;
        Process copyFPfiles = Process.Start(infoStartProcess);
        copyFPfiles.WaitForExit();
    }

    public void RunCopyFileCommand()
    {

        /*        robocopy "C:\AEM" "C:\AEMexport\olddbf" / MIR / Z
        robocopy "\\fs\share\Отдел АСУП\Служба PLM\Маркин А.А\test\newdbf" "c:\aem" / MIR / Z*/

        ProcessStartInfo startRobocopy = new ProcessStartInfo("robocopy.exe > c:\\work\\tetst1.xt");
        // "\"C:\\AEM\" \"C:\\AEMexport\\olddbf\" / MIR / Z > c:\\work\\tetst1.xt"
        startRobocopy.WindowStyle = ProcessWindowStyle.Normal;
        startRobocopy.UseShellExecute = true; // which is the default value.
        /*
         //без окна
         * startRobocopy.CreateNoWindow = true;
        startRobocopy.UseShellExecute = false;*/

        //startRobocopy.Arguments = "\"C:\\AEM\" \"C:\\AEMexport\\olddbf\" / MIR / Z";
        /*        Process.Start(startRobocopy);
                startRobocopy.WaitForExit();*/
        // startRobocopy.CreateNoWindow=
        Process copyFPfiles = Process.Start(startRobocopy);
   
        copyFPfiles.WaitForExit();

        /*        ProcessStartInfo infoStartProcess = new ProcessStartInfo();
                infoStartProcess.WorkingDirectory = pathToDirectory;
                infoStartProcess.FileName = pathToFileString;
                infoStartProcess.CreateNoWindow = false;
                infoStartProcess.WindowStyle = ProcessWindowStyle.Hidden;
                Process copyFPfiles = Process.Start(infoStartProcess);
                copyFPfiles.WaitForExit();*/
    }



    /// <summary>
    /// Обновляем dbf до актуальной версии в рабочей папке
    /// </summary>

    public void copyTableDbf(string filename, bool testdata = true)
    {
        //+
        //лужба PLM\Маркин А.А\test
        ///*
        ///
        // FileInfo newdbf=new FileInfo();

        string newdbfpath;
        string olddbfpath;

        if (testdata)
        {
            newdbfpath = @$"\\fs\share\Отдел АСУП\Служба PLM\Маркин А.А\test\newdbf\{filename}";
            olddbfpath = @$"\\fs\share\Отдел АСУП\Служба PLM\Маркин А.А\test\olddbf\{filename}";
            //*/

            /*
            FileInfo newdbf = new FileInfo(@$"\\fs\share\Отдел АСУП\Служба PLM\update_analog\newdbf\{filename}");
            FileInfo olddbf = new FileInfo(@$"\\fs\share\Отдел АСУП\Служба PLM\update_analog\olddbf\{filename}");
            */
        }
        else
        {
            //\\fs\share\Отдел АСУП\Служба PLM\Маркин А.А\test\pdbf\
            newdbfpath = @$"\\fs\FoxProDB\COMDB\PROIZV\{filename}";
            //newdbfpath = @$"\\fs\share\Отдел АСУП\Служба PLM\Маркин А.А\test\pdbf\{filename}";
            olddbfpath = @$"{pathnewdbf}\{filename}";
        }

        FileInfo newdbf = new FileInfo(newdbfpath);
        FileInfo olddbf = new FileInfo(olddbfpath);

        if (olddbf.Exists)
        {
            olddbf.CopyTo(patholddbf + $"\\{filename}", true);
        }
        if (newdbf.Exists)
        {
            newdbf.CopyTo(pathnewdbf + $"\\{filename}", true);
        }
    }

    public void runImportAnalogToSystemRef()
    {
        //RefTab klasm = new RefTab(TableName.KLASM);
        //RefTab spec = new RefTab(TableName.KLASM);


        RefTab kat_ediz = new RefTab(TableName.KAT_EDIZ);
        RefTab osnast = new RefTab(TableName.OSNAST);
        
        RefTab kat_ins = new RefTab(TableName.KAT_INS);
        
        RefTab norm = new RefTab(TableName.NORM);
        
        RefTab spec = new RefTab(TableName.SPEC);
        RefTab poluf = new RefTab(TableName.Poluf);
        RefTab kat_izd = new RefTab(TableName.KAT_IZD);
        RefTab klasm = new RefTab(TableName.KLASM);
        RefTab klas = new RefTab(TableName.KLAS);
        RefTab marchp = new RefTab(TableName.MARCHP);


        //UpdateAnalog(klasm);
        var poluf_obj = GetAnalog(poluf.guid);
        var kat_izd_obj = GetAnalog(kat_izd.guid);
        var klas_obj = GetAnalog(klas.guid);
        var klasm_obj = GetAnalog(klasm.guid);
        var spec_obj = GetAnalog(spec.guid);
        var norm_obj = GetAnalog(norm.guid);

/*
        RunImportRole(klasm_obj, Tab.KLASM.RoleChange.RoleKLASMMaterial);
        RunImportRole(poluf_obj, Tab.Poluf.RoleChange.RoleListNum);
        RunImportRole(kat_izd_obj, Tab.KAT_IZD.RoleChange.RoleListNum);
        RunImportRole(klas_obj, Tab.KLAS.RoleChange.RoleListNum);
        RunImportRole(klasm_obj, Tab.KLASM.RoleChange.RoleListNum);
 */       
        
        RunImportRole(spec_obj, Tab.SPEC.RoleChange.RoleListNum);
        RunImportRole(spec_obj, Tab.SPEC.RoleChange.RoleListNumIzd);
        RunImportRole(spec_obj, Tab.SPEC.RoleChange.RoleConnect);
        RunImportRole(norm_obj, Tab.NORM.RoleChange.RoleConnect);



        //var rowNum = new List<int> { 1, 2, 3, 4, 5, 6 };
        //var klasmrn = GetAnalog(rowNum, klasm.guid, klasm.ROW_NUMBER);
        //var specnm = GetAnalog(rowNum, spec.guid, spec.ROW_NUMBER);




        //RunImportRole(klasmrn, Tab.RoleChange.RoleKLASMMaterial);
        // ОбменДанными.ИмпортироватьОбъекты("Создание материалов из аналогов 2", klasmrn, показыватьДиалог: true);
        //, показыватьДиалог: true        
        // ОбменДанными.ИмпортироватьОбъекты("Список номенклатуры SPEC", null);
        //ОбменДанными.ИмпортироватьОбъекты("Список номенклатуры SPEC", klasmrn);
        // ОбменДанными.ИмпортироватьОбъекты("fba9104c-14ac-4c6a-a1eb-9823f717a1df", klasmrn);
        //ОбменДанными.ИмпортироватьОбъекты()
        //ОбменДанными.ИмпортироватьОбъекты("88c9749b-5f0a-408d-bf62-989ca641672a", klasmrn);

        //ОбменДанными.ИмпортироватьОбъекты("Создание материалов из аналогов 3", klasmrn, показыватьДиалог: true);

    }


    public void connecting()
    {


        //  Console.WriteLine("Hello World!");
        StringBuilder texttable = new StringBuilder();

        // Build connection string using parameters from portal
        //
        string connString =
            String.Format(
                "Server={0}; User Id={1}; Database={2}; Port={3}; Password={4};SSLMode=Prefer",
                Host,
                User,
                DBname,
                Port,
                Password);

        using (var conn = new NpgsqlConnection(connString))
        {

            Console.Out.WriteLine("Opening connection");
            conn.Open();

            //dbo."SPEC"
            using (var command = new NpgsqlCommand("SELECT \"ROW_NUMBER\", \"POS\",\"SHIFR\",\"NAIM\",\"PRIM\",\"IZD\" FROM dbo.\"SPEC\"", conn))
            {

                var reader = command.ExecuteReader();
                //  /*   
                while (reader.Read())
                {

                    texttable.Append(
                           string.Format(
                            "{0}|{1}|{2}|{3}|{4}|{5}\n",
                            reader.GetValue(0).ToString(),
                            reader.GetValue(1).ToString(),
                            reader.GetValue(2).ToString(),
                            reader.GetValue(3).ToString(),
                            reader.GetValue(4).ToString(),
                            reader.GetValue(5).ToString()
                            )
                        );

                }
                //  */

                /*var i = reader.FieldCount;
                reader.GetFieldValue
                reader.Read();
                reader.Rows
                reader.Close();*/
            }
        }

        Console.WriteLine(texttable.ToString());
        Save(texttable, @"C:\temp\spec_analog.tab2");
        Console.ReadLine();


    }

    /// <summary>
    /// Сравнение и применение изменений
    /// </summary>
    /// 
    public (Dictionary<int, Dictionary<string, string>>, Dictionary<int, Dictionary<string, string>>) runCompare(bool change = true)
    {
        //+
        


        //UpgredeLink(kat_ediz.guid, Tab.NORM.Link.LinkПодключения, Tab.NORM.RoleChange.RoleConnect);


        // var rowNum = new List<int> { 1, 2, 3, 4, 5, 6 };
        // var klasmrn = GetAnalog(rowNum, klasm.guid, klasm.ROW_NUMBER);
        //  RunImportRole(klasmrn , Tab.RoleChange.RoleKLASMMaterial);

        //   /*
        // TODO таблицы KLASM и TRUD изменить поля с датами чтобы они могли принимать NULL

///*
        RefTab spec = new RefTab(TableName.SPEC);
        runCompareRef(spec);
        if (change)
        {
            UpgredeLink(spec.guid, Tab.SPEC.Link.LinkСписокНоменклатуры, Tab.SPEC.RoleChange.RoleListNum);
            UpgredeLink(spec.guid, Tab.SPEC.Link.LinkСписокНоменклатурыСборка, Tab.SPEC.RoleChange.RoleListNumIzd);
            UpgredeLink(spec.guid, Tab.SPEC.Link.LinkПодключения, Tab.SPEC.RoleChange.RoleConnect);
        }
//*/




        RefTab norm = new RefTab(TableName.NORM);
        runCompareRef(norm);
        if (change)
            UpgredeLink(norm.guid, Tab.NORM.Link.LinkПодключения, Tab.NORM.RoleChange.RoleConnect);



        /*
         RefTab klas = new RefTab(TableName.KLAS);
         runCompareRef(klas);
        if (change)
            UpgredeLink(klas.guid, Tab.KLAS.Link.LinkСписокноменклатурыFoxPro, Tab.KLAS.RoleChange.RoleListNum);
       */

        /*
        RefTab klasm = new RefTab(TableName.KLASM);
        runCompareRef(klasm);
        if (change)
        {
            UpgredeLink(klasm.guid, Tab.KLASM.Link.LinkМатериал, Tab.KLASM.RoleChange.RoleKLASMMaterial);
            UpgredeLink(klasm.guid, Tab.KLASM.Link.LinkСписокноменклатурыFoxPro, Tab.KLASM.RoleChange.RoleListNum);
        }
        */

        /*
        RefTab marchp = new RefTab(TableName.MARCHP);
        runCompareRef(marchp);
        */

        /*
        RefTab trud = new RefTab(TableName.TRUD);
        runCompareRef(trud);
        */

 



        var rowold = UnionDict(RowDbfChangeOld, RowDbfDel);
        var rownew = UnionDict(RowDbfChangeNew, RowDbfAdd);


        Save(rownew, @"C:\aemexport\union_rownew.txt");
        Save(rowold, @"C:\aemexport\union_rowold.txt");
        return (rowold, rownew);
    }


    public void RunImpotDbftoAnalog()
    {
        RefTab kat_ediz = new RefTab(TableName.KAT_EDIZ);
        RefTab kat_ins = new RefTab(TableName.KAT_INS);
        RefTab kat_razm = new RefTab(TableName.KAT_RAZM);
        RefTab kat_stol = new RefTab(TableName.KAT_STOL);
        RefTab osnast = new RefTab(TableName.OSNAST);
        
        RefTab norm = new RefTab(TableName.NORM);
        RefTab spec = new RefTab(TableName.SPEC);
        RefTab poluf= new RefTab(TableName.Poluf);
        RefTab kat_izd = new RefTab(TableName.KAT_IZD);
        RefTab klasm = new RefTab(TableName.KLASM);
        RefTab klas = new RefTab(TableName.KLAS);
        RefTab marchp = new RefTab(TableName.MARCHP);
        RefTab trud = new RefTab(TableName.TRUD);

        //00:09:25.2437517
        //00:40:26.5286103 klas+klasm
        //Загрузка завершена 00:39:34.1170924 marchp
        //*/

        ///*
        ImpotDbftoAnalog(kat_ediz);
        ImpotDbftoAnalog(kat_ins);
        ImpotDbftoAnalog(kat_izd);
        ImpotDbftoAnalog(kat_stol);
        ImpotDbftoAnalog(osnast);
        ImpotDbftoAnalog(norm);
        ImpotDbftoAnalog(kat_razm);
        ImpotDbftoAnalog(poluf);
        ImpotDbftoAnalog(spec);
        ImpotDbftoAnalog(klas);
        ImpotDbftoAnalog(klasm);
        ImpotDbftoAnalog(marchp);
        ImpotDbftoAnalog(trud);



        /*
         if (RowDbfDel.Count() > 0)
                    DelobjRef(RowDbfDel, refTab);
        */
    }

    public void ImpotDbftoAnalog(RefTab refTab)
    {

        runImportDbf(refTab);

        RowDbfDel_all = null;
        string log = $"{pathdir}\\";
        string newdata = $"{pathnewdbf}\\{refTab.table_name}";
        dbf_new_dict = GetDbfData(newdata);
        var delrow = new List<int>();


/*        for (int i = rowcount; i <= dbf_new_dict.Count; i++)
        {
            if (dbf_new_dict[i].Equals("Delete"))
                delrow.Add(i);
        }*/
        


        foreach (var item in dbf_new_dict)
        {
            if (item.Value.Equals("Delete"))
                delrow.Add(item.Key);
        }


        RowDbfDel_all = GetDbfDataDic(newdata, delrow);
        Save(RowDbfDel_all, log + @$"\{refTab.table_name.Replace(".dbf", "")}_del_dictdbf_ALL.txt");

        if (RowDbfDel_all.Count() > 0)
                 DelobjRefAnalog(RowDbfDel_all, refTab);
    }




    public void runCompareRef(RefTab reftab, bool analogSql = false, bool updateanalog = false)
    {
        if (analogSql)
        {
            analog_old_dict = null;
            analog_old_dict = ConnectAnalog(reftab);
        }
        /* //copyTableDbf(reftab.table_name,true);
        if (reftab.table_name.Replace(".dbf", "").ToLower().Equals("trud"))
        {
            copyTableDbf(reftab.table_name.Replace(".dbf", ".FPT"), true);
        }*/

        Compare(reftab.table_name);
        if (updateanalog)
            UpdateAnalog(reftab,false);
    }


    public void runImportDbf(RefTab reftab)
    {

        //ОбменДанными.Импортировать(Tab.KAT_EDIZ.RoleChange.RoleImportKatEdizDbf);
        //ОбменДанными.Импортировать(Tab.KAT_EDIZ.RoleChange.RoleImportDbf,показыватьДиалог:false);
        
        ОбменДанными.Импортировать(reftab.roleimportdbf, показыватьДиалог: false);
        // ОбменДанными.Импортировать()
        //ОбменДанными.Импортировать();
        // analog_old_dict = null;
        // analog_old_dict = ConnectAnalog(reftab);
        //copyTableDbf(reftab.table_name);
        //Compare(reftab.table_name);
        //UpdateAnalog(reftab, true);

    }




    /// <summary>
    /// Получение списка объектов без связей
    /// </summary>
    public List<ReferenceObject> GetNotLinkObj()
    {
        var not_link_num = FilterReference(Tab.SPEC.guid, Tab.SPEC.Link.LinkСписокНоменклатуры);
        var not_link_numencl_bild = FilterReference(Tab.SPEC.guid, Tab.SPEC.Link.LinkСписокНоменклатурыСборка);
        var not_link_connect = FilterReference(Tab.SPEC.guid, Tab.SPEC.Link.LinkПодключения);

        not_link_num.AddRange(not_link_numencl_bild);
        not_link_num.AddRange(not_link_connect);
        return not_link_num;
    }

    public void UpgredeLink(List<ReferenceObject> updlink)
    {

        List<ReferenceObject> oneobject = new List<ReferenceObject>();
        var spec_group = GroupAnalog(updlink);
        foreach (var group in spec_group)
        {
            int countgr = group.Count();
            if (group.Count() > 1)
            {
                RunImportRole(group.AsEnumerable(), Tab.SPEC.RoleChange.RoleListNum);
                RunImportRole(group.AsEnumerable(), Tab.SPEC.RoleChange.RoleListNumIzd);
            }
            else
                oneobject.AddRange(group.AsEnumerable());
        }
        RunImportRole(oneobject, Tab.SPEC.RoleChange.RoleListNum);
        RunImportRole(updlink, Tab.SPEC.RoleChange.RoleConnect);
    }


    /// <summary>
    /// Обновление данных в справочниках по правилам обмена
    /// </summary>
    /// 


    public void UpgredeLink(Guid reference, Guid Link, string role)
    {
        //+
        var precount = 0;
        var result_link_count = UpgradeLinkRoles(reference, Link, role);
        while (result_link_count > 0)
        {

            if (precount == (result_link_count))
            {
                break;
            }
            precount = result_link_count;

            result_link_count = UpgradeLinkRoles(reference, Link, role);
        }
    }


    public int UpgradeLinkRoles(Guid reference, Guid Link, string role)
    {

        List<ReferenceObject> not_link = FilterReference(reference, Link);

        if (not_link != null)
            RunImportRole(not_link, role);

        not_link.AddRange(FilterReference(reference, Link));

        return not_link.Count();
    }




    public int UpgradeLinkRoles()
    {
        List<ReferenceObject> oneobject = new List<ReferenceObject>();
        var not_link_num = FilterReference(Tab.SPEC.guid, Tab.SPEC.Link.LinkСписокНоменклатуры);
        var not_link_numencl_izd = FilterReference(Tab.SPEC.guid, Tab.SPEC.Link.LinkСписокНоменклатурыСборка);
        not_link_num.AddRange(not_link_numencl_izd);


        var not_link_connect = FilterReference(Tab.SPEC.guid, Tab.SPEC.Link.LinkПодключения);
        var not_link_connect_norm = FilterReference(Tab.NORM.guid, Tab.NORM.Link.LinkПодключения);



        if (not_link_num.Count > 0)
        {
            var spec_group = GroupAnalog(not_link_num);
            foreach (var group in spec_group)
            {
                var test = group.AsEnumerable();
                var testcount = group.Count();
                if (group.Count() > 1)
                {

                    RunImportRole(group.AsEnumerable(), Tab.SPEC.RoleChange.RoleListNum);
                    RunImportRole(group.AsEnumerable(), Tab.SPEC.RoleChange.RoleListNumIzd);
                }
                else
                    oneobject.AddRange(group.AsEnumerable());

            }
            RunImportRole(oneobject, Tab.SPEC.RoleChange.RoleListNum);
            RunImportRole(oneobject, Tab.SPEC.RoleChange.RoleListNumIzd);

        }
        if ((not_link_connect.Count + not_link_connect_norm.Count) > 0)
        {
            RunImportRole(not_link_connect, Tab.SPEC.RoleChange.RoleConnect);
            RunImportRole(not_link_connect_norm, Tab.NORM.RoleChange.RoleConnect);
        }

        int result_link = not_link_connect.Count + not_link_num.Count + not_link_connect_norm.Count;//+ not_link_klasm_num.Count;



        return result_link;
    }



    public void UpgredeLink()
    {
        var not_link_num = FilterReference(Tab.SPEC.guid, Tab.SPEC.Link.LinkСписокНоменклатуры);
        var not_link_numencl_bild = FilterReference(Tab.SPEC.guid, Tab.SPEC.Link.LinkСписокНоменклатурыСборка);
        var not_link_connect = FilterReference(Tab.SPEC.guid, Tab.SPEC.Link.LinkПодключения);




        foreach (var item in not_link_num)
        {
            string shifr = item.GetObjectValue("SHIFR").ToString();
            if (!shifr.Equals(""))
            {
                try
                {
                    var find_shifr = GetAnalog(ShifrInsertDot(shifr), Tab.Список_номенклатуры_FoxPro.guid, Tab.Список_номенклатуры_FoxPro.Обозначение);
                    item.BeginChanges();

                    if (find_shifr.Count == 0)
                    {
                        //var test2 = CreateObj(item, "номенклатура");
                        item.SetLinkedObject(Tab.SPEC.Link.LinkСписокНоменклатуры, CreateObj(item, "номенклатура"));
                        var test = item.GetObject(Tab.SPEC.Link.LinkПодключения);
                    }
                    else if (find_shifr.Count == 1)
                    {
                        item.SetLinkedObject(Tab.SPEC.Link.LinkСписокНоменклатуры, find_shifr[0]);
                    }
                    item.EndChanges();
                }
                catch (Exception e)
                {
                    Error($"Ошибка {e.Message}");
                }
            }
        }
        //*/

        //поиск в списке номенклатуре по изделию


        //var test01 = GetAnalog(izd, Tab.SPEC.guid, Tab.SPEC.IZD);

        foreach (var item in not_link_numencl_bild)
        {
            string izd = item.GetObjectValue("izd").ToString();
            if (!izd.Equals(""))
            {
                try
                {
                    var find_izd = GetAnalog(ShifrInsertDot(izd), Tab.Список_номенклатуры_FoxPro.guid, Tab.Список_номенклатуры_FoxPro.Обозначение);
                    item.BeginChanges();

                    if (find_izd.Count == 0)
                    {
                        //var test2 = CreateObj(item, "номенклатура");
                        item.SetLinkedObject(Tab.SPEC.Link.LinkСписокНоменклатуры, CreateObj(item, "номенклатура"));
                    }
                    else if (find_izd.Count == 1)
                    {
                        item.SetLinkedObject(Tab.SPEC.Link.LinkСписокНоменклатуры, find_izd[0]);
                    }
                    item.EndChanges();
                }
                catch (Exception e)
                {
                    Error($"Ошибка {e.Message}");
                }
            }
        }




    }

    /* public void runCompareDebag()
     {
         ДиалогОжидания.Показать("Пожалуйста, подождите", true);
         ДиалогОжидания.СледующийШаг("Идёт  загрузка изменений из FOXPRO");



         ДиалогОжидания.Скрыть();
         var rowold = UnionDict(RowDbfChangeOld, RowDbfDel);
         var rownew = UnionDict(RowDbfChangeNew, RowDbfAdd);

         // Запуск обновления данных в TFlex
         ЗагрузитьИзмененияСтруктурыИзделияИзFox(rowold, rownew);



         //CheckLink(Tab.SPEC.guid, Tab.SPEC.ROW_NUMBER, rownew);

         Save(rownew, @"C:\aemexport\union_rownew.txt");
         Save(rowold, @"C:\aemexport\union_rowold.txt");

     }*/

    /// <summary>
    /// Настройка связи уже существующих записей в справочнике материалы
    /// </summary>
    /// 
    public void setLinkMaterialToKlasm ()
    {
        Объект текущийОбъект = ТекущийОбъект;
        var OKP = текущийОбъект["Код / обозначение"].ToString();
        Объект klasmobj = НайтиОбъект("KLASM", String.Format("[OKP] = '{0}'", OKP));
        
        Message("", klasmobj.ToString());
        klasmobj.Изменить();
        klasmobj.Подключить("Материал", текущийОбъект);
        klasmobj.Сохранить();
    }


    #region TestMethod()

    public void TestMethod() {
        // Проверка правильности сортировки данных, полученных от Саши
        Dictionary<int, Dictionary<string, string>> testOldDict = new Dictionary<int, Dictionary<string, string>>();

        testOldDict.Add(1, new Dictionary<string, string> () {
                {"SHIFR", "Потомок на удаление 1"},
                {"IZD", "Родитель на удаление 1"},
                });
        
        testOldDict.Add(2, new Dictionary<string, string> () {
                {"SHIFR", "Потомок на удаление 2"},
                {"IZD", "Родитель на удаление 2"},
                });

        testOldDict.Add(3, new Dictionary<string, string> () {
                {"SHIFR", "Потомок на изменение 1"},
                {"IZD", "Родитель на изменение 1"},
                });

        testOldDict.Add(4, new Dictionary<string, string> () {
                {"SHIFR", "Потомок на изменение 2"},
                {"IZD", "Родитель на изменение 2"},
                });

        testOldDict.Add(5, new Dictionary<string, string> () {
                {"SHIFR", "Потомок на изменение 3"},
                {"IZD", "Родитель на изменение 3"},
                });
        
        
        Dictionary<int, Dictionary<string, string>> testNewDict = new Dictionary<int, Dictionary<string, string>>();

        testNewDict.Add(6, new Dictionary<string, string> () {
                {"SHIFR", "Потомок на добавление 1"},
                {"IZD", "Родитель на добавление 1"},
                });
        
        testNewDict.Add(7, new Dictionary<string, string> () {
                {"SHIFR", "Потомок на добавление 2"},
                {"IZD", "Родитель на добавление 2"},
                });

        testNewDict.Add(3, new Dictionary<string, string> () {
                {"SHIFR", "Потомок на изменение 1"},
                {"IZD", "Родитель на изменение 1"},
                });

        testNewDict.Add(4, new Dictionary<string, string> () {
                {"SHIFR", "Потомок на изменение 2"},
                {"IZD", "Родитель на изменение 2"},
                });

        testNewDict.Add(5, new Dictionary<string, string> () {
                {"SHIFR", "Потомок на изменение 3"},
                {"IZD", "Родитель на изменение 3"},
                });
    }

    #endregion TestMethod()

    public void TestMethod2() {
        // Пробуем получить объекты, которые нужно перенести из одного справочинка в другой
        string denotation = "6013329677";

        ReferenceObject testObject = nomenclatureReference.FindOne(shifrOfNomenclature, denotation);
        if (testObject == null)
            Error("Не удалось найти тестовый объект");

        LinksOfObject links = new LinksOfObject(testObject);
        links.ReverseLinksDirection();

        Message("Информация", links.ToString());

        // Получение с объекта всех основных параметров
        string shifr = (string)testObject[Guids.Parameters.Номенклатура.Обозначение].Value;
        string name = (string)testObject[Guids.Parameters.Номенклатура.Наименование].Value;

        // Удаление объекта
        Desktop.CheckOut(testObject, true);
        Desktop.CheckIn(testObject, "Удаление объекта с целью изменения его типа на тип другого справочинка", false);
        Desktop.ClearRecycleBin(testObject);
        Message("Информация", "Удаление объекта завершено");

        // Пытаемся создать объект в справочнике "Электронные компоненты и произвести подключение всех объектов"
        ReferenceObject newComponent = componentReference.CreateReferenceObject(dictTypesToClassObject[TypeOfObject.ЭлектронныйКомпонент]);
        newComponent[Guids.Parameters.ЭлектронныеКомпоненты.Обозначение].Value = shifr;
        newComponent[Guids.Parameters.ЭлектронныеКомпоненты.Наименование].Value = name;
        newComponent.EndChanges();
        Message("Информация", "Создание электронного компонента завершено");

        // Создаем номенклатурный объект
        NomenclatureReference nomReference = nomenclatureReference as NomenclatureReference;
        ReferenceObject newNomenclature = nomReference.CreateNomenclatureObject(newComponent);
        Message("Информация", "Создание номенклатурного объекта на основе документа завершено");

        // Применяем изменения сразу для того, чтобы они случайно не отменились (и объект не был утерян навсегда)
        Desktop.CheckIn(newComponent, "Создание объекта в справочнике 'Электронные компоненты' взамен неправильно выгруженного в справочник 'Документы'", false);
        //Desktop.CheckIn(newNomenclature, "Создание объекта в справочнике 'Электронные компоненты' взамен неправильно выгруженного в справочник 'Документы'", false);

        // Прописываем ему все необходимые связи
        links.CloneLinksTo(newNomenclature);

        Message("Информация", "Клонирование подключений завершено");
    }

    #endregion Entry points

    #region Sevrice methods

    #region Gukov methods

    #region ЗагрузитьИзмененияСтруктурыИзделияИзFox()
    // Метод для применения изменений
    private void ЗагрузитьИзмененияСтруктурыИзделияИзFox(
            Dictionary<int, Dictionary<string, string>> oldData,
            Dictionary<int, Dictionary<string, string>> newData) {

        // Производим обработку входных данных
        SpecRecordsContainer records = new SpecRecordsContainer(oldData, newData);

        // Запускаем поиск позиций
        ObjectInTFlex.AddToSearch(records.ShifrsOnCreate, TypeOfAction.Создание);
        ObjectInTFlex.AddToSearch(records.ShifrsOnChange, TypeOfAction.Редактирование);
        ObjectInTFlex.AddToSearch(records.ShifrsOnDelete, TypeOfAction.Удаление);

        ObjectInTFlex.PreparatorySearch(); // Производим предварительную подготовку (поиск наименований и поиск объектов в справочниках DOCs по фильтам)
        ObjectInTFlex.Search();

        // Пишем лог по результатам проведения поиска
        string pathToLogDirectory = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "Тестирование выгрузки FoxPro");

        // Получаем словарь со всеми подключениями, которые есть в Spec по выгружаемым изделиям
        //allConnections = ПолучитьВсеПодключенияИзSpec(oldData, newData);
        // Создаем словарь для хранения записей, по которым возникли ошибки

        ApplyAddingsToTflex(records); // Производим добавление новых записей
        ApplyChangesToTflex(records); // Производим изменение существующих записей
        ApplyDeletingToTflex(records); // Производим удаление существующих записей

        СоздатьЛогРезультатовПоиска(pathToLogDirectory);
        СоздатьЛогПодключенияОбъектов(pathToLogDirectory, records);
    }
    #endregion ЗагрузитьИзмененияСтруктурыИзделияИзFox()

    #region СоздатьЛогРезультатовПоиска()

    private void СоздатьЛогРезультатовПоиска(string pathToDirectory) {
        // Производим проверку на существование пути файла лога
        CheckAndCreateDirectory(pathToDirectory);
        string pathToWrongRecordsLogFile = Path.Combine(pathToDirectory, "Ошибки в найденных позициях.txt");
        string pathToCorrectRecordsLogFile = Path.Combine(pathToDirectory, "Найденные позиции.txt");

        // Запись лог файлов
        File.WriteAllText(
                pathToWrongRecordsLogFile,
                string.Join("\n", ObjectInTFlex.Objects.Where(kvp => kvp.Value.HasError).Select(kvp => kvp.Value.Status))
                );
        File.WriteAllText(
                pathToCorrectRecordsLogFile,
                string.Join("\n", ObjectInTFlex.Objects.Where(kvp => !kvp.Value.HasError).Select(kvp => kvp.Value.Status))
                );

        // Вывод сообщения о завершении поиска
        string messageTemplate = 
            "Поиск позиций завершен.\nВсего позиций - {0}, из них:\n" +
            "Успешно найдено - {1}\n\n" +
            "(Создано в Списке номенклатуры FoxPro - {12})\n" +
            "(Создано в ЭСИ - {13})\n" +
            "(Создано в Документах и Электронных компонентах - {14})\n" +
            "Исключено - {2}\n" +
            "По причине признака конечного изделия - {3}\n" +
            "По причине того, что изделие не требуется создавать - {4}\n\n" +
            "С ошибками - {5}\n" +
            "Отсутствует наименование - {6}\n" +
            "Неоднозначныый выбор - {7}\n" +
            "Неопределенный тип - {8}\n" +
            "Некорректный номенклатурный объект - {9}\n" +
            "Ошибка при создании объекта - {10}\n" +
            "Требуется нормализация типов - {11}\n";

        messageTemplate = string.Format(
                messageTemplate,
                ObjectInTFlex.Objects.Count,
                ObjectInTFlex.Objects.Where(kvp => kvp.Value.Error == TypeOfError.ОшибкиНет).Count(),
                ObjectInTFlex.Objects.Where(kvp => kvp.Value.HasExcluded).Count(),
                ObjectInTFlex.Objects.Where(kvp => kvp.Value.Exclusion == TypeOfExclusion.ПризнакКонечногоИзделия).Count(),
                ObjectInTFlex.Objects.Where(kvp => kvp.Value.Exclusion == TypeOfExclusion.НеТребуетсяСоздавать).Count(),
                ObjectInTFlex.Objects.Where(kvp => kvp.Value.HasError).Count(),
                ObjectInTFlex.Objects.Where(kvp => kvp.Value.Error == TypeOfError.ОтсутствуетНаименование).Count(),
                ObjectInTFlex.Objects.Where(kvp => kvp.Value.Error == TypeOfError.НеоднозначныйВыбор).Count(),
                ObjectInTFlex.Objects.Where(kvp => kvp.Value.Error == TypeOfError.НеопределенныйТип).Count(),
                ObjectInTFlex.Objects.Where(kvp => kvp.Value.Error == TypeOfError.НекорректныйНоменклатурныйОбъект).Count(),
                ObjectInTFlex.Objects.Where(kvp => kvp.Value.Error == TypeOfError.СозданиеОбъекта).Count(),
                ObjectInTFlex.Objects.Where(kvp => kvp.Value.Error == TypeOfError.РазныеТипы).Count(),
                ObjectInTFlex.Objects.Where(kvp => (!kvp.Value.HasError) && (kvp.Value.ConnectionObjectStatus == StatusOfObject.Создан)).Count(),
                ObjectInTFlex.Objects.Where(kvp => (!kvp.Value.HasError) && (kvp.Value.NomenclatureObjectStatus == StatusOfObject.Создан)).Count(),
                ObjectInTFlex.Objects.Where(kvp => (!kvp.Value.HasError) && (kvp.Value.DocumentObjectStatus == StatusOfObject.Создан)).Count()
                );

        Message("Информация", messageTemplate);
    }

    #endregion СоздатьЛогРезультатовПоиска()

    #region СоздатьЛогПодключенияОбъектов()

    private void СоздатьЛогПодключенияОбъектов(string pathToDirectory, SpecRecordsContainer records) {
        // TODO Переработать метод под другой контейнер информации
        CheckAndCreateDirectory(pathToDirectory);

        string pathToErrorLog = Path.Combine(pathToDirectory, "Ошибки при подключении.txt");
        string pathToCompleteLog = Path.Combine(pathToDirectory, "Удачные подключения.txt");

        File.WriteAllText(pathToErrorLog, records.ToStringErrors());
        File.WriteAllText(pathToCompleteLog, records.ToStringComplete());


        string messageTemplate = "Подключение объектов завершено. Обработано {0} записей, из них успешно {1} позиций, {2} с ошибками";
        Message("Информация", string.Format(messageTemplate, records.CountError + records.CountComplete, records.CountComplete, records.CountError));
    }

    #endregion СоздатьЛогПодключенияОбъектов()

    #region Методы добавления, удаления и изменения элементов структуры изделия

    #region ApplyChangesToTflex()

    private void ApplyChangesToTflex(SpecRecordsContainer records) {
        string oldChild, newChild, oldParent, newParent; int oldPos, newPos; double oldPrim, newPrim;

        foreach (int id in records.IdToChange) {

            oldChild = records.OldData[id]["SHIFR"];
            oldParent = records.OldData[id]["IZD"];
            oldPos = int.Parse(records.OldData[id]["POS"]);
            oldPrim = double.Parse(records.OldData[id]["PRIM"]);

            newChild = records.NewData[id]["SHIFR"];
            newParent = records.NewData[id]["IZD"];
            newPos = int.Parse(records.OldData[id]["POS"]);
            newPrim = double.Parse(records.OldData[id]["PRIM"]);

            var objs = ObjectInTFlex.Objects;

            if ((!objs.ContainsKey(oldChild)) || (!objs.ContainsKey(oldParent)) || (!objs.ContainsKey(newChild)) || (!objs.ContainsKey(newParent))) {
                records.AddErrorRecord(id, TypeOfAction.Создание, TypeOfLinkError.ОтсутствуетОбъект);
                continue;
            }

            ObjectInTFlex oldChildObject = objs[oldChild];
            ObjectInTFlex oldParentObject = objs[oldParent];
            ObjectInTFlex newChildObject = objs[newChild];
            ObjectInTFlex newParentObject = objs[newParent];

            if ((oldChildObject.HasError) || (newChildObject.HasError) || (oldParentObject.HasError) || (newParentObject.HasError)) {
                records.AddErrorRecord(id, TypeOfAction.Редактирование, TypeOfLinkError.ОбъектСодержитОшибку);
                continue;
            }


            var incl = TypeOfExclusion.Отсутствует;
            if ((oldChildObject.Exclusion != incl) || (oldParentObject.Exclusion != incl) || (newChildObject.Exclusion != incl) || (newParentObject.Exclusion != incl)) {
                records.AddErrorRecord(id, TypeOfAction.Редактирование, TypeOfLinkError.ОбъектИсключен);
                continue;
            }

            if (oldChild == newChild) {
                if (oldParent == newParent) {
                    // Самый простой случай - требуется изменить либо позицию, либо количество
                    try {
                        oldChildObject.ChangeConnectionLinkTo(oldParentObject, oldPos, newPos, oldPrim, newPrim);
                        records.AddCompleteRecord(id, TypeOfAction.Редактирование);
                    }
                    catch (Exception e) {
                        records.AddErrorRecord(id, TypeOfAction.Редактирование, TypeOfLinkError.ОшибкаВПроцессеПодключения);
                    }
                }
                else {
                    bool linkDelete = false;
                    try {
                        oldChildObject.DeleteConnectionLinkTo(oldParentObject, oldPos, oldPrim);
                        linkDelete = true;
                        oldChildObject.CreateConnectionLinkTo(newParentObject, newPos, newPrim);
                        records.AddCompleteRecord(id, TypeOfAction.Редактирование);
                    }
                    catch (Exception e) {
                        string mes = linkDelete ? "Старая связь удалена" : "Старая связь не удалена";
                        mes = string.Format("{0}\n Текст ошибки:\n{1}", mes, e.Message);
                        records.AddErrorRecord(id, TypeOfAction.Редактирование, TypeOfLinkError.ОшибкаВПроцессеПодключения, mes);
                    }
                }
            }
            else {
                bool linkDelete = false;
                try {
                    oldChildObject.DeleteConnectionLinkTo(oldParentObject, oldPos, oldPrim);
                    linkDelete = true;
                    newChildObject.CreateConnectionLinkTo(newParentObject, newPos, newPrim);
                    records.AddCompleteRecord(id, TypeOfAction.Редактирование);
                }
                catch (Exception e) {
                    string mes = linkDelete ? "Старая связь удалена" : "Старая связь не удалена";
                    mes = string.Format("{0}\n Текст ошибки:\n{1}", mes, e.Message);
                    records.AddErrorRecord(id, TypeOfAction.Редактирование, TypeOfLinkError.ОшибкаВПроцессеПодключения, mes);
                }
            }
        }
    }

    #endregion ApplyChangesToTflex()

    #region ApplyAddingsToTflex()

    private void ApplyAddingsToTflex(SpecRecordsContainer records) {

        string child, parent; int pos; double prim;

        foreach (int id in records.IdToCreate) {
            child = records.NewData[id]["SHIFR"];
            parent = records.NewData[id]["IZD"];
            pos = int.Parse(records.NewData[id]["POS"]);
            prim = double.Parse(records.NewData[id]["PRIM"]);

            if ((!ObjectInTFlex.Objects.ContainsKey(child)) || (!ObjectInTFlex.Objects.ContainsKey(parent))) {
                records.AddErrorRecord(id, TypeOfAction.Создание, TypeOfLinkError.ОтсутствуетОбъект);
                continue;
            }

            ObjectInTFlex childObject = ObjectInTFlex.Objects[child];
            ObjectInTFlex parentObject = ObjectInTFlex.Objects[parent];

            if ((childObject.HasError) || (parentObject.HasError)) {
                records.AddErrorRecord(id, TypeOfAction.Создание, TypeOfLinkError.ОбъектСодержитОшибку);
                continue;
            }

            if ((childObject.Exclusion != TypeOfExclusion.Отсутствует) || (parentObject.Exclusion != TypeOfExclusion.Отсутствует)) {
                records.AddErrorRecord(id, TypeOfAction.Создание, TypeOfLinkError.ОбъектИсключен);
                continue;
            }

            try {
                childObject.CreateConnectionLinkTo(parentObject, pos, prim);
                records.AddCompleteRecord(id, TypeOfAction.Создание);
            }
            catch (Exception e) {
                records.AddErrorRecord(id, TypeOfAction.Создание, TypeOfLinkError.ОшибкаВПроцессеПодключения, e.Message);
            }
        }
    }

    #endregion ApplyAddingsToTflex()

    #region ApplyDeletingToTflex()

    private void ApplyDeletingToTflex(SpecRecordsContainer records) {
        string child, parent; int pos; double prim;
        foreach (int id in records.IdToDelete) {
            child = records.OldData[id]["SHIFR"];
            parent = records.OldData[id]["IZD"];
            pos = int.Parse(records.OldData[id]["POS"]);
            prim = double.Parse(records.OldData[id]["PRIM"]);


            if ((!ObjectInTFlex.Objects.ContainsKey(child)) || (!ObjectInTFlex.Objects.ContainsKey(parent))) {
                records.AddErrorRecord(id, TypeOfAction.Удаление, TypeOfLinkError.ОтсутствуетОбъект);
                continue;
            }

            ObjectInTFlex childObject = ObjectInTFlex.Objects[child];
            ObjectInTFlex parentObject = ObjectInTFlex.Objects[parent];

            if ((childObject.HasError) || (parentObject.HasError)) {
                records.AddErrorRecord(id, TypeOfAction.Удаление, TypeOfLinkError.ОбъектСодержитОшибку);
                continue;
            }

            if ((childObject.Exclusion != TypeOfExclusion.Отсутствует) || (parentObject.Exclusion != TypeOfExclusion.Отсутствует)) {
                records.AddErrorRecord(id, TypeOfAction.Удаление, TypeOfLinkError.ОбъектИсключен);
                continue;
            }

            try {
                childObject.DeleteConnectionLinkTo(parentObject, pos, prim);
                records.AddCompleteRecord(id, TypeOfAction.Создание);
            }
            catch (Exception e) {
                records.AddErrorRecord(id, TypeOfAction.Удаление, TypeOfLinkError.ОшибкаВПроцессеПодключения, e.Message);
            }
        }
    }

    #endregion ApplyDeletingToTflex()

    #endregion Методы добавления, удаления и изменения элементов структуры изделия

    #region CheckAndCreateDirectory()

    private void CheckAndCreateDirectory(string pathToDirectory) {
        if (!Directory.Exists(pathToDirectory))
            Directory.CreateDirectory(pathToDirectory);
    }

    #endregion CheckAndCreateDirectory()

    #endregion Gukov methods

    #region Markin methods

    public Dictionary<int, Dictionary<string, string>> UnionDict(Dictionary<int, Dictionary<string, string>> one, Dictionary<int, Dictionary<string, string>> two)
    {
        Dictionary<int, Dictionary<string, string>> onecopy = new Dictionary<int, Dictionary<string, string>>(one);
        foreach (var iterator in two)
        {
            if (!onecopy.ContainsKey(iterator.Key))
                onecopy.Add(iterator.Key, iterator.Value);
        }
        return onecopy;
    }


    public void UpdateAnalog(IEnumerable<ReferenceObject> refobj, string role)
    {




        // групируем по обозначению
        try
        {
            if (!role.Equals("Подключения SPEC") || !role.Equals("Подключения NORM"))
            {
                var spec_new_group = GroupAnalog(refobj);

                foreach (var group in spec_new_group)
                {
                    RunImportRole(group.AsEnumerable(), role);

                }
            }
            if (role.Equals("Подключения SPEC") || role.Equals("Подключения NORM"))
                RunImportRole(refobj, role);

        }
        catch (Exception e)
        {

            Message("Внимание", e.Message);

        }


        /*    DataExchange: Exception type: TFlex.DOCs.Synchronization.SyncData.Exceptions.SyncDataWarningException
        Message: Для оптимальной загрузки данных нужно добавить индекс для параметра 'Обозначение' справочника 'Список номенклатуры FoxPro'*/

    }

    public void UpdateAnalog(RefTab refTab, bool changeall = false)
    {
        //+
        

        ///*


        ////DelobjRef(RowDbfDel, refTab.guid, refTab.ROW_NUMBER, refTab.table_name);
        ///


        if (RowDbfDel.Count()>0)
            DelobjRef(RowDbfDel, refTab);
        //DelobjRef(RowDbfDel_all, refTab);
        if (RowDbfAdd.Count() > 0)
            CreateObj(RowDbfAdd, refTab);

        // //CreateObj(RowDbfAdd, refTab.guid, refTab.ROW_NUMBER, refTab.table_name);


        //var rowNumChange = new List<int> { 18088, 43285, 43633, 43637, 43648, 43652, 43653, 44023, 44288, 46248 };
        //var rowNumChange = new List<int> { 38181, 76061, 76062, 139093, 222726, 222813, 93133, 93134, 93137, 93136, 188087, 203650, 203764, 78020, 222719, 188088, 189842, 189925, 222807, 189715, 190007, 203651, 203765, 78015, 77759, 77852, 77914, 183439, 78032, 188089, 203652, 203766, 93160, 98248, 190008, 188090, 189716, 79647, 189843, 78021, 79955, 189926, 77760, 77853, 203653, 222808, 203767, 203881, 78027, 78023, 77915, 188092, 78043, 189844, 203655, 190009, 203883, 189717, 77761, 79808, 93162, 79649, 222721, 191054, 78024, 79731, 191050, 203884, 188093, 189845, 203656, 203770, 189928, 190010, 222810, 203885, 189718, 203657, 188094, 77784, 98263, 77762, 203658, 188095, 77949, 203772, 203886, 203659, 203773, 188096, 78026, 222723, 189929, 79959, 189846, 79810, 79651, 79733, 188097, 189719, 203660, 203888, 77950, 190011, 203774, 188099, 203776, 77763, 188100, 203777, 203664, 203778, 203892, 93125, 189720, 188101, 190012, 222812, 79652, 79734, 77913, 189848, 190013, 77952, 79889, 79735, 189721, 188103, 98264, 189932, 190014, 189722, 189849, 79654, 188104, 203781, 203895, 77921, 79736, 79962, 77765, 79813, 179558, 203667, 188105, 203668, 77766, 77859, 79655, 188108, 203671, 77768, 77861, 203899, 222720, 203785, 188109, 222722, 222724, 222725, 203672, 78030, 222809, 77770, 222814, 78029, 222883, 222884, 222859, 222881, 222882, 188111, 189723, 189850, 190015, 222727, 222815, 222728, 222816, 78035, 78034, 78036, 93163, 78037, 222729, 222817, 183436, 222731, 189724, 222818, 189934, 189851, 190016, 93127, 188112, 93128, 203790, 203904, 188113, 203676, 78039, 98253, 78041, 78042, 222732, 222819, 188202, 189158, 190017, 98254, 77887, 98256, 93130, 77917, 98257, 93170, 77703, 228323, 93131, 236928, 230678, 93114, 93115, 98247, 93116, 93158, 243409, 98245, 93140, 93110, 98246, 93176, 230679, 140805, 93111, 93139, 93142, 93143, 98259, 98260, 93141, 98261, 93172, 240972, 147503, 240940, 244612, 244614, 244618, 244620, 77702, 77758, 77764, 77767, 77769, 77771, 77772, 77786, 77796, 77856, 77857, 77858, 77860, 77862, 77863, 77864, 77865, 77946, 77953, 77955, 77956, 77961, 78038, 79648, 79650, 79656, 79729, 79730, 79732, 79737, 79738, 79807, 79809, 79811, 79812, 79814, 79815, 79884, 79885, 79886, 79887, 79888, 79890, 79891, 79956, 79960, 79963, 79964, 77795, 77851, 77854, 77855, 77877, 77879, 77916, 77922, 77948, 77951, 77959, 78028, 78040, 194246, 218490, 147507, 257950, 240971, 79653, 79806, 79883, 79892, 79957, 79958, 79961, 137949, 189847, 189852, 189931, 203662, 203663, 203666, 203674, 203675, 203769, 203771, 203782, 203786, 203878, 203879, 203880, 203887, 203900, 257529, 257673, 257949, 258033, 258034, 93175, 93103, 189725, 189927, 189930, 189933, 189935, 191052, 191056, 93104, 203780, 203788, 203789, 203890, 203891, 203894, 203896, 203902, 203903, 197210, 93132, 93135, 93145, 93147, 93157, 93168, 93171, 93177, 110949, 38142, 38149, 38185 };
       
        List<int> rowNumChange = new List<int>();
        
        if (changeall)
            rowNumChange.AddRange(dbf_old_dict.Keys.ToList()); //все записи считать измененными
        else
            rowNumChange.AddRange(RowDbfChangeNew.Keys.ToList());   //измененыые записи
        // записи вручную
        //var rowNumChange = new List<int> { 1 };
        if (rowNumChange.Count() > 0)
            ChangeObj(rowNumChange, refTab);

        var analog = GetAnalog(rowNumChange, refTab.guid, refTab.ROW_NUMBER);
        if (refTab.name_refer.Equals("SPEC"))
            UpgredeLink(analog);
        if (refTab.name_refer.Equals("NORM"))
            RunImportRole(analog, Tab.NORM.RoleChange.RoleConnect);
        if (refTab.name_refer.Equals("KLASM"))
        {
            RunImportRole(analog, Tab.KLASM.RoleChange.RoleKLASMMaterial);
            RunImportRole(analog, Tab.KLASM.RoleChange.RoleListNum);
        }
        if (refTab.name_refer.Equals("KLAS"))
        {
            RunImportRole(analog, Tab.KLAS.RoleChange.RoleListNum);
        }




    }




    /// <summary>
    /// Запускает импорт объектов по правилу обмена
    /// </summary>


    public void RunImportRole(IEnumerable<ReferenceObject> analog, string rolename)
    {



        //List<int> row_number = new List<int> { 263632, 263633, 263635, 263636, 263637 };
        // Объекты sp = НайтиОбъекты("SPEC",String.Format("[ROW_NUMBER] Входит в список '{0}'", String.Join(",", row_number.Distinct()) ));



        /*
               int countdev = analog.Count() / 1;
               var devAnalogList = DivideList.Partition(analog, countdev);
               foreach(var analogpart in devAnalogList)
                   ОбменДанными.ИмпортироватьОбъекты(rolename, analogpart, показыватьДиалог: false);
       */
        //ОбменДанными.ИмпортироватьОбъекты(rolename, analog, показыватьДиалог: false);
        ОбменДанными.ИмпортироватьОбъекты(rolename, analog, показыватьДиалог: false);

        //   ОбменДанными.ИмпортироватьОбъекты("Список номенклатуры SPEC", analog, показыватьДиалог: true);
        // ОбменДанными.ИмпортироватьОбъекты("Подключения SPEC", analog, показыватьДиалог: false);

    }



  


    public ReferenceObject CreateObj(ReferenceObject Object, string rolename)
    {
        ReferenceObject createobj = null;
        //Object.GetObjectValue("IZD").GetValue<int>()
        var izd = Object.GetObjectValue("IZD").ToString();
        var shifr = Object.GetObjectValue("SHIFR").ToString();
        var naim = Object.GetObjectValue("NAIM").ToString();
        var pos = Object.GetObjectValue("POS").ToString();
        var prim = Object.GetObjectValue("PRIM").ToString();

        if (rolename.Equals("номенклатура"))
        {
            try
            {
                var obj = CreateObject(Tab.Список_номенклатуры_FoxPro.guid.ToString(), Tab.Список_номенклатуры_FoxPro.nameTip);
                //obj[Tab.Список_номенклатуры_FoxPro.Обозначение.ToString()] = Object.GetObjectValue("SHIFR").ToString();
                obj["Обозначение"] = ShifrInsertDot(Object.GetObjectValue("SHIFR").ToString());
                obj["Наименование"] = Object.GetObjectValue("NAIM").ToString();
                obj.Save();
                createobj = (ReferenceObject)obj;
            }
            catch
            {
                Сообщение("", "ERR");
            }
        }


        if (rolename.Equals("номенклатура_сборка"))
        {
            try
            {
                var obj = CreateObject(Tab.Список_номенклатуры_FoxPro.guid.ToString(), Tab.Список_номенклатуры_FoxPro.nameTip);
                //obj[Tab.Список_номенклатуры_FoxPro.Обозначение.ToString()] = Object.GetObjectValue("SHIFR").ToString();
                obj["Обозначение"] = ShifrInsertDot(Object.GetObjectValue("IZD").ToString());
                obj["Наименование"] = Object.GetObjectValue("NAIM").ToString();
                obj.Save();
                createobj = (ReferenceObject)obj;
            }
            catch
            {
                Сообщение("", "ERR");
            }
        }

        if (rolename.Equals("подключения"))
        { }


        return createobj;
    }



    public void CreateObj(Dictionary<int, Dictionary<string, string>> getrowadd, Guid guidrefer, string filename)
    {
        StringBuilder stringBuilder = new StringBuilder();

        foreach (var row in getrowadd)
        {
            //поиск доавляемой row_num в Аналогах

            try
            {
                stringBuilder.Append(row.Key + " ");





                var obj = CreateObject(guidrefer.ToString(), "Запись");



                //obj["ROW_NUMBER"] = row.Key;

                obj["Обозначение"] = row.Value["SHIFR"];
                obj["Наименование"] = row.Value["NAIM"];
                /*foreach (var columnname in row.Value.Keys)
                {
                    stringBuilder.Append(row.Value[columnname]);
                    obj[columnname] = row.Value[columnname];
                }*/

                obj.Save();

                //обновляем ссылки
                //  UpgredeLink(analog);

            }
            catch (Exception exp)
            {
                Message("", exp.Message);
            }

        }
    }

    /// <summary>
    /// Определение типа номенклатуры
    /// </summary>

    public int SetTipNomenclature(string Shifr,string Naim)
    {
        List<string> Sborka = new List<string> { "8А1", "8А2", "8А3", "8А4", "8А5", "8А6", "УЯИС30"};
        List<string> Izd = new List<string> { "УЯИС79", "ЭМЗ" };
        List<string> Detal = new List<string> { "8А7", "8А8", "8А9","УЯИС71", "УЯИС72", "УЯИС73", "УЯИС74", "УЯИС75", "УЯИС76" };
        List<string> StandartIzd = new List<string> { "759"};
        List<string> ElectronKomponent = new List<string> { "МИКРО", "К-Р", "РЕЗ", "ТРАНЗИСТОР", "ДИОД" };
        int tip = 0;
        // Сборочная единица=1
        foreach (var item in Sborka)
        {
            if (Shifr.StartsWith(item))
                tip = 1;
        };


        //Стандартное изделие 2

        foreach (var item in StandartIzd)
        {
            if (Shifr.StartsWith(item))
                tip = 2;
        };


        //Прочее изделие  3
        //Изделие 4
        foreach (var item in Izd)
        {
            if (Shifr.StartsWith(item))
                tip = 4;
        };
        //Деталь  5
        foreach (var item in Detal)
        {
            if (Shifr.StartsWith(item))
                tip = 5;
        };


        //Электронный компонент   6

        



        foreach (var item in ElectronKomponent)
        {
            if (Shifr.StartsWith("629") || (Shifr.StartsWith("6") && Naim.StartsWith(item)))
                tip = 6;
        };
        

        

        //Материал    7

        return tip;
    }


    public void CreateObj(Dictionary<int, Dictionary<string, string>> getrowadd, RefTab refTab)
    {

        //Guid SPEC = new Guid("c587b002-3be2-46de-861b-c551ed92c4c1");
        StringBuilder stringBuilder = new StringBuilder();
        string pathnew = $"C:\\aemexport\\analog_add_{refTab.name_refer}.txt";



        var analog = GetAnalog(getrowadd.Keys.ToList(), refTab.guid, refTab.ROW_NUMBER);


        List<int> row_num_list = new List<int>();
        foreach (var item in analog)
        {
            // row_num_list.Add(item.GetObjectValue("ROW_NUMBER").GetValue<int>());
            row_num_list.Add(item.GetObjectValue(refTab.name_row_num).GetValue<int>());

        }


        foreach (var row in getrowadd)
        {
            //поиск доавляемой row_num в Аналогах
            if (!row_num_list.Contains(row.Key))
            {
                try
                {
                    stringBuilder.Append(row.Key + " ");





                    var obj = CreateObject(refTab.guid.ToString(), refTab.nameTip);


                    obj[refTab.name_row_num] = row.Key;


                    foreach (var columnname in row.Value.Keys)
                    {
                        stringBuilder.Append(row.Value[columnname]);
                        obj[columnname] = row.Value[columnname];

                    }
                    if (refTab.name_refer.Equals("SPEC"))
                        {
                        obj["Тип Номенклатуры"] = SetTipNomenclature(row.Value["SHIFR"], row.Value["NAIM"]);
                        }
                    obj.Save();




                    stringBuilder.Append("\n");
                }
                catch
                {
                    stringBuilder.Append("ERROR");
                }


            }
        }
        //обновляем ссылки
        // UpgredeLink(analog);



        stringBuilder.Append("\n");
        Save(stringBuilder, pathnew);
    }


    

    public void DelobjRef(Dictionary<int, Dictionary<string, string>> delrow, Guid refer, Guid parametr, string tablename)
    {

        StringBuilder stringBuilder = new StringBuilder();
        string pathnew = $"C:\\aemexport\\analog_del_{tablename}.txt";
        //var delrow_test= new List<int> { 36039, 36065, 36067, 36078, 36137, 36138 };
        //var analog = GetAnalog(delrow_test, refer, parametr);
        var analog = GetAnalog(delrow.Keys.ToList(), refer, parametr);


        //  List<int> row_num_list = new List<int>();
        var dataStart1 = DateTime.Now;
        foreach (var item in analog)
        {

            try
            {


                if (tablename.ToLower().Equals(Tab.SPEC.name))
                {

                    var text = String.Format($"{item.SystemFields.Id.ToString()}" +
                                            $" {item.GetObjectValue("ROW_NUMBER").ToString()} " +
                                            $" {item.GetObjectValue("SHIFR").ToString()} " +
                                            $"{item.GetObjectValue("IZD").ToString()}");
                    var getlink_connect = item.GetObject(Tab.SPEC.Link.LinkПодключения);
                    if (getlink_connect != null)
                    {
                        var text2 = String.Format($"{getlink_connect.SystemFields.Id.ToString()}" +
                                                  $" {getlink_connect.GetObjectValue("Комплектующая").ToString()} " +
                                                  $"{getlink_connect.GetObjectValue("Сборка").ToString()} " +
                                                  $"{getlink_connect.GetObjectValue("Сводное обозначение").ToString()}");
                        getlink_connect.Delete();
                    }
                  
                    item.Delete();


                }

                //   row_num_list.Add(item.GetObjectValue("ROW_NUMBER").GetValue<int>());
                stringBuilder.Append($"ID {item.Id}  {item.GetObjectValue("ROW_NUMBER").GetValue<int>()} Удалена \n");
            }
            catch
            {
                stringBuilder.Append($"ID {item.Id} {item.GetObjectValue("ROW_NUMBER").GetValue<int>()} Ошибка удаления \n");
            }
        }
        var dataStop1 = DateTime.Now;
        //   Сообщение("", String.Format("Загрузка изменений {0}", (dataStop1 - dataStart1).ToString()));
        stringBuilder.Append(String.Format("Загрузка изменений {0}", (dataStop1 - dataStart1).ToString()));
        stringBuilder.Append("\n");


        Save(stringBuilder, pathnew);

    }

    public void DelobjRefAnalog(Dictionary<int, Dictionary<string, string>> delrow, RefTab refTab)
    {

        StringBuilder stringBuilder = new StringBuilder();
        string pathnew = $"C:\\aemexport\\analog_del_Analog{refTab.table_name}.txt";
        //var delrow_test= new List<int> { 36039, 36065, 36067, 36078, 36137, 36138 };
        //var analog = GetAnalog(delrow_test, refer, parametr);
        var analog = GetAnalog(delrow.Keys.ToList(), refTab.guid, refTab.ROW_NUMBER);


        foreach (var item in analog)
        {
            try
            {
                item.Delete();
            }
            catch
            {
              //  Console.WriteLine("");
            }
        }
    }



    public void DelobjRef(Dictionary<int, Dictionary<string, string>> delrow, RefTab refTab)
    {

        StringBuilder stringBuilder = new StringBuilder();
        string pathnew = $"C:\\aemexport\\analog_del_{refTab.table_name}.txt";
        //var delrow_test= new List<int> { 36039, 36065, 36067, 36078, 36137, 36138 };
        //var analog = GetAnalog(delrow_test, refer, parametr);
        var analog = GetAnalog(delrow.Keys.ToList(), refTab.guid, refTab.ROW_NUMBER);


        //  List<int> row_num_list = new List<int>();
        var dataStart1 = DateTime.Now;
        foreach (var item in analog)
        {

            try
            {


                if (refTab.table_name.ToLower().Equals(Tab.SPEC.name))
                {
                    var getlink_connect = item.GetObject(Tab.SPEC.Link.LinkПодключения);
                    if (getlink_connect != null)
                    {
                        getlink_connect.Delete();
                    }
                    item.Delete();
                }
                else if (refTab.table_name.ToLower().Equals(Tab.NORM.table_name))
                {
                    var getlink_connect = item.GetObject(Tab.NORM.Link.LinkПодключения);
                    if (getlink_connect != null)
                    {
                        getlink_connect.Delete();
                    }

                    item.Delete();
                }
                else
                {
                    item.Delete();
                }
                //   row_num_list.Add(item.GetObjectValue("ROW_NUMBER").GetValue<int>());
                stringBuilder.Append($"ID {item.Id}  {item.GetObjectValue(refTab.name_row_num).GetValue<int>()} Удалена \n");
            }
            catch
            {
                stringBuilder.Append($"ID {item.Id} {item.GetObjectValue(refTab.name_row_num).GetValue<int>()} Ошибка удаления \n");
            }
        }
        var dataStop1 = DateTime.Now;
        //   Сообщение("", String.Format("Загрузка изменений {0}", (dataStop1 - dataStart1).ToString()));
        stringBuilder.Append(String.Format("Загрузка изменений {0}", (dataStop1 - dataStart1).ToString()));
        stringBuilder.Append("\n");


        Save(stringBuilder, pathnew);

    }


    public void DeleteObj(Dictionary<int, Dictionary<string, string>> GetRowDelete, string tablename)
    {
        //  Guid SPEC = new Guid("c587b002-3be2-46de-861b-c551ed92c4c1");
        StringBuilder stringBuilder = new StringBuilder();
        string pathnew = $"C:\\aemexport\\analog_del_{tablename.Replace(".dbf", "")}.txt";


        var AnalogDelete = GetAnalogRefObj(GetRowDelete.Keys.ToList());

        //  ДиалогОжидания.Показать("Пожалуйста, подождите", true);

        int count = AnalogDelete.Count;
        var dataStart1 = DateTime.Now;
        foreach (RefObj itemobj in AnalogDelete)
        {
            try
            {
                itemobj.Delete();
                stringBuilder.Append($"ID {itemobj.Id}  Удалена \n");
            }
            catch
            {
                stringBuilder.Append($"ID {itemobj.Id} Ошибка удаления \n");
            }

        }
        //   ДиалогОжидания.Скрыть();
        var dataStop1 = DateTime.Now;
        //    Сообщение("", String.Format("Загрузка изменений {0}", (dataStop1 - dataStart1).ToString()));
        foreach (var row in GetRowDelete)
        {
            stringBuilder.Append(row.Key + " ");
            //    var obj = CreateObject(SPEC.ToString(), "SPEC");
            foreach (var columnname in row.Value.Keys)
            {
                stringBuilder.Append(row.Value[columnname]);

            }
            //   obj.Save();
            stringBuilder.Append("\n");
            stringBuilder.Append(String.Format("Загрузка изменений {0}", (dataStop1 - dataStart1).ToString()));

            Save(stringBuilder, pathnew);


        }



    }



   

    public void ChangeObj( List<int> ChangeListRowNum, RefTab refTab)
    {
        //   Guid SPEC = new Guid("c587b002-3be2-46de-861b-c551ed92c4c1");
        StringBuilder stringBuilder = new StringBuilder();
        string pathnew = $"C:\\aemexport\\analog_chenge_{refTab.name_refer}.txt";


        //var analogachange = GetAnalogRefObj(RowDbfChangeNew.Keys.ToList(), refTab.name_refer, refTab.name_row_num);

        
        var analogachange = GetAnalogRefObj(ChangeListRowNum, refTab.name_refer, refTab.name_row_num);
        RowDbfChangeNew = GetDbfDataDic($"{pathnewdbf}\\{refTab.table_name}", ChangeListRowNum);
        //RowDbfChangeNew();
        //int count = analogachange.Count;
        //var dataStart1 = DateTime.Now;

        foreach (RefObj itemobj in analogachange)
        {
            var row_num = itemobj[refTab.name_row_num];
            Dictionary<string, string> rowdic = new Dictionary<string, string>();
            if (RowDbfChangeNew.TryGetValue(itemobj[refTab.name_row_num], out rowdic))
            {

                try
                {
              /*      var getlink_connect = item.GetObject(Tab.SPEC.LinkПодключения);
                    if (getlink_connect != null)
                    {
                        var text2 = String.Format($"{getlink_connect.SystemFields.Id.ToString()}" +
                                                  $" {getlink_connect.GetObjectValue("Комплектующая").ToString()} " +
                                                  $"{getlink_connect.GetObjectValue("Сборка").ToString()} " +
                                                  $"{getlink_connect.GetObjectValue("Сводное обозначение").ToString()}");
                        getlink_connect.Delete();
                    }*/

                    itemobj.BeginChanges();
                    //item.GetObject;
                    //itemobj.li
                    string test = itemobj["Номер записи"].ToString();
                    if (refTab.name_refer.Equals("SPEC"))
                    {
                        itemobj.RemoveLink("Список номенклатуры");
                        itemobj.RemoveLink("Список номенклатуры Сборка");
                        itemobj.RemoveLink("Подключения");
                        var trst=itemobj.LinkedObject;
                    }

                    if (refTab.name_refer.Equals("NORM"))
                    {
                       // itemobj.RemoveLink("Список номенклатуры");
                      //  itemobj.RemoveLink("Список номенклатуры Сборка");
                     //   itemobj.RemoveLink("Подключения");
                        var trst = itemobj.LinkedObject["Подключение"];
                        
                    }


                    foreach (var item in rowdic)
                    {

                        //GetTypeRef(item.Key,itemobj.Reference)
                        var type_fild = GetTypeRef(item.Key, refTab.guid);
                        var a = itemobj.Reference;
                        var oo = a.Guid;
                       // Reference ref = Reference()

                        var type = itemobj[item.Key];
                        var ref11 = itemobj.Reference;
                        //var ref12 = ref11
                        var type1 = type.GetType();
                        var type2 = itemobj[item.Key];
                        var type3 = (ReferenceObject)itemobj;




                        //var type31 = type3.item
                        var test4 = type3.ParameterValues;
                        var test5 = type3.GetObjectValue(item.Key);
                        var test6 = test5.Value;

                        if ((item.Value != null) && !item.Value.ToString().Equals(""))
                        {
                            //      itemobj[item.Key] = item.Value;

                            //       if (test6 == null)
                            //----------


                            // TODO перелать условие с GetType возвращает тип DinamicT
                            if (type_fild != null)
                            {
                                // DateTime.Parse("08.04.2033 0:00:00");
                                if (type_fild.LanguageType.Name == typeof(System.DateTime).Name)
                                {

                                    if (item.Value == null || item.Value.ToString().Equals(""))
                                        itemobj[item.Key] = null;
                                    else
                                        itemobj[item.Key] = DateTime.Parse(item.Value);
                                }
                                if (type_fild.LanguageType.Name != typeof(System.DateTime).Name)
                                    itemobj[item.Key] = item.Value;
                            }

                        }
                        else
                        {
                            if (type_fild.LanguageType.Name == typeof(System.String).Name)
                                itemobj[item.Key] = "";
                            else
                                itemobj[item.Key] = null;
                        }  
                        

                    }
                    stringBuilder.Append($"{row_num.ToString()} Изменена \n");

                    itemobj.Save();


                }
                catch
                {
                    stringBuilder.Append($"{row_num.ToString()} Произошла ОШИБКА \n");
                }
                //  var analogachange = GetAnalog(ro_test);
               
                //guidparametr);
            }
            /*else
                stringBuilder.Append("Ошибка " + row_num.ToString());*/
           

        }

        /*      var analog = GetAnalog(RowDbfChangeNew.Keys.ToList(), refTab.guid, refTab.ROW_NUMBER);
              if (refTab.name_refer.Equals("SPEC"))
                  UpgredeLink(analog);
              if (refTab.name_refer.Equals("NORM"))
                  RunImportRole(analog, Tab.RoleChange.RoleConnect);*/
        //var dataStop1 = DateTime.Now;
        //   Сообщение("", String.Format("Загрузка изменений {0}", (dataStop1 - dataStart1).ToString()));
        foreach (var row in RowDbfChangeNew)
        {
            stringBuilder.Append(row.Key + " ");

            foreach (var columnname in row.Value.Keys)
            {
                stringBuilder.Append(row.Value[columnname]);



            }

            stringBuilder.Append("\n");

            Save(stringBuilder, pathnew);


        }


        /*        if (tablename.ToLower().Equals(Tab.SPEC.name))
                {            
                    ОбменДанными.ЭкспортироватьОбъекты("Список номенклатуры SPEC", analogachange, показыватьДиалог: false);
                    ОбменДанными.ЭкспортироватьОбъекты("Подключения SPEC", analogachange, показыватьДиалог: false);

                }*/

    }








    public void ChangeObj(string tablename)
    {
        //   Guid SPEC = new Guid("c587b002-3be2-46de-861b-c551ed92c4c1");
        StringBuilder stringBuilder = new StringBuilder();
        string pathnew = $"C:\\aemexport\\analog_chenge_{tablename.Replace(".dbf", "")}.txt";


        //var analogachange = GetAnalogRefObj(RowDbfChangeNew.Keys.ToList());
        var analogachange = GetAnalogRefObj(RowDbfChangeNew.Keys.ToList(), tablename,"ROW_NUMBER");
        //32965, 35928, 35930, 35931, 35932, 35933

        //var ro_test = new List<int> { 32965, 35928, 35930, 35931, 35932, 35933 };
        //var analogachange = GetAnalogRefObj(ro_test);

        int count = analogachange.Count;
        var dataStart1 = DateTime.Now;

        foreach (RefObj itemobj in analogachange)
        {
            var row_num = itemobj["ROW_NUMBER"];
            Dictionary<string, string> rowdic = new Dictionary<string, string>();
            if (RowDbfChangeNew.TryGetValue(itemobj["ROW_NUMBER"], out rowdic))
            {

                try
                {
                    itemobj.BeginChanges();
                    //item.GetObject;
                    //itemobj.li
                    itemobj.RemoveLink("Список номенклатуры");
                    itemobj.RemoveLink("Список номенклатуры Сборка");
                    itemobj.RemoveLink("Подключения");


                    foreach (var item in rowdic)
                    {
                        itemobj[item.Key] = item.Value;

                    }
                    stringBuilder.Append($"{row_num.ToString()} Изменена \n");

                    itemobj.Save();


                }
                catch
                {
                    stringBuilder.Append($"{row_num.ToString()} Произошла ОШИБКА \n");
                }

            }



        }


        var dataStop1 = DateTime.Now;
        //   Сообщение("", String.Format("Загрузка изменений {0}", (dataStop1 - dataStart1).ToString()));
        foreach (var row in RowDbfChangeNew)
        {
            stringBuilder.Append(row.Key + " ");

            foreach (var columnname in row.Value.Keys)
            {
                stringBuilder.Append(row.Value[columnname]);



            }

            stringBuilder.Append("\n");

            Save(stringBuilder, pathnew);


        }


        /*        if (tablename.ToLower().Equals(Tab.SPEC.name))
                {            
                    ОбменДанными.ЭкспортироватьОбъекты("Список номенклатуры SPEC", analogachange, показыватьДиалог: false);
                    ОбменДанными.ЭкспортироватьОбъекты("Подключения SPEC", analogachange, показыватьДиалог: false);

                }*/

    }
/*
    public void ChangeObject(List<RefObj> analogachange, RefTab refTab)
    {
        RowDbfChangeNew = GetDbfDataDic($"{pathdir}\\{pathnewdbf}\\{refTab.table_name}", ChangeListRowNum);
        foreach (RefObj itemobj in analogachange)
        {
            var row_num = itemobj[refTab.name_row_num];
            Dictionary<string, string> rowdic = new Dictionary<string, string>();
            if (RowDbfChangeNew.TryGetValue(itemobj[refTab.name_row_num], out rowdic))
            {

                try
                {
                    *//*      var getlink_connect = item.GetObject(Tab.SPEC.LinkПодключения);
                          if (getlink_connect != null)
                          {
                              var text2 = String.Format($"{getlink_connect.SystemFields.Id.ToString()}" +
                                                        $" {getlink_connect.GetObjectValue("Комплектующая").ToString()} " +
                                                        $"{getlink_connect.GetObjectValue("Сборка").ToString()} " +
                                                        $"{getlink_connect.GetObjectValue("Сводное обозначение").ToString()}");
                              getlink_connect.Delete();
                          }*//*

                    itemobj.BeginChanges();
                    //item.GetObject;
                    //itemobj.li
                    string test = itemobj["Номер записи"].ToString();
                    if (refTab.name_refer.Equals("SPEC"))
                    {
                        itemobj.RemoveLink("Список номенклатуры");
                        itemobj.RemoveLink("Список номенклатуры Сборка");
                        itemobj.RemoveLink("Подключения");
                        var trst = itemobj.LinkedObject;
                    }

                    if (refTab.name_refer.Equals("NORM"))
                    {
                        // itemobj.RemoveLink("Список номенклатуры");
                        //  itemobj.RemoveLink("Список номенклатуры Сборка");
                        //   itemobj.RemoveLink("Подключения");
                        var trst = itemobj.LinkedObject["Подключение"];

                    }

                    foreach (var item in rowdic)
                    {
                        itemobj[item.Key] = item.Value;

                    }
                    stringBuilder.Append($"{row_num.ToString()} Изменена \n");

                    itemobj.Save();


                }
                catch
                {
                    stringBuilder.Append($"{row_num.ToString()} Произошла ОШИБКА \n");
                }
                //  var analogachange = GetAnalog(ro_test);

                //guidparametr);
            }
            *//*else
                stringBuilder.Append("Ошибка " + row_num.ToString());*//*


        }

        *//*      var analog = GetAnalog(RowDbfChangeNew.Keys.ToList(), refTab.guid, refTab.ROW_NUMBER);
              if (refTab.name_refer.Equals("SPEC"))
                  UpgredeLink(analog);
              if (refTab.name_refer.Equals("NORM"))
                  RunImportRole(analog, Tab.RoleChange.RoleConnect);*//*
        //var dataStop1 = DateTime.Now;
        //   Сообщение("", String.Format("Загрузка изменений {0}", (dataStop1 - dataStart1).ToString()));
        foreach (var row in RowDbfChangeNew)
        {
            stringBuilder.Append(row.Key + " ");

            foreach (var columnname in row.Value.Keys)
            {
                stringBuilder.Append(row.Value[columnname]);



            }

            stringBuilder.Append("\n");

            Save(stringBuilder, pathnew);


        }


        *//*        if (tablename.ToLower().Equals(Tab.SPEC.name))
                {            
                    ОбменДанными.ЭкспортироватьОбъекты("Список номенклатуры SPEC", analogachange, показыватьДиалог: false);
                    ОбменДанными.ЭкспортироватьОбъекты("Подключения SPEC", analogachange, показыватьДиалог: false);

                }*//*

    }*/





    public void SaveChange(RefObj itemobj, Dictionary<string, string> rowdic)
    {
        itemobj.BeginChanges();
        foreach (var item in rowdic)
        {
            itemobj[item.Key] = item.Value;

        }


    }


    public RefObjList GetAnalogRefObj(List<int> row_number)

    {

        string filter = String.Join(", ", row_number);
        // string filter = ("79604, 79646, 79728, 79805, 79882, 79908");
        //79909, 79954, 194897, 204001, 210572, 215851, 217235, 217244, 225416, 225421, 225429, 225430, 225431, 225432, 225433, 225434, 225435, 225552, 225569, 225570, 225572, 225573, 225574, 225575, 225576, 225577, 225578, 225579, 225584, 225585, 225586, 225588, 225590, 225591, 225592, 225595, 225598, 225601, 225605, 225607, 225608, 225609, 225610, 225611, 225613, 225614, 225615, 225616, 225617, 225618, 225619, 225620, 225621, 225622, 225623, 225624, 225625, 225626, 225628, 225629, 225630, 225631, 225632, 225633, 225634, 225637, 225638, 225639, 225641, 225642, 225643, 225644, 225645, 225646, 225648, 225649, 225652, 225655, 225656, 225657, 225658, 225659, 225660, 225661, 225662, 225663, 225664, 225665, 225666, 225667, 225668, 225670, 225671, 225672, 225673, 225674, 225675, 225676, 225677, 225678, 225679, 225681, 225682, 225683, 225684, 225685, 225736, 228076, 228081, 230300, 232886, 232896, 232897, 232901, 232903, 233758, 237147, 238663, 240687, 240695, 240698, 240699, 240938, 241892, 241948, 243331, 243432, 243993, 245235, 255187, 259954, 261365, 261367, 261368, 261369, 261370, 261371, 261372, 261373, 261374, 261375, 261376, 261377, 261378, 261379, 261380, 261381, 261384, 261385, 261386, 261387, 261388, 261523, 261564, 261602, 261609, 261656, 261827
        var result = FindObjects("SPEC",
        String.Format("[ROW_NUMBER] Входит в список '{0}'", filter));

        return result;
    }


    public RefObjList GetAnalogRefObj(List<int> row_number,string tablename,string param)

    {

        string filter = String.Join(", ", row_number);
        // string filter = ("79604, 79646, 79728, 79805, 79882, 79908");
        //79909, 79954, 194897, 204001, 210572, 215851, 217235, 217244, 225416, 225421, 225429, 225430, 225431, 225432, 225433, 225434, 225435, 225552, 225569, 225570, 225572, 225573, 225574, 225575, 225576, 225577, 225578, 225579, 225584, 225585, 225586, 225588, 225590, 225591, 225592, 225595, 225598, 225601, 225605, 225607, 225608, 225609, 225610, 225611, 225613, 225614, 225615, 225616, 225617, 225618, 225619, 225620, 225621, 225622, 225623, 225624, 225625, 225626, 225628, 225629, 225630, 225631, 225632, 225633, 225634, 225637, 225638, 225639, 225641, 225642, 225643, 225644, 225645, 225646, 225648, 225649, 225652, 225655, 225656, 225657, 225658, 225659, 225660, 225661, 225662, 225663, 225664, 225665, 225666, 225667, 225668, 225670, 225671, 225672, 225673, 225674, 225675, 225676, 225677, 225678, 225679, 225681, 225682, 225683, 225684, 225685, 225736, 228076, 228081, 230300, 232886, 232896, 232897, 232901, 232903, 233758, 237147, 238663, 240687, 240695, 240698, 240699, 240938, 241892, 241948, 243331, 243432, 243993, 245235, 255187, 259954, 261365, 261367, 261368, 261369, 261370, 261371, 261372, 261373, 261374, 261375, 261376, 261377, 261378, 261379, 261380, 261381, 261384, 261385, 261386, 261387, 261388, 261523, 261564, 261602, 261609, 261656, 261827
        var result = FindObjects(tablename,
        String.Format("[{0}] Входит в список '{1}'",param, filter));

        return result;
    }





    /// <summary>
    /// Ищет объекты в справочнике у которых нет ссылки
    /// </summary>
    /// <param>
    /// guidref - справочник в котором ищем
    /// parametr - guid связи
    /// </param>
    /// <returns>
    /// Возвращает список объектов без связи
    /// </returns>




    public List<ReferenceObject> FilterReference(Guid guidref, Guid parametr)

    {

        // Guid Guid_ROW_NUMBER = new Guid("b6c7f1e2-d94c-4fde-966b-937a9979813f");
        //  Guid Guid_SHIFR = new Guid("2a855e3d-b00a-419f-bf6f-7f113c4d62a0");
        //    StringBuilder stringBuilder = new StringBuilder();
        // string referenceGuid = "c587b002-3be2-46de-861b-c551ed92c4c1";
        /*  ReferenceObject LinkПодключенияObj = item.Links.ToOne[Tab.SPEC.LinkПодключения].LinkedObject;
          ReferenceObject LinkСписокНоменклатурыObj = item.Links.ToOne[Tab.SPEC.LinkСписокНоменклатуры].LinkedObject;
          ReferenceObject LinkСписокНоменклатурыObjСборка = item.Links.ToOne[Tab.SPEC.LinkСписокНоменклатурыСборка].LinkedObject;*/

        ReferenceInfo info = Context.Connection.ReferenceCatalog.Find(guidref);
        Reference reference = info.CreateReference();
        //  ReferenceObjectCollection listObj = reference.Objects;
        //  ParameterInfo parameterInfo = reference.ParameterGroup[parametr];

        List<ReferenceObject> result = new List<ReferenceObject>();
        using (Filter filter = new Filter(info))
        {
            filter.Terms.AddTerm(parametr.ToString(), ComparisonOperator.IsNull, null);
            result = reference.Find(filter);
            var count = result.Count();
        }



        /////
        //        List<ReferenceObject> result = reference.Find(Tab.SPEC.LinkПодключения, ComparisonOperator.IsNull, null);
        //       var count2 = result.Count();

        Console.WriteLine("GetAnalog");

        return result;
    }




   



    /// <summary>
    /// Фильтр по списку ROW_NUMBER
    /// </summary>
   
    public List<ReferenceObject> GetAnalog(List<int> row_number, Guid guidref, Guid parametr)

    {
        ReferenceInfo info = Context.Connection.ReferenceCatalog.Find(guidref);
        Reference reference = info.CreateReference();
        ParameterInfo parameterInfo = reference.ParameterGroup[parametr];
        List<ReferenceObject> result = reference.Find(parameterInfo, ComparisonOperator.IsOneOf, row_number);
        Console.WriteLine("GetAnalog");
        return result;
    }

    /// <summary>
    /// Фильтр по списку строке
    /// </summary>

    public ParameterType GetTypeRef()

    {
        var fobject = FindObject("TRUD", "SHIFR", "УЯИС798133021-ВАР3");
        var rr = (ReferenceObject)fobject;
        var classs1 = rr.Class;
        ParameterType tip=null;
        foreach (var param_item in classs1.ParameterGroups[0].Parameters)
        {
            if (param_item.FieldName=="SHIFR")
                tip = param_item.Type;
        }

        //var param1 = classs1.ParameterGroups[0].Parameters[0].FieldName;
        //var param1 = classs1.ParameterGroups[0].Parameters[0].Type;
        return tip;
    }


    public ParameterType GetTypeRef(string field, Guid guidref)

    {

        ReferenceInfo info = Context.Connection.ReferenceCatalog.Find(guidref);
        Reference reference = info.CreateReference();
        ParameterInfo parameterInfo = reference.ParameterGroup[field];
        ParameterType tip = null;
        if (parameterInfo != null)
        {
            tip = parameterInfo.Type;
        }
        /*foreach (var param_item in refer.ParameterGroups[0].Parameters)
        {
            if (param_item.FieldName == field)
                tip = param_item.Type;
        }
        */

        return tip;
    }

    public ParameterType GetTypeRef(Guid guidref, Guid parametr)

    {

        // RefObj itemobj = new RefObj()

        //var folder = FindObject(Settings.СправочникНоменклатура.ToString(), "Наименование", "Импорт");
        // folder = CreateObject(Settings.СправочникНоменклатура.ToString(), "Папка");
        var fobject = FindObject("TRUD","SHIFR", "УЯИС798133021-ВАР3");
        var rr = (ReferenceObject)fobject;
        var rr2 = rr.Reference;
        var rr3 = rr2.Classes;
        var rr4 = rr.ParameterValues;
        var rr5 = rr.GetObjectValue("SHIFR");

        //var rr511 = rr5.Parameter;
        var rr51 = rr5.Value;
        //var rr52 = rr5.GetTypeCode();
        //var rr51 = rr5.Value;
        var rr6 = rr5.GetType();
        //var rr7 = rr5.Pa;

        var classs = fobject.Class;

        var classs1 = rr.Class;
        var param = classs.Parameters;
        var param1 = classs1.ParameterGroups[0].Parameters[0].FieldName;
       // var param1 = classs1.ParameterGroups[0].Parameters[0].Type;
        var param11 = classs1.ParameterGroups;
        var param12 = classs1.ParameterGroups;
        var param2 = classs1.RequiredParameters;
        var param3 = classs.Parameters[0];

        //var param = classs.Parameters
        //var fdw = rr.Load(param) ;
        //var param3 = 
        //var equipmentName = FindParameterInfo(guidref, Settings.НаименованиеОснащения);
        //var p = param2.
        //Reference

        ReferenceInfo info = Context.Connection.ReferenceCatalog.Find(guidref);
        Reference reference = info.CreateReference();
        ParameterInfo parameterInfo = reference.ParameterGroup[parametr];
       // ParameterInfo parameterInfo = reference.ParameterGroup;
      //ParameterInfo parameterInfo2 = reference.pa
        var tip = parameterInfo.Type;
        //List<ReferenceObject> result = reference.Find(parameterInfo, ComparisonOperator.IsOneOf, row_number);
        //Console.WriteLine("GetAnalog");
        return tip;
    }


    /// <summary>
    /// Получает объекты справочника по условию если parametr содердит str
    /// </summary>
    public List<ReferenceObject> GetAnalog(String str, Guid guidref, Guid parametr)

    {
        ReferenceInfo info = Context.Connection.ReferenceCatalog.Find(guidref);
        Reference reference = info.CreateReference();
        ParameterInfo parameterInfo = reference.ParameterGroup[parametr];
        List<ReferenceObject> result = reference.Find(parameterInfo, ComparisonOperator.Equal, str);
        Console.WriteLine("GetAnalog");
        return result;
    }

    /// <summary>
    /// Получает все объекты справочника
    /// </summary>
    /// 
    public List<ReferenceObject> GetAnalog(Guid guidref)
    {
        ReferenceInfo info = Context.Connection.ReferenceCatalog.Find(guidref);
        Reference reference = info.CreateReference();
        List<ReferenceObject> result = reference.Objects.ToList();
        return result;
    }


    public Dictionary<int, string> ReftoDict(List<ReferenceObject> AllRefObjects)
    {
        Dictionary<int, string> DictRef = new Dictionary<int, string>();
        foreach (var obj in AllRefObjects)
        {
            string IZD = obj.GetObjectValue("IZD").Value.ToString();
            string NAIM = obj.GetObjectValue("NAIM").Value.ToString();
            string POS = obj.GetObjectValue("POS").Value.ToString();
            string PRIM = obj.GetObjectValue("PRIM").Value.ToString();
            string SHIFR = obj.GetObjectValue("SHIFR").Value.ToString();
            DictRef.Add((int)obj.GetObjectValue("ROW_NUMBER").Value, String.Format($"{POS}{SHIFR}{NAIM}{PRIM}{IZD}"));
        }
        return DictRef;
    }


    /// <summary>
    /// Групирует объекты справочника Spec по полю SHIFR
    /// </summary>
    public IEnumerable<IGrouping<string, ReferenceObject>> GroupAnalog(IEnumerable<ReferenceObject> analog)
    {
        StringBuilder s = new StringBuilder();
        ReferenceInfo info = Context.Connection.ReferenceCatalog.Find(Tab.SPEC.guid);
        Reference reference = info.CreateReference();
        ReferenceObjectCollection listObj = reference.Objects;
        ParameterInfo parameterInfo = reference.ParameterGroup[Tab.SPEC.SHIFR];

        //var test = analog.Where(s => s[parameterInfo].ToString().Equals("8А6425149"));

        var groupanalog = analog.GroupBy(s => s[parameterInfo].ToString());
        //Where(s => s[parameterInfo].ToString().Equals("8А6425149"));

        /*  foreach (var iterator in test)
                  {
                      s.Append(iterator.Key.ToString());
                  };
          var a  = test.
          //var test2 = test.ToList();
          Message("",s.ToString());*/

        //{ 60604,60529,60749,60823,60677,9264,9276,9288,9190};
        // 8А6873058 и 8А6425149
        //obj[Подключения_Обозначение_Guid].GetString()
        return groupanalog;
    }

    /// <summary>
    /// Получает родителя
    /// </summary>
    public void GetParent(List<int> row_number)
    {

        //List<int> row_number = new List<int> { 25637, 225638, 225639, 225641 };
        IEnumerable<ReferenceObject> analog = GetAnalog(row_number, Tab.SPEC.guid, Tab.SPEC.ROW_NUMBER);

        Console.WriteLine(analog.Count());

        foreach (var item in analog)
        {
            RecursSpec(item.GetObjectValue("SHIFR").Value.ToString());
        }
    }


    public void RecursSpec(string item)
    {

        ReferenceInfo info = Context.Connection.ReferenceCatalog.Find(Tab.SPEC.guid);
        Reference reference = info.CreateReference();

        using (Filter filter = new Filter(reference.ParameterGroup.ReferenceInfo))
        {
            filter.Terms.AddTerm(reference.ParameterGroup[Tab.SPEC.SHIFR], ComparisonOperator.Equal, item);
            liststruct = reference.Find(filter);
        }




        if (liststruct.Count() > 0)
        {
            if (liststruct[0].GetObjectValue("POS").Value.ToString().Equals("0"))
            {

                if (!RowNumParent.Contains(liststruct[0][Tab.SPEC.ROW_NUMBER].GetString()))
                    RowNumParent.Add($"{liststruct[0][Tab.SPEC.ROW_NUMBER].GetString()}");
            }
            RecursSpec(liststruct[0][Tab.SPEC.IZD].GetString());
        }


    }


    public Dictionary<int,string> ConnectAnalog(RefTab refTab)
     {
    Dictionary<int, string> analogtable = new Dictionary<int, string>();
    string Host = "tflex-test";//"localhost";
    string User = "analog_foxpro_read"; //"analog_foxpro_read";
    string DBname = "tflex-analog-import-dbf";
    string Password = "tdocs528"; //"postgres529"; //"tdocs528";
    string Port = "5449"; //"5433";


    string connString =
    String.Format(
        "Server={0}; User Id={1}; Database={2}; Port={3}; Password={4};SSLMode=Prefer",
        Host,
        User,
        DBname,
        Port,
        Password);

        using (var conn = new NpgsqlConnection(connString))
        {

            Console.Out.WriteLine("Opening connection");
            conn.Open();

            //dbo."SPEC"
            using (var command = new NpgsqlCommand(refTab.quertystr, conn))
            {
                string str="";
                var reader = command.ExecuteReader();
                //  /*   
                while (reader.Read())
                {
                    //int a= int.Parse()
                    //analogtable.Add(int.Parse(reader.GetValue(0).ToString()),String.Format("{0}",reader.GetValue(1).ToString()));                    
                    int count = reader.FieldCount;
                    str = "";
                    for (int i = 1; i < count; i++)
                    {
                        var testt = reader.GetFieldType(i);
                        if (reader.GetFieldType(i) == typeof(System.Decimal))
                        {
                            var t = reader.GetDecimal(i);
                            if (reader.GetDecimal(i) == 0)
                                str += "0";
                            else
                            {
                                //Decimal.Parse("3.211");
                                //str += Decimal.Round(reader.GetDecimal(i), 3).ToString();
                                str += (Double)Decimal.Round(Decimal.Parse(reader.GetValue(i).ToString()), 3);

                            }
                        }
                       
                        else if (reader.GetFieldType(i) == typeof(System.DateTime))
                        {
                           // var test0001 = reader.GetValue(i).ToString();
                          //  var test0002 = reader.GetValue(i).ToString().Length;
                            if (reader.GetValue(i) == null || reader.GetValue(i).ToString().Equals(""))
                                str += "";
                            else
                                str += reader.GetDateTime(i).ToString();
                        }


                        else
                            str += reader.GetValue(i).ToString();
                        str += "|";
                    }
                    analogtable.Add(int.Parse(reader.GetValue(0).ToString()), str);
                }

                //analogtable.(int.Parse(reader.GetValue(0).ToString());
                //  */

                /*var i = reader.FieldCount;
                reader.GetFieldValue
                reader.Read();
                reader.Rows
                reader.Close();*/
            }
        }
        SaveCsv(analogtable, $"{ pathdir}\\{refTab.name_refer}_analog.csv");
        return analogtable;
    }



    /// <summary>
    /// Сравнение dbf
    /// </summary>
    public void Compare(string tablename, bool analogSql = false, bool savecsv = false)
    {
        //++
        // Compare($"{pathdir}\\{patholddbf}\\{tablename}", $"{pathdir}\\{pathnewdbf}\\{tablename}");
        string olddata = $"{patholddbf}\\{tablename}";
        string newdata = $"{pathnewdbf}\\{tablename}";
        string log = $"{pathdir}\\";
        /*   Dictionary<int, string> dbf_old_dict;
           Dictionary<int, string> dbf_new_dict;*/

        ///*        
       // analog_old_dict
        dbf_old_dict = null;
        dbf_new_dict = null;

        RowDbfChangeNew = null;
        RowDbfChangeOld = null;
        RowDbfDel = null;
        RowDbfDel_all = null;
        RowDbfAdd = null;
//*/


        List<int> change = new List<int>();
        List<int> delrow = new List<int>();
        List<int> delrow_old = new List<int>();  // все удаленные записи в старой таблице
        //  List<int> delrowold = new List<int>();
        List<int> addRow = new List<int>();

        if (analogSql)
            dbf_old_dict = analog_old_dict;  //данные берем из DOCS
        else
            dbf_old_dict = GetDbfData(olddata); // данные из dbf
        
        dbf_new_dict = GetDbfData(newdata);



        int rowcount = (dbf_old_dict.Keys.Max() < dbf_new_dict.Keys.Max()) ? dbf_old_dict.Keys.Max() : dbf_new_dict.Keys.Max();
        for (int i = 1; i <= rowcount; i++)
        {
            if (dbf_old_dict.ContainsKey(i))
            {
                if (!dbf_old_dict[i].Equals(dbf_new_dict[i]) && !dbf_new_dict[i].Equals("Delete"))
                    change.Add(i);
                if (dbf_new_dict[i].Equals("Delete"))
                    delrow_old.Add(i);
                if (dbf_new_dict[i].Equals("Delete") && !dbf_new_dict[i].Equals(dbf_old_dict[i]))
                    delrow.Add(i);
            }
            else if (!dbf_new_dict[i].Equals("Delete"))
            {
                addRow.Add(i);
            }
        }


        for (int i = rowcount; i <= dbf_new_dict.Count; i++)
        {
            if (dbf_new_dict[i].Equals("Delete"))
                delrow.Add(i);
        }





        int maxrow = dbf_old_dict.Keys.Max();



        foreach (var item in dbf_new_dict)
        {
            if (item.Key > maxrow && !item.Value.Equals("Delete"))
                addRow.Add(item.Key);
        }

        StringBuilder stringBuilder = new StringBuilder();


        foreach (var item in addRow)
        {
            stringBuilder.Append(item + " ");
        }


        Save(stringBuilder, log + @$"\{tablename.Replace(".dbf", "")}_add_dictdbf0.txt");

        // измененач запись новая версия
        RowDbfChangeNew = GetDbfDataDic(newdata, change);
        //
        Save(RowDbfChangeNew, log + @$"\{tablename.Replace(".dbf", "")}_change_dictdbf_new.txt");

        // измененач запись старая версия
        RowDbfChangeOld = GetDbfDataDic(olddata, change);
        //
        Save(RowDbfChangeOld, log + @$"\{tablename.Replace(".dbf", "")}_change_dictdbf_old.txt");

        // удалення запись 
        RowDbfDel = GetDbfDataDic(newdata, delrow);
        RowDbfDel_all = GetDbfDataDic(olddata, delrow_old);
        //
        Save(RowDbfDel, log + @$"\{tablename.Replace(".dbf", "")}_del_dictdbf.txt");

        // навые записи 
        RowDbfAdd = GetDbfDataDic(newdata, addRow);
        //
        Save(RowDbfAdd, log + @$"\{tablename.Replace(".dbf", "")}_add_dictdbf.txt");


        Save(dbf_old_dict, dbf_new_dict, log + @$"\{tablename.Replace(".dbf", "")}_change_compare_2dbf.txt");

        // получение всех записей из dbf для сохранения в csv
        if (savecsv)
        {
            var RowDbfNew = GetDbfDataDic(newdata, dbf_new_dict.Keys.ToList());
            SaveCsv(RowDbfNew, log + @$"\{tablename.Replace(".dbf", "")}.csv");
        }
        ///*

        /*
        Save(dbf_old_dict, dbf_new_dict, log + @$"\{tablename.Replace(".dbf", "")}_change_compare_2dbf.txt");

        GetParent(change);
        Save(RowNumParent, log + @$"\{tablename.Replace(".dbf", "")}_izdeliechange.txt");
        GetParent(addRow);
        Save(RowNumParent, log + @$"\{tablename.Replace(".dbf", "")}_izdelieadd.txt");
        */

        //*/

    }





    /// <summary>
    /// Для переданых row_num создает словарь из dbf 
    /// </summary>
    static Dictionary<int, Dictionary<string, string>> GetDbfDataDic(string path, List<int> get_row_num)
    {
        var NumberDecimalSeparator = CultureInfo.InvariantCulture.NumberFormat.NumberDecimalSeparator;
        var options = new DbfDataReaderOptions
        {
            SkipDeletedRecords = false,
            Encoding = Encoding.GetEncoding("windows-1251")

        };

        //var skipDeleted = true;
        //List<string> list1 = new List<string>();

        Dictionary<int, Dictionary<string, string>> gettable = new Dictionary<int, Dictionary<string, string>>();

        var dbfPath = path;

        //Encoding encoding = Encoding.GetEncoding(866);
        //Encoding encoding = Encoding.GetEncoding("windows-1251");
        using (var dbfDataReader = new DbfDataReader.DbfDataReader(dbfPath, options))
        {
            int rownum = 1;

            
            //var dbfRecord = new DbfRecord(dbfTable);

            while (dbfDataReader.Read())
            {
                Dictionary<string, string> row = new Dictionary<string, string>();
                //dbfRecord
                //string rowstring = "";


                if (get_row_num.Contains(rownum))
                {

                    for (int i2 = 0; i2 < dbfDataReader.FieldCount; i2++)
                    {
                        if (dbfDataReader.DbfTable.Columns[i2].DataType == typeof(System.String))
                        {
                            if (dbfDataReader.GetValue(i2) != null && !dbfDataReader.GetValue(i2).ToString().Equals(""))
                                row.Add(dbfDataReader.DbfTable.Columns[i2].ColumnName.ToString(), dbfDataReader.GetValue(i2).ToString());
                            else
                                row.Add(dbfDataReader.DbfTable.Columns[i2].ColumnName.ToString(), "");
                        }
                        if (dbfDataReader.DbfTable.Columns[i2].DataType == typeof(System.Int32))
                        {
                            if (dbfDataReader.GetValue(i2) != null)
                                row.Add(dbfDataReader.DbfTable.Columns[i2].ColumnName.ToString(), dbfDataReader.GetValue(i2).ToString());
                            else
                                row.Add(dbfDataReader.DbfTable.Columns[i2].ColumnName.ToString(), "0");
                        }

                        if (dbfDataReader.DbfTable.Columns[i2].DataType == typeof(System.Decimal))
                        {
                            if (dbfDataReader.GetValue(i2) != null)
                                row.Add(dbfDataReader.DbfTable.Columns[i2].ColumnName.ToString(), dbfDataReader.GetValue(i2).ToString().Replace(",", NumberDecimalSeparator));
                            else
                                row.Add(dbfDataReader.DbfTable.Columns[i2].ColumnName.ToString(), "0");
                        }

                        if (dbfDataReader.DbfTable.Columns[i2].DataType == typeof(System.DateTime))
                        {
                            if (dbfDataReader.GetValue(i2) != null)
                                row.Add(dbfDataReader.DbfTable.Columns[i2].ColumnName.ToString(), dbfDataReader.GetValue(i2).ToString());
                            else
                                row.Add(dbfDataReader.DbfTable.Columns[i2].ColumnName.ToString(), "");
                        }
                    }
                    gettable.Add(rownum, row);
                }



                //row = null;
                rownum++;
               
            }

        }
        return gettable;

        //
    }


   /* public void getdbf(string dbfPath)
    {
        //DbDataReader dbfDataReader= new     

        *//*
        var options = new DbfDataReaderOptions
        {
            SkipDeletedRecords = true
            // Encoding = EncodingProvider.GetEncoding(1252);
        };
        //var dbfPath = "path\\file.dbf";
        Encoding encoding = Encoding.GetEncoding("windows-1251");

        using (var dbfDataReader = new DbfDataReader.DbfDataReader(dbfPath, encoding))
        {
            while (dbfDataReader.Read())
            {
               var colum1 = dbfDataReader["SHIFR"];
                var colum2 =  dbfDataReader["OKP"];
            }
        }
               *//*



        var options = new DbfDataReaderOptions
        {
            SkipDeletedRecords = true
            // Encoding = EncodingProvider.GetEncoding(1252);
        };


        using (var dbfDataReader = new DbfDataReader.DbfDataReader(dbfPath, options))
        {
            while (dbfDataReader.Read())
            {
                var colum1 = dbfDataReader["SHIFR"];
                var colum2 = dbfDataReader["OKP"];
                //var valueCol1 = dbfDataReader.GetString(0);
                //var valueCol2 = dbfDataReader.GetDecimal(1);
                //var valueCol3 = dbfDataReader.GetDateTime(2);
                //var valueCol4 = dbfDataReader.GetInt32(3);
            }
        }
    }*/


    /// <summary>
    /// Получает данные из dbf в виде строки для сравнения
    /// </summary>
    
    static Dictionary<int, string> GetDbfData(string path)
    {
        var options = new DbfDataReaderOptions
        {
            SkipDeletedRecords = false,
            Encoding = Encoding.GetEncoding("windows-1251")
        };
        //var skipDeleted = true;
        //List<string> list1 = new List<string>();

        Dictionary<int, string> dic1 = new Dictionary<int, string>();
        var dbfPath = path;
        //Encoding encoding = Encoding.GetEncoding(866);
        //Encoding encoding = Encoding.GetEncoding("windows-1251");
        using (var dbfDataReader = new DbfDataReader.DbfDataReader(dbfPath, options))
        {
            int rownum = 1;

            while (dbfDataReader.Read())
            {

                string rowstring = "";


                for (int i = 0; i < dbfDataReader.FieldCount; i++)
                {

                    if (!dbfDataReader.DbfRecord.IsDeleted)
                    {
                        if (dbfDataReader.DbfTable.Columns[i].DataType == typeof(System.String))
                            if (dbfDataReader.GetValue(i) == null)
                                rowstring += "0";
                            else
                                rowstring += dbfDataReader.GetValue(i).ToString();
                        if (dbfDataReader.DbfTable.Columns[i].DataType == typeof(System.Int32))
                            if (dbfDataReader.GetValue(i) == null)
                                rowstring += "0";
                            else
                                rowstring += dbfDataReader.GetInt32(i).ToString();
                        if (dbfDataReader.DbfTable.Columns[i].DataType == typeof(System.Decimal))
                            if (dbfDataReader.GetValue(i) == null)
                                rowstring += "0";
                            else if (dbfDataReader.GetDecimal(i) == 0)
                                rowstring += "0";
                            else
                                rowstring += (Double)Decimal.Round(dbfDataReader.GetDecimal(i), 3);
                     
                            if (dbfDataReader.DbfTable.Columns[i].DataType == typeof(System.DateTime))
                            {
                            //    var test0001 = dbfDataReader.GetValue(i).ToString();
                             //   var test0002 = dbfDataReader.GetValue(i).ToString().Length;
                                if (dbfDataReader.GetValue(i) == null || dbfDataReader.GetValue(i).ToString().Equals(""))
                                    rowstring += "";
                                else
                                    rowstring += dbfDataReader.GetDateTime(i).ToString();
                            }
                            rowstring += "|";
                   
                        /*if (dbfDataReader.GetValue(i) != null)
                        rowstring += dbfDataReader.GetValue(i).ToString();
                    if (dbfDataReader.GetValue(i) == null)
                        rowstring += "";*/
                    }
                    else
                    {
                        rowstring = "Delete";
                    }

                }
                dic1.Add(rownum, rowstring);
                if (rownum == 4762)
                    Console.WriteLine("1111");
                rownum++;
            }
        }
        return dic1;
    }


    static Dictionary<int, string> GetDbfData2(string path)
    {
        var options = new DbfDataReaderOptions
        {
            SkipDeletedRecords = false,
            Encoding = Encoding.GetEncoding("windows-1251")
        };
        //var skipDeleted = true;
        //List<string> list1 = new List<string>();

        Dictionary<int, string> dic1 = new Dictionary<int, string>();
        var dbfPath = path;
        //Encoding encoding = Encoding.GetEncoding(866);
        //Encoding encoding = Encoding.GetEncoding("windows-1251");
        using (var dbfDataReader = new DbfDataReader.DbfDataReader(dbfPath, options))
        {
            int rownum = 1;

            while (dbfDataReader.Read())
            {

                string rowstring = "";


                for (int i = 0; i < dbfDataReader.FieldCount; i++)
                {

                    if (!dbfDataReader.DbfRecord.IsDeleted)
                    {
                        if (dbfDataReader.GetValue(i) != null)
                            rowstring += dbfDataReader.GetValue(i).ToString();
                        if (dbfDataReader.GetValue(i) == null)
                            rowstring += "";
                    }
                    else
                    {
                        rowstring = "Delete";
                    }

                }
                dic1.Add(rownum, rowstring);
                rownum++;
            }
        }
        return dic1;
    }



    /// <summary>
    /// Проставляе точки в обозначении
    /// </summary>
    public String ShifrInsertDot(string shifr)
    {

        int num;
        bool isNum = int.TryParse(shifr, out num);
        if (!isNum && shifr.IndexOf('-') > 0 && (shifr.Substring(0, shifr.IndexOf('-')).Length == 6 || shifr.Substring(0, shifr.IndexOf('-')).Length == 7))
        {
            isNum = int.TryParse(shifr.Substring(0, shifr.IndexOf('-')), out num);
            shifr = shifr.Insert(3, ".");
        }

        if ((shifr.StartsWith("УЯИС") || shifr.StartsWith("ШЖИФ") || shifr.StartsWith("УЖИЯ")) && shifr.Length > 12)
        {
            // Изменяем значение
            shifr = shifr.Insert(4, ".").Insert(11, ".");
        }
        if (shifr.StartsWith("ЭСКИЗ") && shifr.Length > 13)
        {
            shifr = shifr.Insert(5, ".").Insert(9, ".").Insert(13, ".");
        }


        if (shifr.StartsWith("8А") && shifr.Length > 7)
        {
            // Изменяем значение
            shifr = shifr.Insert(3, ".").Insert(7, ".");
        }

        if (isNum && (shifr.Length == 6 || shifr.Length == 7))
        {

            shifr = shifr.Insert(3, ".");
        }



        return shifr;

    }

    /// <summary>
    /// Проверяет наличие ссылки у объекта
    /// </summary>
    public void CheckLink(Guid guidtab, Guid guidparam, Dictionary<int, Dictionary<string, string>> dict)
    {

        var checkobj = new Dictionary<int, Dictionary<string, string>>(dict);
        StringBuilder str = new StringBuilder();
        //SHIFR
        IEnumerable<ReferenceObject> analog = GetAnalog(dict.Keys.ToList(), guidtab, guidparam);



        foreach (var item in analog)
        {


            ReferenceObject LinkПодключенияObj = item.Links.ToOne[Tab.SPEC.Link.LinkПодключения].LinkedObject;
            ReferenceObject LinkСписокНоменклатурыObj = item.Links.ToOne[Tab.SPEC.Link.LinkСписокНоменклатуры].LinkedObject;
            ReferenceObject LinkСписокНоменклатурыObjСборка = item.Links.ToOne[Tab.SPEC.Link.LinkСписокНоменклатурыСборка].LinkedObject;
            var shifr = item.GetObjectValue("SHIFR").Value.ToString();
            var row_num = item.GetObjectValue("ROW_NUMBER").Value;


            if (checkobj.ContainsKey((int)row_num))
                checkobj.Remove((int)row_num);




            if (LinkПодключенияObj == null)
                str.Append(String.Format($" У объекта {row_num.ToString()} " +
                                                    $"{shifr} " +
                                                    $"отсутствует связь на подключения\r\n"));

            if (LinkСписокНоменклатурыObj == null)
                str.Append(String.Format($" У объекта {row_num.ToString()}" +
                                                   $" {shifr} " +
                                                   $"отсутствует связь на список номенклатуры\r\n"));

            if (LinkСписокНоменклатурыObjСборка == null)
                str.Append(String.Format($" У объекта {row_num.ToString()} " +
                                                   $" {shifr} " +
                                                   $"отсутствует связь на список номенклатуры сборка \r\n"));


        }

        foreach (var iterator in checkobj)
            str.Append(String.Format($" Объект  {iterator.Key} {iterator.Value["SHIFR"]}  не найден в справочнике SPEC \r\n"));

        Save(str, @"C:\aemexport\checklink.txt");
    }




    

    /// <summary>
    /// Сохранение StringBilder в файл
    /// </summary>
    public static void Save(StringBuilder stringBuilder, string pathfaile)
    {



        try
        {
            //System.Text.UTF8Encoding(false) System.Text.Encoding.UTF8
            //new System.Text.UTF8Encoding(false);
            using (StreamWriter sw = new StreamWriter(pathfaile, false, new System.Text.UTF8Encoding(false)))
            {
                sw.Write(stringBuilder.ToString());
            }


        }
        catch (Exception e)
        {
            Console.WriteLine(e.Message);
        }
    }

    /// <summary>
    /// Сохранение List в файл
    /// </summary>
    public static void Save(List<string> strList, string pathfaile)
    {


        //string text = "text";
        try
        {
            using (StreamWriter sw = new StreamWriter(pathfaile, false, System.Text.Encoding.Default))
            {
                foreach (var str in strList)
                    sw.WriteLine(str.ToString());
            }


        }
        catch (Exception e)
        {
            Console.WriteLine(e.Message);
        }
    }

    
    /// <summary>
    /// Сохраняет Dictionary список измененных записей
    /// </summary>
    public static void Save(Dictionary<int, string> dic_old, Dictionary<int, string> dic_new, string file)
    {
        var change = RowDbfChangeNew.Keys.ToList();
        var delrow = RowDbfDel.Keys.ToList();


        //string file = 

        try
        {
            using (StreamWriter sw = new StreamWriter(file, false, System.Text.Encoding.UTF8))
            {
                foreach (var item in change)
                    sw.WriteLine(String.Format("{0}  {1} - {2}", item.ToString() + " ", dic_old[item], dic_new[item]));

                sw.WriteLine("\nИзмененные записи count:" + change.Count + "\n");
                foreach (var item in change)
                    sw.Write(item.ToString() + " ");

                sw.WriteLine("\nУдаленные записи count:" + delrow.Count + "\n");
                foreach (var item in delrow)
                    sw.Write(item.ToString() + " ");

                sw.WriteLine("\nколличество записей было :" + dic_old.Max().Key);
                sw.WriteLine("\nколличество записей cтало :" + dic_new.Max().Key);
            }


            Console.WriteLine("Запись выполнена");
        }
        catch (Exception e)
        {
            Console.WriteLine(e.Message);
        }

    }



    /// <summary>
    /// Сохранения словаря
    /// </summary>
    public static void Save(Dictionary<int, Dictionary<string, string>> dicttable, string pathsave)
    {
        StringBuilder strBuilder = new StringBuilder();
        foreach (var row in dicttable)
        {
            strBuilder.Append(row.Key + " ");
            foreach (var columnname in row.Value.Keys)
            {
                strBuilder.Append(row.Value[columnname]);
            }
            strBuilder.Append("\n");
        }

        strBuilder.Append(String.Join(", ", dicttable.Keys.ToList()));
        strBuilder.Append($"Количество: {dicttable.Count()}");

        Save(strBuilder, pathsave);
    }


    public static void Save(Dictionary<int, string> dicttable, string pathsave)
    {
        StringBuilder strBuilder = new StringBuilder();
        foreach (var row in dicttable)
        {
            strBuilder.Append(row.Key + " " + row.Value);
            
            strBuilder.Append("\n");
        }

        strBuilder.Append(String.Join(", ", dicttable.Keys.ToList()));
        strBuilder.Append($"Количество: {dicttable.Count()}");

        Save(strBuilder, pathsave);
    }

    /// <summary>
    /// Сохранения dbf в csv
    /// </summary>
    public static void SaveCsv(Dictionary<int, Dictionary<string, string>> dicttable, string pathsave)
    {
        StringBuilder strBuilder = new StringBuilder();
        foreach (var row in dicttable)
        {
            if (strBuilder.Length != 0)
                strBuilder.Append("\n");

            strBuilder.Append(row.Key);
            foreach (var columnname in row.Value.Keys)
                strBuilder.Append("|" + row.Value[columnname].ToString());                        
        }
        Save(strBuilder, pathsave);
    }


    public static void SaveCsv(Dictionary<int, string> dicttable, string pathsave)
    {
        StringBuilder strBuilder = new StringBuilder();
        foreach (var row in dicttable)
        {
            if (row.Value.EndsWith("|"))
            {
               strBuilder.Append(row.Key + "|" + row.Value.Substring(0,row.Value.Length-1));
            } 

            strBuilder.Append("\n");
        }

         Save(strBuilder, pathsave);
    }


    #endregion Markin methods

    #endregion Sevrice methods

    #region Service classes

    #region Gukov Classes

    #region Enums

    private enum TypeOfObject {
        НеОпределено,
        СборочнаяЕдиница,
        СтандартноеИзделие,
        ПрочееИзделие,
        Изделие,
        Деталь,
        ЭлектронныйКомпонент,
        Материал,
        Другое,
        НеОбработано
    }

    private enum TypeOfReference {
        ИсходныйОбъект, // Может быть объект справочника документы или электронные компоненты
        НоменклатурныйОбъект, // Номенклатурный объект, созданный на основе исходного объекта
        СвязующийОбъект // Объект справочника 'Список номенклатуры FoxPro'
    }

    private enum StatusOfObject {
        Отсутствует,
        Найден,
        ПолученПоСвязи,
        Создан,
        Пересоздан
    }

    private enum TypeOfError {
        ОшибкиНет,
        ОтсутствуетНаименование,
        ОтсутствуетПодключение,
        НеоднозначныйВыбор,
        НеопределенныйТип,
        НекорректныйНоменклатурныйОбъект,
        СозданиеОбъекта,
        РазныеТипы
    }

    private enum TypeOfLinkError {
        ОтсутствуетОбъект,
        ОбъектСодержитОшибку,
        ОбъектИсключен,
        ОшибкаВПроцессеПодключения
    }

    private enum TypeOfAction {
        Создание,
        Редактирование,
        Удаление
    }

    private enum TypeOfExclusion {
        Отсутствует,
        ПризнакКонечногоИзделия,
        НеТребуетсяСоздавать
    }

    #endregion Enums

    #region ObjectInTFlex

    private class ObjectInTFlex {

        #region Fields and Properties

        public static Dictionary<string, ObjectInTFlex> Objects { get; set; } = new Dictionary<string, ObjectInTFlex>(); // статический словарь, в котором будут содержаться все объекты

        // Найденные в T-Flex позиции
        private static List<ReferenceObject> nomenclatureFox { get; set; } = null;
        private static List<ReferenceObject> nomenclatureTflex { get; set; } = null;
        private static List<ReferenceObject> documents { get; set; } = null;
        private static List<ReferenceObject> electricalComponents { get; set; } = null;

        // Словарь, в котором будут храниться названия для обрабатываемых ДСЕ
        private static Dictionary<string, string> names = null;

        private static bool IsPreparatorySearchComplete = false;

        public string Shifr { get; private set; } // Шифр объекта
        public string Name { get; private set; }
        public string LogSearch { get; private set; } // Лог для процесса поиска объектов
        public string LogError { get; private set; } // Лог для сообщения об ошибках
        public string LogLinks { get; private set; } // Лог для сообщения о процессе работы с подключениями

        // Поля для определения типов найденных объектов (и общего типа, который должен быть
        public TypeOfObject Type { get; private set; } = TypeOfObject.НеОбработано; // Тип объекта
        public TypeOfObject TypeOfConnection { get; private set; } = TypeOfObject.НеОбработано; // Тип, полученнный с поля объекта справочника 'Список номенклатуры FoxPro'
        public TypeOfObject TypeOfDocument { get; private set; } = TypeOfObject.НеОбработано; // Тип, полученный с документа (или электронного компонента)
        public TypeOfObject TypeOfNomenclature { get; private set; } = TypeOfObject.НеОбработано; // Тип, полученный с номенклатурного объекта 

        public TypeOfError Error { get; private set; } = TypeOfError.ОшибкиНет; // Тип ошибки
        public TypeOfExclusion Exclusion { get; private set; } = TypeOfExclusion.Отсутствует; // Тип исключения позиции из дальнейшей работы

        // Найденные объекты
        public ReferenceObject ConnectionObject { get; private set; } = null; // Объект справочника "Список номенклатуры"
        public ReferenceObject NomenclatureObject { get; private set; } = null; // Объект справочника "ЭСИ"
        public ReferenceObject DocumentObject { get; private set; } = null; // Объект справочника "Документы" или "Электронные компоненты", который привязан к объекту ЭСИ

        // Флаги объекта
        public bool HasError { get; private set; } = false; // Указывает, возникала ли на данном объекте ошибка
        public bool HasExcluded { get; private set; } = false; // Указывает, что позиция была исключена
        public bool HasBeenValidated { get; private set; } = false; // Указывает, что была проведена валидация
        public bool HasSameShifr { get; private set; } = false; // Результат валидации данных по соответствию шифра у всех найденных объектов
        public bool HasSameName { get; private set; } = false; // Результат валидации данных по соответствию наименования у всех найденных объектов

        // Свойство, которое показывает, успешно ли был произведен поиск объектов
        public bool IsDataComplete => ((this.DocumentObject != null) && (this.NomenclatureObject != null) && (this.ConnectionObject != null));
        public bool IsRootRecord => this.Exclusion == TypeOfExclusion.ПризнакКонечногоИзделия;
        public bool IsNotCreates => this.Exclusion == TypeOfExclusion.НеТребуетсяСоздавать;
        public bool IsExcluded => this.Exclusion != TypeOfExclusion.Отсутствует;

        // Статус объектов справочников T-Flex
        public StatusOfObject ConnectionObjectStatus { get; private set; } = StatusOfObject.Отсутствует;
        public StatusOfObject NomenclatureObjectStatus { get; private set; } = StatusOfObject.Отсутствует;
        public StatusOfObject DocumentObjectStatus { get; private set; } = StatusOfObject.Отсутствует;

        // Параметр, который будет отвечать за то, что будет в дальнейшем происходить с данным объектом
        public List<TypeOfAction> Actions { get; private set; } = new List<TypeOfAction>();

        public string Status =>
            string.Format(
                    "Объект '{0}' - '{1}':\n" +
                    "- Список номенклатуры: {2} ({8}) (Тип: {17})\n" +
                    "- Номенклатура: {3} ({9}) (Тип: {18})\n" +
                    "- Исходный объект: {4} ({10}) (Тип: {19})\n" +
                    "Ошибки: {5}({12}), Исключение: {16}, Валидация: {6}, Общий тип: {7}\n" +
                    "Действия: Создание - {13}; Редактирование - {14}; Удаление - {15}\n\n" +
                    "Лог поиска записей:\n{20}\n\n" +
                    "Лог подключения записей:\n{21}\n\n" +
                    "Сообщение об ошибке:\n{11}\n\n",
                    this.Shifr, // 0
                    this.Name, // 1
                    this.ConnectionObject != null ? this.ConnectionObject.SystemFields.Guid.ToString() : "Отсутствует", // 2
                    this.NomenclatureObject != null ? this.NomenclatureObject.SystemFields.Guid.ToString() : "Отсутствует", // 3
                    this.DocumentObject != null ? this.DocumentObject.SystemFields.Guid.ToString() : "Отсутствует", // 4
                    this.HasError.ToString(), // 5
                    this.HasBeenValidated.ToString(), // 6
                    this.Type.ToString(), // 7
                    this.ConnectionObjectStatus.ToString(), // 8
                    this.NomenclatureObjectStatus.ToString(), // 9
                    this.DocumentObjectStatus.ToString(), // 10
                    this.LogError == string.Empty ? "- данные отсутствуют" : this.LogError, // 11
                    this.Error.ToString(), // 12
                    this.Actions.Contains(TypeOfAction.Создание) ? "Да" : "Нет", // 13
                    this.Actions.Contains(TypeOfAction.Редактирование) ? "Да" : "Нет", // 14
                    this.Actions.Contains(TypeOfAction.Удаление) ? "Да" : "Нет", // 15
                    this.Exclusion.ToString(), // 16
                    this.TypeOfConnection.ToString(), // 17
                    this.TypeOfNomenclature.ToString(), // 18
                    this.TypeOfDocument.ToString(), // 19
                    this.LogSearch == string.Empty ? "- данные отсутствуют" : this.LogSearch, // 20
                    this.LogLinks == string.Empty ? "- данные отсутствуют" : this.LogLinks // 21
                    );

        #endregion Fields and Properties

        #region Constructors

        private ObjectInTFlex() {
        }

        private ObjectInTFlex(TypeOfAction action, string shifr) {
            this.Shifr = shifr;
            this.Actions.Add(action);
        }

        #endregion Constructors

        #region static PreparatorySearch()

        public static void PreparatorySearch() {

            if (Objects.Count == 0)
                throw new Exception(
                        "ObjectInTFlex не содержит объектов, необходимых для предварительной подготовки информации для произведения поиска. " +
                        "Возможная причина - не проведено добавление объектов при помощи AddToSearch"
                        );

            // Получаем все обозначения
            string[] shifrs = Objects.Select(kvp => kvp.Key).ToArray();

            // Производим предварительный поиск объектов в необходимых справочниках DOCs для сокращения времени работы программы
            SearchObjectsInTFlex(shifrs);
            // Пробуем получить все названия для задействованных в процессе выгрузки обозначений изделий
            GetNameOfProducts(shifrs);

            // Отмечаем, что предварительная загрузка была выполнены для того, чтобы можно было запускать поиск
            // существующих позиций и создание отсутствующих
            IsPreparatorySearchComplete = true;

        }

        #region static SearchObjectsInTFlex()

        private static void SearchObjectsInTFlex(string[] shifrs) {
            nomenclatureFox = listOfNomenclatureReference.Find(shifrOfListNomenclature, ComparisonOperator.IsOneOf, shifrs);
            nomenclatureTflex = nomenclatureReference.Find(shifrOfNomenclature, ComparisonOperator.IsOneOf, shifrs);
            documents = documentReference.Find(shifrOfDocument, ComparisonOperator.IsOneOf, shifrs);
            electricalComponents = componentReference.Find(shifrOfComponent, ComparisonOperator.IsOneOf, shifrs);
        }

        #endregion static SearchObjectsInTFlex()

        #region static GetNameOfProducts()

        private static void GetNameOfProducts(string[] shifrs) {

            // Создаем контейнет для результатов поиска
            ReferenceObject tempObject = null;

            // Инициализируем словарь с названиями объектов
            names = new Dictionary<string, string>();

            // Получаем параметры справочинка Spec для поиска в нем необходимой информации
            ParameterInfo shifrOfSpec = specReference.ParameterGroup[Guids.Parameters.Spec.Shifr];
            ParameterInfo izdOfSpec = specReference.ParameterGroup[Guids.Parameters.Spec.Izd];

            foreach (string shifr in shifrs) {
                string name = null;
                // Сначала пробегаем по списку номенклатуры FoxPro
                tempObject = nomenclatureFox.FirstOrDefault(rec => (string)(rec[shifrOfListNomenclature].Value) == shifr);
                if (tempObject != null) {
                    name = (string)(tempObject[nameOfListNomenclature].Value);

                    if (!string.IsNullOrWhiteSpace(name)) {
                        names[shifr] = name;
                        continue;
                    }
                }

                // Пробегаем по списку номенклатуры
                tempObject = nomenclatureTflex.FirstOrDefault(rec => (string)(rec[shifrOfNomenclature].Value) == shifr);
                if (tempObject != null) {
                    name = (string)(tempObject[nameOfNomenclature].Value);

                    if (!string.IsNullOrWhiteSpace(name)) {
                        names[shifr] = name;
                        continue;
                    }
                }

                // Пробегаем по списку документов
                tempObject = documents.FirstOrDefault(rec => (string)(rec[shifrOfDocument].Value) == shifr);
                if (tempObject != null) {
                    name = (string)(tempObject[nameOfDocument].Value);

                    if (!string.IsNullOrWhiteSpace(name)) {
                        names[shifr] = name;
                        continue;
                    }
                }

                // Пробегаем по списку компонентов
                tempObject = electricalComponents.FirstOrDefault(rec => (string)(rec[shifrOfComponent].Value) == shifr);
                if (tempObject != null) {
                    name = (string)(tempObject[nameOfComponent].Value);

                    if (!string.IsNullOrWhiteSpace(name)) {
                        names[shifr] = name;
                        continue;
                    }
                }

                // Пробуем найти данное изделие в таблице SPEC.
                // Если его там не получается найти по шифру, значит данной позиции не существует и ее нужно пометить как признак
                // конечного изделия
                ReferenceObject findedObjectInSpec = specReference.FindOne(shifrOfSpec, shifr);
                if (findedObjectInSpec == null) {

                    names[shifr] = "Признак конечного изделия";
                    continue;
                }
                else {
                    names[shifr] = (string)(findedObjectInSpec[Guids.Parameters.Spec.Naim].Value);
                    continue;
                }

                names[shifr] = string.Empty;
            }
        }

        #endregion static GetNameOfProducts()

        #endregion static PreparatorySearch()

        #region static AddToSearch()

        public static void AddToSearch(string[] shifrs, TypeOfAction type) {
            // Метод для добавления шифров и создания их новых объектов ObjectInTFlex
            foreach (string shifr in shifrs) {
                if (shifr != string.Empty) {
                    if (Objects.ContainsKey(shifr)) {
                        Objects[shifr].AddAction(type);
                    }
                    else {
                        Objects.Add(shifr, new ObjectInTFlex(type, shifr));
                    }
                }
            }
        }

        #endregion static AddToSearch()

        #region static Search()

        public static void Search () {

            string errorMessage = string.Empty;
            errorMessage += !IsPreparatorySearchComplete ? "- перед применением Search необходимо выполнить PreparatorySearch" : string.Empty;
            errorMessage += Objects.Count == 0 ? "- отсутствуют объекты для выполнения метода Search()" : string.Empty;

            if (errorMessage != string.Empty)
                throw new Exception(string.Format("При выполнении метода Search возникли ошибки:\n{0}", errorMessage));

            // Приступаем к поиску объектов

            // Сначала находим объекты в справочнике "Список номенклатуры FoxPro".
            // Если не находим - создаем новый объект
            SearchOrCreateConnectionObjects();

            // На втором заходе пытемся получить связанный номенклатурный объект с объекта
            // "Списка номенклатуры FoxPro".
            // Если не получается, пытаемся при помощи найти данный объект в спраовочнике
            SearchOrCreateNomenclatureObjects();

            // На третьем этапе пытаемся найти исходный объект (которым помимо документа может быть объект
            // справочника "Электронные компоненты"
            // При отсутствии пытаемся создать необходимые документы
            SearchOrCreateDocumentObjects();

            // Проходим по позициям, у которых есть несовпадение по типам c целью приведения к общему знаменателю
            NormalizeTypes();

            // Проходим по позициям, у которых сломаны номенклатурные объекты для того, чтобы попытаться снова их воссоздать
            RecreateBrokenNomenclatureObjects();
        }

        #region static SearchOrCreateConnectionObjects()

        private static void SearchOrCreateConnectionObjects() {

            // Контейнер для найденных объектов
            List<ReferenceObject> findedObjects = null;

            // Производим первичное наполнение словаря Objects, пытаемся наполнить его объектами
            // из справочника "Список номенклатуры FoxPro"
            foreach (KeyValuePair<string, ObjectInTFlex> kvp in Objects) {
                // Если у объекта пустое наименование, отмечаем позицию как ошибочкую
                kvp.Value.Name = names[kvp.Key];
                if (kvp.Value.Name == string.Empty) {
                    kvp.Value.SetError(
                            TypeOfError.ОтсутствуетНаименование,
                            "Не удалось получить наименование на основе существующих в T-Flex данных"
                            );
                }
                if (kvp.Value.Name == "Признак конечного изделия") {
                    kvp.Value.SetExclusion(TypeOfExclusion.ПризнакКонечногоИзделия);
                }

                // Пробуем по данному шифру найти объект в справочнике "Список номенклатуры"
                findedObjects = nomenclatureFox
                    .Where(nom => (string)(nom[Guids.Parameters.СписокНоменклатуры.Обозначение].Value) == kvp.Key)
                    .ToList<ReferenceObject>();

                switch (findedObjects.Count) {
                    case 1:
                        kvp.Value.LinkObject(findedObjects[0], TypeOfReference.СвязующийОбъект, StatusOfObject.Найден);
                        break;
                    case 0:
                        kvp.Value.CreateObject(TypeOfReference.СвязующийОбъект);
                        break;
                    default:
                        kvp.Value.SetError(
                                TypeOfError.НеоднозначныйВыбор,
                                string.Format(
                                "В справочнике 'Список номенклатуры FoxPro' найдено более одного совпадения по обозначению:\n{0}",
                                    string.Join("\n", findedObjects.Select(refobj => string.Format("- {0}", refobj.ToString())))
                                    )
                                );
                        break;
                }
            }
        }

        #endregion static SearchOrCreateConnectionObjects()

        #region static SearchOrCreateNomenclatureObjects()

        private static void SearchOrCreateNomenclatureObjects() {

            // Контейнер для найденных объектов
            List<ReferenceObject> findedObjects = null;

            // Второй заход, пытаемся подцепить объект справочника ЭСИ.
            // Обрабатываем только те записи, по которым не возникло ошибок на предыдущем этапe
            foreach (KeyValuePair<string, ObjectInTFlex> kvp in ObjectInTFlex.Objects.Where(obj => (!obj.Value.HasError) && (!obj.Value.HasExcluded))) {
                // Пытаемся получить объект по связи
                kvp.Value.LinkObject(
                        kvp.Value.ConnectionObject.GetObject(Guids.Links.СписокНоменклатуры.Номенклатура),
                        TypeOfReference.НоменклатурныйОбъект,
                        StatusOfObject.ПолученПоСвязи
                        );

                if (kvp.Value.NomenclatureObject != null) {

                    NomenclatureObject nom = kvp.Value.NomenclatureObject as NomenclatureObject;
                    kvp.Value.LinkObject(nom.LinkedObject, TypeOfReference.ИсходныйОбъект, StatusOfObject.ПолученПоСвязи);

                    if (kvp.Value.DocumentObject == null) {
                        kvp.Value.SetError(
                                TypeOfError.НекорректныйНоменклатурныйОбъект,
                                "У номенклатурного объекта отсутствует связанный исходный объект"
                                );
                    }
                }
                else {
                    // Случай, когда у объекта в "Список номенклатуры FoxPro" нет подключенного объекта ЭСИ
                    // Пробуем произвести поиск по обозначению в справочнике ЭСИ
                    
                    findedObjects = nomenclatureTflex
                        .Where(nom => (string)(nom[Guids.Parameters.Номенклатура.Обозначение].Value) == kvp.Key)
                        .ToList<ReferenceObject>();

                    switch (findedObjects.Count) {
                        case 1:
                            kvp.Value.LinkObject(findedObjects[0], TypeOfReference.НоменклатурныйОбъект, StatusOfObject.Найден);

                            NomenclatureObject nom = findedObjects[0] as NomenclatureObject;
                            if (nom.LinkedObject != null) {
                                kvp.Value.LinkObject(nom.LinkedObject, TypeOfReference.ИсходныйОбъект, StatusOfObject.ПолученПоСвязи);
                            }
                            else {
                                kvp.Value.SetError(
                                        TypeOfError.НекорректныйНоменклатурныйОбъект,
                                        "У номенклатурного объекта отсутствует связанный исходный объект"
                                        );
                            }

                            break;
                        case 0:
                            break;
                        default:
                            kvp.Value.SetError(
                                    TypeOfError.НеоднозначныйВыбор,
                                    string.Format(
                                        "В справочнике 'ЭСИ' найдено более одного совпадения по обозначению:\n{0}",
                                        string.Join("\n", findedObjects.Select(refobj => string.Format("- {0}", refobj.ToString())))
                                        )
                                    );
                            break;
                    }
                }
            }
        }

        #endregion static SearchOrCreateNomenclatureObjects()

        #region static SearchOrCreateDocumentObjects()

        private static void SearchOrCreateDocumentObjects() {

            // Контейнер для найденных объектов
            List<ReferenceObject> findedObjects = null;

            // Третий заход - пробуем подцепить объекты справочника "Документы"
            // Обрабатываем только те записи, которые были отработаны без ошибок и у которых еще не полная информация
            foreach (KeyValuePair<string, ObjectInTFlex> kvp in ObjectInTFlex.Objects.Where(obj => (!obj.Value.HasError) && (!obj.Value.IsDataComplete) && (!obj.Value.HasExcluded))) {
                // Данный код должен срабатывать только в том случае, если в справочнике ЭСИ на предыдущих этапах не был обнаружен номенклатурный объект
                // и следовательно его нужно получить при помощи справочника "Документы" или "Электронные компоненты"
                //MessageBox.Show(string.Format("Обработка случая, когда не был найден объект электронной структуры изделия шифр ('{0}')", kvp.Key), kvp.Key);

                switch (kvp.Value.Type) {
                    case TypeOfObject.ЭлектронныйКомпонент:
                        // Случай, когда мы точно знаем, что ищем электронный компонент
                        findedObjects = electricalComponents
                            .Where(comp => (string)(comp[Guids.Parameters.ЭлектронныеКомпоненты.Обозначение].Value) == kvp.Key)
                            .ToList<ReferenceObject>();
                        break;
                    case TypeOfObject.НеОпределено:
                        // Случай, когда мы не знаем, что ищем
                        findedObjects = documents
                            .Where(doc => (string)(doc[Guids.Parameters.Документы.Обозначение].Value) == kvp.Key)
                            .ToList<ReferenceObject>();

                        if (findedObjects.Count == 0) {
                            findedObjects = electricalComponents
                                .Where(comp => (string)(comp[Guids.Parameters.ЭлектронныеКомпоненты.Обозначение].Value) == kvp.Key)
                                .ToList<ReferenceObject>();
                            // Если данный объект получилось найти в справочнике Электронные компоненты, тогда мы меняем его тип
                            if (findedObjects.Count != 0) {
                                kvp.Value.SetTypeObject(TypeOfObject.ЭлектронныйКомпонент);
                            }
                        }
                        break;
                    default:
                        // Случай, когда мы ищем документ
                        findedObjects = documents
                            .Where(doc => (string)(doc[Guids.Parameters.Документы.Обозначение].Value) == kvp.Key)
                            .ToList<ReferenceObject>();
                        break;
                }


                switch (findedObjects.Count) {

                    case 1:
                        kvp.Value.LinkObject(findedObjects[0], TypeOfReference.ИсходныйОбъект, StatusOfObject.Найден);

                        // Пытаемся получить связанный номенклатурный объект с документа
                        NomenclatureObject nom = findedObjects[0].GetLinkedNomenclatureObject();
                        if (nom == null) {
                            kvp.Value.CreateObject(TypeOfReference.НоменклатурныйОбъект);
                            kvp.Value.AddLogSearchMessage(
                                    "- Создан объект в справочнике 'ЭСИ' на основе" +
                                    " существующего объекта в справочнике 'Документы'"
                                    );
                        }
                        else {
                            kvp.Value.LinkObject(
                                    nom as ReferenceObject,
                                    TypeOfReference.НоменклатурныйОбъект,
                                    StatusOfObject.ПолученПоСвязи
                                    );

                            if (kvp.Value.NomenclatureObject == null)
                                throw new Exception(string.Format(
                                            "Ошибка при попытке привести номерклатурный объект к объекту справочника\n\n{0}",
                                            kvp.Value.Status
                                            ));
                        }
                        break;

                    case 0:
                        kvp.Value.CreateObject(TypeOfReference.ИсходныйОбъект);
                        break;

                    default:
                        kvp.Value.SetError(
                                TypeOfError.НеоднозначныйВыбор,
                                "В справочнике 'Документы' найдено более одного совпадения по обозначению"
                                );
                        break;
                }
            }
        }

        #endregion static SearchOrCreateDocumentObjects()

        #region static NormalizeTypes()
        
        private static void NormalizeTypes() {
            // Метод для изменения типов объектов, если они были заданы неправильно
            foreach (KeyValuePair<string, ObjectInTFlex> kvp in ObjectInTFlex.Objects.Where(kvp => kvp.Value.Error == TypeOfError.РазныеТипы)) {
                // Проверка на ошибки
                string errorMessage = string.Empty;

                errorMessage += (kvp.Value.Type != kvp.Value.TypeOfConnection) ? "- основной тип и тип связующего объекта на совпадают\n" : string.Empty;
                errorMessage += (kvp.Value.TypeOfNomenclature != kvp.Value.TypeOfDocument) ? "- типы номенклатурного и исходного объекта не совпадают\n" : string.Empty;
                errorMessage += (kvp.Value.Type == TypeOfObject.НеОбработано) ? "- общий тип не обработан\n" : string.Empty;
                errorMessage += (kvp.Value.TypeOfConnection == TypeOfObject.НеОбработано) ? "- тип связующего объекта не обработан\n" : string.Empty;
                errorMessage += (kvp.Value.TypeOfNomenclature == TypeOfObject.НеОбработано) ? "- тип номенклатурного объекта не обработан\n" : string.Empty;
                errorMessage += (kvp.Value.TypeOfDocument == TypeOfObject.НеОбработано) ? "- тип исходного документа не обработан\n" : string.Empty;

                if (errorMessage != string.Empty) {
                    throw new Exception(string.Format("В процессе нормализации типов возникли следующие ошибки:\n{0}\n\n{1}", errorMessage, kvp.Value.Status));
                }

                // Если основной тип указан как "Не определено" или как "Другое", то мы просто проиводим корректировку
                // значения типа объекта в поле справочника "Список номенклатуры FoxPro"
                if ((kvp.Value.Type == TypeOfObject.НеОпределено) || (kvp.Value.Type == TypeOfObject.Другое)) {

                    // Изменяем поле соотвутствующим образом
                    kvp.Value.ConnectionObject.BeginChanges();
                    kvp.Value.ConnectionObject[Guids.Parameters.СписокНоменклатуры.ТипНоменклатуры].Value =
                        dictTypesToInt[kvp.Value.TypeOfNomenclature];
                    kvp.Value.ConnectionObject.EndChanges();

                    // Приводим в порядок типы объекта ObjectInTFlex
                    kvp.Value.Type = kvp.Value.TypeOfConnection = kvp.Value.TypeOfNomenclature;
                    
                    // Снимаем ошибку с данной позиции и приступаем к обработке следующей позиции
                    kvp.Value.UnsetError();
                    continue;
                }

                // Все остальные позиции нужно приводить к общему типу
                if ((kvp.Value.Type != TypeOfObject.ЭлектронныйКомпонент) && (kvp.Value.TypeOfNomenclature != TypeOfObject.ЭлектронныйКомпонент)) {
                    // В данном случае мы просто меняем тип объекта
                    if (!kvp.Value.DocumentObject.IsCheckedOut)
                        kvp.Value.DocumentObject.CheckOut();

                    kvp.Value.DocumentObject = kvp.Value.DocumentObject.BeginChanges(dictTypesToClassObject[kvp.Value.Type]);
                    kvp.Value.DocumentObject.EndChanges();
                    //  Получаем обновленную версию номенклатурного объекта
                    kvp.Value.NomenclatureObject = kvp.Value.DocumentObject.GetLinkedNomenclatureObject();

                    /*
                    string message = string.Format("Изменение типа объекта '{0}' c типа '{1}' на тип '{2}'", kvp.Value.Shifr, kvp.Value.TypeOfDocument.ToString(), kvp.Value.Type.ToString());
                    Desktop.CheckIn(kvp.Value.DocumentObject, message, false);
                    */
                }
                else {
                    kvp.Value.RecreateObject();
                }

                kvp.Value.UnsetError();

            }
        }
        
        #endregion static NormalizeTypes()

        #region static RecreateBrokenNomenclatureObjects()

        private static void RecreateBrokenNomenclatureObjects() {
            foreach (KeyValuePair<string, ObjectInTFlex> kvp in ObjectInTFlex.Objects.Where(kvp => kvp.Value.Error == TypeOfError.НекорректныйНоменклатурныйОбъект)) {

                kvp.Value.RecreateObject();
                kvp.Value.UnsetError();
            }
        }

        #endregion static RecreateBrokenNomenclatureObjects()

        #endregion static Search()

        #region FindAllOccurences()
        // Метод для поиска всех совпадений по данному шифру в справочниках "ЭСИ", "Документы", "Список номенклатуры".
        // Данный метод пишет сообщение в параметр LogSearch
        //
        // Метод крайне затратный и его лучше не использовать при формировании выгрузки (его лучше использовать при отладке)

        public List<ReferenceObject> FindAllOccurences() {
            List<ReferenceObject> result = new List<ReferenceObject>();
            
            // Производим поиск всех совпадений в справочнике "Список номенклатуры"
            List<ReferenceObject> connections = listOfNomenclatureReference.Find(shifrOfListNomenclature, ComparisonOperator.Equal, this.Shifr);
            if (connections.Count != 0) {
                this.AddLogSearchMessage("Объекты, найденные в справочнике 'Список номенклатуры':");
                foreach (ReferenceObject connection in connections) {
                    string linkToNomenclature = "отсутствует";
                    // Пробуем получить по связи объект справочника ЭСИ
                    ReferenceObject nom = connection.GetObject(Guids.Links.СписокНоменклатуры.Номенклатура);
                    if (nom != null) {
                        linkToNomenclature = nom.SystemFields.Guid.ToString();
                    }
                    
                    this.AddLogSearchMessage(string.Format("- {0} ({1}), связь на ЭСИ - {2}", connection.ToString(), connection.SystemFields.Guid.ToString(), linkToNomenclature));
                }
            }
            else {
                this.AddLogSearchMessage("В справочнике 'Список номенклатуры' ничего не было найдено");
            }
            this.AddLogSearchMessage("\n");

            // Производим поиск всех совпадений в справочнике "ЭСИ"
            List<ReferenceObject> nomenclatures = nomenclatureReference.Find(shifrOfNomenclature, ComparisonOperator.Equal, this.Shifr);
            if (nomenclatures.Count != 0) {
                this.AddLogSearchMessage("Объекты, найденные в справочнике 'ЭСИ':");
                foreach (ReferenceObject nomenclature in nomenclatures) {
                    // Проверяем, есть ли связанный документ
                    NomenclatureObject nom = nomenclature as NomenclatureObject;
                    string linkToDocuments = "отсутствует";
                    if (nom != null) {
                        // Пытаемся получить документ по связи
                        ReferenceObject document = nom.LinkedObject;
                        if (document != null)
                            linkToDocuments = document.SystemFields.Guid.ToString();
                    }
                    this.AddLogSearchMessage(string.Format("- {0} ({1}), связь на Документы - {2}", nomenclature.ToString(), nomenclature.SystemFields.Guid.ToString(), linkToDocuments));
                }
            }
            else {
                this.AddLogSearchMessage("В справочнике 'ЭСИ' ничего не было найдено");
            }
            this.AddLogSearchMessage("\n");

            // Производим поиск всех совпадений в справочнике "Документы"
            List<ReferenceObject> documents = documentReference.Find(shifrOfDocument, ComparisonOperator.Equal, this.Shifr);
            if (documents.Count != 0) {
                this.AddLogSearchMessage("Объекты, найденные в справочнике 'Документы':");
                foreach (ReferenceObject document in documents) {
                    // Проверяем, если ли связанный номенклатурный объект
                    DocumentReferenceObject doc = document as DocumentReferenceObject;
                    string linkToNom = "отсутствует";
                    if (doc != null) {
                        // пробуем получить связь на ЭСИ
                        ReferenceObject nomenclature = doc.GetLinkedNomenclatureObject();
                        if (nomenclature != null)
                            linkToNom = nomenclature.SystemFields.Guid.ToString();
                    }
                    this.AddLogSearchMessage(string.Format("- {0} ({1}), связь на ЭСИ - {2}", document.ToString(), document.SystemFields.Guid.ToString(), linkToNom));
                }
            }
            else {
                this.AddLogSearchMessage("В справочнике 'Документы' ничего не было найдено");
            }

            return result;
        }

        #endregion FindAllOccurences()

        #region AddLogSearchMessage()

        public void AddLogSearchMessage(string message) {
            if (this.LogSearch == null)
                this.LogSearch = string.Empty;
            this.LogSearch = string.Format("{0}{1}\n", this.LogSearch, message);
        }

        #endregion AddLogSearchMessage()

        #region AddLogLinksMessage()

        public void AddLogLinksMessage(string message) {
            if (this.LogLinks == null)
                this.LogLinks = string.Empty;
            this.LogLinks = string.Format("{0}{1}\n", this.LogLinks, message);
        }

        #endregion AddLogLinksMessage()

        #region LinkObject()

        private void LinkObject(ReferenceObject obj, TypeOfReference type, StatusOfObject status) {
            if (obj == null)
                return;

            switch (type) {
                case TypeOfReference.НоменклатурныйОбъект:
                    this.NomenclatureObject = obj;
                    this.NomenclatureObjectStatus = status;
                    // Если объект был создан или найден, производим его подключение к списку номенклатуры
                    if ((status == StatusOfObject.Найден) || (status == StatusOfObject.Создан)) {
                        this.ConnectionObject.BeginChanges();
                        this.ConnectionObject.SetLinkedObject(Guids.Links.СписокНоменклатуры.Номенклатура, this.NomenclatureObject);
                        this.ConnectionObject.EndChanges();
                    }
                    break;
                case TypeOfReference.СвязующийОбъект:
                    this.ConnectionObject = obj;
                    this.ConnectionObjectStatus = status;
                    break;
                case TypeOfReference.ИсходныйОбъект:
                    this.DocumentObject = obj;
                    this.DocumentObjectStatus = status;
                    break;
                default:
                    throw new Exception(string.Format("Ошибка при добавлении объекта во время поиска. Тип объекта {0} не известен\n\n{0}", type.ToString(), this.Status));
            }

            this.GetType(type); // Определяем тип прилинкованного изделия

            // Если объект уже содержит в себе всю информацию, производим проверку типов
            if (this.IsDataComplete)
                this.CheckTypes();
        }

        #endregion LinkObject()

        #region SetError()

        private void SetError(TypeOfError type, string message = "") {
            this.HasError = true; // Отмечаем, что данная позиция имеет ошибку
            this.Error = type; // Отмечаем тип ошибки, который возник
            if (message != string.Empty) {
                this.LogError = message;
            }
        }

        private void UnsetError() {
            this.HasError = false;
            this.Error = TypeOfError.ОшибкиНет;
            this.LogError = string.Empty;
        }

        private void SetExclusion(TypeOfExclusion type) {
            this.HasExcluded = true;
            this.Exclusion = type;
        }

        #endregion SetError()

        #region AddAction()

        private void AddAction(TypeOfAction action) {
            if (!this.Actions.Contains(action))
                this.Actions.Add(action);
        }

        #endregion AddAction()

        #region GetType()

        private void GetType(TypeOfReference type) {

            switch (type) {
                case TypeOfReference.СвязующийОбъект:
                    if (this.ConnectionObject == null)
                        throw new Exception(string.Format("При попытке определить тип объекта возникла ошибка: отсутствует связующий объект\n\n{0}", this.Status));
                    GetTypeOfConnectionObject();
                    break;
                case TypeOfReference.НоменклатурныйОбъект:
                    if (this.NomenclatureObject == null)
                        throw new Exception(string.Format("При попытке определить тип объекта возникла ошибка: отсутствует номенклатурный объект\n\n{0}", this.Status));
                    GetTypeOfNomenclatureObject();
                    break;
                case TypeOfReference.ИсходныйОбъект:
                    if (this.DocumentObject == null)
                        throw new Exception(string.Format("При попытке определить тип объекта возникла ошибка: отсутствует исходный объект\n\n{0}", this.Status));
                    GetTypeOfDocumentObject();
                    break;
                default:
                    throw new Exception(string.Format("При определении типа возникла ошибка. '{0}' - не поддерживается", type.ToString()));
            }
        }

        #region GetTypeOfConnectionObject()

        private void GetTypeOfConnectionObject() {
            int intType = (int)this.ConnectionObject[Guids.Parameters.СписокНоменклатуры.ТипНоменклатуры].Value;

            if (dictIntToTypes.ContainsKey(intType)) {
                this.TypeOfConnection = dictIntToTypes[intType];
            }
            else {
                throw new Exception(string.Format(
                            "Для параметра типа изделия не поддерживаетя значение -'{0}'\n\n{1}",
                            intType,
                            this.Status
                            ));
            }

            this.Type = this.TypeOfConnection;
        }

        #endregion GetTypeOfConnectionObject()
        
        #region GetTypeOfNomenclatureObject()

        private void GetTypeOfNomenclatureObject() {
            string typeString = this.NomenclatureObject.Class.ToString().ToLower();

            if (dictStringToTypes.ContainsKey(typeString)) {
                this.TypeOfNomenclature = dictStringToTypes[typeString];
            }
            else {
                throw new Exception(string.Format(
                            "При определении типа подкюченного номенклатурного объекта возникла ошибка: " +
                            "тип '{0}' не поддерживается\n\n{1}",
                            typeString,
                            this.Status
                            ));
            }
        }

        #endregion GetTypeOfNomenclatureObject()

        #region GetTypeOfDocumentObject()

        private void GetTypeOfDocumentObject() {
            string typeString = this.DocumentObject.Class.ToString().ToLower();

            if (dictStringToTypes.ContainsKey(typeString)) {
                this.TypeOfDocument = dictStringToTypes[typeString];
            }
            else {
                throw new Exception(string.Format(
                            "При определении типа подкюченного номенклатурного объекта возникла ошибка: " +
                            "тип '{0}' не поддерживается\n\n{1}",
                            typeString,
                            this.Status
                            ));
            }
        }

        #endregion GetTypeOfDocumentObject()

        #endregion GetType()

        #region CheckTypes()

        private void CheckTypes() {
            // Метод для проверки типов объектов
            if ((this.Type != this.TypeOfConnection) || (this.Type != this.TypeOfNomenclature) || (this.Type != this.TypeOfDocument)) {
                this.SetError(TypeOfError.РазныеТипы, "Для данного объекта требуется произвести нормализацию типов");
            }
        }

        #endregion CheckTypes()

        #region MakeGuessTheType()

        private void MakeGuessTheType() {
            // Метод для обработки позиций, у которых с списке номенклатуры стоит тип "Не определено"

            try {
                System.Convert.ToInt64(this.Shifr);

                // Делаем предположение по наименованию изделия о типе
                foreach (string key in keyWordsElectricalComponents) {
                    if (this.Name.ToLower().Contains(key)) {
                        this.Type = TypeOfObject.ЭлектронныйКомпонент;
                        this.TypeOfConnection = TypeOfObject.ЭлектронныйКомпонент;

                        this.ConnectionObject.BeginChanges();
                        this.ConnectionObject[Guids.Parameters.СписокНоменклатуры.ТипНоменклатуры].Value = 
                            dictTypesToInt[TypeOfObject.ЭлектронныйКомпонент];
                        this.ConnectionObject.EndChanges();
                        return;
                    }
                }

                foreach (string key in keyWordsStandartParts) {
                    if (this.Name.ToLower().Contains(key)) {
                        this.Type = TypeOfObject.СтандартноеИзделие;
                        this.TypeOfConnection = TypeOfObject.СтандартноеИзделие;

                        this.ConnectionObject.BeginChanges();
                        this.ConnectionObject[Guids.Parameters.СписокНоменклатуры.ТипНоменклатуры].Value = 
                            dictTypesToInt[TypeOfObject.СтандартноеИзделие];
                        this.ConnectionObject.EndChanges();
                        return;
                    }
                }
            }
            catch (FormatException) {
                // Случай, когда обозначение состоит не только из цифр
                this.Type = TypeOfObject.Другое;
                this.TypeOfConnection = TypeOfObject.Другое;

                // Изменяем тип, указанный в объекта "Списка номенклатуры FoxPro"
                this.ConnectionObject.BeginChanges();
                this.ConnectionObject[Guids.Parameters.СписокНоменклатуры.ТипНоменклатуры].Value =
                    dictTypesToInt[TypeOfObject.Другое];
                this.ConnectionObject.EndChanges();
            }
        }

        #endregion MakeGuessTheType()

        #region SetTypeObject()

        private void SetTypeObject(TypeOfObject type) {
            // Метод предназначен для установки типа для объекта.
            // Данный метод помимо установки типа на объекта ObjectInTFlex будет так же устанавливать значение в справочнике
            // "Список номенклатуры FoxPro"
            this.Type = type;

            if (dictTypesToInt.ContainsKey(type)) {
                if (this.ConnectionObject != null) {
                    this.ConnectionObject.BeginChanges();
                    this.ConnectionObject[Guids.Parameters.СписокНоменклатуры.ТипНоменклатуры].Value =
                        dictTypesToInt[this.Type];
                    this.ConnectionObject.EndChanges();
                }
                else {
                    throw new Exception(string.Format(
                                "При попытке установить тип для объекта возникла ошибка. " +
                                "Отсутствует соотвутствующий объект справочника " +
                                "'Список номенклатуры FoxPro'\n\n{0}",
                                this.Status
                                ));
                }
            }
            else {
                throw new Exception(string.Format(
                            "При попытке получить цифровое представление типа '{0}' возникла ошибка:\n" +
                            "Данный тип отсутствует в справочнике 'dictTypesToInt'",
                            type.ToString()
                            ));
            }
        }

        #endregion SetTypeObject()

        #region Методы по работе с подключениями

        #region CreateConnectionLinkTo()

        public void CreateConnectionLinkTo(ObjectInTFlex parent, int pos, double prim) {
            string templateError = 
                "Ошибка в процессе создания нового подключения.\n" +
                "Среди объектов для подключения ('{0}'; '{1}') есть исключенные (Исключения: дочерний - '{2}', родительский - '{3}')";

            if ((this.Exclusion != TypeOfExclusion.Отсутствует) || (parent.Exclusion != TypeOfExclusion.Отсутствует)) {
                this.AddLogLinksMessage(string.Format("- попытка создать подключение ('{0}') на исключенный объект", parent.Shifr));
                throw new Exception(string.Format(templateError, this.Shifr, parent.Shifr, this.Exclusion.ToString(), parent.Exclusion.ToString()));
            }
            
            // Производим поиск подкючения
            List<ComplexHierarchyLink> hLinks = this.NomenclatureObject.GetParentLinks(parent.NomenclatureObject).ToList<ComplexHierarchyLink>();
            ComplexHierarchyLink hLink = hLinks.Where(link => ((int)link[Guids.hLinks.Позиция].Value == pos) && ((double)link[Guids.hLinks.Количество].Value == prim)).FirstOrDefault();

            templateError =
                "Ошибка в процессе создания нового подключения.\n" +
                "Создание подключения '{0}' к '{1}' было отменено так как среди подключений в ЭСИ " +
                "было обнаружено подкючение с  позицией - '{2}' и количеством '{3}'\n" +
                "Найденные подлючения:\n{4}";

            if (hLink != null) {
                this.AddLogLinksMessage(string.Format("- попытка создать подключение ('{0}'), которое уже существует", parent.Shifr));
                string findedConnections = string.Join("\n", hLinks.Select(hlink => hlink.ToString()));
                throw new Exception(string.Format(templateError, this.Shifr, parent.Shifr, pos, prim, findedConnections));
            }

            try {
                hLink = this.NomenclatureObject.CreateParentLink(parent.NomenclatureObject);
                hLink[Guids.hLinks.Позиция].Value = pos;
                hLink[Guids.hLinks.Количество].Value = prim;
                hLink.EndChanges();
                this.AddLogLinksMessage(string.Format("- произведено подключение ('{0}') к родительскому объекту", parent.Shifr));
            }
            catch (Exception e) {
                this.AddLogLinksMessage(string.Format("- возникла системная ошибка при попытке создании подключения ('{0}')", parent.Shifr));
                throw new Exception(e.Message);
            }
        }

        #endregion CreateConnectionLinkTo()
        
        #region ChangeConnectionLinkTo()

        public void ChangeConnectionLinkTo(ObjectInTFlex parent, int oldPos, int newPos, double oldPrim, double newPrim) {
            string templateError = 
                "Ошибка в процессе создания нового подключения.\n" +
                "Среди объектов для подключения ('{0}'; '{1}') есть исключенные (Исключения: дочерний - '{2}', родительский - '{3}')";

            if ((this.Exclusion == TypeOfExclusion.НеТребуетсяСоздавать) || (parent.Exclusion == TypeOfExclusion.НеТребуетсяСоздавать)) {
                this.AddLogLinksMessage(string.Format("- попытка изменения подключения ('{0}') при наличии объекта с пометкой 'НеТребуетсяСоздавать'", parent.Shifr));
                throw new Exception(string.Format(templateError, this.Shifr, parent.Shifr, this.Exclusion.ToString(), parent.Exclusion.ToString()));
            }

            // Прозводим поиск подключения
            List<ComplexHierarchyLink> hLinks = this.NomenclatureObject.GetParentLinks(parent.NomenclatureObject).ToList<ComplexHierarchyLink>();
            ComplexHierarchyLink hLink = hLinks.Where(link => ((int)link[Guids.hLinks.Позиция].Value == oldPos) && ((double)link[Guids.hLinks.Количество].Value == oldPrim)).FirstOrDefault();

            templateError =
                "Ошибка в процессе изменения существующего подключения.\n" +
                "Изменение подключения '{0}' к '{1}' было отменено так как среди подключений в ЭСИ " +
                "отсутствовало подключение с исходными значениями позиции - '{2}' и количеством - '{3}'\n" +
                "Найденные подключения:\n{4}";

            if (hLink == null) {
                this.AddLogLinksMessage(string.Format("- попытка изменения подключения ('{0}'), которое отсутствует между объектами", parent.Shifr));
                string findedConnections = string.Join("\n", hLinks.Select(hlink => hlink.ToString()));
                throw new Exception(string.Format(templateError, this.Shifr, parent.Shifr, oldPos, oldPrim, findedConnections));
            }

            try {
                hLink.BeginChanges();
                hLink[Guids.hLinks.Позиция].Value = newPos;
                hLink[Guids.hLinks.Количество].Value = newPrim;
                hLink.EndChanges();
                this.AddLogLinksMessage(string.Format("- произведено изменение подключения ('{0}') к родительскому объекту", parent.Shifr));
            }
            catch (Exception e) {
                this.AddLogLinksMessage(string.Format("- возникла системная ошибка при попытке корректировки подключения ('{0}')", parent.Shifr));
                throw new Exception(e.Message);
            }
        }

        #endregion ChangeConnectionLinkTo()

        #region DeleteConnectionLinkTo()

        public void DeleteConnectionLinkTo(ObjectInTFlex parent, int pos, double prim) {
            string templateError = 
                "Ошибка в процессе удаления существующего подключения.\n" +
                "Один из объектов подключения ('{0}', '{1}') имеет признак 'Конечное изделие'";

            if ((this.Exclusion == TypeOfExclusion.ПризнакКонечногоИзделия) || (parent.Exclusion == TypeOfExclusion.ПризнакКонечногоИзделия)) {
                this.AddLogLinksMessage(string.Format("- попытка удаления подключения ('{0}') на объект, который является признаком конечного изделия", parent.Shifr));
                throw new Exception(string.Format(templateError, this.Shifr, parent.Shifr));
            }

            // Обработка случая, когда на удаление попадают строки, которые не создавались в T-Flex
            if ((this.Exclusion == TypeOfExclusion.НеТребуетсяСоздавать) || (parent.Exclusion == TypeOfExclusion.НеТребуетсяСоздавать)) {
                this.AddLogLinksMessage(string.Format("- попытка удаления подключения ('{0}') на объект, который не создавался в системе", parent.Shifr));
                throw new Exception(string.Format("Попытка удаления подключения ('{0}') на объект, который не создавался в системе", parent.Shifr));
            }

            List<ComplexHierarchyLink> hLinks = this.NomenclatureObject.GetParentLinks(parent.NomenclatureObject).ToList<ComplexHierarchyLink>();
            ComplexHierarchyLink hLink = hLinks.Where(link => ((int)link[Guids.hLinks.Позиция].Value == pos) && ((double)link[Guids.hLinks.Количество].Value == prim)).FirstOrDefault();

            templateError =
                "Ошибка в процессе удаления существующего подключения.\n" +
                "Удаление подключения '{0}' к '{1}' было отменено так как среди подключений в ЭСИ " +
                "отсутствовало подключение с исходными значениями позиции - '{2}' и количеством - '{3}'\n" +
                "Найденные подключения:\n{4}";
                
            if (hLink == null) {
                this.AddLogLinksMessage(string.Format("- попытка удаления подключения ('{0}'), которое не удалось найти", parent.Shifr));
                string findedConnections = string.Join("\n", hLinks.Select(hlink => hlink.ToString()));
                throw new Exception(string.Format(templateError, this.Shifr, parent.Shifr, pos.ToString(), prim.ToString(), findedConnections));
            }

            try {
                this.NomenclatureObject.DeleteLink(hLink);
                this.AddLogLinksMessage(string.Format("- произведено удаление подключения ('{0}') к родительскому объекту", parent.Shifr));
            }
            catch (Exception e) {
                this.AddLogLinksMessage(string.Format("- возникла системная ошибка при попытке удаления подключения ('{0}')", parent.Shifr));
                throw new Exception(e.Message);
            }
        }

        #endregion DeleteConnectionLinkTo()

        #endregion Методы по работе с подключениями

        #region Методы по созданию объектов в справочниках

        #region CreateObject()

        public void CreateObject(TypeOfReference type) {

            if ((this.Actions.Contains(TypeOfAction.Удаление)) && (this.Actions.Count == 1)) {
                this.SetExclusion(TypeOfExclusion.НеТребуетсяСоздавать);
                return;
            }

            switch (type) {
                case TypeOfReference.СвязующийОбъект:
                    this.CreateInListOfNomenclature();
                    this.GetType(type);
                    break;
                case TypeOfReference.НоменклатурныйОбъект:
                    // Случай, когда получилось найти объект сравочника Документы, и для него нужно создать номенклатуру
                    if (this.CreateInNomenclature()) {
                        this.GetType(type);
                    }
                    break;
                case TypeOfReference.ИсходныйОбъект:
                    // Случай, когда не удалось ничего найти, и нужно создать документ и к нему создать номенклатуру
                    if (this.CreateInDocumentsAndNomenclature()) {
                        this.GetType(type);
                        this.GetType(TypeOfReference.НоменклатурныйОбъект);
                    }
                    break;
                default:
                    throw new Exception(
                            string.Format(
                                "При попытке создания нового объекта возникла ошибка:\n" +
                                "Опция '{0}' не поддерживаетcя\n\n{1}",
                                type.ToString(),
                                this.Status
                                ));
                    break;
            }
        }

        #endregion CreateObject()

        #region CreateInListOfNomenclature()
        // Метод для создания нового объекта в справочнике 'Список номенклатуры'
        private void CreateInListOfNomenclature() {
            // Создаем новый объект
            try {
                this.LinkObject(
                        listOfNomenclatureReference.CreateReferenceObject(),
                        TypeOfReference.СвязующийОбъект,
                        StatusOfObject.Создан
                        );

                // Присваиваем ему обозначение и наименование
                this.ConnectionObject[shifrOfListNomenclature].Value = this.Shifr;
                this.ConnectionObject[nameOfListNomenclature].Value = this.Name;
                this.ConnectionObject.EndChanges();
            }
            catch (Exception e) {
                this.SetError(
                        TypeOfError.СозданиеОбъекта,
                        string.Format(
                            "При попытке создать объект в справочнике 'Список номенклатуры FoxPro' возникла ошибка:\n{0}",
                            e.Message
                            )
                        );
            }
            
            // TODO Здесь можно предусмотреть вариант подключения нового созданного объекта к объекту в Spec
        }

        #endregion CreateInListOfNomenclature()

        #region CreateInDocumentsAndNomenclature()

        // Метод для создания нового объекта в справочнике 'Документы'
        private bool CreateInDocumentsAndNomenclature() {
            
            // Обработка ошибочных сценариев
            string errorMessage = string.Empty;
            
            if (this.DocumentObject != null)
                errorMessage += "- документ уже существует\n";
            if (this.Type == TypeOfObject.НеОбработано)
                errorMessage += "- тип объекта не обработан\n";
            if (errorMessage != string.Empty)
                throw new Exception(string.Format(
                            "При попытке создать документ и связанный номенклатурный объект возникла ошибка:\n{0}\n\n{1}",
                            errorMessage,
                            this.Status
                            ));

            if (this.Type == TypeOfObject.НеОпределено) {
                this.MakeGuessTheType();
            }

            // В зависимости от того, какой тип объекта указан, создаем соотсвуетсвующий номенклатурный объект
            ClassObject typeOfNomenclature = null;
            Reference referenceToCreateObject = null;
            ParameterInfo shifrParameter = null;
            ParameterInfo nameParameter = null;
            
            switch (this.Type) {
                case TypeOfObject.Деталь:
                    typeOfNomenclature = documentReference.Classes.Find(Guids.Types.Документы.Деталь);
                    referenceToCreateObject = documentReference;
                    shifrParameter = shifrOfDocument;
                    nameParameter = nameOfDocument;
                    break;
                case TypeOfObject.СборочнаяЕдиница:
                    typeOfNomenclature = documentReference.Classes.Find(Guids.Types.Документы.СборочнаяЕдиница);
                    referenceToCreateObject = documentReference;
                    shifrParameter = shifrOfDocument;
                    nameParameter = nameOfDocument;
                    break;
                case TypeOfObject.СтандартноеИзделие:
                    typeOfNomenclature = documentReference.Classes.Find(Guids.Types.Документы.СтандартноеИзделие);
                    referenceToCreateObject = documentReference;
                    shifrParameter = shifrOfDocument;
                    nameParameter = nameOfDocument;
                    break;
                case TypeOfObject.ЭлектронныйКомпонент:
                    typeOfNomenclature = componentReference.Classes.Find(Guids.Types.ЭлектронныеКомпоненты.ЭлектронныйКомпонент);
                    referenceToCreateObject = componentReference;
                    shifrParameter = shifrOfComponent;
                    nameParameter = nameOfComponent;
                    break;
                case TypeOfObject.Другое:
                    typeOfNomenclature = documentReference.Classes.Find(Guids.Types.Документы.Другое);
                    referenceToCreateObject = documentReference;
                    shifrParameter = shifrOfDocument;
                    nameParameter = nameOfDocument;
                    break;
                case TypeOfObject.ПрочееИзделие:
                    typeOfNomenclature = documentReference.Classes.Find(Guids.Types.Документы.ПрочееИзделие);
                    referenceToCreateObject = documentReference;
                    shifrParameter = shifrOfDocument;
                    nameParameter = nameOfDocument;
                    break;
                case TypeOfObject.Изделие:
                    typeOfNomenclature = documentReference.Classes.Find(Guids.Types.Документы.Изделие);
                    referenceToCreateObject = documentReference;
                    shifrParameter = shifrOfDocument;
                    nameParameter = nameOfDocument;
                    break;
                default:
                    this.SetError(TypeOfError.НеопределенныйТип, string.Format(
                                "При попытке создания объекта возникла ошибка.\n" +
                                "Тип объекта {0} не поддерживается",
                                this.Type.ToString()
                                ));
                    break;
            }

            // Проверка на то, что удалось получить ClassObject
            if (typeOfNomenclature == null)
                return false;

            // Создаем документ
            try {
                this.LinkObject(
                        referenceToCreateObject.CreateReferenceObject(typeOfNomenclature),
                        TypeOfReference.ИсходныйОбъект,
                        StatusOfObject.Создан
                        );
                this.DocumentObject[shifrParameter].Value = this.Shifr;
                this.DocumentObject[nameParameter].Value = this.Name;
                this.DocumentObject.EndChanges();
            }
            catch (Exception e) {
                this.SetError(
                        TypeOfError.СозданиеОбъекта,
                        string.Format(
                            "Ошибка при создании исходного объекта для номенклатуры:\n{0}",
                            e.Message
                            )
                        );
            }

            // Создаем номенклатурный объект
            if (CreateInNomenclature()) {
                return true;
            }
            else {
                return false;
            }
        }

        #endregion CreateInDocumentsAndNomenclature()

        #region CreateInNomenclature()
        // Метод для создания нового объекта в справочнике 'ЭСИ' на основе найденного документа
        private bool CreateInNomenclature() {

            // Проверка наличия исходных данных
            if (this.DocumentObject == null)
                throw new Exception(string.Format(
                        "Ошибка при попытке создания номенклатурного объекта на основе существующего исходного объекта. Отсутствует исходный объект\n\n{0}",
                        this.Status
                        ));
        
            try {
                NomenclatureReference nomReference = nomenclatureReference as NomenclatureReference;
                this.LinkObject(
                        nomReference.CreateNomenclatureObject(this.DocumentObject),
                        TypeOfReference.НоменклатурныйОбъект,
                        StatusOfObject.Создан
                        );
            }
            catch (Exception e) {
                this.SetError(
                        TypeOfError.СозданиеОбъекта,
                        string.Format(
                            "Ошибка при подключении исходного документа к электронному составу изделия:\n{0}",
                            e.Message
                            )
                        );
                return false;
            }

            return true;
        }

        #endregion CreateInNomenclature()

        #endregion Методы по созданию объектов в справочниках

        #region RecreateObject()

        private void RecreateObject() {
            // Метод для повторного создания объекта

            // Обработка ошибок
            string errorMessage = string.Empty;
            if ((this.Type == TypeOfObject.НеОпределено) || (this.Type == TypeOfObject.НеОбработано))
                errorMessage += "- не указан тип для создания объекта\n";

            if (this.Type == TypeOfObject.Материал)
                errorMessage += string.Format("- тип {0} не поддерживается\n", this.Type);

            if (this.NomenclatureObject == null)
                errorMessage += "- отсутствует номенклатурный объект\n";

            if (errorMessage != string.Empty) {
                throw new Exception(string.Format(
                            "При попытке пересоздания объекта возникла ошибка\n{0}\n\n{1}",
                            errorMessage,
                            this.Status
                            ));
            }

            // Получаем все подключения текущего объекта для последующего клонирования их в новый объект
            LinksOfObject links = new LinksOfObject(this.NomenclatureObject);

            // Удаляем объект
            Desktop.CheckOut(this.NomenclatureObject, true);
            Desktop.CheckIn(this.NomenclatureObject, "Удаление объекта для его последующего пересоздания", false);
            Desktop.ClearRecycleBin(this.NomenclatureObject);

            // Приступаем к созданию объекта
            Reference reference = null;
            ClassObject classObject = null;
            Guid shifrGuid;
            Guid nameGuid;

            switch (this.Type)  {
                case TypeOfObject.ЭлектронныйКомпонент:
                    reference = componentReference;
                    classObject = dictTypesToClassObject[this.Type];
                    shifrGuid = Guids.Parameters.ЭлектронныеКомпоненты.Обозначение;
                    nameGuid = Guids.Parameters.ЭлектронныеКомпоненты.Наименование;
                    break;
                default:
                    reference = documentReference;
                    classObject = dictTypesToClassObject[this.Type];
                    shifrGuid = Guids.Parameters.Документы.Обозначение;
                    nameGuid = Guids.Parameters.Документы.Наименование;
                    break;
            }

            // Создаем новый объект
            this.LinkObject(
                    reference.CreateReferenceObject(classObject),
                    TypeOfReference.ИсходныйОбъект,
                    StatusOfObject.Пересоздан
                    );

            // Заполняем параметры объекта
            this.DocumentObject[shifrGuid].Value = this.Shifr;
            this.DocumentObject[nameGuid].Value = this.Name;
            this.DocumentObject.EndChanges();

            // Создаем номенклатурный объект и клонируем в него все подключения
            NomenclatureReference nomReference = nomenclatureReference as NomenclatureReference;
            this.LinkObject(
                    nomReference.CreateNomenclatureObject(this.DocumentObject),
                    TypeOfReference.НоменклатурныйОбъект,
                    StatusOfObject.Пересоздан
                    );

            links.CloneLinksTo(this.NomenclatureObject); // Применяем к новому объекту сохраненные ранее подключения

            // Сразу применяем изменения, для того, чтобы не потерять объект, если вдруг пользователь решит отменить изменения
            Desktop.CheckIn(this.DocumentObject, "Пересоздание объекта в процессе выгрузки изменений из SPEC", false);
        }

        #endregion RecreateObject()

        #region ToString()

        public string ToString() {
            return string.Format("\"{0}\" - \"{1}\"", this.Shifr, this.Name);
        }

        #endregion ToString()

        #region Validate()

        // Метод для проведения валидации найденных данных.
        // Данный метод будет дополняться, если придется проводить проверки на какие-то дополнительные параметры.
        // На данный момент проверка должна включать в себя проверку на то, чтобы обозначение у всех объектов было одинаковым
        public void Validate() {
            this.HasBeenValidated = true;
            // Позиции, которые содержат ошибки, или данные у которых неполные валидироваться не будут
            if ((this.HasError) || (this.HasExcluded) || (!this.IsDataComplete))
                return;

            // Проверка на то, что все объекты данной позиции имеют одинаковое обозначение
            if (
                    ((string)(this.ConnectionObject[Guids.Parameters.СписокНоменклатуры.Обозначение].Value) ==
                    (string)(this.NomenclatureObject[Guids.Parameters.Номенклатура.Обозначение].Value)) &&
                    ((string)(this.NomenclatureObject[Guids.Parameters.Номенклатура.Обозначение].Value) == 
                    (string)(this.DocumentObject[Guids.Parameters.Документы.Обозначение].Value))
                    ) {
                this.HasSameShifr = true;
            }

            // Проверка на то, что все объекта данной позиции имеют одинаковое наименование
            if (
                    ((string)(this.ConnectionObject[Guids.Parameters.СписокНоменклатуры.Наименование].Value) ==
                    (string)(this.NomenclatureObject[Guids.Parameters.Номенклатура.Наименование].Value)) &&
                    ((string)(this.NomenclatureObject[Guids.Parameters.Номенклатура.Наименование].Value) == 
                    (string)(this.DocumentObject[Guids.Parameters.Документы.Наименование].Value))
                    ) {
                this.HasSameName = true;
            }

            // TODO Рассмотреть, требуется ли валидирование записей на данный момент
        }

        #endregion Validate()
    }

    #endregion ObjectInTFlex

    #region LinksOfObject
    // Объект, который будет хранить информацию обо всех подключениях данного объекта
    private class LinksOfObject {
        public ReferenceObject InitialObject { get; private set; }
        public string Name { get; private set; }
        public string Denotation { get; private set; }
        public List<TflexHlink> ChildLinks { get; private set; } = new List<TflexHlink>();
        public List<TflexHlink> ParentLinks { get; private set; } = new List<TflexHlink>();


        #region constructors

        public LinksOfObject(ReferenceObject initialObject) {
            // Проверка на то, что объект не нулевой
            if (initialObject == null) {
                throw new Exception(
                        "При попытке получить подключения объекта возникла ошибка. " +
                        "Исходный объект отсутстует"
                        );
            }
            if ((initialObject as NomenclatureObject) == null) {
                throw new Exception(string.Format(
                            "При попытке получить подключения объекта '{0}' возникла ошибка. " +
                            "Исходный объект не является объектом 'ЭСИ'",
                            initialObject.ToString()
                        ));
            }

            this.Name = (string)initialObject[Guids.Parameters.Номенклатура.Наименование].Value;
            this.Denotation = (string)initialObject[Guids.Parameters.Номенклатура.Обозначение].Value;

            // Формируем перечень всех дочерних подключений
            double amount;
            int position;

            if (initialObject.Children != null) {
                foreach (var hlink in initialObject.Children.GetHierarchyLinks()) {
                    // Пробуем получить значения параметров
                    amount = hlink[Guids.hLinks.Количество].IsEmpty ? 1.0 : (double)hlink[Guids.hLinks.Количество].Value;
                    position = hlink[Guids.hLinks.Позиция].IsEmpty ? 0 : (int)hlink[Guids.hLinks.Позиция].Value;

                    this.ChildLinks.Add(new TflexHlink(hlink.ChildObject, position, amount));
                }
            }
            
            // Формируем перечень всех родительских подключений
            if (initialObject.Parents != null) {
                foreach (var hlink in initialObject.Parents.GetHierarchyLinks()) {
                    // Пробуем получить значения параметров
                    amount = hlink[Guids.hLinks.Количество].IsEmpty ? 1.0 : (double)hlink[Guids.hLinks.Количество].Value;
                    position = hlink[Guids.hLinks.Позиция].IsEmpty ? 0 : (int)hlink[Guids.hLinks.Позиция].Value;

                    this.ParentLinks.Add(new TflexHlink(hlink.ParentObject, position, amount));
                }
            }
        }

        #endregion constructors

        #region CloneLinksTo()

        public void CloneLinksTo(ReferenceObject newObject) {
            // Метод для создания подключений на новом объекте
            // Проверка на то, что целевой объект не нулевой
            if (newObject == null) {
                throw new Exception (string.Format(
                            "При попытке клонирования подключений с объекта '{0}' - '{1}' возникла ошибка. Целевой объект отсутствует",
                            this.Denotation,
                            this.Name
                            ));
            }

            // Проверка, что целевой объект является объектом номенклатуры
            if ((newObject as NomenclatureObject) == null) {
                throw new Exception (string.Format(
                            "При попытке клонирования подключений с объекта '{0}' - '{1}' возникла ошибка. " +
                            "Целевой объект '{2}' не является объектом ЭСИ",
                            this.Denotation,
                            this.Name,
                            newObject.ToString()
                            ));
            }


            // Создаем дочерние подключения
            foreach (TflexHlink link in this.ChildLinks) {
                ComplexHierarchyLink newLink = newObject.CreateChildLink(link.TargetObject);
                newLink[Guids.hLinks.Позиция].Value = link.Position;
                newLink[Guids.hLinks.Количество].Value = link.Amount;
                newLink.EndChanges();
            }

            // Создаем родительские подключения
            foreach (TflexHlink link in this.ParentLinks) {
                ComplexHierarchyLink newLink = newObject.CreateParentLink(link.TargetObject);
                newLink[Guids.hLinks.Позиция].Value = link.Position;
                newLink[Guids.hLinks.Количество].Value = link.Amount;
                newLink.EndChanges();
            }
        }

        #endregion CloneLinksTo()

        #region ReverseLinksDirection()

        public void ReverseLinksDirection() {
            // Метод нужен для исправления ошибочного применения ссылок
            (this.ParentLinks, this.ChildLinks) = (this.ChildLinks, this.ParentLinks);
        }

        #endregion ReverseLinksDirection()

        #region ToString()

        public override string ToString() {
            string result = string.Empty;
            if (this.ParentLinks.Count != 0) {
                result += string.Format("Объект '{0}' - '{1}' имеет следующие родительские подключения:\n", this.Denotation, this.Name);
                foreach (TflexHlink link in this.ParentLinks) {
                    result += string.Format("- {0}\n", link.ToString());
                }
                result += "\n";
            }
            if (this.ChildLinks.Count != 0) {
                result += string.Format("Объект '{0}' - '{1}' имеет следующие дочерние подключения:\n", this.Denotation, this.Name);
                foreach (TflexHlink link in this.ChildLinks) {
                    result += string.Format("- {0}\n", link.ToString());
                }
            }
            return result;
        }

        #endregion ToString()
    }

    #region class TflexHlink
    // Объект для хранения необходимой информации о подключении
    private class TflexHlink {
        public ReferenceObject TargetObject { get; private set; }
        public int Position { get; private set; }
        public double Amount { get; private set; }

        public TflexHlink (ReferenceObject target, int position, double amount) {
            this.TargetObject = target;
            this.Position = position;
            this.Amount = amount;
        }

        public override string ToString() {
            return string.Format(
                    "к объекту '{0}' в количестве '{1}' штук (позиция - '{2}')",
                    this.TargetObject.ToString(),
                    this.Amount.ToString(),
                    this.Position.ToString()
                    );
        }
    }

    #endregion class TflexHlink

    #endregion LinksOfObject

    #region SpecRecordsContainer

    private class SpecRecordsContainer {

        #region Fields and Properties

        public Dictionary<int, Dictionary<string, string>> OldData { get; private set; }
        public Dictionary<int, Dictionary<string, string>> NewData { get; private set; }

        private Dictionary<int, string> ErrorContainer { get; set; }
        private Dictionary<int, string> CompleteContainer { get; set; }

        public int CountError => ErrorContainer.Count;
        public int CountComplete => CompleteContainer.Count;


        private List<int> IDs { get; set; }

        public List<int> IdToCreate { get; private set; } = new List<int>();
        public List<int> IdToChange { get; private set; } = new List<int>();
        public List<int> IdToDelete { get; private set; } = new List<int>();

        public string[] ShifrsOnCreate { get; private set; }
        public string[] ShifrsOnChange { get; private set; }
        public string[] ShifrsOnDelete { get; private set; }

        #endregion Fields and Properties

        #region Constructors

        public SpecRecordsContainer(
                Dictionary<int, Dictionary<string, string>> oldData,
                Dictionary<int, Dictionary<string, string>> newData
                ) {
            
            this.OldData = oldData;
            this.NewData = newData;

            this.ErrorContainer = new Dictionary<int, string>();
            this.CompleteContainer = new Dictionary<int, string>();

            // Производим обработку входной информации
            FilterEmptyRows(); // Производим фильтрацию пустых строк
            NormalizeShifrsInInputData(); // Проставляем точки в обозначениях
            GetAllIDs(); // Получаем полный перечень id записей, которые участвуют в выгрузке
            DistributeIDsByOperations(); // Распределяем записи по назначению (создание, изменение, удаление)
            GetAllShifrs(); // Получаем все шифры для поиска объектов в DOCs
        }

        #endregion Constructors

        #region Public Methods

        #region AddErrorRecord()

        public void AddErrorRecord(int id, TypeOfAction action, TypeOfLinkError error, string errorMessage = "") {

            if (errorMessage != string.Empty)
                errorMessage = string.Format("Текст ошибки:\n", errorMessage);
            
            string templateMessage = "{0} Error:\n{1}\nУчастники:\n{2}\n{3}";
            string method = string.Empty;
            string description = string.Empty;
            string participatingObjects = string.Empty;

            switch (action) {
                case TypeOfAction.Создание:
                    method = "ApplyAddingsToTflex";
                    break;
                case TypeOfAction.Редактирование:
                    method = "ApplyChangesToTflex";
                    break;
                case TypeOfAction.Удаление:
                    method = "ApplyDeletingToTflex";
                    break;
                default:
                    throw new Exception(string.Format("Метод AddErrorRecord не поддерживает '{0}'", action.ToString()));
            }

            switch (error) {
                case TypeOfLinkError.ОтсутствуетОбъект:
                    description =
                        "Один или несколько объектов, необходимых для импорта подкючений отсутствует";
                    break;
                case TypeOfLinkError.ОбъектИсключен:
                    description =
                        "Один или несколько объектов были исключены из обработки";
                    break;
                case TypeOfLinkError.ОбъектСодержитОшибку:
                    description =
                        "Один или несколько объектов, необходимых для импорта подключений содержит ошибку";
                    break;
                case TypeOfLinkError.ОшибкаВПроцессеПодключения:
                    description =
                        "В процессе импортирования подключения возникла ошибка";
                    break;
                default:
                    throw new Exception(string.Format("Метод AddErrorRecord не поддерживает '{0}'", error.ToString()));
            }

            if (this.IdToChange.Contains(id)) {
                participatingObjects = string.Format(
                        "Старый потомок: '{0}'; Старый родитель: '{1}'; Новый потомок: '{2}' Новый родитель: '{3}';",
                        this.OldData[id]["SHIFR"],
                        this.OldData[id]["IZD"],
                        this.NewData[id]["SHIFR"],
                        this.NewData[id]["IZD"]
                        );
            }
            else if (this.IdToCreate.Contains(id)) {
                participatingObjects = string.Format(
                        "потомок: '{0}'; родитель: '{1}';",
                        this.NewData[id]["SHIFR"],
                        this.NewData[id]["IZD"]
                        );
            }
            else if (this.IdToDelete.Contains(id)) {
                participatingObjects = string.Format(
                        "потомок: '{0}'; родитель: '{1}';",
                        this.OldData[id]["SHIFR"],
                        this.OldData[id]["IZD"]
                        );
            }
            else
                throw new Exception(string.Format("В процессе работы метода AddErrorRecord не удалось найти id: {0}", id));

            this.ErrorContainer[id] = string.Format(templateMessage, method, description, participatingObjects, errorMessage);

        }

        #endregion AddErrorRecord()

        #region AddCompleteRecord()

        public void AddCompleteRecord(int id, TypeOfAction action) {
            string templateMessage = "{0} Success:\n{1}\n";
            string method = string.Empty;
            string description = string.Empty;
            string templateDescription = string.Empty;

            switch (action) {
                case TypeOfAction.Создание:
                    method = "ApplyAddingsToTflex";
                    templateDescription = "Произведено добавление связи '{0}' -> '{1}' (поз. {2}; кол. {3})";
                    description = string.Format(
                            templateDescription,
                            this.NewData[id]["SHIFR"],
                            this.NewData[id]["IZG"],
                            this.NewData[id]["POS"],
                            this.NewData[id]["PRIM"]
                            );
                    break;
                case TypeOfAction.Редактирование:
                    method = "ApplyChangesToTflex";
                    templateDescription = 
                        "Произведено корректировка связи с '{0}' -> '{1}' на '{2}' '{3}' (поз. с {4} на {5}; кол. с {6} на {7})";
                    description = string.Format(
                            templateDescription,
                            this.OldData[id]["SHIFR"],
                            this.OldData[id]["IZD"],
                            this.NewData[id]["SHIFR"],
                            this.NewData[id]["IZD"],
                            this.OldData[id]["POS"],
                            this.NewData[id]["POS"],
                            this.OldData[id]["PRIM"],
                            this.NewData[id]["PRIM"]
                            );
                    break;
                case TypeOfAction.Удаление:
                    method = "ApplyDeletingToTflex";
                    templateDescription = "Произведено удаление связи '{0}' -> '{1}' (поз. {2}; кол. {3})";
                    description = string.Format(
                            templateDescription,
                            this.OldData[id]["SHIFR"],
                            this.OldData[id]["IZG"],
                            this.OldData[id]["POS"],
                            this.OldData[id]["PRIM"]
                            );
                    break;
                default:
                    throw new Exception(string.Format("Метод AddCompleteRecord не поддерживает {0}", action.ToString()));
            }

            this.CompleteContainer[id] = string.Format(templateMessage, method, description);
        }

        #endregion AddCompleteRecord()

        #region ToString methods

        public string ToStringErrors() {
            string template = "ID: {0} Message:\n {1}";
            return string.Join("\n\n", this.ErrorContainer.Select(kvp => string.Format(template, kvp.Key, kvp.Value)));
        }

        public string ToStringComplete() {
            string template = "ID: {0} Message:\n {1}";
            return string.Join("\n\n", this.CompleteContainer.Select(kvp => string.Format(template, kvp.Key, kvp.Value)));
        }

        #endregion ToString methods

        #endregion Public Methods

        #region Private Methods

        #region FilterEmptyRows()

        private void FilterEmptyRows() {
            this.OldData = this.OldData.Where(kvp => kvp.Value["SHIFR"] != string.Empty).ToDictionary(kvp => kvp.Key, kvp => kvp.Value);
            this.NewData = this.NewData.Where(kvp => kvp.Value["SHIFR"] != string.Empty).ToDictionary(kvp => kvp.Key, kvp => kvp.Value);
        }

        #endregion FilterEmptyRows()

        #region NormalizeShifrsInInputData()

        private void NormalizeShifrsInInputData() {
            foreach (KeyValuePair<int, Dictionary<string, string>> kvp in this.OldData) {
                kvp.Value["SHIFR"] = NormalizeShifrs(kvp.Value["SHIFR"]);
                kvp.Value["IZD"] = NormalizeShifrs(kvp.Value["IZD"]);
            }
            foreach (KeyValuePair<int, Dictionary<string, string>> kvp in this.NewData) {
                kvp.Value["SHIFR"] = NormalizeShifrs(kvp.Value["SHIFR"]);
                kvp.Value["IZD"] = NormalizeShifrs(kvp.Value["IZD"]);
            }
        }

        #endregion NormalizeShifrsInInputData()

        #region NormalizeShifrs()

        private string NormalizeShifrs(string valueString) {
            // Список исключений, в которых не нужно производить нормализации
            if (valueString.ToLower().Contains("северо"))
                return valueString;

            // Создаем переменную для того, чтобы можно было использовать метод int.TryParse()
            long valueInt;

            // Узнаем, цифра ли содержится в значении
            bool isNum = long.TryParse(valueString, out valueInt);

            // Если входные данные не являются числовым значением и содержат в себе символ "-", и часть строки до тире состоит из шести или из семи символов
            if (!isNum && valueString.IndexOf('-') > 0 && (valueString.Substring(0, valueString.IndexOf('-')).Length == 6 || valueString.Substring(0, valueString.IndexOf('-')).Length == 7)) {
                isNum = long.TryParse(valueString.Substring(0, valueString.IndexOf('-')), out valueInt);
                valueString = valueString.Insert(3, ".");
            }

            // Случай, когда по определению обозначение не состоит из цифровых значений
            if (valueString.StartsWith("УЯИС") || valueString.StartsWith("ШЖИФ") || valueString.StartsWith("УЖИЯ")) {
                valueString = valueString.Insert(4, ".").Insert(11, ".");
            }

            if (valueString.StartsWith("3905")) {
                valueString = valueString.Insert(4, ".");
                return valueString;
            }

            if (valueString.StartsWith("8А") && valueString.Length > 7) {
                valueString = valueString.Insert(3, ".").Insert(7, ".");
            }

            // В данном случае рассматривается и вариант, когда значение полностью состоит из цифр, и вариант, когда в обозначении присутствовало тире
            if (isNum && (valueString.Length == 6 || valueString.Length == 7)) {
                valueString = valueString.Insert(3, ".");
            }

            return valueString;
        }

        #endregion NormalizeShifrs()

        #region GetAllIDs()

        private void GetAllIDs() {
            List<int> result = this.OldData.Select(kvp => kvp.Key).ToList<int>();
            result.AddRange(this.NewData.Select(kvp => kvp.Key).ToList<int>());
            this.IDs = result.Distinct().OrderBy(row => row).ToList<int>();
        }

        #endregion GetAllIDs()

        #region DistributeIDsByOperations()

        private void DistributeIDsByOperations() {
            foreach (int id in this.IDs) {
                // Если запись содержится в двух словаряе, значит ее нужно корректировать
                if ((this.OldData.ContainsKey(id)) && (this.NewData.ContainsKey(id)))
                    this.IdToChange.Add(id);
                // Запись отсутствует в новых данных, следовательно ее нужно удалить
                else if (!this.NewData.ContainsKey(id))
                    this.IdToDelete.Add(id);
                // Запись отсутствует в старых данных, следовательно ее нужно  создать
                else
                    this.IdToCreate.Add(id);
            }
        }

        #endregion DistributeIDsByOperations()

        #region GetAllShifrs()

        private void GetAllShifrs() {
            this.ShifrsOnCreate = GetShifrs(this.IdToCreate);
            this.ShifrsOnChange = GetShifrs(this.IdToChange);
            this.ShifrsOnDelete = GetShifrs(this.IdToDelete);
        }

        #endregion GetAllShifrs()

        #region GetShifrs()

        private string[] GetShifrs(List<int> ids) {

            var sliceOldData = new Dictionary<int, Dictionary<string, string>>();
            var sliceNewData = new Dictionary<int, Dictionary<string, string>>();
            foreach (int id in ids) {
                if (this.OldData.ContainsKey(id))
                    sliceOldData.Add(id, this.OldData[id]);
                if (this.NewData.ContainsKey(id))
                    sliceNewData.Add(id, this.NewData[id]);
            }

            // Получаем все обозначения, которые есть в таблице oldData
            var allShifrsInOldData = sliceOldData
                .Select(row => row.Value["SHIFR"])
                .Union(sliceOldData.Select(row => row.Value["IZD"]));

            var allShifrsInNewData = sliceNewData
                .Select(row => row.Value["SHIFR"])
                .Union(sliceNewData.Select(row => row.Value["IZD"]));

            return allShifrsInOldData.Union(allShifrsInNewData).OrderBy(row => row).ToArray();
        }

        #endregion GetShifrs()

        #endregion Private Methods

    }

    #endregion SpecRecordsContainer

    #endregion Gukov Classes

    #region Markin Classes

    public class Param
    {
        public Dictionary<int, Dictionary<string, string>> getrowadd;
        public Guid guidrefer;
        public Guid guidparametr;
        public string filename;
    }


    public class RefTab
    {
        public readonly string table_name;
        public readonly string name_refer;
        public readonly string nameTip;
        public readonly string name_row_num;
        public readonly string quertystr;
        public readonly string roleimportdbf;
        public readonly Guid guid;
        public readonly Guid ROW_NUMBER;
        

        //public RefTab(string refname)
        public RefTab(TableName tableName)
        {
            switch (tableName)
            {
                case TableName.SPEC:
                    table_name = Tab.SPEC.table_name;
                    name_refer = Tab.SPEC.name_refer;
                    nameTip = Tab.SPEC.nameTip;
                    name_row_num = Tab.SPEC.name_row_num;
                    guid = Tab.SPEC.guid;
                    ROW_NUMBER = Tab.SPEC.ROW_NUMBER;
                    quertystr = Tab.SPEC.quertystr;
                    roleimportdbf = Tab.SPEC.RoleChange.RoleImportDbf;
                    break;
                case TableName.NORM:
                    table_name = Tab.NORM.table_name;
                    name_refer = Tab.NORM.name_refer;
                    nameTip = Tab.NORM.nameTip;
                    name_row_num = Tab.NORM.name_row_num;
                    guid = Tab.NORM.guid;
                    ROW_NUMBER = Tab.NORM.ROW_NUMBER;
                    quertystr = Tab.NORM.quertystr;
                    roleimportdbf = Tab.NORM.RoleChange.RoleImportDbf;
                    break;
                case TableName.KLASM:
                    table_name = Tab.KLASM.table_name;
                    name_refer = Tab.KLASM.name_refer;
                    nameTip = Tab.KLASM.nameTip;
                    name_row_num = Tab.KLASM.name_row_num;
                    guid = Tab.KLASM.guid;
                    ROW_NUMBER = Tab.KLASM.ROW_NUMBER;
                    quertystr = Tab.KLASM.quertystr;
                    roleimportdbf = Tab.KLASM.RoleChange.RoleImportDbf;
                    break;
                case TableName.KLAS:
                    table_name = Tab.KLAS.table_name;
                    name_refer = Tab.KLAS.name_refer;
                    nameTip = Tab.KLAS.nameTip;
                    name_row_num = Tab.KLAS.name_row_num;
                    guid = Tab.KLAS.guid;
                    ROW_NUMBER = Tab.KLAS.ROW_NUMBER;
                    quertystr = Tab.KLAS.quertystr;
                    roleimportdbf = Tab.KLAS.RoleChange.RoleImportDbf;
                    break;
                case TableName.MARCHP:
                    table_name = Tab.MARCHP.table_name;
                    name_refer = Tab.MARCHP.name_refer;
                    nameTip = Tab.MARCHP.nameTip;
                    name_row_num = Tab.MARCHP.name_row_num;
                    guid = Tab.MARCHP.guid;
                    ROW_NUMBER = Tab.MARCHP.ROW_NUMBER;
                    quertystr = Tab.MARCHP.quertystr;
                    roleimportdbf = Tab.MARCHP.RoleChange.RoleImportDbf;
                    break;
                case TableName.TRUD:
                    table_name = Tab.TRUD.table_name;
                    name_refer = Tab.TRUD.name_refer;
                    nameTip = Tab.TRUD.nameTip;
                    name_row_num = Tab.TRUD.name_row_num;
                    guid = Tab.TRUD.guid;
                    ROW_NUMBER = Tab.TRUD.ROW_NUMBER;
                    quertystr = Tab.TRUD.quertystr;
                    roleimportdbf = Tab.TRUD.RoleChange.RoleImportDbf;
                    break;
                case TableName.KAT_EDIZ:
                    table_name = Tab.KAT_EDIZ.table_name;
                    name_refer = Tab.KAT_EDIZ.name_refer;
                    name_row_num = Tab.KAT_EDIZ.name_row_num;
                    guid = Tab.KAT_EDIZ.guid;
                    ROW_NUMBER = Tab.KAT_EDIZ.ROW_NUMBER;
                    roleimportdbf = Tab.KAT_EDIZ.RoleChange.RoleImportDbf;
                    break;
                case TableName.KAT_INS:
                    table_name = Tab.KAT_INS.table_name;
                    name_refer = Tab.KAT_INS.name_refer;
                    name_row_num = Tab.KAT_INS.name_row_num;
                    guid = Tab.KAT_INS.guid;
                    ROW_NUMBER = Tab.KAT_INS.ROW_NUMBER;
                    roleimportdbf = Tab.KAT_INS.RoleChange.RoleImportDbf;
                    break;
                case TableName.KAT_IZD:
                    table_name = Tab.KAT_IZD.table_name;
                    name_refer = Tab.KAT_IZD.name_refer;
                    name_row_num = Tab.KAT_IZD.name_row_num;
                    guid = Tab.KAT_IZD.guid;
                    ROW_NUMBER = Tab.KAT_IZD.ROW_NUMBER;
                    roleimportdbf = Tab.KAT_IZD.RoleChange.RoleImportDbf;
                    break;
                case TableName.KAT_STOL:
                    table_name = Tab.KAT_STOL.table_name;
                    name_refer = Tab.KAT_STOL.name_refer;
                    name_row_num = Tab.KAT_STOL.name_row_num;
                    guid = Tab.KAT_STOL.guid;
                    ROW_NUMBER = Tab.KAT_STOL.ROW_NUMBER;
                    roleimportdbf = Tab.KAT_STOL.RoleChange.RoleImportDbf;
                    break;
                case TableName.Poluf:
                    table_name = Tab.Poluf.table_name;
                    name_refer = Tab.Poluf.name_refer;
                    name_row_num = Tab.Poluf.name_row_num;
                    guid = Tab.Poluf.guid;
                    ROW_NUMBER = Tab.Poluf.ROW_NUMBER;
                    roleimportdbf = Tab.Poluf.RoleChange.RoleImportDbf;
                    break;
                case TableName.OSNAST:
                    table_name = Tab.OSNAST.table_name;
                    name_refer = Tab.OSNAST.name_refer;
                    name_row_num = Tab.OSNAST.name_row_num;
                    guid = Tab.OSNAST.guid;
                    ROW_NUMBER = Tab.OSNAST.ROW_NUMBER;
                    roleimportdbf = Tab.OSNAST.RoleChange.RoleImportDbf;
                    break;
                case TableName.KAT_RAZM:
                    table_name = Tab.KAT_RAZM.table_name;
                    name_refer = Tab.KAT_RAZM.name_refer;
                    name_row_num = Tab.KAT_RAZM.name_row_num;
                    guid = Tab.KAT_RAZM.guid;
                    ROW_NUMBER = Tab.KAT_RAZM.ROW_NUMBER;
                    roleimportdbf = Tab.KAT_RAZM.RoleChange.RoleImportDbf;
                    break;
            }


        }

    }

    private static class Tab
    {
        public static class SPEC
        {
            public static readonly string table_name = "spec.dbf";
            public static readonly string name = "spec.dbf";
            public static readonly string name_refer = "SPEC";
            public static readonly string nameTip = "SPEC";
            public static readonly string name_row_num = "ROW_NUMBER";
            public static readonly string quertystr = "SELECT \"ROW_NUMBER\", \"POS\",\"SHIFR\",\"NAIM\",\"PRIM\",\"IZD\" FROM dbo.\"SPEC\"";
            public static readonly List<string> column = new List<string> { "ROW_NUMBER", "POS", "SHIFR", "NAIM", "PRIM", "IZD" };


            public static readonly Guid guid = new Guid("c587b002-3be2-46de-861b-c551ed92c4c1");
            public static readonly Guid ROW_NUMBER = new Guid("b6c7f1e2-d94c-4fde-966b-937a9979813f");
            public static readonly Guid SHIFR = new Guid("2a855e3d-b00a-419f-bf6f-7f113c4d62a0");
            public static readonly Guid IZD = new Guid("817eca21-1e7e-46e0-b80a-e61685bef5f7");
            public static class Link
            {
                public static readonly Guid LinkПодключения = new Guid("a3973f89-fdb9-47a3-87ef-536e91224e38"); //standart (tdocs)
                                                                                                                //public static readonly Guid LinkПодключения = new Guid("cfc46bc9-e931-4f7c-bc91-de29a2da4635"); //mod (tflex-docs)
                public static readonly Guid LinkСписокНоменклатуры = new Guid("a356fe70-1a77-4e8a-9a90-38ba118431a8");
                public static readonly Guid LinkСписокНоменклатурыСборка = new Guid("86110e2f-0be5-414a-a98b-a925a201e9c9");
            }
            public static class RoleChange
            {

                public static readonly string RoleListNum = "fba9104c-14ac-4c6a-a1eb-9823f717a1df"; //"Список номенклатуры SPEC"; 
                public static readonly string RoleListNumIzd = "03749ea7-bd7d-48cd-8cad-bff6dab65ce5"; //"Список номенклатуры SPEC для сборок";
                public static readonly string RoleConnect = "9c10c0b3-0c9c-4b7b-ace5-2fee7132edaf"; //"Подключения SPEC";
           //   public static readonly string RoleConnect = "b57e28ab-297c-439e-a5ab-d687d0effb48"; //"Подключения SPEC_v2";
                public static readonly string RoleImportDbf = "5c84fa1b-331a-4b1d-a7d1-898e39a6da6d"; //Структура изделий SPEC
            }

        }


        public static class NORM
        {
            public static readonly string table_name = "norm.dbf";
            public static readonly string name_refer = "NORM";
            public static readonly string nameTip = "NORM";
            public static readonly string name_row_num = "Номер записи";
            public static readonly List<string> column = new List<string> { "Nomer_zapisi", "SHIFR","OKP","EDIZM","NMAT","NVES","NOTH","NPOT","POKR_S","POKR_H","MAS"};
            //public static readonly string quertystr = $"SELECT \"Nomer_zapisi\" , b.\"SHIFR\" || b.\"OKP\" || b.\"EDIZM\"::decimal(3, 0) || b.\"NMAT\"::decimal(11, 2) || b.\"NVES\"::decimal(10, 4) || b.\"NOTH\"::decimal(10, 4)  || b.\"NPOT\"::decimal(10, 4)  || b.\"POKR_S\"::decimal(12, 3)  || b.\"POKR_H\"::decimal(3, 0)  || b.\"MAS\"::decimal(11, 3) as NORM FROM dbo.\"NORM\" as b";
            public static readonly string quertystr = $"SELECT \"Nomer_zapisi\" , b.\"SHIFR\",b.\"OKP\" ,b.\"EDIZM\"::decimal(3, 0) , round(b.\"NMAT\"::decimal(10, 4),3) , b.\"NVES\"::decimal(10, 4) , b.\"NOTH\"::decimal(10, 4) , b.\"NPOT\"::decimal(10, 4) , b.\"POKR_S\"::decimal(12, 3) , b.\"POKR_H\"::decimal(3, 0) , b.\"MAS\"::decimal(11, 3) as normanalog FROM dbo.\"NORM\" as b ORDER BY \"Nomer_zapisi\"";
            //public static readonly string quertystr = $"SELECT \"Nomer_zapisi\" , b.\"SHIFR\",b.\"OKP\" ,b.\"EDIZM\"::decimal(3, 0) ,'|' || b.\"NMAT\"::decimal(11, 3) || '|', b.\"NVES\"::decimal(10, 4) , b.\"NOTH\"::decimal(10, 4) , b.\"NPOT\"::decimal(10, 4) , b.\"POKR_S\"::decimal(12, 3) , b.\"POKR_H\"::decimal(3, 0) , b.\"MAS\"::decimal(11, 3) as normanalog FROM dbo.\"NORM\" as b ORDER BY \"Nomer_zapisi\"";
            //b."SHIFR"||'|'||b."OKP"||'|'||b."EDIZM"::decimal(3,0)||'|'||b."NMAT"::decimal(11,2) ||'|'||b."NVES"::decimal(10,4)||'|'||b."NOTH"::decimal(10,4)||'|'||b."NPOT"::decimal(10,4)||'|'||b."POKR_S"::decimal(12,3)||'|'||b."POKR_H"::decimal(3,0)||'|'||b."MAS"::decimal(11,3) as normanalog
            public static readonly Guid guid = new Guid("7e59b31a-eab5-4da5-938f-812703046345");
            public static readonly Guid ROW_NUMBER = new Guid("46a081c5-d993-4c60-af71-162bb11c1927");
            public static class Link
            {
                public static readonly Guid LinkПодключения = new Guid("5065e6bd-77c6-492f-a8f1-0f19a709dd21");
            }

            public static class RoleChange
            {
                public static readonly string RoleConnect = "2310db86-d1c8-46c3-9d6e-c2e740e678ba"; //"Подключения NORM";
                public static readonly string RoleImportDbf = "32158420-541b-45a8-8c46-2545909a29a2"; //Нормы материалов NORM
            }
        }

        public static class KLASM
        {
            public static readonly string table_name = "klasm.dbf";
            public static readonly string name_refer = "KLASM";
            public static readonly string nameTip = "KLASM";
            public static readonly string name_row_num = "Row_num";
            public static readonly string quertystr = $"SELECT k.\"Nomer_zapisi\",  k.\"OKP\" , k.\"NAME_2\" , k.\"GOST\" , k.\"EDIZM\" , k.\"SKLAD\" , k.\"PRAC\",k.\"PRACD\",k.\"PRACE\", k.\"MARKA\", k.\"VID\" as KLASM FROM dbo.\"KLASM\" as k ORDER BY  k.\"Nomer_zapisi\"";
            public static readonly Guid guid = new Guid("e253835e-260e-45f3-8ae8-600a7a8907a2");
            public static readonly Guid ROW_NUMBER = new Guid("9ec4fb11-0be1-4602-931a-18b479d9ac94");
            public static class Link
            {
                public static readonly Guid LinkСписокноменклатурыFoxPro = new Guid("27afe853-380c-4178-8841-b1eb9ca28c6c");
                public static readonly Guid LinkМатериал = new Guid("2031c179-5304-411f-bf35-902d600e6650");
            }

            public static class RoleChange
            {
                public static readonly string RoleListNum = "871f83bd-bd23-4424-9486-f796c2b98ba1"; //"Список номенклатуры KLASM"; //
                public static readonly string RoleKLASMMaterial = "88c9749b-5f0a-408d-bf62-989ca641672a"; //"Создание материалов из аналогов 2";
                public static readonly string RoleImportDbf = "ded1e3a7-b189-4b41-9b93-e69468d89966"; //Материалы, Номенклатура и изделия с типом Материал KLASM


            }

        }

        public static class KLAS
        {
            public static readonly string table_name = "klas.dbf";
            public static readonly string name = "klas.dbf";
            public static readonly string name_refer = "KLAS";
            public static readonly string nameTip = "KLAS";
            public static readonly string name_row_num = "ROW_NUMBER";
            //public static readonly string quertystr = $"SELECT b.\"ROW_NUMBER\", b.\"OKP\" || b.\"NAME_2\" || b.\"GOST\" || b.\"EDIZM\" || b.\"SKLAD\"::decimal(7, 0) || b.\"PRAC\" || b.\"PRACE\"::decimal(10, 2) || b.\"MARKA\" || b.\"VID\" as KLAS FROM dbo.\"KLAS\" b";
            public static readonly string quertystr = $"SELECT b.\"ROW_NUMBER\", b.\"OKP\" , b.\"NAME_2\" , b.\"GOST\" , b.\"EDIZM\" , b.\"SKLAD\"::decimal(7, 0) , b.\"PRAC\", b.\"PRACD\" , b.\"PRACE\"::decimal(10, 2) , b.\"MARKA\" , b.\"VID\" as KLAS FROM dbo.\"KLAS\" b";
            //public static readonly string quertystr = $"SELECT b.\"ROW_NUMBER\", b.\"OKP\" , b.\"NAME_2\" , b.\"GOST\" , b.\"EDIZM\" , b.\"SKLAD\"::decimal(7, 0) , b.\"PRAC\", \'\' as PRACD , b.\"PRACE\"::decimal(10, 2) , b.\"MARKA\" , b.\"VID\" as KLAS FROM dbo.\"KLAS\" b";
            public static readonly Guid guid = new Guid("443d8e71-1385-459d-879e-5e03c19b9d49");
            public static readonly Guid ROW_NUMBER = new Guid("41439758-0dbb-4d2d-8296-fdd8376ffb89");
            public static class Link
            {
                public static readonly Guid LinkСписокноменклатурыFoxPro = new Guid("b1eb7e7c-19be-4688-8a00-c2cbb366f1d9");
            }

            public static class RoleChange
            {
                public static readonly string RoleListNum = "8687ef5a-fda8-4288-9b2c-d3538b05a2a6"; //"Список номенклатуры KLAS";                                                                
                public static readonly string RoleImportDbf = "1f2d7e0e-fc21-4c6b-bb47-1be400b2bcc7"; //Номенклатура и изделия, типы Стандартные изделия и Прочие изделия KLAS
            }

        }


        public static class TRUD
        {
            public static readonly string table_name = "trud.dbf";
            public static readonly string name_refer = "TRUD";
            public static readonly string nameTip = "TRUD";
            public static readonly string name_row_num = "Row_num";
            public static readonly string quertystr = $"select \"Row_num\",\"SHIFR\",\"IZG\",\"NUM_OP\",\"SHIFR_OP\",\"RAZR\",\"NORM_T\",\"NORM_M\",\"DATA_OP\",\"K_OPER\",\"NPER\",\"VREM\",trim(\"OP_OP\") as \"OP_OP\",\"NAIM_ST\",\"PROF\",\"OSN_TARA\",\"SK_TXT\",\"NORM_T_P\",\"NORM_M_P\"  from dbo.\"TRUD\" order by \"Row_num\"";
            public static readonly Guid guid = new Guid("8324a9a7-51a9-4946-bde2-31fb458edd64");
            public static readonly Guid ROW_NUMBER = new Guid("7eb529fd-8d7a-4f9b-a1f0-1e9fd6c1be60");
            /*            public static class Link
                        {
             
                        }
            */
            public static class RoleChange
            {
                public static readonly string RoleImportDbf = "29e13c68-f108-4192-92f9-19524e37cf12"; //Аналог Trud_dbf
            }
            
        }


        public static class MARCHP
        {
            public static readonly string table_name = "marchp.dbf";
            public static readonly string name_refer = "MARCHP";
            public static readonly string nameTip = "MARCHP";
            public static readonly string name_row_num = "Row_Number";
            public static readonly string quertystr = $"SELECT \"Row_Number\", \"SHIFR\", \"NPER\", \"IZG\",\"PER1\", \"POTR\", \"PER2\", \"SDAT\", \"VREM\", \"NORM\"  FROM dbo.\"MARCHP\"";
            public static readonly Guid guid = new Guid("dda8e2f7-67e9-41ea-a8db-cbe6c9b7eb9b");
            public static readonly Guid ROW_NUMBER = new Guid("6214f997-8450-4b3e-ad71-f9631e102a54");
            /*            public static class Link
                        {

                        }
            */
            public static class RoleChange
            {
                public static readonly string RoleImportDbf = "c001ceff-1e30-414f-a76e-2e78282d73e5";
            }
            
        }

        public static class KAT_IZD
        {
            public static readonly string table_name = "kat_ediz.dbf";
            public static readonly string name_refer = "KAT_IZD";
            public static readonly string nameTip = "KAT_IZD";
            public static readonly string name_row_num = "ROW_NUMBER";
            public static readonly Guid guid = new Guid("a7b0d67a-9da3-441b-aad4-e6d951a188f4");
            public static readonly Guid ROW_NUMBER = new Guid("14991c7e-f30a-47f2-9206-f4baeb253d27");
            /*            public static class Link
            {

            }
          */
            public static class RoleChange
            {
                public static readonly string RoleImportDbf = "15c6d1c1-d38d-4bc9-b098-c7795a85e5a5";
                public static readonly string RoleListNum = "d92d152b-5ae6-4c61-8031-a2a34875cad9";  //Список номенклатуры KAT_IZD

            }
              
        }


        public static class KAT_INS
        {
            public static readonly string table_name = "kat_ins.dbf";
            public static readonly string name_refer = "KAT_INS";
            public static readonly string nameTip = "KAT_INS";
            public static readonly string name_row_num = "ROW_NUMBER";
            public static readonly Guid guid = new Guid("748a9ca7-931e-4db3-88da-4f42d9c0157b");
            public static readonly Guid ROW_NUMBER = new Guid("eae6a018-f628-41f2-94da-628fa47e792d");
            /*            public static class Link
            {

            }
          */
            public static class RoleChange
            {
                public static readonly string RoleImportDbf = "f97726de-fbbd-4781-b7b3-288483bbabdd"; //"Единицы измерения Kat_ediz";                                                                
            }

        }


        public static class KAT_EDIZ
        {
            public static readonly string table_name = "kat_ediz.dbf";
            public static readonly string name_refer = "KAT_EDIZ";
            public static readonly string nameTip = "KAT_EDIZ";
            public static readonly string name_row_num = "ROW_NUMBER";
            public static readonly Guid guid = new Guid("22ee7b25-0744-405c-986e-01d62bfed86d");
            public static readonly Guid ROW_NUMBER = new Guid("6cd153d5-aed0-4d3d-b9de-59a37190dce1");
            /*            public static class Link
            {

            }
            */
            public static class RoleChange
            {
                public static readonly string RoleImportDbf = "4f3db9b5-eaac-4562-a956-8befcfc12e5c"; //"Единицы измерения Kat_ediz";                                                                
                
            }

        }

        public static class OSNAST
        {
            public static readonly string table_name = "OSNAST.dbf";
            public static readonly string name_refer = "OSNAST";
            public static readonly string nameTip = "OSNAST";
            public static readonly string name_row_num = "ROW_NUMBER";
            public static readonly Guid guid = new Guid("e557245a-88eb-4fcd-93d4-7ee47ce9bd07");
            public static readonly Guid ROW_NUMBER = new Guid("3dba4c2c-a603-4413-96e6-251dd899fb0a");
            /*            public static class Link
            {

            }
            */
            public static class RoleChange
            {
                public static readonly string RoleImportDbf = "29dfed0b-9753-4c89-9238-810f6710193e"; //"Единицы измерения Kat_ediz";                                                                

            }

        }


        public static class Poluf
        {
            public static readonly string table_name = "Poluf.dbf";
            public static readonly string name_refer = "Poluf";
            public static readonly string nameTip = "poluf";
            public static readonly string name_row_num = "ROW_NUMBER";
            public static readonly Guid guid = new Guid("610f6a2d-a019-4174-a0f5-769ea0740607");
            public static readonly Guid ROW_NUMBER = new Guid("ed0b094d-fe42-4d98-817d-6aeaad779656");
            /*            public static class Link
            {

            }
            */
            public static class RoleChange
            {
                
                public static readonly string RoleImportDbf = "c0bd9777-bb60-4984-b578-16f95ea5b1a5"; //"Единицы измерения Kat_ediz";                                                                
                public static readonly string RoleListNum = "2b31ed4b-7663-49cc-9005-6c27f418cafa"; //Список номенклатуры Poluf


            }
        }

        

        public static class KAT_RAZM
        {
            public static readonly string table_name = "KAT_RAZM.dbf";
            public static readonly string name_refer = "KAT_RAZM";
            public static readonly string nameTip = "KAT_RAZM";
            public static readonly string name_row_num = "ROW_NUMBER";
            public static readonly Guid guid = new Guid("cfb91e14-5499-4c63-ab9c-e908e048c92e");
            public static readonly Guid ROW_NUMBER = new Guid("71cd2fd5-e763-44f9-829c-d7b77ad3d710");
            /*            public static class Link
            {

            }
            */
            public static class RoleChange
            {
                public static readonly string RoleImportDbf = "67da3c81-6445-4090-9799-4ff9383560ca";
            }
        }

        

        public static class KAT_STOL
        {
            public static readonly string table_name = "KAT_STOL.dbf";
            public static readonly string name_refer = "KAT_STOL";
            public static readonly string nameTip = "KAT_STOL";
            public static readonly string name_row_num = "ROW_NUMBER";
            public static readonly Guid guid = new Guid("ddd085f9-8c74-4b61-b248-d66cd8a2e35e");
            public static readonly Guid ROW_NUMBER = new Guid("66775dbe-ad47-4b34-8e68-2dcd01bfd09e");
            /*            public static class Link
            {

            }
            */
            public static class RoleChange
            {
                public static readonly string RoleImportDbf = "dd42b2ce-081e-4ee8-beb3-b4fa9fd63989";
            }
        }

        public static class Список_номенклатуры_FoxPro
        {
            //public static readonly string name = "kat_ediz.dbf";
            public static readonly string name_refer = "Список номенклатуры FoxPro";
            public static readonly string nameTip = "Запись";
            public static readonly Guid guid = new Guid("c9d26b3c-b318-4160-90ae-b9d4dd7565b6");
            public static readonly Guid Обозначение = new Guid("1bbb1d78-6e30-40b4-acde-4fc844477200");
        }


        public static class Подключения
        {
            //public static readonly string name = "kat_ediz.dbf";
            public static readonly string name_refer = "Подключения";
            public static readonly string nameTip = "Подключение";
            public static readonly Guid guid = new Guid("9dd79ab1-2f40-41f3-bebc-0a8b4f6c249e");
            public static readonly Guid Обозначение = new Guid("1bbb1d78-6e30-40b4-acde-4fc844477200");
            public static readonly Guid Сводное_обозначение = new Guid("9dd79ab1-2f40-41f3-bebc-0a8b4f6c249e");
        }


    }
}

/*    private static class TableName
    {
        public static string SPEC = "SPEC";

    }*/

public enum TableName
{
    SPEC,
    NORM,
    KLAS,
    KLASM,
    TRUD,
    MARCHP,
    //-------------------
    OSNAST,
    Poluf,
    KAT_EDIZ,
    KAT_INS,
    KAT_IZD,
    KAT_IZV,
    KAT_IZVM,
    KAT_IZVP,
    KAT_IZVT,
    KAT_OPER,
    KAT_PODR,
    KAT_RAZM,
    KAT_STOL,
    KAT_VHOD
}





    #endregion Markin Classes

    #endregion Service classes






#region extension methods 

public static class DivideList
{
    
    // Делит коллекцию на n частей
    public static IEnumerable<IEnumerable<T>> Partition<T>(this IEnumerable<T> source, int n)
    {
        var count = source.Count();

        return source.Select((x, i) => new { value = x, index = i })
            .GroupBy(x => x.index / (int)Math.Ceiling(count / (double)n))
            .Select(x => x.Select(z => z.value));
    }
}

#endregion extension methods 

