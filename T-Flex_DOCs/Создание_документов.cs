using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using TFlex.DOCs.Model.Classes;
using TFlex.DOCs.Model.Desktop;
using TFlex.DOCs.Model.Macros;
using TFlex.DOCs.Model.References;
using TFlex.DOCs.Model.References.Documents;
using TFlex.DOCs.Model.References.Files;
using TFlex.DOCs.Model.References.Nomenclature;
using TFlex.DOCs.Model.Search;
using FolderObject = TFlex.DOCs.Model.References.Files.FolderObject;
using System.Text.RegularExpressions;

public class MacroAEM : MacroProvider
{
    public MacroAEM(MacroContext context)
        : base(context)
    {
/*        if (Вопрос("Хотите запустить в режиме отладки?"))
        {
            System.Diagnostics.Debugger.Launch();
            System.Diagnostics.Debugger.Break();
        }*/
    }

    /// <summary>
    /// Строка для записи лога ошибок
    /// </summary>
    private readonly StringBuilder _logRemarks = new StringBuilder();

    public override void Run()
    {
    }
    public void СоздатьДокументы()
    {
        //Проверка на тип
        var folderObject = Context.ReferenceObject as FolderObject;
        if (folderObject == null)
            Ошибка("Предупреждение", "Выбранный объект не является папкой.");

        //Находим и заполняем все настройки, такие как справочник документов итд.
        var contextImport = CreateContextImport(folderObject);

        ДиалогОжидания.Показать("Подождите, идет обработка", true);

        WriteLogText(new string('-', 72));

        //Перебираем все дочерние объекты 
        foreach (var fileReferenceObject in folderObject.Children)
        {
            ProcessingFileObject(fileReferenceObject, contextImport);
        }

        //Сохранение объектов
        ReferenceObjectSaveSet.RunWithSaveSet(set =>
        {
            foreach (ReferenceObject saveobj in contextImport.SaveObjects)
            {
                if (saveobj.Changing)
                    set.Add(saveobj);
            }

            set.EndChanges();
        });

        //Применение изменений объектов
        ReferenceObjectSaveSet.RunWithSaveSet(set =>
        {
            //Пакетное сохранение объектов на сервер
            foreach (ReferenceObject saveObj in contextImport.SaveObjects)
            {
                if(saveObj.Changing)
                set.Add(saveObj);
            }

            set.EndChanges();
        });

        //Reference.EndChanges(contextImport.SaveObjects.Where(obj => obj.Changing));

        //Применение изменений объектов
        Desktop.CheckIn(contextImport.SaveObjects, "Подключение электронного архива из файлов", false);
        WriteLogText(Environment.NewLine + "Завершение процесса создания документов");

        //Запись файла во временную папку
        string logFilePath = string.Format("{0}{1} {2:d.MM.yyyy HH-mm-ss}.txt", Path.GetTempPath(), "Загрузка файлов в DOCs", DateTime.Now);
        File.AppendAllText(logFilePath, Environment.NewLine + _logRemarks);


        FileReference fileReference = folderObject.Reference;
        //Загрузка файла в докс
        FileObject logObjPath = fileReference.AddFile(logFilePath, folderObject);

        ДиалогОжидания.Скрыть();
        Сообщение("Создание завершено", "Создание объектов документов было завершено. Лог можно найти по пути: " + logObjPath.Path);

    }

    private void ProcessingFileObject(FileReferenceObject fileReferenceObject, ContextImport contextImport)
    {
        //Получаем наименование файла, без расширения. т.е. Name^Code.grb => Name^Code
        string fileObjectName = Path.GetFileNameWithoutExtension(fileReferenceObject.Name);

        //Проверка, что объект является файлом, иначе пишем в лог
        if (!fileReferenceObject.IsFile)
        {
            WriteLogText(string.Format("Объект {0}, не является файлом", fileObjectName));
            return;
        }

        string docName;
        string docCode;

        //В этом методе он проверяет есть ли символ ^ так что возвращает true если все хорошо.
        //Записывает наименование обозначение в переменные docName, docCode
        if (!TryParseFileName(fileObjectName, out docName, out docCode))
            return;

        //Получаем связанные документы Settings.DocumentFiles это гуид связи
        var linkDokuments = fileReferenceObject.GetObjects(Settings.ДокументФайлы);

        //Находим первый связанный документ у которого наименование и обозначение совпадают
        var findDocument = linkDokuments.FirstOrDefault(docObj =>
            (docObj[Settings.ДокументНаименование].GetString() == docName) && (docObj[Settings.ДокументНаименование].GetString() == docCode));

        //Если нашли такой документ, то подключать не нужно, пишем в лог что такой документ уже подключен
        if (findDocument != null)
        {
            WriteLogText(string.Format("Файл: '{0}' уже подключен к документу с наименование: '{1}' и обозначением: '{2}'",
                fileObjectName, docName, docCode));
            return;
        }

        //Находим документ по наименование и обозначению
        var documentReferenceObject = FindDocumentReferenceObject(contextImport, docName, docCode);
        if (documentReferenceObject == null)
        {
            //Разбирает обозначение и получает у него тип объекта
            string typeObject = TryParseTypeOfName(docCode);

            //Создаем объект справочника документы с наименованием обозначением и указанным типом
            documentReferenceObject = CreateDocumentObject(contextImport, docName, docCode, typeObject);
        }

        //Берем в редактирование файл
        if (!TryBeginChangesReferenceObject(fileReferenceObject))
        {
            WriteLogText(string.Format("Невозможно взять в редактирование объект '{0}'", fileReferenceObject));
            return;
        }

        //Добавляем документ по связи к файлу
        fileReferenceObject.AddLinkedObject(Settings.ДокументФайлы, documentReferenceObject);

        contextImport.SaveObjects.Add(fileReferenceObject);
    }

    /// <summary>
    /// Преобразует из строки тип, возвращает строку с типом, если удалось разобрать
    /// Если не удалось разобрать, то возвращает тип "Деталь"
    /// </summary>
    /// <param name="docCode"></param>
    /// <returns></returns>
    private string TryParseTypeOfName(string docCode)
    {
        // Проверяем не сборочный чертеж ли это.
        // Есть такая строка "АБВГ.002003.003 СБ"
        // надо узнать оканчивается ли она на "СБ"
      //  Regex.IsMatch(input, @" РР\d*$")
  /*     
       	if (docCode.EndsWith(" СБ"))
            return Settings.ДокументТипЧертежСб;
        if (docCode.EndsWith(" ГЧ"))
            return Settings.ДокументТипЧертежГч;
          if (docCode.EndsWith(" МЧ"))
            return Settings.ДокументТипЧертежМч;
        if (docCode.EndsWith(" МЭ"))
            return Settings.ДокументТипЧертежМэ;
        if (docCode.EndsWith(" ОВ"))
            return Settings.ДокументТипЧертеж;
        if (docCode.EndsWith(" ВС"))
            return Settings.ДокументТипВедомостьВС;
        if (docCode.EndsWith(" ВП"))
            return Settings.ДокументТипВедомостьВП;
        if (docCode.EndsWith(" РЭ"))
            return Settings.ДокументТипРуководствоРЭ;
        if (docCode.EndsWith( "РО"))
            return Settings.ДокументТипРуководство;
        if (docCode.EndsWith(" ПС"))
            return Settings.ДокументТипПаспорт;
        if (docCode.EndsWith(" ТУ"))
            return Settings.ДокументТипТехУсловия;
        if (docCode.EndsWith(" УЧ"))
            return Settings.ДокументТипЧертежУч;
        if (docCode.EndsWith(" ДП"))
            return Settings.ДокументТипВедомостьДП;
        if (docCode.EndsWith(" ЗИ"))
            return Settings.ДокументВедомостьЗИП;
        if (docCode.EndsWith(" ВД"))
            return Settings.ДокументТипВедомостьСсылкаДок;
        if (docCode.EndsWith(" ТБ"))
            return Settings.ДокументТипТаблицы;
        if (docCode.EndsWith(" Д") || docCode.EndsWith(" ОПД"))
            return Settings.ДокументТипДокументыПрочие;
        if (docCode.EndsWith(" ПЗ"))
            return Settings.ДокументТипПояснительнаяЗаписка;
        if (docCode.EndsWith(" ПМ"))
            return Settings.ДокументТипПрограммаМетодикаИспытаний;
        if (docCode.EndsWith(" ВИ"))
            return Settings.ДокументТипВедомостьРазрешенияПримененияПокупныхИзделий;
        if (docCode.EndsWith(" ВИ"))
            return Settings.ДокументТипВедомость;
        if (docCode.EndsWith(" РР1"))
            return Settings.ДокументТипРасчеты;
        if (docCode.EndsWith(" РР2"))
            return Settings.ДокументТипРасчеты;
        if (docCode.EndsWith(" РР3"))
            return Settings.ДокументТипРасчеты;
        if (docCode.EndsWith(" РР4"))
            return Settings.ДокументТипРасчеты;
        if (docCode.EndsWith(" К") || docCode.EndsWith(" К3") || docCode.EndsWith(" С"))
        	return Settings.ДокументТипСхема;
        if (docCode.EndsWith(" Э1"))
            return Settings.ДокументТипСхемаЭП;
        if (docCode.EndsWith(" Э2"))
            return Settings.ДокументТипСхемаЭП;
        if (docCode.EndsWith(" Э3"))
            return Settings.ДокументТипСхемаЭП;
        	
    */ 

       	if (Regex.IsMatch(docCode, @" СБ\d*$"))
            return Settings.ДокументТипЧертежСб;
        if (Regex.IsMatch(docCode, @" ГЧ\d*$"))
            return Settings.ДокументТипЧертежГч;
          if (Regex.IsMatch(docCode, @" МЧ\d*$"))
            return Settings.ДокументТипЧертежМч;
        if (Regex.IsMatch(docCode, @" МЭ\d*$"))
            return Settings.ДокументТипЧертежМэ;
        if (Regex.IsMatch(docCode, @" ОВ\d*$"))
            return Settings.ДокументТипЧертеж;
        if (Regex.IsMatch(docCode, @" ВС\d*$"))
            return Settings.ДокументТипВедомостьВС;
        if (Regex.IsMatch(docCode, @" ВП\d*$"))
            return Settings.ДокументТипВедомостьВП;
        if (Regex.IsMatch(docCode, @" РЭ\d*$"))
            return Settings.ДокументТипРуководствоРЭ;
        if (Regex.IsMatch(docCode, @" РО\d*$"))
            return Settings.ДокументТипРуководство;
        if (Regex.IsMatch(docCode, @" ПС\d*$"))
            return Settings.ДокументТипПаспорт;
        if (Regex.IsMatch(docCode, @" ТУ\d*$"))
            return Settings.ДокументТипТехУсловия;
        if (Regex.IsMatch(docCode, @" УЧ\d*$"))
            return Settings.ДокументТипЧертежУч;
        if (Regex.IsMatch(docCode, @" ДП\d*$"))
            return Settings.ДокументТипВедомостьДП;
        if (Regex.IsMatch(docCode, @" ЗИ\d*$"))
            return Settings.ДокументВедомостьЗИП;
        if (Regex.IsMatch(docCode, @" ВД\d*$"))
            return Settings.ДокументТипВедомостьСсылкаДок;
        if (Regex.IsMatch(docCode, @" ТБ\d*$"))
            return Settings.ДокументТипТаблицы;
        if (Regex.IsMatch(docCode, @" Д\d*$") || Regex.IsMatch(docCode, @" ОПД\d*$"))
            return Settings.ДокументТипДокументыПрочие;
        if (Regex.IsMatch(docCode, @" ПЗ\d*$"))
            return Settings.ДокументТипПояснительнаяЗаписка;
        if (Regex.IsMatch(docCode, @" ПМ\d*$"))
            return Settings.ДокументТипПрограммаМетодикаИспытаний;
        if (Regex.IsMatch(docCode, @" ВИ\d*$"))
            return Settings.ДокументТипВедомостьРазрешенияПримененияПокупныхИзделий;
        if (Regex.IsMatch(docCode, @" ВИ\d*$"))
            return Settings.ДокументТипВедомость;
        if (Regex.IsMatch(docCode, @" РР\d*$"))
            return Settings.ДокументТипРасчеты;
        if (Regex.IsMatch(docCode, @" К\d*$") || Regex.IsMatch(docCode, @" С\d*$"))
        	return Settings.ДокументТипСхема;
        if (Regex.IsMatch(docCode, @" Э\d*$"))
            return Settings.ДокументТипСхемаЭП;        	
        	
            
        
        //Дальше проверяем не сборка ли это.
        //Надо найти число которое стоит за первой точкой т.е. "АБВГ.3" нам нужна 3
        //Первое условие, находим точку в строке
        if (docCode.Contains('.'))
        {
            //Разделяем строку на 2 составляющие, получаем вторую т.е. 301002
            string firstSplitStirng = docCode.Split('.')[0];
            string secondSplitStirng = docCode.Split('.')[1];

            

            //Если указанная строка начинается с "3" то это тип Сборочная единица
            if (firstSplitStirng.StartsWith("8А2")  || 
                firstSplitStirng.StartsWith("8А3")  ||
                firstSplitStirng.StartsWith("8А4")  ||
                firstSplitStirng.StartsWith("8А5")  ||
                firstSplitStirng.StartsWith("8А6"))
                return Settings.ДокументТипСборка;
            
            
            //Если указанная строка начинается с "3" то это тип Сборочная единица
            if (secondSplitStirng.StartsWith("3")  || 
                secondSplitStirng.StartsWith("46") || 
                secondSplitStirng.StartsWith("52") || 
                secondSplitStirng.StartsWith("56") || 
                secondSplitStirng.StartsWith("64") || 
                secondSplitStirng.StartsWith("65") || 
                secondSplitStirng.StartsWith("67") || 
                secondSplitStirng.StartsWith("68") || 
                secondSplitStirng.StartsWith("79") ||
                secondSplitStirng.StartsWith("УЯИС.3"))
                return Settings.ДокументТипСборка;
        }

        //Во всех остальных случаях это "Деталь"
        return Settings.ДокументТипДеталь;
    }

    /// <summary>
    /// Создает объект в справочнике документы, тип по умолчанию Деталь
    /// </summary>
    /// <param name="contextImport"></param>
    /// <param name="name"></param>
    /// <param name="code"></param>
    /// <param name="classObjectName"></param>
    /// <returns></returns>
    private ReferenceObject CreateDocumentObject(ContextImport contextImport, string name, string code, string classObjectName = "Деталь")
    {
        //Получаем наименование папки которое является текущим объектом
        var folderName = contextImport.FileFolderObject.Name;

        //Находим есть ли в контексте уже папка документов
        if (contextImport.DocumentFolderReferenceObject == null)
        {
            //Если нет папки, то мы находим папку в документах по наименованию
            contextImport.DocumentFolderReferenceObject = FindDocumentReferenceObject(contextImport, folderName);
            
            //Если не нашли папку в справочнике, то создаем папку
            if (contextImport.DocumentFolderReferenceObject == null)
                contextImport.DocumentFolderReferenceObject = CreateDocumentFolder(contextImport, folderName);
        }

        //Находим указанный тип
        ClassObject createdClassObject = contextImport.DocumentReference.Classes.Find(classObjectName);
        if (createdClassObject == null)
        {
            WriteLogText(string.Format(
                "Не найден тип '{0}' в справочнике документы",
                folderName));
            return null;
        }

        //Создаем объект в справочнике документы
        var createDocumentObject = contextImport.DocumentReference.CreateReferenceObject
            (contextImport.DocumentFolderReferenceObject, createdClassObject);

        //Заполняем параметры
        createDocumentObject[Settings.ДокументНаименование].Value = name;
        createDocumentObject[Settings.ДокументОбозначение].Value = code;

        createDocumentObject.EndChanges();

        //Находим папку, в номенклатуре, для указания родителя

/*      Создаём в корне Номенклатуры
        if (contextImport.NomenclatureFolderObject == null)
            contextImport.NomenclatureFolderObject = CreateNomenclatureFolderObject(contextImport);
*/

        //Создаем объект номенклатуры
        var createNomenclatureObject = contextImport.NomenclatureReference.CreateNomenclatureObject
            (createDocumentObject);//, contextImport.NomenclatureFolderObject);//Создаём в корне Номенклатуры

        if(createNomenclatureObject.Changing)
            createNomenclatureObject.EndChanges();

        //Добавляем объект для применения изменений объектов
        contextImport.SaveObjects.Add(createDocumentObject);

        return createDocumentObject;
    }

    /// <summary>
    /// Создает папку в справочнике номенклатура
    /// </summary>
    /// <param name="contextImport"></param>
    /// <returns></returns>
    private NomenclatureReferenceObject CreateNomenclatureFolderObject(ContextImport contextImport)
    {
        var папка = НайтиОбъект(Settings.СправочникНоменклатура.ToString(), "Наименование", "Импорт");

        if (папка == null)
        {
            папка = СоздатьОбъект(Settings.СправочникНоменклатура.ToString(), "Папка");
            папка["Наименование"] = "Импорт";
            папка.Сохранить();
        }

        return (NomenclatureReferenceObject)папка;
    }

    /// <summary>
    /// Находит объект в справочнике Документы, по совпадению наименование обозначение
    /// </summary>
    /// <param name="contextImport"></param>
    /// <param name="name"></param>
    /// <param name="code"></param>
    /// <returns></returns>
    private ReferenceObject FindDocumentReferenceObject(ContextImport contextImport, string name, string code = "")
    {
        //Формируем базовую строку фильтра по наименованию
        string filterString = string.Format("[{0}] = '{1}'", "Наименование", name);

        //Если указанно обозначение, его так же добавляем в фильтр
        if (code != string.Empty)
            filterString += string.Format(" И [{0}] = '{1}'", "Обозначение", code);

        Filter filter;

        //Преобразуем строку фильтра в фильтр для справочника
        if (!Filter.TryParse(filterString, contextImport.DocumentReference.ParameterGroup, out filter))
            return null;

        //Находим и возвращаем первый найденный объект
        return contextImport.DocumentReference.Find(filter).FirstOrDefault();
    }

    /// <summary>
    /// Преобразует строку наименования файла в две строки наименование и обозначение
    /// </summary>
    /// <param name="fileName"></param>
    /// <param name="name"></param>
    /// <param name="code"></param>
    /// <returns>Возвращает tru если преобразование произошло</returns>
    private bool TryParseFileName(string fileName, out string name, out string code)
    {
        name = string.Empty;
        code = string.Empty;

        //Проверка на наличие спец символа
        if (!fileName.Contains(Settings.СтрокаРазделитель))
        {
            WriteLogText(string.Format(
                "Невозможно разобрать наименование объекта {0}{1}В наименовании отсутствует символ: '{2}'",
                fileName, Environment.NewLine, Settings.СтрокаРазделитель));
            return false;
        }

        //Разбиваем строку на массив строк
        var splitFileName = fileName.Split(new string[] { Settings.СтрокаРазделитель }, StringSplitOptions.RemoveEmptyEntries);

        //Проверяем что при разбитии строк получилось 2 значения
        if (splitFileName.Length > 2)
        {
            WriteLogText(string.Format
            ("Невозможно разобрать наименование объекта {0}{1}В наименовании присутствуют больше двух символов: '{2}'",
                fileName, Environment.NewLine, Settings.СтрокаРазделитель));
            return false;
        }

        //Записываем значения из массива строк
        name = splitFileName[1];
        code = splitFileName[0];

        return true;
    }

    private ReferenceObject CreateDocumentFolder(ContextImport contextImport, string folderName)
    {
        var createFolderReferenceObject = contextImport.DocumentReference.CreateReferenceObject(contextImport.DocumentFolderClassObject);

        createFolderReferenceObject[Settings.ДокументНаименование].Value = folderName;

        createFolderReferenceObject.EndChanges();

        return createFolderReferenceObject;
    }

    private void WriteLogText(string text)
    {
        _logRemarks.AppendLine(string.Format("{0:d.MM.yyyy HH-mm-ss}: {1}{2}", DateTime.Now, text, Environment.NewLine));
    }

    private ContextImport CreateContextImport(FolderObject folderObject)
    {
        var documentInfo = Context.Connection.ReferenceCatalog.Find(Settings.СправочникДокумент);
        if (documentInfo == null)
            Ошибка("Не найден справочник 'Документы'");

        var nomenclatureInfo = Context.Connection.ReferenceCatalog.Find(Settings.СправочникНоменклатура);
        if(nomenclatureInfo == null)
            Ошибка("Не найден справочник 'Электронная структура изделий'");

        var nomenclatureReference = (NomenclatureReference)nomenclatureInfo.CreateReference();
        
        var documentReference = (DocumentReference)documentInfo.CreateReference();

        var classObject = documentReference.Classes.Find("Папка");

        var contextImport = new ContextImport()
        {
            DocumentReference = documentReference,
            FileFolderObject = folderObject,
            DocumentFolderClassObject = classObject,
            SaveObjects = new List<ReferenceObject>(),
            NomenclatureReference = nomenclatureReference
        };

        return contextImport;
    }

    private static bool TryBeginChangesReferenceObject(ReferenceObject referenceObject)
    {
        if (referenceObject.IsCheckedOut)
        {
            if (referenceObject.IsCheckedOutByCurrentUser)
            {
                if (referenceObject.CanEdit)
                {
                    referenceObject.BeginChanges();
                    return true;
                }

                return referenceObject.Changing;
            }
        }
        else if (referenceObject.CanCheckOut)
        {
            referenceObject.CheckOut(false);

            if (referenceObject.CanEdit)
            {
                referenceObject.BeginChanges();
                return true;
            }
        }
        return false;
    }


    private class ContextImport
    {
        public Reference DocumentReference;
        public DocumentType DocumentFolderClassObject;

        public NomenclatureReference NomenclatureReference;
        public NomenclatureReferenceObject NomenclatureFolderObject;

        /// <summary>
        /// Папка в справочнике документы в которую будут создаваться новые документы
        /// </summary>
        public ReferenceObject DocumentFolderReferenceObject;

        /// <summary>
        /// ТкущийОбъект только с переопределенным типом
        /// </summary>
        public FileReferenceObject FileFolderObject;

        /// <summary>
        /// Для пакетного сохранения объектов
        /// </summary>
        public List<ReferenceObject> SaveObjects;
    }

    private static class Settings
    {
        public static Guid СправочникДокумент = new Guid("ac46ca13-6649-4bbb-87d5-e7d570783f26");
        public static Guid ДокументНаименование = new Guid("7e115f38-f446-40ce-8301-9b211e6ce5fd");
        public static Guid ДокументОбозначение = new Guid("b8992281-a2c3-42dc-81ac-884f252bd062");
        public static Guid ДокументФайлы = new Guid("9eda1479-c00d-4d28-bd8e-0d592c873303");
        public static Guid СправочникНоменклатура = new Guid("853d0f07-9632-42dd-bc7a-d91eae4b8e83");
        
       public static string ДокументТипДеталь = "7c41c277-41f1-44d9-bf0e-056d930cbb14";
        public static string ДокументТипСборка = "dd2cb8e8-48fa-4241-8cab-aac3d83034a7";
        public static string ДокументТипЧертежСб = "d6324424-a39e-4112-a207-7e96a1971852";
		public static string ДокументТипЧертежГч = "2314da75-13ef-4b5d-914e-9d250cf7f3de";
		public static string ДокументТипЧертежМч = "1fcad810-17b4-42ad-9eec-0e0f1a8edc0c";
		public static string ДокументТипЧертежМэ = "1334d2d2-58ce-48fc-8e7c-0669f5db79a8";
		public static string ДокументТипЧертеж = "d4caa669-e807-42cb-8679-09f933d8683e";
		
		public static string ДокументТипСхема = "63673a22-0a54-4d84-af96-976b9d2ccaa9";
	   public static string ДокументТипСхемаЭП ="cf9e58e9-97e8-43cd-b0fa-69ba8ce74d99"; //Э3
		
		public static string ДокументТипВедомостьВС = "57bbce5b-7fcc-40b0-a119-3cd13b41ad1b";
		public static string ДокументТипВедомостьВП = "e600bdc4-6fa5-4607-8586-f0a552bd4f45";
		
		public static string ДокументТипРуководствоРЭ = "4dee11fc-5e66-4fe1-b6da-f9f7bbdb0ac1";
		public static string ДокументТипРуководство = "85080a8a-1569-47c3-9f07-d4f7214fe234";
		
		public static string ДокументТипПаспорт = "e9eb5933-298e-4323-a815-e306dda68d64";
		
		public static string ДокументТипТехУсловия = "8f572e14-48a5-4078-a34b-fadb76cf95e7";
				
		public static string ДокументТипЧертежУч = "f1670a1f-eaef-459a-b5b7-f255dc87dba6";
		public static string ДокументТипВедомостьДП = "80836c0d-050d-4605-92b5-23f7f6912cf7";
		public static string ДокументВедомостьЗИП = "972ab13f-0535-45f3-a38b-f06dd09be0e0";
		public static string ДокументТипВедомостьСсылкаДок = "353f37be-de9f-4e0c-8c0a-1e6c32f62617";
		public static string ДокументТипТаблицы = "20100110-5a6f-4ef9-856b-7aa90b7cf535";
		public static string ДокументТипРасчеты = "0d29918a-a281-42db-ae22-08eb32e51455";
		public static string ДокументТипДокументыПрочие = "036be417-f560-4113-80d4-3f4c8c309781";
		public static string ДокументТипПояснительнаяЗаписка = "7815d2bd-18aa-4b1d-93b5-6963d74df6d6";
		public static string ДокументТипПрограммаМетодикаИспытаний = "b54d074f-ed43-4a6e-9c7b-6b550f1d34a7";
		public static string ДокументТипВедомостьРазрешенияПримененияПокупныхИзделий = "282646c6-c324-4f9c-9372-0a39009b9791";
		public static string ДокументТипВедомость= "12f9df69-42d3-4c3e-952e-5c208a9a2f85";
	
		
		
		
        

        public static string СтрокаРазделитель = "^";
    }
}
