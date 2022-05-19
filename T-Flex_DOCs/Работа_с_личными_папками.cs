/*
TFlex.DOCs.UI.Objects.dll
TFlex.DOCs.UI.Common.dll
TFlex.DOCs.UI.Types.dll
TFlex.DOCs.Common.dll
*/

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using TFlex.DOCs.Model.Macros;
using TFlex.DOCs.Model.Access;
using TFlex.DOCs.Model.Classes;
using TFlex.DOCs.Model.Macros.ObjectModel;
using TFlex.DOCs.Model.References;
using TFlex.DOCs.Model.References.Users;
using TFlex.DOCs.Model.Desktop;
using TFlex.DOCs.Model.References.Files;
using TFlex.DOCs.Model.References.Nomenclature;
using TFlex.DOCs.Model.References.Documents;

public class MacroCreateAEM : MacroProvider
{
    public MacroCreateAEM(MacroContext context)
        : base(context)
    {
#if DEBUG
        System.Diagnostics.Debugger.Launch();
        System.Diagnostics.Debugger.Break();
#endif
    }
    private string _редакторский = "Редакторский"; //Назначаемый доступ на папку
    private string _редакторский_корзина = "Редакторский+Очистка корзины+Документовед"; //Назначаемый доступ на папку
    //private string _просмотрИсоздание = "Просмотр и создание"; //Назначаемый доступ на папку
    private string _наименованиеПапкиДляХраненияЛичныхПапок = "Личные папки";
    private readonly Guid ОтносительныйПуть = new Guid("adda774c-dbdf-48ba-bcf6-87bb42a67e90");
    //private string _наименованиеПапкиДляСогласования = "Согласования";

    /// <summary> Поддерживаемые расширения файлов </summary>
    private static readonly string[] SupportedExtensions = { "tif", "tiff", "pdf" };
   
    
    public List<ReferenceObject> FilesAddedBefore = new List<ReferenceObject>();

    public override void Run()
    {
        //User user = Context.ReferenceObject as User;
        var user = ТекущийОбъект;
        if (user == null)
            return;
        //Назначение наименования личной папки пользователя
        string наименованиеЛичнойПапки = ТекущийОбъект.Параметр["Логин"];
        Объект папка = НайтиОбъект("Файлы", "Наименование", наименованиеЛичнойПапки);
        //Проверка на то, что пользователь запустивший макрос, Администратор
        if ((ТекущийПользователь.Параметр["Тип"] == "Администратор") && (папка == null))
        {
            //Проверка и создание папки "Личные папки" в корне справочника
            Объект личныеПапки = НайтиОбъект("Файлы", "Наименование", _наименованиеПапкиДляХраненияЛичныхПапок);
            if (личныеПапки == null)
            {
                личныеПапки = СоздатьОбъект("Файлы", "Папка");
                личныеПапки["Наименование"] = _наименованиеПапкиДляХраненияЛичныхПапок;
                личныеПапки.Сохранить();
            }

            папка = СоздатьОбъект("Файлы", "Папка", личныеПапки);//Создаем личную папку пользователя
            папка.Параметр["Наименование"] = наименованиеЛичнойПапки;
            папка.Сохранить();
            НазначитьДоступНаОбъект(папка, user, _редакторский_корзина, true);
            Desktop.CheckIn((ReferenceObject)папка, "Создание", false);//Применение изменений
        }
    }

    //private Guid FolderGuid = new Guid("9d05651b-2676-4788-866b-e39a79a1e2f3");           // БД "Аэроэлектромаш"
    //private Guid FolderGuid = new Guid("f94907ce-dd01-4983-ba63-74721861dfe8");             // БД "Топ Системы"
    private Guid FolderGuid = new Guid("0f38fa33-cc34-4af8-9e99-60dc5aff81bf");           // БД "Макет на АЭМ"

    public bool ПолучитьОпубликованныеФайлы(Объекты объекты)
    {
        ReferenceObject folderObj = FindFolder();
        if (folderObj == null)
            throw new Exception("Не найдена папка Согласование");
        List<ReferenceObject> listObj = объекты.Select(ob => (ReferenceObject)ob).ToList();
        var перенесенныеОбъекты = ПеренестиВСогласование(listObj, folderObj);

        if (FilesAddedBefore.Any())
        {
            ВывестиИнформацию(false);
            return false;
        }
        return true;//Объекты.CreateInstance(перенесенныеОбъекты, Context);
                    //перенесенныеОбъекты.Select(ob => Объект.CreateInstance(ob, Context)).ToList();
    }

    public void ОпубликоватьФайл()
    {
        ReferenceObject folderObj = FindFolder();
        if (folderObj == null)
            Error("Не найдена папка с ID = " + FolderGuid.ToString());

        if (!Вопрос("Переместить файлы в папку \"Согласования\"?"))
            return;

        WaitingDialog.Show("Подождите, идет обработка", true);

        var currentObjs = Context.GetSelectedObjects().ToList();
        //Коллекции для сохранеия/изменения объектов
        ПеренестиВСогласование(currentObjs, folderObj, true);

        if (FilesAddedBefore.Any())
            ВывестиИнформацию(true);
    }

    public void ВернутьФайлы(List<string> paths)
    {
        /*using (StreamWriter sw = new StreamWriter(@"C:\1\text.txt"))
        {
            sw.WriteLine(paths.Count.ToString());
        }*/

        foreach (var path in paths)
        {
            int index = path.IndexOf("^^^");
            int indexLastSlash = path.LastIndexOf('\\');

            string Name = path.Substring(0, index);
            string originalFolder = path.Substring(index + 3, indexLastSlash - index - 3);

            var sourceFolder = НайтиОбъект("Файлы", string.Format("[Относительный путь] = '{0}'", originalFolder));
            var file = НайтиОбъект("Файлы", string.Format("[Наименование] = '{0}' И [Относительный путь] содержит 'Согласовани'", Name));
            file.Изменить();
            file.РодительскийОбъект = sourceFolder;

            file.Сохранить();
            Desktop.CheckIn((ReferenceObject)file, "Возврат при прерывании процесса", false);
        }
    }

    private ReferenceObject FindFolder()
    {
        //Получаем справочник файлов
        var fileReference = new FileReference(Context.Connection);
        //Находим папку
        var folderObj = fileReference.Find(FolderGuid);
        return folderObj;
    }

    private List<ReferenceObject> ПеренестиВСогласование(List<ReferenceObject> currentObjs, ReferenceObject folderObj, bool useDialog = false)
    {
        var saveObjs = new List<ReferenceObject>();
        var deleteObjs = new List<ReferenceObject>();
        if (useDialog)
            WaitingDialog.NextStep("Подождите идет копирование объектов");
        var childFolderObjs = folderObj.Children;

        УдалитьСуществующиеЛокальныеФайлы(currentObjs, folderObj);

        foreach (var currentObj in currentObjs.Where(t => t.Parent != folderObj))
        {
            //Находим дочерний объект в папке согласования
            var findObj = childFolderObjs.FirstOrDefault(n => n.ToString() == currentObj.ToString());
            if (findObj != null)
            {
                FilesAddedBefore.Add(currentObj);
                continue;
            }

            ClassObject classObj = currentObj.Class;
            //Создает копию объекта
            ReferenceObject createObj = currentObj.CreateCopy(classObj, folderObj);
            Guid fileDocLinkGuid = new Guid("9eda1479-c00d-4d28-bd8e-0d592c873303");
            //Связанные документы
            var docObjs = currentObj.GetObjects(fileDocLinkGuid);
            foreach (var objObj in docObjs)
            {
                createObj.AddLinkedObject(fileDocLinkGuid, objObj);
            }
            saveObjs.Add(createObj);
            //Отменяет изменения объекта, добавляет в список для удаления
            if (currentObj.IsCheckedOut)
                currentObj.UndoCheckOut();
            deleteObjs.Add(currentObj);

        }
        if (useDialog)
            WaitingDialog.NextStep("Подождите идет сохранение новых объектов на сервер");

        Reference.EndChanges(saveObjs);

        // ReferenceObjectSaveSet.RunWithSaveSet(set =>
        //  {
        //       //Пакетное сохранение объектов на сервер
        //       foreach (ReferenceObject saveObj in saveObjs)
        //           set.Add(saveObj);
        //        set.EndChanges();
        //});
        WaitingDialog.NextStep("Подождите идет применение изменений новых объектов на сервер");
        Desktop.CheckIn(saveObjs, "Копирование объектов", false);

        WaitingDialog.NextStep("Подождите идет удаление исходных объектов с сервера");
        //Удаление старых файлов
        Desktop.CheckOut(deleteObjs, true);
        Desktop.CheckIn(deleteObjs, "Удаление объектов после копирования", false);
        try
        {
            Desktop.ClearRecycleBin(deleteObjs);
        }
        catch
        { }
        WaitingDialog.Hide();
        return saveObjs;
    }

    private void УдалитьСуществующиеЛокальныеФайлы(List<ReferenceObject> currentObjs, ReferenceObject folderObj)
    {
        var destFolderPath = (folderObj as TFlex.DOCs.Model.References.Files.FolderObject)?.LocalPath;
        if (string.IsNullOrWhiteSpace(destFolderPath))
            return;

        foreach (ReferenceObject currentObj in currentObjs.Where(t => t.Parent != folderObj))
        {
            var destFilePath = (currentObj as FileObject)?.Name;
            if (string.IsNullOrWhiteSpace(destFilePath))
                continue;



            string curObjDestPath = Path.Combine(destFolderPath, destFilePath);

            if (File.Exists(curObjDestPath))
            {
                File.SetAttributes(curObjDestPath, FileAttributes.Normal);
                File.Delete(curObjDestPath);
            }
        }
    }

    public void ПринятьНаХранениеИзБП(Объекты объектыБП)
    {
        foreach (var объект in объектыБП)
        {
            var nomenReferenceObj = (ReferenceObject)объект as NomenclatureReferenceObject;
            if (nomenReferenceObj == null)
                continue;
            
            ВыполнитьМакрос("АЭМ Номенклатура Изменения", "ПринятьНаХранение", nomenReferenceObj);
        }
    }


    public void ПринятьНаХранениеИзБПОГТ(Объекты объектыБП)
    {
        foreach (var объект in объектыБП)
        {
            var nomenReferenceObj = (ReferenceObject)объект as NomenclatureReferenceObject;
            if (nomenReferenceObj == null)
                continue;
            ВыполнитьМакрос("PDM-Номенклатура. Изменения ОГТ", "ПринятьНаХранение", nomenReferenceObj);
        }
    }

    /// <summary>
    /// Назначает на объект выбранный доступ, есть возможность очистить все доступы
    /// </summary>
    /// <param name="объектНазначенияДоступа">Объект для изменения доступа</param>
    /// <param name="пользователь">Пользователь на которого будет назначатся доступ</param>
    /// <param name="наименованиеДоступа">Наименование доступа, которое необходимо установить на объект</param>
    /// <param name="очищатьСтарыеДоступы">Значение позволяющее удалить все доступы(для исспользования личных папок)</param>
    private void НазначитьДоступНаОбъект(Объект объектНазначенияДоступа, Объект пользователь, string наименованиеДоступа, bool очищатьСтарыеДоступы = false)
    {
        User user = (ReferenceObject)пользователь as User;
        if (user == null)
            throw new Exception(string.Format("Объект: {0} не является пользователем", пользователь));
        //Получем группу доступов
        AccessGroup accessEdit = AccessGroup.GetGroups(Context.Connection).FirstOrDefault(ag => ag.Type.IsObject && ag.Name == наименованиеДоступа);
        //Получаем менеджер доступа на объект
        AccessManager accessManager = AccessManager.GetReferenceObjectAccess((ReferenceObject)объектНазначенияДоступа);
        //Убираем наследование доступа от родителя
        accessManager.SetInherit(true, true);
        accessManager.Save();
        accessManager.SetInherit(false, false);
        //Очищаем все доступы
        if (очищатьСтарыеДоступы)
            accessManager.ToList().Clear();
        //Устанавливаем доступ пользователю
        accessManager.SetAccess(user, accessEdit);
        //Сохраняем изменения
        accessManager.Save();
    }
    
    
    public void Права_на_созданнулю_личную_папку()
    {        
        Объект папка = ТекущийОбъект;        
        string login =папка["Наименование"].ToString();
        Объект user=  НайтиОбъект("Группы и пользователи","Логин",login);
        НазначитьДоступНаОбъект(папка, user, _редакторский_корзина);
    }

    private void ВывестиИнформацию(bool show)
    {
        string res = "Следующие файлы уже опубликованы:\r\n";
        res += string.Join("\r\n", FilesAddedBefore.Select(t => t[ОтносительныйПуть].ToString()));
        if (show)
            Сообщение("Информация", res);
        else
        {
            using (StreamWriter sw = new StreamWriter(Path.Combine(Path.GetTempPath(), "TFlex.DOCs.Client.log"), append: true)) //(@"C:\1\информация.txt"))
            {
                res = "\r\n------------------------------------------------------------------------------------------------------------------------\r\n";
                res += "Следующие файлы уже опубликованы:\r\n";
                res += string.Join("\r\n", FilesAddedBefore.Select(t => t[ОтносительныйПуть].ToString()));
                sw.WriteLine(res);
            }
        }
    }

    public void ЗапуситьПроцедуруКДНаТехОснащение()
    {
        ПроверитьИЗапуситьПроцедуру("PDM. Согласование КД на технологическое оснащение", ВыбранныеОбъекты);
    }
    
    public void ЗапуситьПроцедуруКЦСогласованиеКД()
    {
        System.Collections.Generic.List<Объект> обрабатываемыеОбъекты = (List<Объект>)ВыбранныеОбъекты;
        ДиалогВвода диалог = СоздатьДиалогВвода("Выберите обрабатываемые объекты");
        диалог.ДобавитьВыборИзСписка("Список:", "Выбранные объекты", true, "Выбранные объекты", "Выбранные и дочерние на один уровень", "Выбранные и дочерние на все уровни");
        
        if (диалог.Показать())
        {
            switch(диалог["Список:"])
            {
        		case "Выбранные и дочерние на один уровень":
            	foreach(Объект обрабатываемыйОбъект in ВыбранныеОбъекты)
            	{
            		обрабатываемыеОбъекты.AddRange(обрабатываемыйОбъект.ДочерниеОбъекты);
            	}
                break;
                
                case "Выбранные и дочерние на все уровни":
                foreach(Объект обрабатываемыйОбъект in ВыбранныеОбъекты)
                {
                    обрабатываемыеОбъекты.AddRange(обрабатываемыйОбъект.ВсеДочерниеОбъекты);
                }
                break;
                
                default: break;
            }
            ПроверитьИЗапуситьПроцедуру("КЦ. Согласование КД", обрабатываемыеОбъекты);
        }
    }

// Отключил, чтобы сохранить этот метод, но не запускался. Запускать метод выше "ЗапуситьПроцедуруКЦСогласованиеКД"
/*
    public void ЗапуститьПроцедуруКЦСогласованиеКД()
    {
        ПроверитьИЗапуситьПроцедуру("КЦ. Согласование КД", ВыбранныеОбъекты);
    }
*/

    public void ЗапуситьПроцедуруКЦСогласованиеСтруктуры()
    {
    	ПроверитьИЗапуситьПроцедуру("КЦ. Согласование структуры", ВыбранныеОбъекты);
    }

    private void ПроверитьИЗапуситьПроцедуру(string имяПроцедуры, Объекты обрабатываемыеОбъекты)
    {
        var fileReference = new FileReference(Context.Connection);
        ReferenceObject folderObj = FindFolder();
        List<string> notAvailableStages = new List<string>() { "Хранение", "Аннулировано", "Утверждено" };

        System.Collections.Generic.List<Объект> объектыДляБП = new System.Collections.Generic.List<Объект>();
        System.Collections.Generic.List<Объект> подключаемыеФайлы = new System.Collections.Generic.List<Объект>();
        bool нетФайлов = false;
        bool запускать = true;
        string стиНеВХранении = "";
        
        foreach (Объект объектБП in обрабатываемыеОбъекты)
        {
            if (notAvailableStages.Contains(объектБП["Стадия"].ToString()))
                continue;

            if (объектБП.Тип.ПорожденОт("Стандартное изделие") || объектБП.Тип.ПорожденОт("Прочее изделие") || объектБП.Тип.ПорожденОт("Электронный компонент") || объектБП.Тип.ПорожденОт("Материал"))
            {
                стиНеВХранении += объектБП.Тип + " - " + объектБП["Объект"] + "\r\n";
                continue;
            }

            объектыДляБП.Add(объектБП);

            Объект документ = null;
            Объекты файлыДокумента = null;

            if (объектБП.Тип.ПорожденОт(new Guid("7688d7fc-5524-4b5f-9bea-f02ad1e8e0a7"))) //Объект номенклатуры - справочник Номенклатура
            {
                Объект связанныйОбъект = объектБП.СвязанныйОбъект["Связанный объект"];
                if (связанныйОбъект != null && связанныйОбъект.Тип.ПорожденОт(new Guid("89e45926-0f0f-4c36-b649-3784d274e348")))//Конструкторско-технологический документ - справочник "Документы"
                    документ = связанныйОбъект;
            }
            else if (объектБП.Тип.ПорожденОт(new Guid("89e45926-0f0f-4c36-b649-3784d274e348"))) //Конструкторско-технологический документ - справочник "Документы"
                документ = объектБП;

            if (документ == null)
                continue;

            файлыДокумента = документ.СвязанныеОбъекты["Файлы"].Where(файл => файл["Стадия"] != "Хранение").ToList();
            if (файлыДокумента.Count() > 0)
                подключаемыеФайлы.AddRange(файлыДокумента);
            else
                нетФайлов = true;
        }
        var файлыНаРедактировании = подключаемыеФайлы.Where(t => t.ВзятНаРедактирование || t.ВзятНаРедактированиеТекущимПользователем);
        if (файлыНаРедактировании.Any())
        {
            string resError = "";
            foreach (var file in файлыНаРедактировании)
                resError += string.Format("{0}\r\n", file.ToString());
            Ошибка("Следующие файлы не сохранены:\r\n" + resError);
        }
        if (!String.IsNullOrEmpty(стиНеВХранении))
        {
            Ошибка("В составе выбранных объектов есть не согласованные стандартные изделия:\r\n" + стиНеВХранении);
        }
        if (нетФайлов)
        {
            if (!Вопрос("У одного или нескольких выбранных объектов нет файлов для согласования.\r\n Запустить процедуру согласования КД?"))
                запускать = false;
        }

        if (подключаемыеФайлы.Any())
        {
            bool errorRecycleBin = false;
            bool errorExist = false;
            List<string> filesExist = new List<string>();
            List<string> filesRecycleBin = new List<string>();
            var childFolderObjs = folderObj.Children.Select(t => t[new Guid("63aa0058-4a37-4754-8973-ffbc1b88f576")].Value.ToString());      // Наименование
            foreach (var файл in подключаемыеФайлы)
            {
                // проверка на наличие файла в папке "Согласование"
                if (childFolderObjs.Contains(файл["Наименование"].ToString()))
                {
                    errorExist = true;
                    filesExist.Add(файл["Наименование"].ToString());
                }

                // проверка на объекты в корзине
                var filesInRecycleBin = fileReference.GetDeletedObjects();
                if (!filesInRecycleBin.Any())
                    continue;

                var filesInRecycleBinFolderAgreement = filesInRecycleBin.Where(t => t.Parent.SystemFields.Guid == FolderGuid);
                if (!filesInRecycleBinFolderAgreement.Any())
                    continue;

                var filesInRecycleBinFolderAgreementNames = filesInRecycleBinFolderAgreement.Select(t => t[new Guid("63aa0058-4a37-4754-8973-ffbc1b88f576")].Value.ToString());      // Наименование
                if (filesInRecycleBinFolderAgreementNames.Contains(файл["Наименование"].ToString()))
                {
                    errorRecycleBin = true;
                    filesRecycleBin.Add(файл["Наименование"].ToString());
                }
            }

            if (errorExist || errorRecycleBin)
            {
                string resError = "";
                if (errorExist)
                    resError += string.Format("Следующие файлы содержатся в папке Согласования:\r\n{0}", string.Join("\r\n", filesExist));

                if (errorRecycleBin)
                    resError += string.Format("Следующие файлы содержатся в Корзине:\r\n{0}", string.Join("\r\n", filesRecycleBin));

                Error(resError);
            }
        }

        if (запускать)
            БизнесПроцессы.Запустить(имяПроцедуры, объектыДляБП);

    }
    
    public string ПроверитьНаличиеФайлаПодлинника(Объекты объекты)
    {
        List<string> errors = new List<string>();
        string result = "";
        foreach (var объект in объекты)
        {
            if (объект.Справочник.Имя == "Электронная структура изделий")
            {
                var документ = объект.СвязанныйОбъект["Связанный объект"];
                if (объект["Формат"].ToString() == "БЧ")
                	continue;
                
                var файлы = документ.СвязанныеОбъекты["Файлы"];
                bool естьФайлПодлинник = false;
                
                /*var test = НайтиОбъект("test", "Id", "1");
                test.Изменить();
                test["txt"] = result;
                
                test.Сохранить();
                */
                
                if (объект["Стадия"].ToString() != "Утверждено" && объект["Стадия"].ToString() != "Хранение")
                    errors.Add(string.Format("Документ \"{0} ({1})\" не находится на стадии \"Утверждено\"", объект.ToString(), объект.Параметр["Обозначение"].ToString()));
                
                if (!объект.Подписи.Any(t => t.ТипПодписи == "Разраб."))
                    errors.Add(string.Format("У документа \"{0} ({1})\" отсутствует подпись Разраб.", объект.ToString(), объект.Параметр["Обозначение"].ToString()));
                if (!объект.Подписи.Any(t => t.ТипПодписи == "Н. контр."))
                    errors.Add(string.Format("У документа \"{0} ({1})\" отсутствует подпись Н. контр.", объект.ToString(), объект.Параметр["Обозначение"].ToString()));
                
                foreach (var файл in файлы)
                {
                    FileObject file = (ReferenceObject)файл as FileObject;
                    string fileExtension = file.Class.Extension.ToLower();
                    if (SupportedExtensions.Contains(fileExtension) && файл["Стадия"].ToString() == "Утверждено")
                    {
                        естьФайлПодлинник = true;
                        
                        if (!файл.Подписи.Any(t => t.ТипПодписи == "Разраб."))
                            errors.Add(string.Format("У файла \"{0}\" отсутствует подпись Разраб.", файл.ToString()));
                        if (!файл.Подписи.Any(t => t.ТипПодписи == "Н. контр."))
                            errors.Add(string.Format("У файла \"{0}\" отсутствует подпись Н. контр.", файл.ToString()));
                    }
                }
                if (!естьФайлПодлинник)
                    errors.Add(string.Format("У документа \"{0} ({1})\" отсутствует утвержденный файл-подлинник.", документ.ToString(), объект.Параметр["Обозначение"].ToString()));
            }
        }
        
        if (errors.Any())
            result = string.Join("\r\n", errors);
        else
            result = "Проверка пройдена успешно.";
                
        return result;
    }
    
}

