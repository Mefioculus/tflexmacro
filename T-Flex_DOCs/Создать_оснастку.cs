using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using TFlex.DOCs.Model;
//using System.Windows.Forms;
using TFlex.DOCs.Model.Macros;
using TFlex.DOCs.Model.Macros.ObjectModel;
using TFlex.DOCs.Model.References;
using TFlex.DOCs.Model.References.Files;
using TFlex.DOCs.Model.References.Nomenclature;
using TFlex.DOCs.Model.References.Users;
using TFlex.DOCs.Model.Search;
using TFlex.DOCs.Model.Signatures;
using TFlex.DOCs.Model.Structure;
using FolderObject = TFlex.DOCs.Model.References.Files.FolderObject;

public class Macro : MacroProvider
{
    private readonly StringBuilder _logRemarks = new StringBuilder();
    Reference docReference;
    Reference osnReference;
    NomenclatureReference nomReference;
    ReferenceObjectSaveSet saveSet = new ReferenceObjectSaveSet();//Объекты на сохранение
    private static readonly string RelativePath = "Архив ОГТ\\Оснастка";//относительный путь до необходимой папки
    private bool IsNeedToWriteLogs = false;
    public Macro(MacroContext context)
        : base(context)
    {
        //if (Вопрос("Хотите запустить в режиме отладки?"))
        //{
        //    System.Diagnostics.Debugger.Launch();
        //    System.Diagnostics.Debugger.Break();
        //}
    }

    public override void Run()
    {
        //Reference reference = Context.Connection.ReferenceCatalog.Find("")?.CreateReference() ?? throw new MacroException("");
        ReferenceInfo referenceInfo = Context.Connection.ReferenceCatalog.Find(new Guid("ac46ca13-6649-4bbb-87d5-e7d570783f26"));//Справочник документы
        docReference = referenceInfo.CreateReference();

        referenceInfo = Context.Connection.ReferenceCatalog.Find(new Guid("853d0f07-9632-42dd-bc7a-d91eae4b8e83"));//Справочник ЭСИ
        nomReference = referenceInfo.CreateReference() as NomenclatureReference;

        referenceInfo = Context.Connection.ReferenceCatalog.Find(new Guid("a13d8c5a-c76c-4df6-91df-fc5bbf8db86d"));//Справочник Каталог оснащения
        osnReference = referenceInfo.CreateReference();

        List<FileReferenceObject> fileReferenceObjects = new List<FileReferenceObject>();

        FolderObject rootFolderObject = null;

        GetSelectFilesAndFolder(out rootFolderObject, out fileReferenceObjects);

        foreach (var fileReferenceObject in fileReferenceObjects)
        {
            ProcessFileObject(fileReferenceObject);

        }

        saveSet.EndChanges();

        if (IsNeedToWriteLogs)
        {
            string logPath = WriteLogs(rootFolderObject);
            Сообщение("Создание завершено", "Создание Оснастки было завершено. Лог можно найти по пути: " + Environment.NewLine + logPath);
        }
            
    }

    private string WriteLogs(FolderObject rootFolderObject)
    {
        string logFilePath = string.Format("{0}{1} {2:d.MM.yyyy HH-mm-ss}.txt", Path.GetTempPath(), "Создание оснастки в Docs", DateTime.Now);
        File.AppendAllText(logFilePath, Environment.NewLine + _logRemarks);

        FileReference fileReference = rootFolderObject.Reference;
        //Загрузка файла в докс
        var fileObj = fileReference.AddFile(logFilePath, rootFolderObject);

        return fileObj.Path;
    }

    private void ProcessFileObject(FileReferenceObject fileReferenceObject)
    {
        //Получаем наименование файла без расширения
        string fileObjectName = Path.GetFileNameWithoutExtension(fileReferenceObject.Name);

        //Проверка, что объект является файлом, иначе пишем в лог
        if (!fileReferenceObject.IsFile)
        {
            IsNeedToWriteLogs = true;
            AddTextLog(string.Format("Объект {0}, не является файлом", fileObjectName));
            return;
        }

        //string filterString = String.Format("[Обозначение] = '{0}'", fileObjectName);
        //var filter = Filter.Parse(filterString, nomReference.ParameterGroup);
        //var nomSearch = nomReference.Find(filter);
        
        //Находим или создаем объект в ЭСИ
        NomenclatureObject nomObject = null;
        var parameterInfo = nomReference.ParameterGroup.OneToOneParameters.Find(new Guid("ae35e329-15b4-4281-ad2b-0e8659ad2bfb")) ?? throw new MacroException("Параметр \"Обозначение\"не найден в ЭСИ");
        var foundObject = nomReference.FindOne(parameterInfo, fileObjectName);
        if (foundObject != null)
        {
            nomObject = foundObject as NomenclatureObject;
            nomObject.LinkedObject.CheckOut(false);
            nomObject.LinkedObject.BeginChanges();
            //nomObject.CheckOut(false);
            nomObject.BeginChanges();
        }
        else
        {
            if (saveSet.FirstOrDefault(x => x.ToString() == fileReferenceObject.ToString()) != null )//если такой объект уже есть в сейвсете
            {
                IsNeedToWriteLogs = true;
                AddTextLog(string.Format("Объект {0} уже находится на сохранении в ЭСИ", fileReferenceObject.ToString()));
                return;
            }
            else
                nomObject = CreateNewNomenclatureObject(fileObjectName);
        }    


        //Добавляем подпись
        var signatureObjectType = nomObject.Signatures.Types.FirstOrDefault(st => st.Name == "Копировал");
        nomObject.SetSignature((User)CurrentUser, signatureObjectType, "");//на объект ЭСИ и Документ

        //nomObject.Signatures.AddAndSign(signatureObjectType, "");
        //nomObject.Signatures.Add(signatureObjectType, (UserReferenceObject)CurrentUser);
        //Объект номенклатурныйОбъект = Объект.CreateInstance(nomObject, Context);
        //номенклатурныйОбъект.ДобавитьПодпись("Копировал", ТекущийПользователь);

        parameterInfo = osnReference.ParameterGroup.OneToOneParameters.Find(new Guid("13d72d00-d7ec-4428-a065-18e5964f9a52")) ?? throw new MacroException("Параметр \"Обозначение\"не найден в Каталоге оснащения");
        foundObject = osnReference.FindOne(parameterInfo, fileObjectName);

        if (foundObject != null)
            nomObject.AddLinkedObject(new Guid("cf883d14-919e-49ba-b9ba-44a732c06924"), foundObject);

        saveSet.Add(nomObject);

        MoveFile(fileReferenceObject, signatureObjectType, nomObject.LinkedObject);//переносим файл в нужную папку и ставим подпись
    }

    private void MoveFile(FileReferenceObject fileReferenceObject, SignatureType signatureObjectType, ReferenceObject document)
    {
        FileReference fReference = new FileReference(Context.Connection) ?? throw new MacroException("Не найден справочник файлов");
        ReferenceObject destFolder = fReference.FindByRelativePath(RelativePath) ?? throw new Exception("Пути Архив ОГТ\\Оснастка не существует"); //Задаем целевую папку

        destFolder.Children.Reload();
        if (destFolder.Children.FirstOrDefault(f => f.ToString() == fileReferenceObject.Name) == null)
        {
            fileReferenceObject.CheckOut(false);
            fileReferenceObject.BeginChanges();
            fileReferenceObject.AddLinkedObject(new Guid("9eda1479-c00d-4d28-bd8e-0d592c873303"), document);//добавляем связь на документ
            fileReferenceObject.SetSignature((User)CurrentUser, signatureObjectType, "");//ставим подпись
            fileReferenceObject.SetParent(destFolder);//переносим файл в нужную папку
            saveSet.Add(fileReferenceObject);
        }
        else
        {
            IsNeedToWriteLogs = true;
            AddTextLog(string.Format("Объект {0} уже есть по пути Архив ОГТ\\Оснастка", fileReferenceObject.ToString()));
        }
    }

    private NomenclatureObject CreateNewNomenclatureObject(string name)
    {
        // Новый объект Документ с типом оснастка
        var classObject = docReference.Classes.Find(new Guid("a49790e6-11e4-480a-8886-a3e29de85465"));//Тип Оснастка
        ReferenceObject newDocument = docReference.CreateReferenceObject(classObject);

        if (newDocument == null)
            Ошибка("newDocument == null");
        //return null;

        // Установка параметров объекта
        newDocument[new Guid("b8992281-a2c3-42dc-81ac-884f252bd062")].Value = name;//Обозначение
        newDocument[new Guid("7e115f38-f446-40ce-8301-9b211e6ce5fd")].Value = "Создано автоматически";//Наименование

        saveSet.Add(newDocument);
        NomenclatureObject toReturn = nomReference.CreateNomenclatureObject(newDocument);// Подключение к Номенклатуре
        if (toReturn == null)
            Ошибка("toReturn == null");
        return toReturn;
    }

    private void GetSelectFilesAndFolder(out FolderObject rootFolderObject, out List<FileReferenceObject> fileReferenceObjects)
    {
        var selectedObjects = Context.GetSelectedObjects();
        rootFolderObject = null;
        fileReferenceObjects = new List<FileReferenceObject>();

        var first = selectedObjects[0] as FileReferenceObject;

        //Получаем родительскую папку
        rootFolderObject = Context.ReferenceObject.Parent as FolderObject;
        if (rootFolderObject == null)
        {
        	rootFolderObject = Context.ReferenceObject as FolderObject;
        }

        if (first.IsFolder)
        {
            foreach (var selectedObject in selectedObjects)
            {
                var folder = selectedObject as FolderObject;
                fileReferenceObjects.AddRange(folder.Children);
            }
        }
        else
        {
            foreach (var selectedObject in selectedObjects)
            {
                //Приводим выбранные объекты к типу файл
                var fileReferenceObject = selectedObject as FileReferenceObject;
                if (fileReferenceObject != null)
                {
                    fileReferenceObjects.Add(selectedObject as FileReferenceObject);
                }
            }
        }

        ////Если были выбрано несколько объектов
        //if (selectedObjects.Length > 1)
        //{
        //    foreach (var selectedObject in selectedObjects)
        //    {
        //        //Приводим выбранные объекты к типу файл
        //        var fileReferenceObject = selectedObject as FileReferenceObject;
        //        if (fileReferenceObject != null)
        //        {
        //            fileReferenceObjects.Add(selectedObject as FileReferenceObject);
        //        }
        //    }
        //}
        //else
        //{
        //    //Если выбран один объект
        //    var fileReferenceObject = Context.ReferenceObject as FileReferenceObject;
        //    if (fileReferenceObject != null)
        //    {
        //        //Если выбранный объект папка
        //        if (fileReferenceObject.IsFolder)
        //        {
        //            //Добавляем дочерние объекты итд
        //            rootFolderObject = fileReferenceObject as FolderObject;
        //            fileReferenceObjects.AddRange(fileReferenceObject.Children);
        //        }
        //        else
        //        {
        //            //Если файл, то должны файл записать в коллекцию, чтобы его обработать и получить папку
        //            rootFolderObject = fileReferenceObject.Parent;

        //            fileReferenceObjects.Add(fileReferenceObject);
        //        }
        //    }
        //    else
        //        Ошибка("Выбранный объект не является частью справочника 'Файлы'");
        //}
    }
    private void AddTextLog(string text)
    {
        _logRemarks.AppendLine(string.Format("{0:d.MM.yyyy HH-mm-ss}: {1}{2}", DateTime.Now, text, Environment.NewLine));
    }
}
