using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using TFlex.DOCs.Model.Macros;
using TFlex.DOCs.Model.References;
using TFlex.DOCs.Model.References.Files;
using TFlex.DOCs.Model.References.Processes;
using TFlex.DOCs.Model.References.ActiveAction;
using TFlex.DOCs.Model.References.ActiveActions;
using TFlex.DOCs.Model.Processes.Events.Contexts;
using TFlex.DOCs.Model.Processes.Events.Contexts.Data;
// Пространство имен для импортирования документов CAD в другие форматы
using TFlex.DOCs.Model.FilePreview.CADService;
// Пространство имен, которое позволяет работать с классом NomenclatureObject
using Nomenclature = TFlex.DOCs.Model.References.Nomenclature;
// Пространство имен для взаимодействия с рабочим столом
using TFlex.DOCs.Model.Desktop;
// Для работы со стадиями
using TFlex.DOCs.Model.Stages;
// Для использования класса ObjectInfo
using TFlex.DOCs.Model.Processes.Objects;

namespace CreateRealFilesForSTO
{
    public class Macro: MacroProvider
    {
        // Нужно написать код, который будет получать объекты с текущего бизнесс процесса и создавать подлинники в формате
        // tiff, размещать их в директории Архив ОГТ\Подлинники, а так же будет привязывать файлы по связи к объектам
        // электронной структуры изделия
        

        public Macro(MacroContext context)
            : base(context)
        {
        }

        internal static class Guids
        {
            // Guid абстрактного класса "Материальный объект"
            internal static Guid МатериальныйОбъект = new Guid("0ba28451-fb4d-47d0-b8f6-af0967468959");
            // Guid связи файлового объекта с справочником документы
            internal static Guid СвязьНаДокументы = new Guid("9eda1479-c00d-4d28-bd8e-0d592c873303");
            // Guid бизнесс-просесса Согласование КД на технологическое оснащение
            internal static Guid СогласованиеКДнаТО = new Guid("78737786-5d1d-439c-8984-0fff157b1124");
            // Guid стадии "Действие на компьютере пользователя" БП "Согласование КД на технологическое оснащение"
            internal static Guid ДействиеНаКомпьютереПользователя = new Guid("fbc24abd-43a4-40b5-90c3-300f75a564bf");
            // Guid абстрактного класса расширения GRB файлов
            internal static Guid GrbРасширение = new Guid("73c01001-41d6-454d-816f-311528e5115d");
        }

        public void СоздатьПодлинник()
        {
            // Создаем переменную, в которой будет храниться текущий бизнесс процесс
            ProcessReferenceObject process = null;

            // Создаем переменную, в которой будет храниться текущее действие
            ActiveActionReferenceObject activeAction = null;

            // Получаем контекст событий по БП
            EventContext eventContext = Context as EventContext;
            if (eventContext != null)
            {
                var data = eventContext.Data as StateContextData;
                // Получаем данные текущего БП
                process = data.Process;
                // Получаем текущее действие
                activeAction = data.ActiveAction;
            }
            else
            {
                // Данный блок кода срабатывает в том случае, если не удалось получить EventContext
                // В этом случае мы получаем текущий процесс, узнаем его GUID и в соответствии с этим
                // выставляем GUID действия на компьютере пользлователя, которое будет производить формирование
                // подлинников
                process = Context.ReferenceObject as ProcessReferenceObject;

                // Создаем переменную, в которой будем хранить Guid состояния "Действие на компьютере пользователя"
                Guid userMacrosState = Guid.Empty;

                if (process.Procedure.SystemFields.Guid == Guids.СогласованиеКДнаТО)
                {
                    userMacrosState = Guids.ДействиеНаКомпьютереПользователя;
                }
                else
                {
                    Сообщение("Ошибка", "Необходимый бизнес-процесс не был найден, работа макроса будет прекращена");
                    return;
                }

                activeAction = process.ActiveActions.AsList.FirstOrDefault(action => action.StepGuid == userMacrosState);
            }

            // Далее, мы получаем данные текущего действия, а вернее положение курсора в данном действии
            ActiveActionData activeActionData = activeAction.GetData<ActiveActionData>();
            // Получаем объекты, подключенные к БП
            List<ReferenceObject> processObject = activeActionData.GetReferenceObjects().ToList();

            // Формируем подлинники и получаем список сформированных подлинников
            List<ReferenceObject> realFiles = CreateRealFiles(processObject);
            
            // Если будет необходимость добавть подлинники к бизнесс процессу, то далее можно будет это сделать через
            // список файлов realFiles
            if (realFiles.Any())
            {
                // Подключаем к бизнесс-процессу подлинники
                process.BeginChanges();
                // Для начала получаем список guid уже подключенных к бизнес-процессу объектов
                var existingObjectGuids = activeActionData.ObjectsList
                    .Where(oi => oi.IsEngaged)
                    .Select(obj => obj.ObjectGuid)
                    .ToArray();
                // Формируем список объектов, которые нужно добавить к бизнес процессу
                // (отфильтровываем те, которые могли быть уже добавлены)
                ReferenceObject[] addingObjects = realFiles
                    .Where(obj => (obj != null) && (!existingObjectGuids.Contains(obj.SystemFields.Guid)))
                    .ToArray();

                if (addingObjects.Any())
                {
                    activeActionData.ObjectsList.AddRange(addingObjects.Select(obj => ObjectInfo.Create(obj, true)));
                    activeAction.SaveData(activeActionData);
                }
                process.EndChanges();
            }

            if (realFiles.Any())
            {
                string message = "Следующие подлинники были созданы:\n\n";
                foreach(FileObject realFile in realFiles)
                {
                    message += string.Format("{0}\n", realFile.ToString());
                }
                Сообщение("Информация", message);
            }
            else
                Сообщение("Информация", "Для данного процесса не было создано ни одного подлинника");
        }

        private List<ReferenceObject> CreateRealFiles(List<ReferenceObject> referenceObjects)
        {
            // Данный метод обрабатывает прикрепленые к бизнес-процессу объекты и производит предварительные проверки,
            // после чего запускает метод, который создает подлинники (ConvertGrbToTiff()).
            List<ReferenceObject> realFiles = new List<ReferenceObject>();
            
            // Создаем экземпляр провайдера CAD для импорта чертежей. Данный провайдер будет передаваться в функцию,
            // которая будет создавать подлинники. Я инициализирую ее сдесь, чтобы делать это один раз, а не повторять
            // это действие каждый раз для каждого документа.
            CadDocumentProvider provider = CadDocumentProvider.Connect(Context.Connection, ".grb");

            foreach (ReferenceObject referenceObject in referenceObjects)
            {

                // Проверка типа подключенного элемента на то, что он унаследован от абстрактного типа
                // "материальный объект" справочника "Электронная структура изделия"
                if (referenceObject.Class.IsInherit(Guids.МатериальныйОбъект))
                {
                    // Получаем папку для сохранения файлов
                    FolderObject folder = GetFolderForSaving();
                    if (folder == null)
                    {
                        Сообщение("Ошибка", "Директория для сохранения подлинников не была найдена, работа макроса прекращена");
                        return null;
                    }

                    // Получаем файла CAD, для которых необходимо сделать подлинники
                    List<FileObject> filesOnProcessing = GetAllLinkedGRBFiles(referenceObject);
                    if (filesOnProcessing != null)
                    {
                        foreach (FileObject file in filesOnProcessing)
                        {
                            realFiles.Add((ReferenceObject)ConvertGrbToTiff(file, folder, provider));
                        }
                    }
                    else
                        continue;
                }
            }

            return realFiles;
        }

        private FileObject ConvertGrbToTiff(FileObject file, FolderObject folder, CadDocumentProvider provider)
        {
            // Данный метод работает с одним файлом за один раз
            
            // Загружаем последнюю версию файла с файлового сервера (данный метод ничего не возвращает, он всего лишь
            // сохраняет файл в рабочую директорию для последующей с ним работы)
            file.GetHeadRevision();

            // Создаем переменную, в которой в дальнейшем будет храниться tiff файла подлинника
            FileObject tiffFile = null;

            // Получаем пути для временного расположения файла на локальном компьютере пользователя
            string tmpFile = Path.GetTempFileName();
            string tempFileFullPath = Path.ChangeExtension(tmpFile, "tif");

            // Создаем переменную, которая будет хранить путь к результуриющему файлу. Такая переменная уже существует,
            // но цель данной переменной заключается в том, что ее значение будет присваиваться уже после произведения
            // экспорта, так что если экспорт не будет произведен, она будет равна null
            string tiffPath = string.Empty;

            // Создаем tiff файл подлинника на локальной машине
            try
            {
                // Открывает документ на чтение (за это отвечает булевый аргумент readOnly)
                using (var document = provider.OpenDocument(file.LocalPath, true))
                {
                    // Задаем контекст экспорта через путь результирующего файла. Путь должен так же включать расширение
                    // результирующего файла для того, чтобы программа понимала, в какой формат производить экспорт
                    ExportContext exportContext = new ExportContext(tempFileFullPath);
                    // Получаем список страниц документа
                    PageInfo[] pages = document.GetPagesInfo();
                    // Добавляем все страницы в область экспорта документа
                    for (int i = 0; i < pages.Count(); i++)
                    {
                        exportContext.Pages.Add(i);
                    }
                    tiffPath = document.Export(exportContext);
                    document.Close(false);
                }

            }
            catch (Exception e)
            {
                Сообщение("Ошибка", string.Format("Подлинник не был создан\n\n{0}", e.Message));
            }

            // Подключаем сформированный tiff файл подлинника в базу данных
            try
            {
                // Подключаем документ к сборочной единице
                tiffFile = AddAndLinkRealFile(folder, file, tiffPath);
            }
            catch (Exception e)
            {
                Сообщение("Ошибка", string.Format("Ошибка при подключении подлинника в базу данных\n\n{0}", e.Message));
            }

            // Присваиваем сданию подлиннику
            ChangeStageForRealFile(file, tiffFile);

            // Удаляем временные файлы с локального компьютера пользователя
            DeleteTempFile(tmpFile);
            DeleteTempFile(tempFileFullPath);

            return tiffFile;
        }

        private void ChangeStageForRealFile(FileObject sourceFile, FileObject realFile)
        {
            // Метод для изменения стадии на подлиннике на такую же стадию, в которой находится исходный файл
            // CAD системы

            // Для начала получаем стадию, в которой находится исходный файл
            Stage stage = null;
            try
            {
                stage = sourceFile.SystemFields.Stage.Stage;
            }
            catch (Exception e)
            {
                Сообщение("Ошибка", string.Format("При получении стадии с объекта '{0}' возникла ошибка:\n\n{1}",
                            sourceFile.ToString(), e.Message));
                throw new NullReferenceException();
            }

            // Меняем стадию у подлинника
            try
            {
                stage.Change(new List<ReferenceObject>() { (ReferenceObject)realFile });
            }
            catch (Exception e)
            {
                Сообщение("Ошибка", string.Format("При изменении стадии объекта '{0}' возникла ошибка:\n\n{1}",
                            realFile.ToString(), e.Message));
                throw new Exception("Unknown error");
            }
        }

        private FileObject AddAndLinkRealFile(FolderObject folder, FileObject sourceFile, string pathToRealFile)
        {
            // Создаем файловый объект, который будет представлять подлинник в файловом справочнике
            FileObject tiffFile = null;

            // Проверяем, существует ли файл, который необходимо подключить в базу
            if (!string.IsNullOrEmpty(pathToRealFile))
            {
                // Генерируем название файла с исходного файла
                string nameOfRealFileInDB = string.Format("{0}.tif", Path.GetFileNameWithoutExtension(sourceFile.Name));
                
                // Проверяем, есть ли этот файл в базе данных
                FileObject oldRealFile = SearchExistingRealFile(nameOfRealFileInDB, folder);

                if (oldRealFile == null)
                {
                    // Создаем новый файл и подключаем его к ДСЕ

                    tiffFile = folder.CreateFile(pathToRealFile, string.Empty, nameOfRealFileInDB,
                            sourceFile.Reference.Classes.GetFileTypeByExtension("tif"));
                    
                    // Привязываем файл подлинника в ДСЕ, для которой он был создан
                    // Для начала получаем объект, к которому был привязан CAD файл
                    foreach (ReferenceObject dce in sourceFile.GetObjects(Guids.СвязьНаДокументы))
                    {
                        tiffFile.BeginChanges();
                        tiffFile.AddLinkedObject(Guids.СвязьНаДокументы, dce);
                        tiffFile.EndChanges();
                    }
                    if (tiffFile.CanCheckIn)
                        Desktop.CheckIn(tiffFile, string.Format("Создание подлинника из '{0}'", sourceFile.Path), false, null);

                }
                else
                {
                    // Редактируем существующий файл, подключать его к ДСЕ не требуется
                    tiffFile = oldRealFile;

                    tiffFile.GetHeadRevision();
                    string localPath = tiffFile.LocalPath;

                    // Производим замену файла
                    tiffFile.CheckOut(false);
                    tiffFile.BeginChanges();
                    // Удаляем старую версию, копируем новую версию
                    File.Delete(localPath);
                    File.Move(pathToRealFile, localPath);
                    // Заканчиваем изменения объекта
                    tiffFile.EndChanges();
                    // Применяем изменения к объекту
                    Desktop.CheckIn(tiffFile, string.Format("Обновление подлинника {0}", nameOfRealFileInDB), false, null);
                }
            }
            else
            {
                Сообщение("Ошибка", "Попытка добавить несуществующий файл в базу (возникла проблема при экспорте подлинников)");
                throw new FileNotFoundException();
            }
            return tiffFile;
        }

        private FileObject SearchExistingRealFile(string nameFile, FolderObject folder)
        {
            FileObject resultOfSearch = null;

            // Проверяем все объекты, которые входят в папку на соответствие имени файла
            foreach (ReferenceObject child in folder.Children)
            {
                if (child.ToString() == nameFile)
                    resultOfSearch = (FileObject)child;
            }
            
            return resultOfSearch;
        }

        private void DeleteTempFile(string pathToFile)
        {
            try
            {
                // Инициализируем новый объект FileInfo
                FileInfo file = new FileInfo(pathToFile);
                if (file.Exists && !file.IsReadOnly)
                    file.Delete();
            }
            catch
            {
                Сообщение("Ошибка", string.Format("Не удалось удалить временный файл, расположенный по пути:\n{0}", pathToFile));
            }
        }
        
        private List<FileObject> GetAllLinkedGRBFiles (ReferenceObject nomenclature)
        {
            // Данный метод предназначен для получения списка файлов для создания подлинников.
            
            // Для начала проверяем, что файл, переданный в функцию является объектом справочника
            // Электронная структура изделия
            Nomenclature.NomenclatureObject nom = nomenclature as Nomenclature.NomenclatureObject;
            
            if (nom != null)
            {
                // Пробуем получить доступ к файлам данного объекта, если они имеются в наличии
                List<FileObject> GrbFiles = new List<FileObject>();
                foreach (FileObject file in nom.LinkedObject.GetAllLinkedFiles())
                {
                    // Проверяем файлы на соответствие расширениям GRB
                    if (file.Class.IsInherit(Guids.GrbРасширение))
                    {
                        // Добавляем файл на выдачу
                        GrbFiles.Add(file);

                        // Проверка наличия связи на объект документов
                        
                        string message = string.Empty;

                        foreach (var document in file.GetObjects(Guids.СвязьНаДокументы))
                            message += string.Format("{0}\n", document.ToString());
                    }
                }

                if ((GrbFiles == null) | (GrbFiles.Count == 0))
                    return null;
                else
                    return GrbFiles;
            }
            return null;
        }

        private FolderObject GetFolderForSaving()
        {
            string folderPath = Path.Combine("Архив ОГТ", "Подлинники");

            FileReference fileReferenceInstance = new FileReference(Context.Connection);
            FolderObject folder = fileReferenceInstance.FindByRelativePath(folderPath) as FolderObject;

            if (folder == null)
            {
                Сообщение("Ошибка", "Папка для сохранения подлинников не была найдена");
                return null;
            }
            return folder;
        }
    }
}

