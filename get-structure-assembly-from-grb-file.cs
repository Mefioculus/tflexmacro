using System;
using System.Collections;
using System.Collections.Generic;
using System.Windows.Forms;
using System.IO;
using SevenZipNET; // Пространство имен для распакования 7zip архивов
using System.Xml; // Пространство имен для работы с xml файлами

using TFlex.DOCs.Model.Macros;

// Для работы макроса потребуется подключить следующие библиотеки
// 7z.NET.dll


public class Macro : MacroProvider {
    public Macro(MacroContext context)
        : base(context)
    {
    }

    #region Fields and Properties
    private bool Debug = false;
    private string nameOfDirectoryForStoringUnpackFiles = "Unpacked";
    private string pathToTempDirectory = string.Empty;
    #endregion Fields and Properties

    #region EntryPoints
    public override void Run () {
        List<string> grbFiles = GetPathsGrbFiles();
        if (grbFiles.Count == 0)
            return;

        foreach (string file in grbFiles) {
            // Разархивируем файл и получаем из него xml структуру документа
            string xmlStructure = GetXmlStructure(file);
            if (xmlStructure == string.Empty) {
                Message("Ошибка", string.Format("'{0}' не содержит необходимого для работы структурного файла", Path.GetFileName(file)));
                continue;
            }
            
            // Обрабатываем xml документ
            XmlDocument doc = new XmlDocument();
            doc.LoadXml(xmlStructure);
            // Получаем корневой элемент
            XmlElement root = doc.DocumentElement;

            List<string> listOfStandartElement = GetListOfStandartElement(root);

            string message = string.Format("Список стандартных элементов, которые входят в файл '{0}'", Path.GetFileName(file));
            foreach(string element in listOfStandartElement) {
                message += string.Format("\n- {0}", element);
            }
            Message("Информация", message);
        }

        Message("Информация", "Работа макроса завершена");
    }
    #endregion EntryPoints

    #region Service methods
    #region Method GetPathsGrbFiles
    private List<string> GetPathsGrbFiles() {
        // Метод для получения grb файлов, которые находятся в указанной
        // пользователем директории
        List<string> listOfGrbFilePaths = new List<string>();
        // Запрашиваем у пользователя директорию, которая содержит grb файлы
        string directory = GetPathToDirectory();
        // По этому пути будут храниться все временные файлы
        pathToTempDirectory = Path.Combine(directory, nameOfDirectoryForStoringUnpackFiles);
        // Сразу создаем директорию для временного хранения распакованных файлов
        Directory.CreateDirectory(pathToTempDirectory);

        if (!Directory.Exists(directory)) {
            Message("Ошибка", string.Format("Указанная директория \n'{0}'\nне существует", directory));
            return listOfGrbFilePaths;
        }
        
        string[] listOfFiles = Directory.GetFiles(directory);

        foreach (string path in listOfFiles) {
            if (path.ToLower().Contains(".grb")) {
                listOfGrbFilePaths.Add(path);
            }
        }

        if (Debug == true) {
            string message = "Найденные grb файлы:";
            foreach (string file in listOfGrbFilePaths) {
                message += string.Format("\n- {0}", Path.GetFileName(file));
            }
            Message("Информация", message);
        }
        
        // Находим все файлы с расширением grb в данной директории
        return listOfGrbFilePaths;
    }
    #endregion Method GetPathsGrbFiles

    #region Method GetPathToDirectory
    private string GetPathToDirectory() {
        // Запросить у пользователя путь к директории
        string result = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "DirectoryForTestingPurpose");
        if (!Directory.Exists(result))
            Directory.CreateDirectory(result);
        // TODO Реализовать диалог выбора директории при помощи Forms
        if (Debug == true)
            Message("Информация", string.Format("Выбранная директория '{0}'", result));
        return result;
    }
    #endregion Method GetPathToDirectory

    #region Method GetXmlStructure
    private string GetXmlStructure(string pathToFile) {
        // Для начала определяем путь, в котором будет храниться разархивированный файл
        string destination = Path.Combine(pathToTempDirectory, Path.GetFileNameWithoutExtension(pathToFile));
        // Указываем путь, по которому расположена утилита, котоаря будет занимтаься разархивированем
        // (предварительно она должна быть установлена)
        SevenZipExtractor.Path7za = "D:\\Дистрибутивы\\7z1900-extra\\7za.exe";
        SevenZipExtractor extractor = new SevenZipExtractor(pathToFile);
        Directory.CreateDirectory(destination);

        if (Debug == true) {
            string message = string.Format("Файлы, которые содержатся в '{0}':", Path.GetFileNameWithoutExtension(pathToFile));
            foreach (var file in extractor.Files) {
                message += string.Format("\n- {0}", file.Filename);
            }
            Message("Информация", message);
        }

        extractor.ExtractAll(destination);

        // После этого открываем файл и возвращаем его содержимое
        string targetFile = Path.Combine(destination, "TFPDMProductAssemblyDataEx");
        string result = string.Empty;
        if (File.Exists(targetFile))
            result = File.ReadAllText(targetFile);
        else
            return result;

        // Производим обрезку данных в начале строки, которые не имеют ничего общего с xml форматом
        result = result.Remove(0, result.IndexOf("<"));
        // Производим обрезку данных в конце строки
        result = result.Remove(result.IndexOf("</AssembliesList>") + 17);

        if (Debug == true) {
            Message("Данные, которые содержатся в файле", result);
        }

        return result;
    }
    #endregion Method GetXmlStructure

    #region Method GetListOfStandartElement
    private List<string> GetListOfStandartElement (XmlElement item) {
        List<string> result = new List<string>();
        // Метод для получения данных из xml строки
        if (item.LocalName == "TreeNodeData") {
            if (item.Attributes["BomSection"].InnerText == "Стандартные изделия")
                result.Add(string.Format("{0}, Количество - {1}",
                            item.Attributes["Name"].InnerText,
                            item.ParentNode.Attributes["Amount"].InnerText));
        }

        foreach(var element in item.ChildNodes) {
            if (element is XmlElement) {
                result.AddRange(GetListOfStandartElement((XmlElement)element));
            }
        }

        return result;
    }

    #endregion Method GetListOfStandartElement
    #endregion Service methods
}
