using System;
using System.IO;
using System.Collections;
using System.Collections.Generic;
using System.Text; // Для установки кодировки при записи в файл
using Forms = System.Windows.Forms;
using TFlex;
using TFlex.Model;
using TFlex.Model.Model2D;
using TFlex.Model.Model3D;

namespace NewMacroNamespace
{
	public class NewMacroClass {
        #region Entry points
		public static void ВыгрузитьПереченьСтандартныхИзделий() {
			// Метод для получения стандартных изделий

			List<string> files = GetPathsToFiles();
            if (files.Count == 0)
                return;

            CollectionOfStData loadData = new CollectionOfStData();
            foreach (string file in files) {
                loadData.AddRange(GetStandartParts(file));
            }

            WriteResultInFile(loadData);

		}
        #endregion Entry points
		
        #region ServiceMethods
        #region Method GetPathsToFiles
		// Метод для получения grb файлов в директории
		private static List<string> GetPathsToFiles() {
			// Метод для получения путей файлов
			List<string> result = new List<string>();
			string pathToDirectory = GetDirectory();
			if (pathToDirectory == string.Empty)
				return result;

			// Производим чтение grb файлов
			foreach (string pathToFile in Directory.GetFiles(pathToDirectory)) {
				if (pathToFile.ToLower().Contains(".grb"))
					result.Add(pathToFile);
			}

			return result;
		}
        #endregion Method GetPathsToFiles
		
        #region Method GetDirectory
		// Метод для запроса директории у пользователя
		private static string GetDirectory() {
            string result = string.Empty;
            
            Forms.FolderBrowserDialog dialog = new Forms.FolderBrowserDialog();
            dialog.Description = "Укажите директорию, в которой находятся целевые CAD файлы";
            dialog.ShowNewFolderButton = true;
            dialog.RootFolder = Environment.SpecialFolder.MyComputer;
            
            if (dialog.ShowDialog() == Forms.DialogResult.OK)
                result = dialog.SelectedPath;

            return result;
		}
        #endregion Method GetDirectory

        #region Method GetStandartParts
		// Метод для получения данных о входящих стандартных изделиях
		private static CollectionOfStData GetStandartParts(string pathToFile) {
            // Данный метод будет открывать документ, получать из документа необходимые данные
            // размещать их в классе контейнере, закрывать документ
            CollectionOfStData result = new CollectionOfStData();
            // Открываем документ только на чтение в невидимом режиме
            Document document = TFlex.Application.OpenDocument(pathToFile, false, true);
            foreach (ProductStructure specification in document.GetProductStructures()) {
                foreach (var element in specification.GetAllRowElements()) {
                    string message = string.Empty;
                    if (element.GetTextProperty("Раздел") == "Спецификации\\Стандартные изделия") {
                        // Получаем требуемые параметры
                        string name = element.GetTextProperty("Сводное наименование");
                        int quantity = 0;
                        if (element.GetIntProperty("Количество") != null)
                            quantity = (int)element.GetIntProperty("Количество");
                        result.Add(Path.GetFileNameWithoutExtension(pathToFile), name, quantity);
                    }
                }
            }


            // Закрываем документ
            document.Close();
            return result;
		}
        #endregion Method GetStandartParts

        #region Method WriteResultInFile
        private static void WriteResultInFile(CollectionOfStData data) {
            // Запрашиваем у пользователя путь для сохранения результатов выгрузки
            Forms.SaveFileDialog dialog = new Forms.SaveFileDialog();
            dialog.Filter = "Comma-Separated Values files (*.csv)|*.csv|All files (*.txt)|*.txt";
            dialog.FilterIndex = 1;
            dialog.RestoreDirectory = true;
            dialog.AddExtension = true;

            if (dialog.ShowDialog() == Forms.DialogResult.OK) {
                // Получаем кодировку, в которую будем производить запись
                Encoding targetEncoding = Encoding.GetEncoding(1251);

                data.ExportToCsv(dialog.FileName, false, targetEncoding);
            }
        }
        #endregion Method WriteResultInFile
        #endregion ServiceMethods

        #region ServiceClasses
        #region Class StData
		// Класс контейнер
		private class StData {
			public string NameOfAssembly;
			public string NameOfPart;
			public int Quantity;
            public static string CsvTableHead = "Файл;Стандартное изделие;Количество";

			public StData(string nameAssembly, string namePart, int quant) {
                // Для того, чтобы в названиях значений не оказалось символов, которые могут использоваться в
                // качестве разделителей, производим замену этих символов на смежные
                this.NameOfAssembly = nameAssembly.Replace(";", "/");
                this.NameOfPart = namePart.Replace(";", "/").Replace("%%042", "х").Replace("%%S", " ").Replace("%%-", "-");
                this.Quantity = quant;
			}

            public override string ToString() {
                // TODO Улучшить вывод данных
                return string.Format("{0}; {1}; {2}",
                            this.NameOfAssembly,
                            this.NameOfPart,
                            this.Quantity.ToString());
            }

            public string ToString(string option) {
                string result = string.Empty;
                // TODO предусмотреть другие варианты вывода информации
                switch (option) {
                    case "csv":
                        result = string.Format("{0}; {1}; {2}",
                                        this.NameOfAssembly,
                                        this.NameOfPart,
                                        this.Quantity.ToString());
                        break;
                    case "text":
                        result = string.Empty;
                        break;
                    default:
                        break;
                }

                return result;
            }
		}
        #endregion Class StData

        #region Class CollectionOfStData
        private class CollectionOfStData {
            private List<StData> data;

            public int Count {
                get {
                    return data.Count;
                }
            }

            public CollectionOfStData () {
                // Конструктор
                this.data = new List<StData>();
            }

            // Индексатор
            public StData this[int index] {
                get {
                    return data[index];
                }
                set {
                    data[index] = value;
                }
            }

            public void Add(string file, string name, int quantity) {
                StData data = new StData(file, name, quantity);
                this.data.Add(data);
            }

            public void AddRange(CollectionOfStData collection) {
                for (int i = 0; i < collection.Count; i++) {
                    string nameAssembly = collection[i].NameOfAssembly;
                    string namePart = collection[i].NameOfPart;
                    int quantity = collection[i].Quantity;

                    this.Add(nameAssembly, namePart, quantity);
                }
            }

            public override string ToString() {
                // TODO Реализовать вывод данный в строку
                string result = string.Empty;
                return result;
            }

            public void ExportToCsv(string pathToFile, bool append, Encoding encoding) {
                // TODO Метод для экспортирования данных в CSV
                if (this.data.Count != 0) {
                    using (StreamWriter sw = new StreamWriter(pathToFile, append, encoding)) {
                        sw.WriteLine(StData.CsvTableHead);
                        foreach (StData row in this.data) {
                            sw.WriteLine(row.ToString("csv"));
                        }
                    }
                }
                else
                    Forms.MessageBox.Show("Так как файл не содержит информации, он не будет создан", "Ошибка");
            }
        }
        #endregion Class CollectionOfStData
        #endregion ServiceClasses
	}
}
