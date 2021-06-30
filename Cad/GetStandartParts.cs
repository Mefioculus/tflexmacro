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

			string message = string.Empty;

            List<StData> allElementInProductStructure = new List<StData>();
            foreach (string file in files) {
                allElementInProductStructure.AddRange(GetStandartParts(file));
            }

            WriteResultInFile(allElementInProductStructure);

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
		private static List<StData> GetStandartParts(string pathToFile) {
            // Данный метод будет открывать документ, получать из документа необходимые данные
            // размещать их в классе контейнере, закрывать документ
            List<StData> result = new List<StData>();
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
                        result.Add(new StData(Path.GetFileNameWithoutExtension(pathToFile), name, quantity));
                    }
                }
            }


            // Закрываем документ
            document.Close();
            return result;
		}
        #endregion Method GetStandartParts

        #region Method WriteResultInFile
        private static void WriteResultInFile(List<StData> result) {
            string stringOfData = string.Empty;
            // Запрашиваем у пользователя путь для сохранения результатов выгрузки
            Forms.SaveFileDialog dialog = new Forms.SaveFileDialog();
            dialog.Filter = "Comma-Separated Values files (*.csv)|*.csv|All files (*.txt)|*.txt";
            dialog.FilterIndex = 1;
            dialog.RestoreDirectory = true;
            dialog.AddExtension = true;

            if (dialog.ShowDialog() == Forms.DialogResult.OK) {
                // Получаем кодировку, в которую будем производить запись
                Encoding targetEncoding = Encoding.GetEncoding(1251);
                using (StreamWriter sw = new StreamWriter(dialog.FileName, false, targetEncoding)) {
                    // Пишем заголовок
                    sw.WriteLine("Файл;Стандартное изделие;Количество");
                    // Пишем регулярную часть
                    foreach (StData data in result) {
                        sw.WriteLine(data.ToString());
                    }
                }
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
		}
        #endregion Class StData
        #endregion ServiceClasses
	}
}
