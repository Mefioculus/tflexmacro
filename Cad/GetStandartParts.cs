using System;
using System.IO;
using System.Collections;
using System.Collections.Generic;
using TFlex;
using TFlex.Model;
using TFlex.Model.Model2D;
using TFlex.Model.Model3D;

namespace NewMacroNamespace
{
	public class NewMacroClass
	{
		public static void NewMacro()
		{
			// Метод для получения стандартных изделий

			List<string> files = GetPathsToFiles();

			string message = string.Empty;

            List<StData> allElementInProductStructure = new List<StData>();
            foreach (string file in files) {
                allElementInProductStructure.AddRange(GetStandartParts(file));
            }

            WriteResultInFile(allElementInProductStructure);

		}
		
		// Метод для получения grb файлов в директории
		private static List<string> GetPathsToFiles() {
			// Метод для получения путей файлов
			List<string> result = new List<string>();
			string pathToDirectory = GetDirectory();
			if (pathToDirectory == string.Empty)
				return result;

			// Производим чтение grb файлов
			foreach (string pathToFile in Directory.GetFiles(pathToDirectory)) {
				if (pathToFile.ToLower().Contains(".grb")) {
					result.Add(pathToFile);
				}
			}


			return result;
		}
		
		// Метод для запроса директории у пользователя
		private static string GetDirectory() {
			string result = "C:\\Users\\gukovry\\Desktop\\DirectoryForTestingPurpose";
			return result;
		}

		// Метод для получения данных о входящих стандартных изделиях
		private static List<StData> GetStandartParts(string pathToFile) {
            // Данный метод будет открывать документ, получать из документа необходимые данные
            // размещать их в классе контейнере, закрывать документ
            List<StData> result = new List<StData>();
            // Открываем документ
            Document document = Application.OpenDocument(pathToFile);
            foreach (ProductStructure specification in document.GetProductStructures()) {
                foreach (var element in specification.GetAllRowElements()) {
                    result.Add(new StData(Path.GetFileNameWithoutExtension(pathToFile), element.Name, 1));
                }
            }
            // Закрываем документ
            document.Close();
            return result;
		}

        private static void WriteResultInFile(List<StData> result) {
            string stringOfData = string.Empty;

            using (StreamWriter sw = new StreamWriter("C:\\Users\\gukovry\\Desktop\\result.txt")) {
                foreach (StData data in result) {
                    sw.WriteLine(data.ToString());
                }
            }

        }

		// Класс контейнер
		private class StData {
			public string NameOfAssembly;
			public string NameOfPart;
			public int Quantity;

			public StData(string nameAssembly, string namePart, int quant) {
                this.NameOfAssembly = nameAssembly;
                this.NameOfPart = namePart;
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

	}
}
