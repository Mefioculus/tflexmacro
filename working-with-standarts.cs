using System;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Collections;
using System.Collections.Generic;


using TFlex.DOCs.Model.Macros;
using TFlex.DOCs.Model.Macros.ObjectModel;
using TFlex.DOCs.Model.Structure;
using TFlex.DOCs.Model.Search;
using TFlex.DOCs.Model.References;

public class Macro : MacroProvider {
    public Macro(MacroContext context)
        : base (context) {
        }

    private Dictionary<TypeOfDocument, Regex> RegexPatterns = new Dictionary<TypeOfDocument, Regex> () {
        [TypeOfDocument.ГОСТ] = new Regex(""),
        [TypeOfDocument.ГОСТ] = new Regex(""),
        [TypeOfDocument.ГОСТ] = new Regex(""),
        [TypeOfDocument.ГОСТ] = new Regex(""),
        [TypeOfDocument.ГОСТ] = new Regex(""),
        [TypeOfDocument.ГОСТ] = new Regex(""),
        [TypeOfDocument.ГОСТ] = new Regex(""),
        [TypeOfDocument.ГОСТ] = new Regex(""),
        [TypeOfDocument.ГОСТ] = new Regex(""),
        [TypeOfDocument.ГОСТ] = new Regex(""),
        [TypeOfDocument.ГОСТ] = new Regex(""),
        [TypeOfDocument.ГОСТ] = new Regex(""),
        [TypeOfDocument.ГОСТ] = new Regex(""),
    }



    public override void Run() {
        ЗагрузитьГосты();
    }

    public void ЗагрузитьГосты() {
        // Данный метод предназначен для добавления кнопки, которая будет производить импорт pdf из выбранной директории
        
        List<RegulatoryDocument> documents = new List<RegulatoryDocument>();
        foreach (string file in GetFilesFromUser()) {
            documents.Add(new RegulatoryDocument(file, TypeOfDocument.ГОСТ));
        }

        Message("Количество загруженных документов", documents.Count.ToString());
        Message("Загруженные документы", string.Join("\n", documents.Select(doc => doc.LinkedFile.Name)));
    }

    private string[] GetFilesFromUser() {
        // Запросить у пользователя директорию, в которой производить поиск
        string pathToDirectory = @"D:\ГОСТы";
        string searchPattern = "*.pdf";

        return Directory.GetFiles(pathToDirectory, searchPattern, SearchOption.AllDirectories);
    }

    public class RegulatoryDocument {
        public string Name { get; private set; }
        public string Designation { get; private set; }
        public FileInfo LinkedFile { get; private set; }
        public TypeOfDocument Type { get; private set; }
        public string TypeString => this.GetStringRepresentationOfType(this.Type);

        public RegulatoryDocument(string pathToFile, TypeOfDocument type = TypeOfDocument.НеОпределен) {
            // Проверка на то, что файл существует
            if (!File.Exists(pathToFile)) {
                string template =
                    "При инициализации объекта справочника 'Нормативные документы' возникла ошибка\n" +
                    "Исходный файл по пути '{0}' не был обнаружен";
                throw new Exception(string.Format(template, pathToFile));
            }

            this.LinkedFile = new FileInfo(pathToFile);
            this.Type = type;

            if (this.Type == TypeOfDocument.НеОпределен)
                this.Type = this.GetTypeOfDocumentFromFile();
            else
                this.CheckType();

            if ((this.Type == TypeOfDocument.НеОпределен) || (this.Type == TypeOfDocument.ОпределенСОшибкой))
                throw new Exception(string.Format("Не удалось однозначно определить тип документа по названию файла: {0}", this.LinkedFile.Name));

            FillFildsData(this.Type);
        }

        private TypeOfDocument GetTypeOfDocumentFromFile() {
            // TODO: Метод для определения типа документа, если он не был указан при создании нового объекта создания документа
            return TypeOfDocument.НеОпределен;
        }

        private void CheckType() {
            // TODO: Метод для проверки того, является ли полученный файл тем типом, который был задан при инициализации объекта
        }

        private string GetStringRepresentationOfType(TypeOfDocument type) {
            switch (type) {
                case TypeOfDocument.НеОпределен:
                    return "Неизвестно";
                case TypeOfDocument.ГОСТ:
                    return "ГОСТ";
                case TypeOfDocument.ОСТ:
                    return "ОСТ";
                case TypeOfDocument.СТО:
                    return "СТО";
                case TypeOfDocument.СТП:
                    return "СТП";
                case TypeOfDocument.ПИ:
                    return "ПИ";
                case TypeOfDocument.ТУ:
                    return "ТУ";
                case TypeOfDocument.Нормали:
                    return "Нормаль";
                case TypeOfDocument.Метрология:
                    return "Метрологический документ";
                case TypeOfDocument.ОпределенСОшибкой:
                    return "Неизвестно";
                default:
                    throw new Exception(string.Format("Переданный в функцию GetStringRepresentationOfType тип - {0} является неизвестным", type.ToString()));
            }
        }

        public override string ToString() {
            return string.Format("{0} {1} {2}", this.TypeString, this.Designation, this.Name);
        }
    }

    public enum TypeOfDocument {
        НеОпределен,
        ОпределенСОшибкой,
        ГОСТ,
        ОСТ,
        СТО,
        СТП,
        ПИ,
        ТУ,
        Нормали,
        Метрология
    }
}
