using System;
using System.IO;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using Newtonsoft.Json;

namespace LocalStorage {

    public class Storage {

        // Поля экземпляра класса
        private Dictionary<string, string> Fields { get; set; }
        public string Dir { get; private set; }
        public string Name { get; private set; }

        // Флаги
        public bool HasUnsavedChanges { get; private set; } = true;

        // Вычисляемые поля
        public string FileName => GetNameOfSettingsFile(this.Name);
        public string PathToFile => Path.Combine(this.Dir, this.FileName);
        public int Count => this.Fields.Count;

        // Статичные члены класса
        public static string DefaultDirectory { get; } = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Temp");
        private static Dictionary<string, Storage> Instances { get; set; } = new Dictionary<string, Storage>();

        // Конструкторы
        private Storage(string directory, string name) {
            if (name.Contains(".json") || name.Contains("LocalStorageFile-"))
                throw new Exception("Название хранилища не может содержать подстроку '.json' или 'LocalStorageFile-'");
            if (name.Any(x => Char.IsWhiteSpace(x)))
                throw new Exception("Название хранилища должно состоять из одного слова");
            this.Dir = directory;

            // Создаем директорию, если она отсутствовала
            if (!Directory.Exists(this.Dir))
                Directory.CreateDirectory(this.Dir);

            this.Name = name;
            this.Fields = new Dictionary<string, string>();

            this.Load();
        }

        public static Storage CreateInstance(string directory, string name) {
            if (!Instances.ContainsKey(name))
                Instances[name] = new Storage(directory, name);
            return Instances[name];
        }

        public static Storage CreateInstance(string name) =>
            CreateInstance(DefaultDirectory, name);

        public static Storage CreateInstance() =>
            CreateInstance("Unnamed");



        // Статические методы
        public static Dictionary<string, string> GetAllSettingsFiles(string pathToDirectory) {
            if (!Directory.Exists(pathToDirectory))
                throw new Exception($"Директория, расположенная по пути '{pathToDirectory}' не существует");
            DirectoryInfo dirInfo = new DirectoryInfo(pathToDirectory);

            Dictionary<string, string> result = new Dictionary<string, string>();
            foreach (FileInfo fileInfo in dirInfo.GetFiles("LocalStorageFile-*.json"))
                result.Add(
                        fileInfo.Name
                            .Replace("LocalStorageFile-", string.Empty)
                            .Replace(".json", string.Empty),
                        fileInfo.FullName
                        );

            return result;
        }

        public static Dictionary<string, string> GetAllSettingsFiles() =>
            GetAllSettingsFiles(DefaultDirectory);

        public static void RemoveSettingsFile(string pathToDirectory, string nameOfSettings) {
            Dictionary<string, string> allSettings = GetAllSettingsFiles(pathToDirectory);
            if (allSettings.ContainsKey(nameOfSettings)) {
                File.Delete(allSettings[nameOfSettings]);
                Instances.Remove(nameOfSettings);
            }
            else
                throw new Exception($"Файл хранилища с названием {nameOfSettings} отсутствует в системе");
        }

        public static void RemoveSettingsFile(string nameOfSettings) =>
            RemoveSettingsFile(DefaultDirectory, nameOfSettings);

        public static void RemoveSettingsFile() =>
            RemoveSettingsFile("Unnamed");

        public static void RemoveAllSettingsFiles(string pathToDirectory) {
            foreach (KeyValuePair<string, string> kvp in GetAllSettingsFiles(pathToDirectory))
                RemoveSettingsFile(kvp.Key);
        }

        public static void RemoveAllSettingsFiles() =>
            RemoveAllSettingsFiles(DefaultDirectory);

        public static string GetNameOfSettingsFile(string nameOfSettings) =>
            $"LocalStorageFile-{nameOfSettings}.json";


        // Методы
        public bool Load() {
            if (File.Exists(this.PathToFile)) {
                string jsonString = File.ReadAllText(this.PathToFile);
                this.Fields = JsonConvert.DeserializeObject<Dictionary<string, string>>(jsonString) ?? new Dictionary<string, string>();
                return true;
            }
            return false;
        }

        public void Save() {
            if (this.HasUnsavedChanges == true) {
                string jsonString = JsonConvert.SerializeObject(this.Fields);
                File.WriteAllText(this.PathToFile, jsonString);
                this.HasUnsavedChanges = false;
            }
            else
                throw new Exception($"При попытке произвести сохранение локального хранилища возникла ошибка.\nЛокальное хранилище {this.Name} не содержит изменений");
        }

        public T RetrieveOrDefault<T>(string key, T defaultValue) {

            if (defaultValue == null) {
                throw new NullReferenceException();
            }

            if (this.Fields.ContainsKey(key)) {
                string typeOfValue = typeof(T).ToString();
                string[] supportedTypes = new string[] {
                    "System.String",
                    "System.Int32",
                    "System.Int64",
                    "System.DateTime",
                    "System.Boolean",
                    "System.Single",
                    "System.Double",
                    "System.Decimal"
                };

                if (!supportedTypes.Contains(typeOfValue))
                    throw new Exception($"Метод {nameof(RetrieveOrDefault)} не поддерживает тип {typeOfValue} ");

                return (T)Convert.ChangeType(this.Fields[key], typeof(T));
            }
            else
                return defaultValue;
        }

        public string this[string key] {
            get {
                if (this.Fields.ContainsKey(key))
                    return this.Fields[key];
                else
                    throw new Exception($"Локальное хранилище '{this.Name}' не содержит значения для ключа '{key}'");
            }
            set {
                if (this.Fields.ContainsKey(key)) {
                    if (this.Fields[key] != value) {
                        this.Fields[key] = value;
                        this.HasUnsavedChanges = true;
                    }
                }
                else {
                    this.Fields[key] = value;
                    this.HasUnsavedChanges = true;
                }
            }
        }

        public bool ContainsKey(string key) => this.Fields.ContainsKey(key);

        public bool ContainsValue(string value) => this.Fields.ContainsValue(value);

        public void Clear() {
            this.Fields.Clear();
            this.HasUnsavedChanges = true;
            this.Save();
        }
    }


}
