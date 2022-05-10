using Xunit;
using System;
using System.Collections.Generic;
using System.Linq;
using LocalStorage;

namespace LocalStorageTests {

    public class UnitTest1 {
        [Fact]
        public void CreatingSettingsTest() {
            Storage settings = Storage.CreateInstance();
            // Производим очистку словаря перед тестами
            settings.Clear();

            // Тестируем основные параметры исходного нового словаря
            Assert.Equal("Unnamed", settings.Name);
            Assert.Equal("LocalStorageFile-Unnamed.json", settings.FileName);
            Assert.Equal(0, settings.Count);

            // Тестируем создание нового поля
            settings["test"] = "test";
            settings.Save();
            Assert.Equal(1, settings.Count);

            // Инициализируем вторую копию
            Storage secondInstanceOfSettings = Storage.CreateInstance();
            Assert.Equal(1, settings.Count);
            Assert.Equal("test", secondInstanceOfSettings["test"]);

            settings.Clear();
            Assert.Equal(0, settings.Count);
            
            Storage.RemoveSettingsFile();
        }

        [Fact]
        public void TestSingleton() {
            // Данный тест предназначен для проверки корректности работы паттерна Singleton
            string nameOfSettings = "TestSingleton";
            Storage testSettings1 = Storage.CreateInstance(nameOfSettings);
            Storage testSettings2 = Storage.CreateInstance(nameOfSettings);
            Assert.True(testSettings2.HasUnsavedChanges);
            testSettings1.Save();

            // Проверка на то, что объекты одинаковые
            Assert.Equal(testSettings1, testSettings2);

            // Проверка на то, что статут изменений у объектов одинаковый
            Assert.False(testSettings2.HasUnsavedChanges);
            testSettings1["test"] = "value";
            Assert.Equal("value", testSettings2["test"]);
            Storage.RemoveSettingsFile(nameOfSettings);
        }

        [Fact]
        public void TestSaveAndRetrieve() {
            string nameOfSettings = "TestSaveAndRetrieve";
            Storage testSettings = Storage.CreateInstance(nameOfSettings);

            // проверка извлечения текстового значения
            testSettings["string"] = "value";
            Assert.Equal("value", testSettings.RetrieveOrDefault<string>("string", "default"));
            Assert.Equal("default", testSettings.RetrieveOrDefault<string>("dont existed value", "default"));

            // проверка извлечения даты
            DateTime now = DateTime.Now;
            DateTime today = DateTime.Today;
            testSettings["date"] = now.ToString();
            Assert.Equal(now.ToString(), testSettings.RetrieveOrDefault<DateTime>("date", today).ToString());
            Assert.Equal(today.ToString(), testSettings.RetrieveOrDefault<DateTime>("dont existed date", today).ToString());
            testSettings.Save();
            Storage.RemoveSettingsFile(nameOfSettings);
        }

        [Fact]
        public void TestStaticMethods() {
            // Создаем несколько тестовых настроек
            Storage testSettings1 = Storage.CreateInstance("Test1");
            testSettings1.Save();
            Storage testSettings2 = Storage.CreateInstance("Test2");
            testSettings2.Save();
            Storage testSettings3 = Storage.CreateInstance("Test3");
            testSettings3.Save();
            // Проверяем количество найденных настроек
            Assert.Equal(3, Storage.GetAllSettingsFiles().Count);
            // Удаляем одну настройку по имени
            Storage.RemoveSettingsFile("Test3");
            Assert.Equal(2, Storage.GetAllSettingsFiles().Count);
            // Удаляем все оставшиеся настройки
            Storage.RemoveAllSettingsFiles();
            Assert.Equal(0, Storage.GetAllSettingsFiles().Count);
        }

        [Fact]
        public void TestNamingOfSettingsObjects() {

            // Проверка на отсутствие недопустимого текста в названии объекта настроек
            Assert.Throws<Exception>(() => Storage.CreateInstance(".json"));
            Assert.Throws<Exception>(() => Storage.CreateInstance("LocalStorageFile-"));
            Assert.Throws<Exception>(() => Storage.CreateInstance("test.json"));
            Assert.Throws<Exception>(() => Storage.CreateInstance("LocalStorageFile-test"));
            Assert.Throws<Exception>(() => Storage.CreateInstance("LocalStorageFile-test.json"));

            // Проверка на то, что название объекта настроек состоит из одного слова
            Assert.Throws<Exception>(() => Storage.CreateInstance("test 1"));
            Assert.Throws<Exception>(() => Storage.CreateInstance("test  1"));
        }

        [Fact]
        public void TestHasUnsavedChangesProperty() {
            string nameOfSettings = "TestHasUnsavedChangesProperty";
            Storage testSettings = Storage.CreateInstance(nameOfSettings);
            Assert.True(testSettings.HasUnsavedChanges);

            testSettings.Save();
            Assert.False(testSettings.HasUnsavedChanges);

            testSettings["test"] = "value";
            Assert.True(testSettings.HasUnsavedChanges);

            testSettings.Save();
            testSettings["test"] = "value";
            Assert.False(testSettings.HasUnsavedChanges);
            Assert.Throws<Exception>(() => testSettings.Save());

            Storage.RemoveSettingsFile(nameOfSettings);
        }
    }
}
