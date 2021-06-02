using System;
using System.IO;
using System.Collections;
using System.Collections.Generic;
using System.Windows.Forms;
// Пространства имен DOCs
using TFlex.DOCs.Model;
using TFlex.DOCs.Model.Macros;
using TFlex.DOCs.Model.References;
// Дополнительные пространства имен
using Newtonsoft.Json;

public class Macro : MacroProvider {
    public Macro(MacroContext context)
        : base(context)
    {
    }

    // Данный макрос предназначен для выгрузки всех макросов в определенную директорию

    #region Guids

    private static class Guids {
        private static class References {
            private static Guid MacroReference = new Guid("3e6df4d0-b1d8-4375-978c-4da676604cca");
        }

        private static class Properties {
            private static Guid Name = new Guid("8334853d-8b04-4716-b3e2-19bc0a360384");
            private static Guid Code = new Guid("3d654359-8567-49b3-8060-516f5f2f2ad2");
        }
    }

    #endregion Guids

    #region Properties

    private string statusFileName = "status.info";
    private string configDirectoryName = ".config";
    private bool firstExport = true;
    private string pathToExportDirectory = string.Empty();

    #endregion Properties

    #region Entry point

    public override void Run() {
        // Получаем список всех макросов, которые есть в справочнике
        List<MacrosObject> listOfMacros = GetListOfMacro();

        GetFolderForExportMacro();
        Check();
        ExportMacro();
    }

    #endregion Entry point

    #region Service methods

    private string GetFolderForExportMacro() {
        // TODO Реализовать диалог выбора директории для сохранения макросов
        return string.Empty;
    }

    private void Check() {
        // TODO Реализовать метод, который будет проводить все предварительные проверки с целью определения текущего
        // состояния экспорта

    }

    private void CheckEarlyExportedFiles() {
        // TODO Реализовать проверку ранее проведенного экспорта с целью подтверждения соответствия файлов данным
        // о времени проведения экспорта.
    }

    private void ExportMacro() {
        // TODO Реализовать метод экспортирования макросов

    }

    private List<MacrosObject> GetListOfMacro() {
        // TODO Реализовать метод, который получает перечень всех макросов
        List<MacrosObject> listOfMacros = new List<MacrosObject>();
        string template = "Количество найденных макросов в справочнике - {0} шт.";
        string message = string.Empty;

        Reference macroReference = Context.Connection.ReferenceCatalog.Find(Guids.References.MacroReference).CreateReference();

        foreach (ReferenceObject refObj in macroReference.Objects) {
            MacrosObject macros = new MacrosObject();
            
            // Создание объекта, который будет хранить все основные параметры макроса
            macros.GuidOfMacro = refObj.SystemFields.Guid;
            macros.Name = refObj[Guids.Properties.Name];
            macros.Code = refObj[Guids.Properties.Code];
            macros.DateLastModification = refObj.SystemFields.EditDate;

            listOfMacros.Add(macros);
            message += string.Format("{0}\n", macros.Name);
        }

        Message("Информация", string.Format(template, listOfMacros.Count));
        Message("Список макросов", message);

        return listOfMacros;
    }

    #endregion Service methods

    #region Service classes

    private class StatusExport {
        // TODO Реализовать класс для хранения данных об экспорте
    }

    private class MacrosObject {
        // TODO Реализовать класс, которырый будет хранить в себе данные о макросе
        public Guid GuidOfMacro { get; set; }
        public string Name { get; set; }
        public string Code { get; set; }
        public DateTime DateLastModification { get; set; }
    }

    #endregion Service classes
    
}
