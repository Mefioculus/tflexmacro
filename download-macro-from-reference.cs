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


    #region Properties

    private string statusFileName = "status.info";
    private string configDirectoryName = ".config";
    private bool firstExport = true;
    private string pathToExportDirectory = string.Empty();

    #endregion Properties

    #region Entry point

    public override void Run() {
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

    #endregion Service methods

    #region Service classes

    private class StatusExport {
        // TODO Реализовать класс для хранения данных об экспорте
    }

    #endregion Service classes
    
}
