using System;
using System.Collections.Generic;
using System.Linq;
using TFlex.DOCs.Common;
using TFlex.DOCs.Model.Macros;
using TFlex.DOCs.Model.Processes.Events.Contexts;
using TFlex.DOCs.Model.Processes.Events.Contexts.Data;
using TFlex.DOCs.Model.References;
using TFlex.DOCs.Model.References.ActiveActions;
using TFlex.DOCs.Model.References.Documents;
using TFlex.DOCs.Model.References.Nomenclature;

public class PDM_CreateRealFiles : MacroProvider
{
    private static readonly string _stageName = "Утверждение";

    #region Guids

    private static class Guids
    {
        public static readonly Guid ИзвещениеОбИзменении = new Guid("52ccb35c-67c5-4b82-af4f-e8ceac4e8d02");
        public static readonly Guid СвязьИзвещенияИзменения = new Guid("5e46670a-400c-4e36-bb37-d4d651bdf692");
        public static readonly Guid СвязьИзмененияРабочиеФайлы = new Guid("6b65a575-3ca4-4fb0-9bfc-4d1655c2d83e");
    }

    #endregion

    public PDM_CreateRealFiles(MacroContext context)
    : base(context)
    {
    }

    public void СоздатьПодлинники()
    {
        CreateRealFiles(Context.GetSelectedObjects().ToList(), _stageName);
    }

    public void СоздатьПодлинникиИзБП()
    {
        //Контекст событий по БП
        var eventContext = Context as EventContext;
        if (eventContext == null)
            return;

        var data = eventContext.Data as StateContextData;
        // Текущее действие
        var activeAction = data.ActiveAction;
        // Данные текущего действия
        ActiveActionData activeActionData = activeAction.GetData<ActiveActionData>();
        // Объекты, подключенные к БП
        List<ReferenceObject> processObjects = activeActionData.GetReferenceObjects().ToList();

        CreateRealFiles(processObjects, _stageName);
    }

    /// <summary>
    /// Сформировать подлинники
    /// </summary>
    /// <param name="referenceObjects">Объекты, которые содержат файлы по связи</param>
    /// <param name="stageName">Наименование стадии, в которой должны быть объекты и связанные с ними файлы</param>
    private void CreateRealFiles(List<ReferenceObject> referenceObjects, string stageName)
    {
        if (referenceObjects.IsNullOrEmpty())
            return;

        List<ReferenceObject> sourceFiles = new List<ReferenceObject>();
        foreach (ReferenceObject referenceObject in referenceObjects)
        {
            CheckStage(referenceObject, stageName);

            var nomenclatureObject = referenceObject as NomenclatureObject;
            if (nomenclatureObject != null)
            {
                var document = nomenclatureObject.LinkedObject as EngineeringDocumentObject;
                if (document != null)
                {
                    var files = document
                         .GetFiles()
                         .Where(file => CheckStage(file, stageName) == true);
                    if (files.IsNullOrEmpty())
                        continue;

                    sourceFiles.AddRange(files);
                }
            }
            else
            {
                if (referenceObject.Class.IsInherit(Guids.ИзвещениеОбИзменении))
                {
                    var modifications = referenceObject.GetObjects(Guids.СвязьИзвещенияИзменения).ToList();
                    foreach (var modification in modifications)
                    {
                        var modificationFiles = modification
                            .GetObjects(Guids.СвязьИзмененияРабочиеФайлы)
                            .Where(file => CheckStage(file, stageName) == true);
                        if (modificationFiles.IsNullOrEmpty())
                            continue;

                        sourceFiles.AddRange(modificationFiles);
                    }
                }
            }
        }

        if (sourceFiles.Any())
            RunMacro("458858a7-6acf-4727-b581-1e2699aeab28", "ExportToFormat", sourceFiles);
    }

    private bool CheckStage(ReferenceObject referenceObject, string stageName)
    {
        if (referenceObject == null)
            throw new ArgumentNullException("referenceObject");

        if (String.IsNullOrEmpty(stageName))
            return true;

        var schemeStage = referenceObject.SystemFields.Stage;
        if (schemeStage != null)
        {
            var stage = schemeStage.Stage;
            if (stage != null)
                return stage.Name == stageName;
        }

        return false;
    }
}

