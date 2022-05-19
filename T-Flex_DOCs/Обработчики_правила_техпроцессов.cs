/*
TFlex.DOCs.SynchronizerReference.dll
*/

using System;
using TFlex.DOCs.Model.Macros;
using TFlex.DOCs.Model.References;
using TFlex.DOCs.Model.Search.Path;
using TFlex.DOCs.Synchronization.Macros;
using TFlex.DOCs.Synchronization.SyncData;
using TFlex.DOCs.Synchronization.SyncData.Behaviors;
using TFlex.DOCs.Model.References.Links.Extensions;

public class MacroAEM_TP : MacroProvider
{
    public MacroAEM_TP(MacroContext context) : base(context) { }

    #region Создание техпроцессов из Marchp

    public void НачалоТехпроцессы()
    {
        var connection = Context.Connection;
        var context = Context as ExchangeDataMacroContext;
        var behaviors = context.Settings.Behaviors;

        // Технологические процессы / SHIFR <-> Обозначение
        behaviors.Add(new ReplaceParameterValueBatchBehavior("75c1c70c-e831-4758-b936-9253b7b02914", "eb7c439b-8c6f-4a1e-8b10-d31d7ee9ec2b",
            ReferencePath.Parse("[853d0f07-9632-42dd-bc7a-d91eae4b8e83].[ae35e329-15b4-4281-ad2b-0e8659ad2bfb]", connection), // Номенклатура -> Обозначение
            ReferencePath.Parse("[853d0f07-9632-42dd-bc7a-d91eae4b8e83].[45e0d244-55f3-4091-869c-fcf0bb643765]", connection))); // Номенклатура -> Наименование

        behaviors.Add(new ProcessingObjectBeforeSavingBehavior("75c1c70c-e831-4758-b936-9253b7b02914", TechnologicalProcessSaving));
        behaviors.Add(new InputObjectsDistinctBehavior(context.Connection, "MARCHP", "SHIFR", "NORM")); // Distinct объектов по параметру SHIFR NORM
    }

    private void TechnologicalProcessSaving(SyncDataObject dataObject)
    {
        var tp = dataObject.ReceiverObject as ReferenceObject;
        ВыполнитьМакрос("Пересчет Тшт в изготовлении", "AddMaterialsToTechnologicalProcess", tp);
    }

    public void УстановитьКаталогОснащения()
    {
        var context = Context as DataObjectProcessingMacroContext;
        var dataObject = context.DataObject;

        var sourceObject = dataObject.SourceObject as ReferenceObject;
        var recieverObject = dataObject.ReceiverObject as ReferenceObject;

        if (sourceObject == null || recieverObject == null)
            return;

        //string potr = sourceObject[new Guid("ec64608f-b92d-427c-b277-34d30a41b265")].GetString(); // POTR
        string izg = sourceObject[new Guid("a4f4dc8e-78d2-4524-9528-14886629e2f0")].GetString(); // IZG
        string shifr = sourceObject[new Guid("45275ca2-eb5a-47aa-afb6-2f72b3ff3b9a")].GetString(); // SHIFR
        string per1 = sourceObject[new Guid("6c833d92-3d9c-493e-a3b1-7b48c6e45730")].GetString(); // PER1

        //string potr = dataObject.FindParameterByExternalKey("POTR").Value as string;
        //string shifr = dataObject.FindParameterByExternalKey("SHIFR").Value as string;

        recieverObject.RemoveAllStorageLinkedObjects(new Guid("c7af468f-95dd-4835-a562-c0f96e170e4e")); //Каталог оснащения	Связь (список со списком)

        if ((String.IsNullOrEmpty(izg) && String.IsNullOrEmpty(shifr)) || per1 != "1")
            return;

        var оснащения = НайтиОбъекты("Каталог оснащения",
            String.Format("[Обозначение детали] = '{0}' И [Подразделение] = '{1}'", shifr, izg));

        if (оснащения.Count > 0)
            recieverObject.AddStorageLinkedObjects(new Guid("c7af468f-95dd-4835-a562-c0f96e170e4e"), оснащения.To<ReferenceObject>()); //Каталог оснащения	Связь (список со списком)
    }

    #endregion
    #region Создание цехопереходов из Marchp

    public void НачалоЦехопереходы()
    {
        var connection = Context.Connection;
        var context = Context as ExchangeDataMacroContext;
        var behaviors = context.Settings.Behaviors;

        // Технологические процессы / Наименование подразделения <-> Подразделение
        behaviors.Add(new ReplaceParameterValueBatchBehavior("bb82fe1f-7861-4a8b-89f3-07872df0f347", "c25ef93a-fcdc-44d9-9864-8064b7ac77de",
            ReferencePath.Parse("[8ee861f3-a434-4c24-b969-0896730b93ea].[1ff481a8-2d7f-4f41-a441-76e83728e420]", connection), // Группы и пользователи -> Номер
            ReferencePath.Parse("[8ee861f3-a434-4c24-b969-0896730b93ea].[beb2d7d1-07ef-40aa-b45b-23ef3d72e5aa]", connection))); // Группы и пользователи -> Наименование

        // Технологические процессы / Наименование <-> Наименование
        behaviors.Add(new ReplaceParameterValueBatchBehavior("bb82fe1f-7861-4a8b-89f3-07872df0f347", "c6da004c-ed9f-4cea-af58-882e19ef50c6",
            ReferencePath.Parse("[8ee861f3-a434-4c24-b969-0896730b93ea].[1ff481a8-2d7f-4f41-a441-76e83728e420]", connection), // Группы и пользователи -> Номер
            ReferencePath.Parse("[8ee861f3-a434-4c24-b969-0896730b93ea].[beb2d7d1-07ef-40aa-b45b-23ef3d72e5aa]", connection))); // Группы и пользователи -> Наименование
    }

    #endregion
}
