using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using TFlex.DOCs.Model.FilePreview.CADExchange;
using TFlex.DOCs.Model.Macros;
using TFlex.DOCs.Model.References;
using TFlex.DOCs.Model.References.Files;
using TFlex.DOCs.Model.Structure;
using TFlex.Model.Technology.References.SetOfDocuments;
using TFlex.DOCs.Model.Search;
using TFlex.Model.Technology.References.TechnologyElements;
using TFlex.Technology.References;

public class SetOfDocumentsMacros : MacroProvider
{
    private static readonly string[] Extensions = new string[] { ".grb", ".grn", ".grs" };
    private static readonly Guid CadDrawingClassId = new Guid("e440761f-2d0f-465e-a2b7-641427901c9a"); //чертёж системы 2D

    public SetOfDocumentsMacros(MacroContext context)
        : base(context)
    {
    }

    public override void Run()
    {
        if (Context.ReferenceObject is TechnologicalDocument technologicalDocument)
        {
            GenerateSingleTechnologicalDocument(technologicalDocument);
        }
    }

    public void OnCreateSetOfDocuments()
    {
        if (!Context.Reference.IsSlave)
            return;

        var newDocumentSetObject = Context.ReferenceObject as SetOfDocumentsReferenceObject;

        if (newDocumentSetObject.MasterObject.Reference is not TechnologicalProcessReference)
            return;

        var techProcess = newDocumentSetObject.MasterObject as StructuredTechnologicalProcess;

        if (techProcess is null)
            Error("Требуется наличие тех. процесса!");
        
        newDocumentSetObject.Name.Value = techProcess.Name;
        newDocumentSetObject.Denotation.Value = techProcess.Denotation;
    }

    public void MergeDocuments()
    {
        if (!(Context.ReferenceObject is TechnologicalSet setOfDocuments))
            return;

        FolderObject folder = setOfDocuments.Folder;
        if (folder == null)
        {
            Context.ShowMessage("Внимание!", "Не указана папка комплекта.");
            return;
        }

        setOfDocuments.Children.Reload();
        var documents = setOfDocuments.Children
            .OfType<TechnologicalDocument>()
            .ToList();

        if (!documents.Any())
        {
            Context.ShowMessage("Внимание!", "Комплект не содержит документов с файлами.");
            return;
        }

        string outputPath = String.Empty;
        var file = setOfDocuments.File as FileObject;
        if (file != null)
        {
            file.GetHeadRevision();
            string ext = Path.GetExtension(file.LocalPath);
            if (Extensions.Contains(ext))
            {
                if (!Context.ShowQuestion("Комплект документов уже содержит файл. Заменить?"))
                    return;

                outputPath = file.LocalPath;
                if (!file.IsCheckedOutByCurrentUser)
                {
                    file.CheckOut(false);
                }
            }
            else
            {
                file = null;
            }
        }

        if (String.IsNullOrEmpty(outputPath))
            outputPath = CreateUniqueFileName(folder, setOfDocuments.ToString());

        var files = new List<string>();
        foreach (FileObject fileObject in documents.Select(document => document.FileObject).OfType<FileObject>())
        {
            fileObject.GetHeadRevision();
            if (Extensions.Contains(Path.GetExtension(fileObject.LocalPath)))
                files.Add(fileObject.LocalPath);
        }

        if (files.Count == 0)
            return;

        bool isEmbedded = !setOfDocuments.IsPageCountManually.GetBoolean();
        var provider = new CombineFilesProvider(files, outputPath) { IsEmbedded = isEmbedded };
        provider.Execute(setOfDocuments.Reference.Connection);

        bool isNewFile = file == null;
        if (isNewFile)
            file = folder.Reference.AddFile(outputPath, folder);

        if (file == null) 
            return;

        if (isNewFile)
        {
            setOfDocuments.Modify(t => t.File = file);
        }
    }

    public void GenerateSingleTechnologicalDocument(TechnologicalDocument document)
    {
        var parentSet = document.Parent as TechnologicalSet;
        if (parentSet == null)
            throw new InvalidOperationException($"Документ {document} не входит в технологический комплект.");

        LoadSettings.SortField sortField = parentSet.Reference.LoadSettings.AddSortField(
            parentSet.Reference.ParameterGroup.SystemParameters.Find(SystemParameterType.Order),
            SortOrder.Ascending);

        try
        {
            parentSet.Children.Reload();
            var documents = parentSet.Children
                .OfType<TechnologicalDocument>()
                .ToList();

            TechnologicalDocument prevDoc = null;
            int currentIdx = documents.IndexOf(document);
            if (currentIdx > 0)
                prevDoc = documents[currentIdx - 1] as TechnologicalDocument;

            int oldPageCount = document.PageCount;
            if (prevDoc != null)
            {
                document.Generate(prevDoc.FirstPageNumber + prevDoc.PageCount);
            }
            else
            {
                document.Generate(1);
            }

            if (oldPageCount == document.PageCount ||
                !Context.ShowQuestion(
                    "Количество страниц в документе изменилось, обновить общее количество страниц комплекта?")) 
                return;

            int firstPage = document.PageCount + document.FirstPageNumber;

            foreach (TechnologicalDocument doc in documents.Skip(currentIdx + 1))
            {
                int page = firstPage;
                doc.Modify(d => d.FirstPageNumber.Value = page);
                firstPage += doc.PageCount;
            }

            ((TechnologicalSet)document.Parent).Modify(s => s.PageCount.Value = firstPage - 1);
        }
        finally
        {
            parentSet.Reference.LoadSettings.RemoveSortField(sortField);
        }
    }

    public void GenerateSetOfDocuments()
    {
        foreach (var technologicalSet in  Context.GetSelectedObjects().OfType<TechnologicalSet>())
        {
            try
            {
                technologicalSet.Generate();
            }
            catch (Exception e)
            {
                Context.ShowMessage("Внимание!", e.Message);
            }
        }
    }

    public void UpdateTechnologicalSets()
    {
        foreach (var technologicalSet in  Context.GetSelectedObjects().OfType<TechnologicalSet>())
        {
            UpdateCurrentTechnologicalSet(technologicalSet);
        }
    }

    private void UpdateCurrentTechnologicalSet(TechnologicalSet technologicalSet)
    {
        LoadSettings.SortField sortField = technologicalSet.Reference.LoadSettings.AddSortField(
            technologicalSet.Reference.ParameterGroup.SystemParameters.Find(SystemParameterType.Order),
            SortOrder.Ascending);

        try
        {
            technologicalSet.Children.Reload();
            int firstPageNumber = 1;
            foreach (var techDoc in technologicalSet.Children.OfType<TechnologicalDocument>())
            {
                int page = firstPageNumber;
                if (techDoc.FirstPageNumber != firstPageNumber)
                    techDoc.Modify((doc) => doc.FirstPageNumber.Value = page);

                firstPageNumber += techDoc.PageCount;
            }
            technologicalSet.Modify(s => s.PageCount.Value = firstPageNumber - 1);
        }
        finally
        {
            technologicalSet.Reference.LoadSettings.RemoveSortField(sortField);
        }
    }

    private string CreateUniqueFileName(FolderObject folder, string defaultFileName)
    {
        var reference = folder.Reference;
        var localPath = folder.LocalPath;

        var classObject = reference.Classes.TFlexCADFileBase.ChildClasses.Find(CadDrawingClassId) as FileType ?? 
                          reference.Classes.TFlexCADFileBase.ChildClasses
                              .OfType<FileType>()
                              .FirstOrDefault(cl => !cl.Hidden && !cl.IsAbstract && cl.Extension.ToLower() == "grb");

        int iteration = 0;
        string fullName;
        Dictionary<Guid, object> dictionary;
        string nowTime = $"[{DateTime.Now.ToString("yyyy.MM.dd HH.mm")}]";
        Guid parentParameterGuid = reference.ParameterGroup.SystemParameters.Find(SystemParameterType.Parent).Guid;
        do
        {
            fullName = String.Concat(
                localPath, 
                Path.DirectorySeparatorChar,
                defaultFileName,
                iteration == 0 ? String.Empty : nowTime,
                iteration < 2 ? String.Empty : " (" + (iteration - 1) + ")",
                ".grb");

            dictionary = new Dictionary<Guid,object>()
            {
                { parentParameterGuid, folder },
                { FileReferenceObject.FieldKeys.Name, Path.GetFileName(fullName) },
            };

            iteration++;
        }
        while (!reference.CanCreateUniqueObject(classObject, dictionary));

        return fullName;
    }
}
