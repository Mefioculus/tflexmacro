using System.Linq;
using TFlex.DOCs.Common;
using TFlex.DOCs.Model.Macros;
using TFlex.DOCs.Model.References;
using TFlex.DOCs.Model.References.Documents;
using TFlex.DOCs.Model.References.Files;
using TFlex.DOCs.Model.References.Links;
using TFlex.DOCs.Model.References.Nomenclature;

namespace PDM
{
    public class Macro : MacroProvider
    {
        private const string separator = "^";

        public Macro(MacroContext context)
            : base(context)
        {
        }

        public void СозданиеРевизии()
        {
            UpdateLinkedDocumentFiles();
        }

        private void UpdateLinkedDocumentFiles()
        {
            if (!(Context["SourceObject"] is NomenclatureObject sourceObject))
                return;

            if (!((ReferenceObject)CurrentObject is NomenclatureObject nomenclatureRevision) || !nomenclatureRevision.HasLinkedObject)
                return;

            var linkedObject = nomenclatureRevision.LinkedObject;
            if (linkedObject is null)
                return;

            var linkToFiles = linkedObject.Links.ToMany
                .FirstOrDefault(link => link?.LinkGroup?.SlaveGroup?.Guid == SystemParameterGroups.Files) as OneToManyLink;
            if (linkToFiles is null)
                return;

            var sourceLinkToFiles = sourceObject.LinkedObject?.Links.ToMany
                .FirstOrDefault(link => link?.LinkGroup?.SlaveGroup?.Guid == SystemParameterGroups.Files) as OneToManyLink;
            if (sourceLinkToFiles is null)
                return;

            TFlex.DOCs.Model.References.Files.FolderObject defaultFolder = null;

            var fileLinkSettings = linkToFiles != null ? new FileLinkAdditionalSettings(linkToFiles.LinkGroup, linkedObject.Class) : null;
            if (fileLinkSettings != null && fileLinkSettings.Data != null)
                defaultFolder = fileLinkSettings.Data.GetDefaultFolder(linkToFiles);

            if (defaultFolder is null)
            {
                if (linkedObject is EngineeringDocumentObject)
                    defaultFolder = RunMacro("b764aa86-1986-4df4-8dcb-b05b7a701f26", "ПолучитьПапкуДляХраненияФайловДокумента");
            }

            var newRevisionObject = Context.ReferenceObject as NomenclatureObject;

            string denotation = newRevisionObject.Denotation;
            string newfileName = $"{denotation}";

            if (newRevisionObject.Reference.ParameterGroup.SupportsRevisions)
                newfileName = $"{newfileName}{separator}{newRevisionObject.SystemFields.RevisionName}";

            var revisionParentDocument = (nomenclatureRevision.Parent as NomenclatureObject)?.LinkedObject as EngineeringDocumentObject;

            var sourceFiles = sourceLinkToFiles.OfType<FileObject>().ToList();

            foreach (var sourceFile in sourceFiles)
            {
                linkToFiles.RemoveLinkedObject(sourceFile);

                if (sourceFile.Path.GetString().ToLower().Contains(@"архив\подлинники"))
                    continue;

                //%%TODO
                var assemblyFile = revisionParentDocument?
                    .GetFiles()?
                    .FirstOrDefault(f => f.Id == sourceFile.Id);

                if (assemblyFile is null)
                {
                    FileObject fileCopy = null;
                    try
                    {
                        // Создаем копию файла с новым наименованием

                        var parentFolder = defaultFolder ?? sourceFile.Parent;
                        string uniqueFileName = GetUniqueFileName(newfileName, sourceFile.Class.Extension, parentFolder);

                        fileCopy = sourceFile.CreateCopy(uniqueFileName, parentFolder, null);
                        fileCopy.EndChanges();

                        if (fileCopy != null)
                            linkToFiles.AddLinkedObject(fileCopy);
                    }
                    catch
                    {
                        if (fileCopy != null && fileCopy.Changing)
                            fileCopy.CancelChanges();
                    }
                }
                else
                    linkToFiles.AddLinkedObject(assemblyFile);
            }
        }

        private string GetUniqueFileName(string fileName, string extension, TFlex.DOCs.Model.References.Files.FolderObject parentFolder)
        {
            parentFolder.Children.Load();

            var filesNameSet = parentFolder.Children.AsList
                .Where(child => child.IsFile)
                .Select(file => file.Name.Value)
                .ToList();

            string uniqueFileName = $"{fileName}.{extension}";
            int counter = 1;

            while (filesNameSet.Contains(uniqueFileName))
                uniqueFileName = $"{fileName}_{counter++}.{extension}";

            return uniqueFileName;
        }
    }
}

