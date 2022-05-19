using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using DevExpress.XtraRichEdit;
using TFlex.DOCs.Client.Utils;
using TFlex.DOCs.Client.ViewModels.References.Modifications;
using TFlex.DOCs.Common;
using TFlex.DOCs.Model.Macros;
using TFlex.DOCs.Model.References;
using TFlex.DOCs.Model.References.Macros;
using TFlex.DOCs.Model.References.Modifications;
using TFlex.DOCs.Model.References.Nomenclature;

namespace PDM
{
    public class PDM___Формирование_текстов_изменений : MacroProvider
    {
        private static readonly ApplicabilityComparer _applicabilityComparer = new ApplicabilityComparer();
        private const string _changedPlacementText = "изменить положение объекта в сборке";

        public PDM___Формирование_текстов_изменений(MacroContext context) : base(context)
        {
        }

        public void CreateActionTextCore(List<NomenclatureHierarchyLink> checkedLinks)
        {
            var macro = Context[nameof(MacroKeys.MacroSource)] as CodeMacro;
            if (macro.DebugMode.Value == true)
            {
                Debugger.Launch();
                Debugger.Break();
            }

            string text = String.Empty;
            var action = Context.ReferenceObject as ModificationActionReferenceObject;

            if (!checkedLinks.IsNullOrEmpty())
            {
                if (action.Class.IsAddAction)
                {
                    text = $"Добавить {String.Join(", ", checkedLinks.Select(link => $"'{link.ChildObject}'").Distinct())}";
                }
                else if (action.Class.IsEditAction)
                {
                    var content = new StringBuilder();
                    content.AppendLine("Изменить:");

                    var reference = new NomenclatureReference(Context.Connection);
                    var newConfigurationSettings = new ConfigurationSettings(reference.ConfigurationSettings)
                    {
                        ApplyDesignContext = false,
                    };

                    using (reference.ChangeAndHoldConfigurationSettings(newConfigurationSettings))
                    {
                        foreach (var editedLink in checkedLinks.Where(l => l.SubstitutedLinkId > 0))
                        {
                            int sourceLinkId = editedLink.SubstitutedLinkId;
                            var sourceObject = reference.Find(editedLink.ChildObjectId);

                            var sourceLink = sourceObject.Parents.GetHierarchyLinks().OfType<NomenclatureHierarchyLink>()
                                .FirstOrDefault(lnk => lnk.Id == sourceLinkId);

                            string result = CompareLinks(content, sourceLink, editedLink);
                            if (!result.IsNullOrEmpty())
                                content.AppendLine($"Для '{editedLink.ChildObject}' {result}");
                        }
                    }

                    text = content.ToString();
                }
                else if (action.Class.IsDeleteAction)
                {
                    text = $"Удалить {String.Join(", ", checkedLinks.Select(link => $"'{link.ChildObject}'").Distinct())}";
                }
                else if (action.Class.IsReplaceAction)
                {
                    var deletedLinks = checkedLinks.Where(link => link.DeletedInDesignContext);
                    var addedLinks = checkedLinks.Except(deletedLinks).Where(link => link.SubstitutedLinkId == 0);

                    text = $"Заменить {String.Join(", ", deletedLinks.Select(link => $"'{link.ChildObject}'").Distinct())} на " +
                           $"{String.Join(", ", addedLinks.Select(link => $"'{link.ChildObject}'").Distinct())}";
                }
            }

            string formattedText = RtfHtmlUtilites.ConvertPlainToRtf(text.TrimEnd());
            action.Content.Value = formattedText;

            RefreshControls("ActionContent");
        }

        public void CreateActionTextForPdmObjectCore(NomenclatureReferenceObject referenceObject,
            List<NomenclatureHierarchyLink> sourceLinks, List<NomenclatureHierarchyLink> newLinks)
        {
            var macro = Context[nameof(MacroKeys.MacroSource)] as CodeMacro;
            if (macro.DebugMode.Value == true)
            {
                Debugger.Launch();
                Debugger.Break();
            }

            var action = Context.ReferenceObject as ModificationActionReferenceObject;
            if (!action.Class.IsEditObjectAction)
                return;

            var content = new StringBuilder();
            content.AppendLine($"Изменения объекта ЭСИ '{referenceObject}':");

            if (!sourceLinks.IsNullOrEmpty())
            {
                var sourceChild = sourceLinks.FirstOrDefault()?.ChildObject;
                var newChild = newLinks.FirstOrDefault()?.ChildObject;
                if (sourceChild != null && newChild != null)
                {
                    foreach (var sourceParameter in sourceChild.ParameterValues)
                    {
                        var parameterInfo = sourceParameter.ParameterInfo;
                        if (parameterInfo.IsSystem)
                            continue;

                        var newParameter = newChild.ParameterValues.GetParameter(parameterInfo);

                        object sourceValue = sourceParameter.GetValue(parameterInfo.Type.LanguageType);
                        object newValue = newParameter.GetValue(parameterInfo.Type.LanguageType);

                        if (!Equals(newValue, sourceValue))
                        {
                            string parameterText = String.Format($"изменён параметр '{parameterInfo}' с '{sourceValue ?? String.Empty}'" +
                                  $" на '{newValue ?? String.Empty }'");

                            content.AppendLine(parameterText);
                        }
                    }
                }

                foreach (var sourceLink in sourceLinks)
                {
                    int index = sourceLinks.IndexOf(sourceLink);
                    var newLink = newLinks.ElementAtOrDefault(index);

                    string result = CompareLinks(content, sourceLink, newLink);
                    if (!result.IsNullOrEmpty())
                        content.AppendLine(result);
                }
            }

            string text = content.ToString();

            string formattedText = RtfHtmlUtilites.ConvertPlainToRtf(text.TrimEnd());
            action.Content.Value = formattedText;
        }

        private string CompareLinks(StringBuilder content, NomenclatureHierarchyLink sourceLink, NomenclatureHierarchyLink newLink)
        {
            if (sourceLink is null || newLink is null)
                return String.Empty;

            var sourceIntervals = Context.Connection.References.ProductsApplicability.GetIntervals(sourceLink.Guid).ToList();
            var newIntervals = Context.Connection.References.ProductsApplicability.GetIntervals(newLink.Guid).ToList();

            if (newIntervals.IsNullOrEmpty())
            {
                if (!sourceIntervals.IsNullOrEmpty())
                    content.AppendLine($"Для '{newLink.ChildObject}' удалена применяемость");
            }
            else
            {
                var intervalsText = new List<string>();

                if (sourceIntervals.IsNullOrEmpty())
                {
                    foreach (var interval in newIntervals)
                        intervalsText.Add(interval.ToString());

                    content.AppendLine($"Для '{newLink.ChildObject}' добавлена применяемость: {String.Join(", ", intervalsText)}");
                }
                else
                {
                    sourceIntervals.Sort(_applicabilityComparer);
                    newIntervals.Sort(_applicabilityComparer);

                    var differentIntervals = newIntervals.Except(sourceIntervals, _applicabilityComparer);
                    if (differentIntervals.Any())
                    {
                        foreach (var interval in differentIntervals)
                            intervalsText.Add(interval.ToString());

                        content.AppendLine($"Для '{newLink.ChildObject}' изменить применяемость: {String.Join(", ", intervalsText)}");
                    }
                }
            }

            string sourceConditions = Context.Connection.References.ProductsApplicability.GetApplicabilityConditions(sourceLink.Guid)?.Conditions ?? String.Empty;
            string newConditions = Context.Connection.References.ProductsApplicability.GetApplicabilityConditions(newLink.Guid)?.Conditions ?? String.Empty;

            string sourceFilterString = OptionsFilter.Deserialize(sourceConditions, Context.Connection)?.ToString();
            string newFilterString = OptionsFilter.Deserialize(newConditions, Context.Connection)?.ToString();

            if (String.IsNullOrEmpty(newFilterString))
            {
                if (!String.IsNullOrEmpty(sourceFilterString))
                    content.AppendLine($"Для '{newLink.ChildObject}' удалены условия применяемости");
            }
            else
            {
                if (String.IsNullOrEmpty(sourceFilterString))
                    content.AppendLine($"Для '{newLink.ChildObject}' добавлены условия применяемости: {newFilterString}");
                else
                {
                    if (newFilterString != sourceFilterString)
                        content.AppendLine($"Для '{newLink.ChildObject}' изменить условия применяемости: {newFilterString}");
                }
            }

            var changedParametersText = new List<string>();
            foreach (var sourceParameter in sourceLink.ParameterValues.Where(param => !param.ParameterInfo.IsSystem))
            {
                var sourceParameterGuid = sourceParameter.ParameterInfo.Guid;

                if (sourceParameterGuid == NomenclatureHierarchyLink.FieldKeys.DesignContextId ||
                    sourceParameterGuid == NomenclatureHierarchyLink.FieldKeys.SubstitutedLinkId ||
                    sourceParameterGuid == NomenclatureHierarchyLink.FieldKeys.CadObjectIdentifier ||
                    sourceParameterGuid == NomenclatureHierarchyLink.FieldKeys.DeletedInDesignContext
                    )
                    continue;

                var currentParameter = newLink.ParameterValues.FirstOrDefault(parameter
                    => parameter.ParameterInfo.Id == sourceParameter.ParameterInfo.Id);

                object currentValue = currentParameter.GetValue(currentParameter.ParameterInfo.Type.LanguageType);
                object sourceValue = sourceParameter.GetValue(currentParameter.ParameterInfo.Type.LanguageType);
                if (!Equals(currentValue, sourceValue))
                {
                    if (sourceParameterGuid == NomenclatureHierarchyLink.FieldKeys.Placement ||
                        sourceParameterGuid == NomenclatureHierarchyLink.FieldKeys.XMax ||
                        sourceParameterGuid == NomenclatureHierarchyLink.FieldKeys.XMin ||
                        sourceParameterGuid == NomenclatureHierarchyLink.FieldKeys.YMax ||
                        sourceParameterGuid == NomenclatureHierarchyLink.FieldKeys.YMin ||
                        sourceParameterGuid == NomenclatureHierarchyLink.FieldKeys.ZMax ||
                        sourceParameterGuid == NomenclatureHierarchyLink.FieldKeys.ZMin)
                    {
                        if (!changedParametersText.Contains(_changedPlacementText))
                            changedParametersText.Add(_changedPlacementText);
                    }
                    else
                    {
                        string parameterText = String.Format($"изменить '{currentParameter.ParameterInfo}' с '{sourceParameter.Value ?? String.Empty}'" +
                                          $" на '{currentParameter.Value ?? String.Empty }'");
                        changedParametersText.Add(parameterText);
                    }
                }
            }

            return String.Join(", ", changedParametersText);
        }

        public void CreateUsingAreaTextCore(List<NomenclatureHierarchyLink> checkedLinks)
        {
            var macro = Context[nameof(MacroKeys.MacroSource)] as CodeMacro;
            if (macro.DebugMode.Value == true)
            {
                Debugger.Launch();
                Debugger.Break();
            }

            var usingArea = Context.ReferenceObject as ModificationUsingAreaReferenceObject;

            var modification = usingArea.MasterObject;
            var designContext = (DesignContextObject)modification.GetObject(ModificationReferenceObject.RelationKeys.DesignContext);
            var targetRevision = ModificationHelper.GetLinkedObject(modification, ModificationReferenceObject.RelationKeys.PDMObject, designContext);

            var text = new StringBuilder();

            var groupingObjects = checkedLinks.GroupBy(link => link.ChildObject, link => link.ParentObject);
            foreach (var group in groupingObjects)
            {
                text.AppendLine($"Заменить '{group.Key.ToString() ?? String.Empty}'" +
                    $" на '{targetRevision?.ToString() ?? String.Empty}'" +
                    $" в составе {String.Join(", ", group.Select(grp => $"'{grp}'").Distinct())}");
            }
            string formattedText = RtfHtmlUtilites.ConvertPlainToRtf(text.ToString().TrimEnd());
            usingArea.Content.Value = formattedText;

            RefreshControls("UsingAreaContent");
        }

        public void CreateModificationTextsCore()
        {
            var macro = Context[nameof(MacroKeys.MacroSource)] as CodeMacro;
            if (macro.DebugMode.Value == true)
            {
                Debugger.Launch();
                Debugger.Break();
            }

            var modification = Context.ReferenceObject as ModificationReferenceObject;

            var actionObjects = modification.GetObjects(ModificationReferenceObject.RelationKeys.ModificationActions)
                .OfType<ModificationActionReferenceObject>();
            if (actionObjects.Any())
            {
                bool supportAcionOrder = actionObjects.FirstOrDefault()?.Reference.ParameterGroup.SupportsOrder ?? false;
                if (supportAcionOrder)
                    actionObjects = actionObjects.OrderBy(referenceObject => referenceObject.SystemFields.Order).ToList();

                using (var server = new RichEditDocumentServer())
                {
                    server.CreateNewDocument();

                    foreach (var action in actionObjects)
                    {
                        string actionContent = action.Content.ToString();
                        server.Document.AppendRtfText(actionContent, DevExpress.XtraRichEdit.API.Native.InsertOptions.KeepSourceFormatting);
                    }

                    modification.ModificationContent.Value = server.RtfText;
                }
            }
            else
            {
                modification.ModificationContent.Value = String.Empty;
            }

            var usingAreaObjects = modification.GetObjects(ModificationReferenceObject.RelationKeys.UsingArea)
                .OfType<ModificationUsingAreaReferenceObject>();
            if (usingAreaObjects.Any())
            {
                bool supportUsingAreaOrder = usingAreaObjects.FirstOrDefault()?.Reference.ParameterGroup.SupportsOrder ?? false;
                if (supportUsingAreaOrder)
                    usingAreaObjects = usingAreaObjects.OrderBy(referenceObject => referenceObject.SystemFields.Order).ToList();

                using (var server = new RichEditDocumentServer())
                {
                    server.CreateNewDocument();

                    foreach (var usingArea in usingAreaObjects)
                    {
                        string usingAreaContent = usingArea.Content.ToString();
                        server.Document.AppendRtfText(usingAreaContent, DevExpress.XtraRichEdit.API.Native.InsertOptions.KeepSourceFormatting);
                    }

                    modification.UsingAreaContent.Value = server.RtfText;
                }
            }
            else
            {
                modification.UsingAreaContent.Value = String.Empty;
            }
        }

        private class ApplicabilityComparer : IEqualityComparer<ReferenceObject>, IComparer<ReferenceObject>
        {
            public bool Equals(ReferenceObject x, ReferenceObject y)
                => Equals(x?.ToString(), y?.ToString());

            public int GetHashCode(ReferenceObject obj)
                => obj.ToString().GetHashCode();

            public int Compare(ReferenceObject x, ReferenceObject y)
                => String.Compare(x?.ToString(), y?.ToString());
        }
    }
}

