using System;
using System.Collections.Generic;
using TFlex.DOCs.Model.Macros;

public class MacrosGetSpecificationFromReq : MacroProvider
{
    public MacrosGetSpecificationFromReq(MacroContext context)
        : base(context)
    {
    }

    public override void Run()
    {
        // получить по связи спецификацию требований
        if (Context.ReferenceObject == null)
            return;

        CreateRelationsWithProductSpecifications();
        RefreshReferenceWindow();
    }

    public ButtonValidator GetButtonValidator()
    {
        bool visible = Context.Connection.ReferenceCatalog.Find(GuidsHelper.RequirementsReferenceGuid) != null;

        return new ButtonValidator()
        {
            Enable = true,
            Visible = visible
        };
    }

    private void CreateRelationsWithProductSpecifications()
    {
        // собрать существующие связи в hashset
        // вызывать существующий макрос по обновлению/добавлению характеристик в справочник Характеристики изделий
        // вновь пройтись по спецификации требований, собрать связи с характеристиками оттуда
        // проверяя в hashset наличие связи, добавлять те, которых нет.

        // итого три прохода: неприятно, но жить можно. 

        var hashSet = new HashSet<Guid>();
        foreach (var productSpecification in Context.ReferenceObject.GetObjects(GuidsHelper
            .RelationWithProductSpecifications))
        {
            hashSet.Add(productSpecification.Guid);
        }

        var requirementSpecificationReferenceObject =
            Context.ReferenceObject.GetObject(GuidsHelper.RelationWithRequirementSpecification);

        if (requirementSpecificationReferenceObject is null)
            Error($"Не задана спецификация требований для текущего объекта {Context.ReferenceObject}");

        WaitingDialog.Show("Идет получение характеристик", canCancel: true);
        Context.RunMacro("СУТР. Требования и характеристики", "ExportProductSpecificationsFrom",
            requirementSpecificationReferenceObject);
        
        try
        {

            Context.ReferenceObject.BeginChanges();
            foreach (var child in requirementSpecificationReferenceObject.Children.RecursiveLoad())
            {
                if (child.Class.IsInherit(GuidsHelper.RequirementSpecificationTypeGuid))
                {
                    var linkedProductCharacteristicReferenceObject =
                        child.GetObject(GuidsHelper.RequirementSpecificationProductSpecificationRelation);

                    if (linkedProductCharacteristicReferenceObject is null)
                        continue;

                    if (!hashSet.Contains(linkedProductCharacteristicReferenceObject.Guid))
                    {
                        Context.ReferenceObject.AddLinkedObject(GuidsHelper.RelationWithProductSpecifications,
                            linkedProductCharacteristicReferenceObject);
                    }
                }
            }

            Context.ReferenceObject.EndChanges();
        }
        catch
        {
            if (Context.ReferenceObject.Changing)
                Context.ReferenceObject.CancelChanges();

            throw;
        }
        finally
        {
            WaitingDialog.Hide();
        }
    }

    private class GuidsHelper
    {
        /// <summary>
        /// Связь объекта классификатора изделий со спецификацией требований
        /// </summary>
        public static readonly Guid RelationWithRequirementSpecification =
            new Guid("db9ea4a9-1743-4727-ac2a-76441e29500a");

        /// <summary>
        /// Связь объекта классификатора изделий с характеристиками изделия
        /// </summary>
        public static readonly Guid RelationWithProductSpecifications =
            new Guid("2e45705c-cb8b-4712-8fad-1d0618fbd7eb");

        /// <summary>
        /// Guid типа Требуемая характеристика справочника Требования
        /// </summary>
        public static readonly Guid RequirementSpecificationTypeGuid = new Guid("f4edbee1-383e-421d-bbc9-3a00225307a1");

        /// <summary>
        /// Guid связи между требуемой характеристикой и прикрепленной к ней характеристикой изделия
        /// </summary>
        public static readonly Guid RequirementSpecificationProductSpecificationRelation =
            new Guid("daaf8005-e1bf-42f9-9f95-656b7c262f85");

        /// <summary>
        /// Guid справочника Требования
        /// </summary>
        public static readonly Guid RequirementsReferenceGuid =
                new Guid("48c51985-0f22-4315-a965-7b49888f4098");
            
    }
}
