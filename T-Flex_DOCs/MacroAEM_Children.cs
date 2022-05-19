/*
TFlex.DOCs.SynchronizerReference.dll
*/

using System;
using System.Collections.Generic;
using System.Linq;
using TFlex.DOCs.Model.Macros;
using TFlex.DOCs.Model.References;
using TFlex.DOCs.Model.Search;
using TFlex.DOCs.Synchronization.Macros;
using TFlex.DOCs.Synchronization.SyncData;

public class MacroAEM_Children : MacroProvider
{
    public MacroAEM_Children(MacroContext context) : base(context) { }

    private static readonly Guid СписокНоменклатуры_Guid = new Guid("c9d26b3c-b318-4160-90ae-b9d4dd7565b6");
    private static readonly Guid СписокНоменклатуры_Обозначение_Guid = new Guid("1bbb1d78-6e30-40b4-acde-4fc844477200");

    private static readonly Guid Подключения_Guid = new Guid("9dd79ab1-2f40-41f3-bebc-0a8b4f6c249e");
    private static readonly Guid Подключения_Сборка_Guid = new Guid("4a3cb1ca-6a4c-4dce-8c25-c5c3bd13a807");
    private static readonly Guid Подключения_Обозначение_Guid = new Guid("7d1ac031-8c7f-49b5-84b8-c5bafa3918c2");

    public void Начало()
    {
        var context = Context as ExchangeDataMacroContext;
        var objectSettings = context.Settings.ObjectSettings.OfType<DataExchangeReferenceObjectSettings>().FirstOrDefault();

        if (objectSettings == null || objectSettings.ReferenceInfo.Guid != СписокНоменклатуры_Guid)
            return;

        var списокНоменклатурыReference = objectSettings.ReferenceInfo.CreateReference();
        var списокНоменклатуры = списокНоменклатурыReference.Find(objectSettings.Objects);
        var levelObjects = списокНоменклатуры;

        var подключениеReference = Context.Connection.ReferenceCatalog.Find(Подключения_Guid).CreateReference();
        var allObjects = new HashSet<ReferenceObject>(списокНоменклатуры);

        while (levelObjects.Count > 0)
        {
            string[] обозначения = levelObjects.Select(obj => obj[СписокНоменклатуры_Обозначение_Guid].GetString()).ToArray();
            List<ReferenceObject> подключения;

            using (var filter = new Filter(подключениеReference.ParameterGroup))
            {
                filter.Terms.AddTerm(
                    подключениеReference.ParameterGroup.OneToOneParameters.Find(Подключения_Сборка_Guid),
                    ComparisonOperator.IsOneOf,
                    обозначения);

                подключения = подключениеReference.Find(filter);
            }

            string[] дочерниеОбозначения = подключения.Select(obj => obj[Подключения_Обозначение_Guid].GetString())
                .Where(s => !String.IsNullOrEmpty(s)).ToArray();

            if (дочерниеОбозначения.Length == 0)
                break;

            List<ReferenceObject> дочернийСписокНоменклатуры;

            using (var filter = new Filter(списокНоменклатурыReference.ParameterGroup))
            {
                filter.Terms.AddTerm(
                    списокНоменклатурыReference.ParameterGroup.OneToOneParameters.Find(СписокНоменклатуры_Обозначение_Guid),
                    ComparisonOperator.IsOneOf,
                    дочерниеОбозначения);

                дочернийСписокНоменклатуры = списокНоменклатурыReference.Find(filter);
            }

            levelObjects = new List<ReferenceObject>(дочернийСписокНоменклатуры.Count);

            foreach (var obj in дочернийСписокНоменклатуры)
            {
                if (allObjects.Add(obj))
                    levelObjects.Add(obj);
            }

            objectSettings.AddObjects(levelObjects);
        }
    }

}
