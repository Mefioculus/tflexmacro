using System;
using System.Linq;
using System.Windows.Forms;
using TFlex.DOCs.Model.Macros;
using TFlex.DOCs.Model.References;
using TFlex.DOCs.Model.References.Links;
using TFlex.DOCs.Model.Structure;
using TFlex.DOCs.Resources.Strings;

#if (WPFCLIENT)
using TFlex.DOCs.Client.ViewModels;
#else
using TFlex.DOCs.UI.Objects.Managers;
#endif

public static class Guids
{
    // спр. Материалы
    public static readonly Guid СводноеНаименование = new Guid("23cfeee6-57f3-4a1e-9cf0-9040fed0e90c");
    public static readonly Guid НазначениеОписание = new Guid("8b7e28ff-83a2-43ad-bdf1-3ad44a12ef82");
    public static readonly Guid ГПФизическиеСвойстваМатериала = new Guid("03206fdb-8810-4470-8844-bcf85cf37b57"); // ГП - группа параметров

    public static readonly Guid Обозначение1Строка = new Guid("f125ebf1-da65-4234-8724-306f696ac54a");
    public static readonly Guid Обозначение2Строка = new Guid("4eec9f87-ad39-4cd1-b524-7a58da2b5b47");
    public static readonly Guid Обозначение3Строка = new Guid("3b6c2159-7644-49f7-a4c6-1e423884f8ce");
    public static readonly Guid Обозначение4Строка = new Guid("1d57fda7-8fe3-424d-a054-ac997056c5c8");

    public static readonly Guid СвязьМаркаМатериала = new Guid("43300b75-129e-4f62-b5c2-81c4d84ede03");
    public static readonly Guid СвязьСортамент = new Guid("09aa0bd3-c7a1-4a94-8ee6-b911748cfd32");
    public static readonly Guid СвязьТребованияКСортаменту = new Guid("c18f1dbd-59c0-4c57-b2ac-44f48daa8113");
    public static readonly Guid СвязьТехТребования = new Guid("f387600d-e613-4a18-b4e6-af7d0831d8e4");
    public static readonly Guid СвязьТехническиеУсловия = new Guid("1aae8a96-46fa-4d95-a0a7-8f72de68bd5b");

    // спр. Сортамент материалов
    public static readonly Guid ОбозначениеНТДСортамента = new Guid("a41e58da-eb6d-4e14-b8f6-84168c59aa2e");
    public static readonly Guid ТипНТДСортамента = new Guid("e7137a3f-f5bc-4bca-bd98-1f7277f4e49b");
    public static readonly Guid ОбозначениеСортамента = new Guid("e7a357ea-fa09-4988-9141-ea9c6d02dd63");
    public static readonly Guid НаименованиеСортамента = new Guid("c504a6bd-9ed5-4633-8991-061edcab13e2");

    // спр. Требования к сортаменту
    public static readonly Guid ПрефиксТребСорт = new Guid("bafd0988-e966-4df4-b36f-8522ea4b6319");
    public static readonly Guid СуффиксТребСорт = new Guid("60d848fc-9e03-4ec2-b001-a0385c7a4946");

    // спр. Технические требования к материалам
    public static readonly Guid ПрефиксТребМатер = new Guid("4a824d6d-8c55-4c30-b1e9-05e1d8cb4a39");
    public static readonly Guid СуффиксТребМатер = new Guid("c2ac9562-21a8-4de5-bf56-efc376dae5b2");

    // спр. Марки материалов
    public static readonly Guid ОбозначениеНТДМарки = new Guid("f7af6da5-9fd6-46d9-80c8-f686f743eec2");
    public static readonly Guid ТипНТДМарки = new Guid("868da86b-a6bf-4fb2-8515-c0d6c7ab767e");
    public static readonly Guid Марка = new Guid("3fbd153e-096f-44b3-acc2-0613c8091047");
    public static readonly Guid НаименованиеМарки = new Guid("c7a0352d-4a29-4112-8e0e-0966e632f815");
    public static readonly Guid Характеристика = new Guid("814bc54f-b798-4ab1-bad3-997ce3466753"); // Назначение / характеристика
    public static readonly Guid СписокФизическиеСвойства = new Guid("ae609957-7848-4295-928e-f4ccac0a3ff9");
   
    // спр. Технические условия на материалы
    public static readonly Guid ГОСТ = new Guid("552bbdd6-b32b-4504-b1a7-e2ac16896bcb");
    public static readonly Guid СуффиксТУ = new Guid("1ae8b62f-3cb9-4422-9ba8-5d1030fc2811");
}

public class Macro : MacroProvider
{
    public Macro(MacroContext context)
        : base(context)
    {
    }

    private void SetNotationByAssortment(ReferenceObject material)
    {
        ReferenceObject assortment = material.GetObject(Guids.СвязьСортамент);
        if (assortment == null)
        {
            material[Guids.Обозначение1Строка].Value = material[Guids.Обозначение2Строка].Value = string.Empty;
            return;
        }

        material[Guids.Обозначение1Строка].Value = assortment[Guids.НаименованиеСортамента].ToString();

        ReferenceObject assortmentsRequirement = material.GetObject(Guids.СвязьТребованияКСортаменту); // Требования к сортаменту
        string design2Line;
        if (assortmentsRequirement != null)
        {
            design2Line = assortmentsRequirement[Guids.ПрефиксТребСорт] + " " +
                            assortment[Guids.ОбозначениеСортамента] + " " +
                            assortmentsRequirement[Guids.СуффиксТребСорт];
        }
        else
        {
            design2Line = assortment[Guids.ОбозначениеСортамента].ToString(); //Обозначение сортамент
        }

        material[Guids.Обозначение2Строка].Value = design2Line.Trim() + " " // Во 2-ую строку наименование сортамента
        + assortment[Guids.ТипНТДСортамента].ParameterInfo.ValueList.GetName(assortment[Guids.ТипНТДСортамента].Value) + " "   // + Тип НТД
        + assortment[Guids.ОбозначениеНТДСортамента];
    }

    private void SetNotationByBrand(ReferenceObject material, ReferenceObject materialBrand)
    {
        string design3Line = string.Empty;
        ReferenceObject technicalRequirements = material.GetObject(Guids.СвязьТехТребования); // Тех. требования
        if (technicalRequirements != null)
        {
            design3Line = technicalRequirements[Guids.ПрефиксТребМатер] + " "
                + materialBrand[Guids.Марка] + " "
                + technicalRequirements[Guids.СуффиксТребМатер];
        }
        else
        {
            design3Line = materialBrand[Guids.Марка].ToString();
        }

        material[Guids.Обозначение3Строка].Value = design3Line.Trim() + " "
            + materialBrand[Guids.ТипНТДМарки].ParameterInfo.ValueList.GetName(materialBrand[Guids.ТипНТДМарки].Value) + " "  // Тип НТД
            + materialBrand[Guids.ОбозначениеНТДМарки]; // Обозначение НТД;
    }

    private void SetNotationByTechnicalConditions(ReferenceObject material)
    {
        ReferenceObject technicalConditions = material.GetObject(Guids.СвязьТехническиеУсловия); // ТУ
        if (technicalConditions != null)
        {
            material[Guids.Обозначение4Строка].Value = technicalConditions[Guids.ГОСТ]; // В 4-ую строку ГОСТ технических условий 

            string design1Line = material[Guids.Обозначение1Строка].ToString();
            if (!string.IsNullOrWhiteSpace(design1Line))
                material[Guids.Обозначение1Строка].Value = string.Concat(design1Line, " ", technicalConditions[Guids.СуффиксТУ]);
        }
        else
        {
            material[Guids.Обозначение4Строка].Value = string.Empty;
        }
    }


    //Сформировать 4 строки наименования
    public void Сформировать()
    {
        ReferenceObject material = Context.ReferenceObject;
        if (material == null || !material.Changing)
            return;

        ReferenceObject materialBrand = GetMaterialBrand(material);

        SetNotationByAssortment(material);
        SetNotationByBrand(material, materialBrand);
        SetNotationByTechnicalConditions(material);
    }

    private ReferenceObject GetMaterialBrand(ReferenceObject material, bool clearDescription = false)
    {
        if (material == null)
            return null;

        ReferenceObject materialBrand = material.GetObject(Guids.СвязьМаркаМатериала); // Марка
        if (materialBrand != null)
            return materialBrand;

        if (clearDescription)
            material[Guids.СводноеНаименование].Value = string.Empty;

        Error("Не задана марка материала");
        return null;
    }

    // Сформировать строку сводного наименования
    public void СформироватьСтрочку()
    {
        ReferenceObject material = Context.ReferenceObject;
        if (material == null || !material.Changing)
            return;

        ReferenceObject materialBrand = GetMaterialBrand(material, true);

        Сформировать();

        string consolidatedName = string.Empty; // Сводное наименование
        string design1Line = material[Guids.Обозначение1Строка].ToString();
        string design2Line = material[Guids.Обозначение2Строка].ToString();
        string design3Line = material[Guids.Обозначение3Строка].ToString();
        string design4Line = material[Guids.Обозначение4Строка].ToString();

        if (!string.IsNullOrWhiteSpace(design1Line))
            consolidatedName = design1Line;

        if (!string.IsNullOrWhiteSpace(design2Line))
            consolidatedName = string.Concat(consolidatedName, " ", design2Line);

        if (!string.IsNullOrWhiteSpace(consolidatedName))
            consolidatedName = string.Concat(consolidatedName, " / ");

        if (!string.IsNullOrWhiteSpace(design3Line))
            consolidatedName = string.Concat(consolidatedName, design3Line, " ");

        if (!string.IsNullOrWhiteSpace(design4Line))
            consolidatedName = consolidatedName + design4Line;

        consolidatedName.Trim();
        material[Guids.СводноеНаименование].Value = consolidatedName;

        if (string.IsNullOrWhiteSpace(material[Guids.НазначениеОписание].ToString()) && materialBrand.Parent != null)
        {
            string description = materialBrand.Parent[Guids.НаименованиеМарки].ToString();
            string brandCharacteristic = materialBrand[Guids.Характеристика].ToString();
            if (!string.IsNullOrWhiteSpace(brandCharacteristic))
                description += Environment.NewLine + "Применение: " + brandCharacteristic;

            material[Guids.НазначениеОписание].Value = description;
        }
    }

    // Взять физические свойства из марки материала
    public void ВзятьФизическиеСвойства()
    {
        ReferenceObject material = Context.ReferenceObject;
        if (material == null || !material.Changing)
            return;

        ReferenceObject physicalPropertyObject = GetSelectedPhysicalPropertyObject(material);
        if (physicalPropertyObject == null)
            return;

        // Ищем группу параметров "Физические свойства" в справочнике "Материалы"
        ParameterGroup materialPhysicalPropsGroup = material.Reference.ParameterGroup.OneToOneTables.Find(Guids.ГПФизическиеСвойстваМатериала);
        if (materialPhysicalPropsGroup == null)
        {
            Message(Texts.Error, "Не найдена группа параметров 'Физические свойства'");
            return;
        }

        foreach (ParameterInfo parameterInfo in physicalPropertyObject.Reference.ParameterGroup.OneToOneParameters)
        {
            if (parameterInfo.IsSystem)
                continue;

            object value = physicalPropertyObject[parameterInfo].Value;
            if (value == null)
                continue;

            ParameterInfo materialParameter = materialPhysicalPropsGroup.Parameters.Find(parameterInfo.FieldName);
            if (materialParameter != null)
                material[materialParameter].Value = value;
        }
    }

    private ReferenceObject GetSelectedPhysicalPropertyObject(ReferenceObject material)
    {
        OneToOneLink linkMaterialToBrand = material.Links.ToOne[Guids.СвязьМаркаМатериала];
        if (linkMaterialToBrand == null)
        {
            Message(Texts.Error, "Не найдена связь 'Марка материала'");
            return null;
        }

        ReferenceObject materialBrand = linkMaterialToBrand.LinkedObject;
        if (materialBrand == null)
        {
            Message(Texts.Error, "Необходимо выбрать марку материала");
            return null;
        }

        var uiContext = Context as UIMacroContext;
        if (uiContext == null)
            return null;
        
        ParameterGroup markPhysicalPropsLink = materialBrand.Reference.ParameterGroup.OneToManyTables.Find(Guids.СписокФизическиеСвойства);
        if (markPhysicalPropsLink == null)
        {
            Message(Texts.Error, "Физические свойства марки не найдены!");
            return null;
        }
        
        Reference physicalPropsReference = materialBrand.Links.ToMany[markPhysicalPropsLink].LinkReference;
        
        var dialog = uiContext.CreateSelectObjectDialog(physicalPropsReference);
        dialog.MultipleSelect = false;
        
        return dialog.Show() ? dialog.FocusedObject : null;
    }


    #region Копирование в буфер

    //Копировать сводное наименование в буфер обмена
    public void КопироватьВБуфер()
    {
        ReferenceObject material = Context.ReferenceObject;
        if (material == null)
            return;

        string consolidatedName = material[Guids.СводноеНаименование].ToString().Trim();
        if (!string.IsNullOrEmpty(consolidatedName))
            Clipboard.SetDataObject(consolidatedName);
    }

    //Копировать строки в буфер обмена с разделением знаком табуляции. Для вставки в другие пограммы
    public void КопироватьВБуфер1()
    {
        ReferenceObject material = Context.ReferenceObject;
        if (material == null)
            return;

        string[] notationLines = new string[]
        {
            material[Guids.Обозначение1Строка].ToString().Trim(),
            material[Guids.Обозначение2Строка].ToString().Trim(),
            material[Guids.Обозначение3Строка].ToString().Trim(),
            material[Guids.Обозначение4Строка].ToString().Trim()
        };

        if (notationLines.Any(line => !string.IsNullOrWhiteSpace(line)))
            Clipboard.SetDataObject(string.Join("\t", notationLines));
    }

    #endregion
}

