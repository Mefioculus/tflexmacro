using System;
using TFlex.DOCs.Model.Macros;

public class MacroDocNSI : MacroProvider
{
    public MacroDocNSI(MacroContext context)
        : base(context)
    {

    }

    public override void Run()
    {
    }

    /// <summary>
    /// Справочник "Параметры документа НСИ" - Тип "Параметры документа НСИ" - Событие "Сохранение объекта"
    /// </summary>
    public void Сохранение()
    {
        if (Context.ReferenceObject.IsPrototype)
            return;

        var типНСИ = ТекущийОбъект.СвязанныйОбъект["Тип НСИ"];
        if (типНСИ == null)
            Ошибка("Не указан тип НСИ");
    }

    /// <summary>
    /// Справочник "Параметры документа НСИ" - Список объектов "Список значений" - Событие "Сохранение объекта"
    /// </summary>
    public void СохранениеСпискаЗначений()
    {
        var currentObject = Context.ReferenceObject;
        var masterObject = currentObject.MasterObject;
        if (masterObject == null)
            return;

        int masterParameterType = masterObject[DescriptionParametersNSI.ParameterType].GetInt32();
        object newParameterValue = masterParameterType switch
        {
            3 => currentObject[ListValueParameters.Int_Value].Value,
            4 => currentObject[ListValueParameters.Double_Value].Value,
            5 => currentObject[ListValueParameters.Boolean_Value].Value,
            15 => currentObject[ListValueParameters.Double_Value].Value,
            24 => currentObject[ListValueParameters.String_Value].Value,
            26 => currentObject[ListValueParameters.MultilineString_value].Value,
            27 => currentObject[ListValueParameters.Guid_Value].Value,
            36 => currentObject[ListValueParameters.HTML_Value].Value,
            _ => throw new NotSupportedException(masterParameterType.ToString()),
        };

        if (newParameterValue == null)
            return;

        currentObject[ListValueParameters.Value].Value = newParameterValue;
    }

    /// <summary>
    /// Описание параметров НСИ
    /// </summary>
    private static class DescriptionParametersNSI
    {
        /// <summary>
        /// Тип параметр (int)
        /// </summary>
        public static Guid ParameterType = new Guid("5f095520-bc54-4eb6-a6b8-b64c024a433a");
    }

    internal static class ListValueParameters
    {
        //Основные параметры
        public static Guid Ikonka = new Guid("344277d8-f2d0-4d26-9dd6-0d4526c34635");
        public static Guid Name = new Guid("a46662b8-98c9-4234-bf8c-ecf7412b9631");
        public static Guid Value = new Guid("cecab80f-604e-4f86-bd3b-ac74bc2d93b6");
        public static Guid Description = new Guid("2de5e3a0-856f-4a2f-bb99-316dbf910a11");

        //string
        public static Guid String_Value = new Guid("433a901e-35ef-4488-80a8-8dc83ec286a0");
        public static Guid MultilineString_value = new Guid("376ab6ec-1757-4103-a82f-57b1bb792ea0");
        public static Guid HTML_Value = new Guid("957e8d9d-bcb1-4adb-be2c-5af3eb864a92");

        //number
        public static Guid Int_Value = new Guid("4979a853-1914-4249-93d2-b507bbba8ac9");
        public static Guid Double_Value = new Guid("73aee616-25c0-4399-9260-ecf773ceea9b");

        public static Guid Guid_Value = new Guid("31ca0754-5b60-433b-b1e4-d2a56806e42b");
        public static Guid Boolean_Value = new Guid("d3e9fbd8-984b-4be2-a86a-e4ac03595ede");
        public static Guid DateTime_Value = new Guid("4bd902a9-a3b8-4b90-875b-cda3cd08598a");
    }
}
