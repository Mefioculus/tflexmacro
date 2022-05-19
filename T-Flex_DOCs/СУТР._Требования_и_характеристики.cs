using System;
using System.Linq;
using TFlex.DOCs.Model.Macros;
using TFlex.DOCs.Model.References;
using TFlex.DOCs.Model.References.Units;

public class MacroRequirementsManagement: MacroProvider
{
    private static readonly Guid[] _observedRequiredCharacteristicLinks = new Guid[] { Guids.СвязьОтТхНаЕдиницуИзмерения };

    private static readonly Guid[] _observedRequiredCharacteristicParams = new Guid[]
    {
        Guids.ТипТребуемойХарактеристики,
        Guids.НаименованиеТребуемойХарактеристики,
        Guids.ОбозначениеТребуемойХарактеристики,
        Guids.ЗначениеТекстТребуемойХарактеристики,
        Guids.ЗначениеДаНетТребуемойХарактеристики,
        Guids.ЗначениеЧислоТребуемойХарактеристики,
        Guids.ДопустимоеMinТребуемойХарактеристики,
        Guids.ДопустимоеMaxТребуемойХарактеристики,
        Guids.НедопустимоеMinТребуемойХарактеристики,
        Guids.НедопустимоеMaxТребуемойХарактеристики
    };

    public MacroRequirementsManagement(MacroContext context)
        : base(context)
    {
    }

    public void ЗавершениеИзмененияПараметраТребуемойХарактеристики()
    {
        if (!(Context.ModelChangedArgs is ObjectParameterChangedEventArgs objectParameterChangedEventArgs))
            return;

        if (_observedRequiredCharacteristicParams.Contains(objectParameterChangedEventArgs.Parameter.ParameterInfo.Guid))
            ЗаполнитьТекстТребуемойХарактеристики();
    }

    public void ЗавершениеИзмененияСвязиТребуемойХарактеристики()
    {
        if (!(Context.ModelChangedArgs is ObjectLinkChangedEventArgs objectLinkChangedEventArgs))
            return;

        if (_observedRequiredCharacteristicLinks.Contains(objectLinkChangedEventArgs.Link.LinkGroup.Guid))
            ЗаполнитьТекстТребуемойХарактеристики();
    }

    public ButtonValidator GetButtonValidator()
    {
       bool visible = Context.ReferenceObject != null;
       return new ButtonValidator()
            {
                Enable = true,
                Visible = visible
            };
    }
    
    public void ExportProductSpecificationsFrom(ReferenceObject requirementReferenceObject)
    {
    	if (requirementReferenceObject.Class.IsInherit(Guids.ТипТребуемаяХарактеристика))
        {
            UpdateLinkedCharacteristic(Context.ReferenceObject);
            return;
        }
    	
    	if (!requirementReferenceObject.Class.IsInherit(Guids.ReqSpecTypeGuid))
    	    Error($"{requirementReferenceObject} не является типом Спецификация требований");
    	
        var loadSettings = requirementReferenceObject.Reference.LoadSettings.Clone();
        loadSettings.AddRelation(Guids.СвязьХарактеристикаИзделия);
        loadSettings.AddParameters(Guids.ТипТребуемойХарактеристики, Guids.ЗначениеТекстТребуемойХарактеристики,
            Guids.ЗначениеДаНетТребуемойХарактеристики, Guids.ЗначениеЧислоТребуемойХарактеристики);

        var childCharacteristics = requirementReferenceObject.Reference.RecursiveLoad(new[] { requirementReferenceObject },
            RelationLoadSettings.RecursiveLoadDirection.Children, loadSettings);

        foreach (var childRequiredCharacteristic in childCharacteristics)
        {
            if (childRequiredCharacteristic.Class.IsInherit(Guids.ТипТребуемаяХарактеристика))
                UpdateLinkedCharacteristic(childRequiredCharacteristic);
        }
    }
    
    public void ЭкспортироватьХарактеристики()
    {
        if (Context.ReferenceObject == null)
            Error("Нельзя вызывать экспорт характеристик на корневом объекте");

        if (Context.ReferenceObject.Class.IsInherit(Guids.ТипТребуемаяХарактеристика))
        {
            UpdateLinkedCharacteristic(Context.ReferenceObject);
            return;
        }

        var requiredCharacteristic = Context.ReferenceObject;

        var loadSettings = requiredCharacteristic.Reference.LoadSettings.Clone();
        loadSettings.AddRelation(Guids.СвязьХарактеристикаИзделия);
        loadSettings.AddParameters(Guids.ТипТребуемойХарактеристики, Guids.ЗначениеТекстТребуемойХарактеристики,
            Guids.ЗначениеДаНетТребуемойХарактеристики, Guids.ЗначениеЧислоТребуемойХарактеристики);

        var childCharacteristics = requiredCharacteristic.Reference.RecursiveLoad(new[] {requiredCharacteristic},
            RelationLoadSettings.RecursiveLoadDirection.Children, loadSettings);

        foreach (var childRequiredCharacteristic in childCharacteristics)
        {
            if (childRequiredCharacteristic.Class.IsInherit(Guids.ТипТребуемаяХарактеристика))
                UpdateLinkedCharacteristic(childRequiredCharacteristic);
        }
    }

    public void SetCharacteristicTextByRequiredCharacteristic(ReferenceObject requiredCharacteristic, ReferenceObject characteristicObject)
    {
        if (requiredCharacteristic is null || characteristicObject is null)
            return;

        var requiredCharacteristicType = (CharacteristicType)requiredCharacteristic[Guids.ТипТребуемойХарактеристики].GetInt32();
        characteristicObject.Modify(co =>
        {
            co[Guids.ЗначениеХарактеристики].Value = requiredCharacteristicType switch
            {
                CharacteristicType.Text => requiredCharacteristic[Guids.ИтогЗначениеТекстТребуемойХарактеристики].GetString(),
                CharacteristicType.Bool => requiredCharacteristic[Guids.ИтогЗначениеДаНетТребуемойХарактеристики].GetBoolean() ? "Да" : "Нет",
                CharacteristicType.Number => requiredCharacteristic[Guids.ИтогЗначениеЧислоТребуемойХарактеристики].GetString(),
                _ => throw new InvalidOperationException($"Задан неизвестный тип характеристики {requiredCharacteristicType}")
            };
        });
    }

    public void ЗаполнитьТекстТребуемойХарактеристики()
    {
        var requiredCharacteristic = Context.ReferenceObject;
        if (requiredCharacteristic?.Class.IsInherit(Guids.ТипТребуемаяХарактеристика) == false)
            return;

        var characteristicName = requiredCharacteristic[Guids.НаименованиеТребуемойХарактеристики].GetString();
        var characteristicDenotation = requiredCharacteristic[Guids.ОбозначениеТребуемойХарактеристики].GetString();
        var characteristicFullName = $"{characteristicName} ({characteristicDenotation})";

        var textValue = requiredCharacteristic[Guids.ЗначениеТекстТребуемойХарактеристики].GetString();
        var boolValue = requiredCharacteristic[Guids.ЗначениеДаНетТребуемойХарактеристики].GetBoolean() ? "Да" : "Нет";

        var characteristicType = (CharacteristicType)requiredCharacteristic[Guids.ТипТребуемойХарактеристики].GetInt32();
        requiredCharacteristic[Guids.ТекстТребуемойХарактеристики].Value = characteristicType switch
        {
            CharacteristicType.Text => characteristicFullName + Environment.NewLine + $"Требуемое значение: {textValue}",
            CharacteristicType.Bool => characteristicFullName + Environment.NewLine + $"Требуемое значение: {boolValue}",
            CharacteristicType.Number => characteristicFullName + Environment.NewLine + GetRequiredValueOfNumberCharacteristic(requiredCharacteristic),
            _ => throw new InvalidOperationException($"Задан неизвестный тип характеристики {characteristicType}")
        };
    }

    private void UpdateLinkedCharacteristic(ReferenceObject requiredCharacteristic)
    {
        var linkedCharacteristic = requiredCharacteristic.GetObject(Guids.СвязьХарактеристикаИзделия);
        if (linkedCharacteristic is null)
        {
            var characteristicsReference = Context.Connection.ReferenceCatalog.Find(Guids.СправочникХарактеристикаИзделий).CreateReference();
            var characteristic = characteristicsReference.CreateReferenceObject();
            UpdateLinkedCharacteristicParameters(requiredCharacteristic, characteristic);
            if (characteristic.SaveSet != null)
                characteristic.SaveSet.EndChanges();
            else
                characteristic.EndChanges();

            requiredCharacteristic.Modify(rc => rc.SetLinkedObject(Guids.СвязьХарактеристикаИзделия, characteristic));
        }
        else
        {
            UpdateLinkedCharacteristicParameters(requiredCharacteristic, linkedCharacteristic);
        }
    }

    private void UpdateLinkedCharacteristicParameters(ReferenceObject requiredCharacteristic, ReferenceObject characteristicObject)
    {
        if (requiredCharacteristic is null || characteristicObject is null)
            return;

        characteristicObject.Modify((characteristic) =>
        {
            characteristic[Guids.ОбозначениеХарактеристики].Value = requiredCharacteristic[Guids.ОбозначениеТребуемойХарактеристики].GetString();
            characteristic[Guids.НаименованиеХарактеристики].Value = requiredCharacteristic[Guids.НаименованиеТребуемойХарактеристики].GetString();
            SetCharacteristicTextByRequiredCharacteristic(requiredCharacteristic, characteristic);
            characteristic.SetLinkedObject(Guids.СвязьОтХиНаЕдиницуИзмерения, requiredCharacteristic.GetObject(Guids.СвязьОтТхНаЕдиницуИзмерения));
        });
    }

    private string GetRequiredValueOfNumberCharacteristic(ReferenceObject requiredCharacteristic)
    {
        var numberValue = requiredCharacteristic[Guids.ЗначениеЧислоТребуемойХарактеристики].GetString();
        var unit = (requiredCharacteristic.GetObject(Guids.СвязьОтТхНаЕдиницуИзмерения) as Unit)?.ShortName.GetString();
        var acceptableMin = requiredCharacteristic[Guids.ДопустимоеMinТребуемойХарактеристики].GetString();
        var acceptableMax = requiredCharacteristic[Guids.ДопустимоеMaxТребуемойХарактеристики].GetString();
        var unacceptableMin = requiredCharacteristic[Guids.НедопустимоеMinТребуемойХарактеристики].GetString();
        var unacceptableMax = requiredCharacteristic[Guids.НедопустимоеMaxТребуемойХарактеристики].GetString();

        return $@"Требуемое значение: {numberValue} {unit}
Допустимый диапазон значений: от {acceptableMin} {unit} до  {acceptableMax} {unit}
Недопустимые отклонения: менее {unacceptableMin} {unit} и более {unacceptableMax} {unit}";
    }

    private enum CharacteristicType
    {
        Number = 1,
        Text,
        Bool
    }

    private static class Guids
    {
        // Требования 
        internal readonly static Guid ТипТребуемаяХарактеристика = new Guid("f4edbee1-383e-421d-bbc9-3a00225307a1");

        internal readonly static Guid СвязьХарактеристикаИзделия = new Guid("daaf8005-e1bf-42f9-9f95-656b7c262f85");
        internal readonly static Guid СвязьОтТхНаЕдиницуИзмерения = new Guid("24621957-59ab-4644-ba79-e5c26e261475");

        internal readonly static Guid ТипТребуемойХарактеристики = new Guid("af6b7bde-f9ab-4bba-9f02-2a839347bf2a"); // Это параметр
        internal readonly static Guid НаименованиеТребуемойХарактеристики = new Guid("55b07f2e-26fe-40c4-99b6-1561fb648db2");
        internal readonly static Guid ОбозначениеТребуемойХарактеристики = new Guid("ee37fa5b-2900-4467-8a0b-f3ec3089b845");
        internal readonly static Guid ТекстТребуемойХарактеристики = new Guid("49d6731b-9a72-4817-a23a-15bc919752d5");
        internal readonly static Guid ЗначениеТекстТребуемойХарактеристики = new Guid("bce34c0c-4e17-47a9-a317-1f4950f8fa33");
        internal readonly static Guid ЗначениеДаНетТребуемойХарактеристики = new Guid("3144ea92-6f17-466f-992b-3ecf1e854d27");
        internal readonly static Guid ЗначениеЧислоТребуемойХарактеристики = new Guid("39e4b247-9e7f-44ba-a46f-51d0c5420997");
        internal readonly static Guid ИтогЗначениеТекстТребуемойХарактеристики = new Guid("123d3469-852f-4354-a27c-750882e3813e");
        internal readonly static Guid ИтогЗначениеДаНетТребуемойХарактеристики = new Guid("8539a20d-865b-4ad6-8da8-c5db1cb6d8d8");
        internal readonly static Guid ИтогЗначениеЧислоТребуемойХарактеристики = new Guid("34b5c90f-fd2f-42c4-ab71-0232358afc90");
        internal readonly static Guid ДопустимоеMinТребуемойХарактеристики = new Guid("1b828397-b3d9-4aee-a552-6595c4e7963a");
        internal readonly static Guid ДопустимоеMaxТребуемойХарактеристики = new Guid("520c8bd8-e446-4135-881f-9e3dca04676a");
        internal readonly static Guid НедопустимоеMinТребуемойХарактеристики = new Guid("4cece521-81a3-421a-82e9-f52ccd7a2e48");
        internal readonly static Guid НедопустимоеMaxТребуемойХарактеристики = new Guid("f45c2cdf-1bd7-4ebd-b0c0-c73bce5595a4");

        // Характеристики изделий
        internal readonly static Guid СправочникХарактеристикаИзделий = new Guid("f4cf09ef-eb95-4711-ad13-7590d8d479cc");

        internal readonly static Guid СвязьОтХиНаЕдиницуИзмерения = new Guid("81a11fdd-dca9-48c5-b64e-226373f7e418");

        internal readonly static Guid ОбозначениеХарактеристики = new Guid("24ea7395-5f93-4364-8f70-a59fa01e47af");
        internal readonly static Guid НаименованиеХарактеристики = new Guid("c4fa5123-ec8c-4895-b21f-aa627ea9f3c7");
        internal readonly static Guid ЗначениеХарактеристики = new Guid("4df45323-7065-47cd-b3c7-c6d49622089e");
        //Guid типа спецификация требований
        internal readonly static Guid ReqSpecTypeGuid = new Guid("3707fb29-42c9-4a44-b10c-51ff809ffc64");
    }
}
