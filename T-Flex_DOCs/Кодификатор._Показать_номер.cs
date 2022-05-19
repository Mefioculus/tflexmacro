using System;
using System.Collections.Generic;
using System.Linq;
using TFlex.DOCs.Model.Macros;
using TFlex.DOCs.Model.References.Codifier;
using TFlex.DOCs.Model.References.Codifier.NumberElements;
public class MacroShowTestNumber : MacroProvider
{
    public MacroShowTestNumber(MacroContext context)
        : base(context)
    {
    }

    public override void Run()
    {
        int capacity = 4;
        if (Context.ReferenceObject == null)
        {
            Сообщение("Результат", "Невозможно получить код из-за отсутствия объектов номера");
            return;
        }

        // получить текущий объект автонумератора, собрать из него токены номера	
        var referenceObject = Context.ReferenceObject.MasterObject as CodifierReferenceObject;
        if (referenceObject == null)
        {
            Сообщение("Внимание!", "Макрос настроен не на справочник Кодификатор!");
            return;
        }

        var linkedObjects = referenceObject.NumberElements.OfType<NumberElementsReferenceObject>().ToArray();

        // создадим набор строковых значений параметров для заполнения пользователем. Допускается отсутствие элементов в наборе
        var testParametersValues = new List<string>(capacity);
        foreach (var numberReferenceObject in linkedObjects.OfType<ParameterTextElementReferenceObject>())
        {
            testParametersValues.Add(String.Empty);
        }

        if (testParametersValues.Count > 0)
        {
            var inputDialog = CreateInputDialog("Введите значения текстовых параметров");

            for (int i = 0; i < testParametersValues.Count; i++)
                inputDialog.AddString($"Параметр {i}", String.Empty);

            if (inputDialog.Show())
            {
                for (int i = 0; i < testParametersValues.Count; i++)
                    testParametersValues[i] = inputDialog[$"Параметр {i}"];
            }
        }

        string numberExample = referenceObject.GetTestNumber(testParametersValues);
        Сообщение("Результат", numberExample);
    }
}
