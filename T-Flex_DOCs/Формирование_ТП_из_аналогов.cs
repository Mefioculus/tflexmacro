using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using TFlex.DOCs.Model.Macros;
using TFlex.DOCs.Model.Macros.ObjectModel;
using TFlex.DOCs.Model.References;
using TFlex.DOCs.Model.Search;
using TFlex.DOCs.Model.Structure;

public class MacroAEMCreate : MacroProvider
{
    /// <summary>
    /// Строка для записи лога ошибок
    /// </summary>
    //private readonly StringBuilder _logRemarks = new StringBuilder();

    private static readonly Guid _marchpReference = new Guid("dda8e2f7-67e9-41ea-a8db-cbe6c9b7eb9b");
    private static readonly Guid _marchpParameterShifr = new Guid("45275ca2-eb5a-47aa-afb6-2f72b3ff3b9a");
    /// <summary>
    /// Номенклатура параметр шифр
    /// </summary>
    private static readonly Guid _nomenParameterDenotation = new Guid("ae35e329-15b4-4281-ad2b-0e8659ad2bfb");

    public MacroAEMCreate(MacroContext context)
        : base(context)
    {
        /*if (Вопрос("Хотите запустить в режиме отладки?"))
        {
            System.Diagnostics.Debugger.Launch();
            System.Diagnostics.Debugger.Break();
        }*/
    }

    public override void Run()
    {
        //Получаем выбранные объекты
        var selectedObjects = Context.GetSelectedObjects();
        RunImportTP(selectedObjects.ToList());
        /*if (selectedObjects.Length == 0)
            return;

        //Производим поиск уникальных записей по шифру
        var distinctReferenceObjects = selectedObjects.Distinct(new Compare()).ToList();

        RefObjList refObjList = new RefObjList(distinctReferenceObjects, Context);

        ОбменДанными.ImportObjects("Создание техпроцессов из Marchp", refObjList, false);*/
    }
    public void ИмпортироватьЦехопереходыИзФоксПро()
    {
        Reference marcReference = FindReference(_marchpReference);
        ParameterInfo shifrParameter = marcReference.ParameterGroup.Parameters.Find(_marchpParameterShifr);

        if (shifrParameter == null)
            Ошибка(String.Format("Не найден параметр с Guid: {0}", _marchpParameterShifr));

        var errors = new List<string>();

        foreach (var currentObject in Context.GetSelectedObjects())
        {
            var nomDenotation = currentObject[_nomenParameterDenotation].GetString().Replace(".", string.Empty); // Удалить точки
            
            /*            
            if (nomDenotation.StartsWith("3905"))
            {
            	nomDenotation = nomDenotation.Insert(4, ".");
            }
            */

            var marchpObjects = marcReference.Find(shifrParameter, nomDenotation);

            if (marchpObjects.Count == 0)
            {
                errors.Add(String.Format("Не найдены объекты с шифром: {0}", nomDenotation));
                continue;
            }

            ПолучитьДочерниеОбъекты(marchpObjects);
            RunImportTP(marchpObjects);

            RefObjList refObjList = new RefObjList(marchpObjects, Context);

            ОбменДанными.ImportObjects("Создание цехопереходов из Marchp", refObjList, false);
        }

        if (errors.Count > 0)
            Ошибка(String.Join(Environment.NewLine, errors));
    }

    private Reference FindReference(Guid referenceGuid)
    {
        var referenceInfo = Context.Connection.ReferenceCatalog.Find(referenceGuid);
        if (referenceInfo == null)
            Ошибка(String.Format("Не найден справочник с Guid: {0}", referenceGuid));

        return referenceInfo.CreateReference();
    }

    private void RunImportTP(List<ReferenceObject> listObjects)
    {
        if (listObjects.Count == 0)
            return;

        //Производим поиск уникальных записей по шифру
        //var distinctReferenceObjects = listObjects.Distinct(new Compare()).ToList();

        RefObjList refObjList = new RefObjList(listObjects, Context);

        ОбменДанными.ImportObjects("Создание техпроцессов из Marchp", refObjList, false);
    }

    ////Метод который используется в сравнении объектов
    //private class Compare : IEqualityComparer<ReferenceObject>
    //{
    //    public bool Equals(ReferenceObject x, ReferenceObject y)
    //    {
    //        //Сравнивает объекты по параметру Шифр
    //        return x[_marchpParameterShifr].GetString() == y[_marchpParameterShifr].GetString();
    //    }

    //    public int GetHashCode(ReferenceObject obj)
    //    {
    //        return obj[_marchpParameterShifr].GetString().GetHashCode();
    //    }
    //}


    //private static readonly Guid СписокНоменклатуры_Guid = new Guid("dda8e2f7-67e9-41ea-a8db-cbe6c9b7eb9b"); // MARCHP
    private static readonly Guid СписокНоменклатуры_Обозначение_Guid = new Guid("45275ca2-eb5a-47aa-afb6-2f72b3ff3b9a"); // SHIFR

    private static readonly Guid Подключения_Guid = new Guid("9dd79ab1-2f40-41f3-bebc-0a8b4f6c249e");
    private static readonly Guid Подключения_Сборка_Guid = new Guid("4a3cb1ca-6a4c-4dce-8c25-c5c3bd13a807");
    private static readonly Guid Подключения_Обозначение_Guid = new Guid("7d1ac031-8c7f-49b5-84b8-c5bafa3918c2");

    private void ПолучитьДочерниеОбъекты(List<ReferenceObject> списокНоменклатуры)
    {
        var списокНоменклатурыReference = списокНоменклатуры.First().Reference;
        var levelObjects = списокНоменклатуры.ToList();

        var подключениеReference = Context.Connection.ReferenceCatalog.Find(Подключения_Guid).CreateReference();
        var allObjects = new Dictionary<int, ReferenceObject>();

        foreach (var ro in levelObjects)
            allObjects.Add(ro.SystemFields.Id, ro);

        while (levelObjects.Count > 0)
        {
            string[] обозначения = levelObjects.Select(obj => ВставитьТочки(obj[СписокНоменклатуры_Обозначение_Guid].GetString())).ToArray();
            List<ReferenceObject> подключения;

            using (var filter = new Filter(подключениеReference.ParameterGroup))
            {
                filter.Terms.AddTerm(
                    подключениеReference.ParameterGroup.OneToOneParameters.Find(Подключения_Сборка_Guid),
                    ComparisonOperator.IsOneOf,
                    обозначения);

                подключения = подключениеReference.Find(filter);
            }

            string[] дочерниеОбозначения = подключения.Select(
                    obj => obj[Подключения_Обозначение_Guid].GetString()
                        .Replace(".", string.Empty))
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
                if (!allObjects.ContainsKey(obj.SystemFields.Id))
                {
                    allObjects.Add(obj.SystemFields.Id, obj);
                    levelObjects.Add(obj);
                }
            }

            списокНоменклатуры.AddRange(levelObjects);
        }
    }

    private static string ВставитьТочки(string valueString)
    {
        int num;
        bool isNum = int.TryParse(valueString, out num);
        if (!isNum && valueString.IndexOf('-') > 0 && (valueString.Substring(0, valueString.IndexOf('-')).Length == 6 || valueString.Substring(0, valueString.IndexOf('-')).Length == 7))
        {
            isNum = int.TryParse(valueString.Substring(0, valueString.IndexOf('-')), out num);
            valueString = valueString.Insert(3, ".");
        }

        if (valueString.StartsWith("УЯИС") || valueString.StartsWith("ШЖИФ") || valueString.StartsWith("УЖИЯ"))
        {
            // Изменяем значение
            valueString = valueString.Insert(4, ".").Insert(11, ".");
        }
        
        if (valueString.StartsWith("3905")) {
        	valueString = valueString.Insert(4, ".");
        	//System.Windows.Forms.MessageBox.Show(valueString, ""); Для проверки корректности
        	return valueString;
        }

        if (valueString.StartsWith("8А"))
        {
            // Изменяем значение
            valueString = valueString.Insert(3, ".").Insert(7, ".");
        }

        if (isNum && (valueString.Length == 6 || valueString.Length == 7))
        {
            valueString = valueString.Insert(3, ".");
        }
        return valueString;
    }

}
