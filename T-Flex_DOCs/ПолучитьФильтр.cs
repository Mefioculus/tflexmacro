using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using TFlex.DOCs.Model.Macros;
using TFlex.DOCs.Model.Macros.ObjectModel;

public class MacroFilter : MacroProvider
{
    public MacroFilter(MacroContext context)
        : base(context)
    {
        if (Вопрос("Хотите запустить в режиме отладки?"))
        {
            System.Diagnostics.Debugger.Launch();
            System.Diagnostics.Debugger.Break();
        }
    }

    public override void Run()
    {
        ПолучитьФильтрНоменклатурыСборка();
    }

    public List<int> ПолучитьФильтрНоменклатурыСборка()
    {
        return GetFiltersList(_specReferenceGuid, _IZDParameter, _linkNomSborka, _linkNomParameterDenotation);
    }

    public List<int> ПолучитьФильтрНоменклатуры()
    {
        return GetFiltersList(_specReferenceGuid, _shifrParameter, _linkNomSborka, _linkNomParameterDenotation);
    }

    private Guid _specReferenceGuid = new Guid("c587b002-3be2-46de-861b-c551ed92c4c1");
    private Guid _shifrParameter = new Guid("2a855e3d-b00a-419f-bf6f-7f113c4d62a0");
    private Guid _IZDParameter = new Guid("817eca21-1e7e-46e0-b80a-e61685bef5f7");

    private Guid _linkNomSborka = new Guid("37cf3b77-7d38-4a82-97f6-d86766c5aef1");
    private Guid _linkNomParameterDenotation = new Guid("1bbb1d78-6e30-40b4-acde-4fc844477200");

    private Guid _linkNom = new Guid("46e88851-3f3f-4a21-acfa-ac3297baf408");
    //private Guid _linkNomDenotation = new Guid("1bbb1d78-6e30-40b4-acde-4fc844477200");



    public List<int> GetFiltersList(Guid referenceGuid, Guid referenceParameter, Guid objectLink, Guid objectParameter)
    {
        var collection = new List<int>();

        var specReferenceInfo = Context.Connection.ReferenceCatalog.Find(referenceGuid);//справочник spec
        if (specReferenceInfo == null)
            return collection;

        var specReference = specReferenceInfo.CreateReference();
        specReference.LoadSettings.Add(referenceParameter); // shifr

        //Загружем связанныеОбъекеты;
        var linkSetting = specReference.LoadSettings.AddRelation(objectLink);//Список номенклатуры
        linkSetting.Add(objectParameter);//Обозначение

        var objects = specReference.Objects;
        foreach (var obj in objects)
        {
            var linkedObject = obj.GetObject(objectLink);
            if (linkedObject == null)
                continue;

            var linkedValue = linkedObject[objectParameter].GetString();
            var objectValue = obj[referenceParameter].GetString();
            if (linkedValue == objectValue)
                continue;

            if (GetValue(objectValue) != linkedValue)
                collection.Add(obj.SystemFields.Id);
        }
        return collection;
    }
    private string GetValue(string value)
    {
        int num;
        bool isNum = int.TryParse(value, out num);

        if (value.StartsWith("УЯИС") || value.StartsWith("ШЖИФ") || value.StartsWith("УЖИЯ"))
        {
            // Изменяем значение
            return value.Insert(4, ".").Insert(11, ".");
        }

        if (value.StartsWith("8А"))
        {
            // Изменяем значение
            return value.Insert(3, ".").Insert(7, ".");
        }

        if (isNum && (value.Length == 6 || value.Length == 7))
        {

            return value.Insert(3, ".");
        }

        return value;
    }
}
