using System;
using System.Collections.Generic;
using System.Linq;
using TFlex.DOCs.Model.Macros;
using TFlex.DOCs.Model.References;
using TFlex.DOCs.Model.Search;
using TFlex.DOCs.Model.Structure;
using TFlex.DOCs.Model.Structure.Builders;

public class MacroMDMGol : MacroProvider
{
    public MacroMDMGol(MacroContext context)
        : base(context)
    {
        if (Context.Connection.ClientView.HostName == "MOSINS")
            if (Вопрос("Хотите запустить в режиме отладки?"))
            {
                System.Diagnostics.Debugger.Launch();
                System.Diagnostics.Debugger.Break();
            }
    }


    public override void Run()
    {
    }

    /// <summary>
    /// Справочник "Глоссарий параметров" - Событие "Сохранение объекта"
    /// парамытр 
    /// </summary>
    public void СохранениеГлоссария()
    {
        var glossaryObj = Context.ReferenceObject;
        if (glossaryObj == null)
            return;

        string glossaryName = glossaryObj[Parameters.Name].GetString();
        if (String.IsNullOrWhiteSpace(glossaryName))
            Ошибка("Незаполнен параметр 'Наименование'");

        var glossaryExtendedParameterGuid = glossaryObj[Parameters.ExtendedParameterGuid].GetGuid();
        if (glossaryExtendedParameterGuid == Guid.Empty)
            Ошибка("Не запонен параметр 'GUID доп. параметра'. Объект не связан с дополнительным параметром");

        var glossaryReference = Context.Reference;
        Filter filter = CreateFilter(glossaryReference.ParameterGroup, glossaryObj, new Guid[] { Parameters.Name });
        var objects = glossaryReference.Find(filter);

        if (objects.Count != 0)
            Ошибка($"В справочнике глоссарий уже есть объект с наименованием: {glossaryName}");

        var foundExtendedParameter = FindExtendedParameter(glossaryExtendedParameterGuid);
        if (foundExtendedParameter == null)
            Ошибка("В параметре 'Guid допоплнительного параметра' указан несуществующий в системе параметр");

        CopyParametersInGlossaryObj(glossaryObj, foundExtendedParameter);
    }

    private Filter CreateFilter(ParameterGroup parameterGroup, ReferenceObject glossaryObj, Guid[] parameters)
    {
        if (parameterGroup == null)
            throw new ArgumentNullException(nameof(parameterGroup));

        if (glossaryObj == null)
            throw new ArgumentNullException(nameof(glossaryObj));

        if (parameters.Length == 0)
            return null;

        var filter = new Filter(parameterGroup);
        foreach (var parameterGuid in parameters)
        {
            var parameterInfo = FindParameterInfo(parameterGroup, parameterGuid);
            if (parameterInfo == null)
                continue;

            filter.Terms.AddTerm(parameterInfo, ComparisonOperator.Equal, glossaryObj[parameterInfo].Value);
        }

        filter.Terms.AddTerm("ID", ComparisonOperator.NotEqual, glossaryObj.SystemFields.Id);

        return filter;
    }

    /// <summary>
    /// Показывает диалог по выбору доступных дополнительных параметров в системе, написывает указаный параметрв к глоссарию.
    /// Диалог свойств справочника Глосарий
    /// </summary>
    public void ПодключитьДопПараметр()
    {
        var extendedParameters = GetExtendenParameters();
        if (extendedParameters.Count == 0)
            Error("В системе нет ни одного дополнительного параметра");

        CreateInputDialog("Выберите дополнительный параметр");
        var dialog = CreateInputDialog("Выберите дополнительный параметр");
        dialog.AddSelectFromList("Дополнительный параметр", String.Empty, true, extendedParameters.Select(p => p.Name).ToArray());
        if (!dialog.Show())
            return;

        string selectedParameterName = dialog["Дополнительный параметр"];
        var foundParameterInfo = extendedParameters.Find(p => p.Name == selectedParameterName);
        if (foundParameterInfo == null)
            Error("Ошибка поиска параметра");

        CopyParametersInGlossaryObj(Context.ReferenceObject, foundParameterInfo);

        Message("Выполнено", "Дополнительный параметр был подключен");
    }

    /// <summary>
    /// Находит или создает доп параметр при подключении к объекту глассария
    /// Диалог свойств справочника Глосарий
    /// </summary>
    public void СоздатьДопПараметр()
    {
        var glossaryObj = Context.ReferenceObject;
        string glossaryName = glossaryObj[Parameters.Name].GetString();
        if (String.IsNullOrWhiteSpace(glossaryName))
            Error("Не задано наименование параметра");

        int glossaryType = glossaryObj[Parameters.ParameterType].GetInt32();
        ParameterType parameterType = ParameterType.GetType(glossaryType);
        if (parameterType == null)
            Error("Указан некорректный тип параметра");

        var foundExtendedParameter = GetExtendenParameters().Find(p => p.Name == glossaryName && p.Type == parameterType);
        if (foundExtendedParameter != null)
        {
            if (!Question($"В системе найден дополнительный параметр с указанными свойствами{Environment.NewLine}Хотите его подключить?"))
                return;

            CopyParametersInGlossaryObj(glossaryObj, foundExtendedParameter);
            Message("Выполнено", "Дополнительный параметр был подключен");
        }
        else
        {
            foundExtendedParameter = СоздатьДополнительныйПараметрИПсевдоним(Context.Reference, parameterType, glossaryName);
            CopyParametersInGlossaryObj(glossaryObj, foundExtendedParameter);
            Message("Выполнено", "Дополнительный параметр был создан и подключен");
        }
    }

    /// <summary>
    /// Создает дополнительный параметр с указанными свойствами для указанного справочника
    /// </summary>
    /// <param name="parameterGroup"></param>
    /// <param name="parameterType"></param>
    /// <param name="extendedParameterName"></param>
    /// <returns></returns>
    private ParameterInfo СоздатьДополнительныйПараметрИПсевдоним(Reference reference, ParameterType parameterType, string extendedParameterName)
    {
        var baseClass = reference.ParameterGroup.Classes.BaseClasses[0];
        ExtendedParameterInfoBuilder extendedParameterInfoBuilder = new ExtendedParameterInfoBuilder(Context.Connection, reference.ParameterGroup, baseClass)
        {
            Name = extendedParameterName,
            Type = parameterType
        };
        extendedParameterInfoBuilder.Save();
        return extendedParameterInfoBuilder.ParameterInfo;
    }

    private void CopyParametersInGlossaryObj(ReferenceObject glossaryObj, ParameterInfo extendedParameterInfo)
    {
        glossaryObj[Parameters.Name].Value = extendedParameterInfo.Name;
        glossaryObj[Parameters.ExtendedParameterGuid].Value = extendedParameterInfo.Guid;
        glossaryObj[Parameters.ParameterType].Value = extendedParameterInfo.Type.Id;
    }

    private List<ParameterInfo> GetExtendenParameters() => Context.Connection.ExtendedParameters.GetExtendedParameters().ToList();

    private ParameterInfo FindExtendedParameter(Guid extendedParamererGuid) => GetExtendenParameters().Find(p => p.Guid == extendedParamererGuid);

    private ParameterInfo FindParameterInfo(ParameterGroup parameterGroup, Guid guidParameter) => parameterGroup.OneToOneParameters.Find(guidParameter);

    private class Parameters
    {
        /// <summary>
        /// Guid доп. параметра
        /// </summary>
        public static readonly Guid ExtendedParameterGuid = new Guid("ebce6c07-b09b-43c0-a2b7-7e95fbabc8eb");

        /// <summary>
        /// Тип параметра
        /// </summary>
        public static readonly Guid ParameterType = new Guid("b22713d5-5a15-4f62-bb7b-4faf9e8aec27");

        /// <summary>
        /// Тип единицы измерения
        /// </summary>
        public static readonly Guid UnitType = new Guid("edf0ddc9-88b9-43b1-bf64-8ff20488192a");

        /// <summary>
        /// Наименование
        /// </summary>
        public static readonly Guid Name = new Guid("9ae2b9b2-9e66-4088-a83c-d8f52c1969db");

    }
}
