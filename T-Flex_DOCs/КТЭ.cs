using System;
using System.Collections.Generic;
using System.Linq;
using TFlex.DOCs.Model;
using TFlex.DOCs.Model.FilePreview.CADInteraction;
using TFlex.DOCs.Model.Macros;
using TFlex.DOCs.Model.References;
using TFlex.DOCs.Model.Search;
using TFlex.DOCs.Model.Structure;
using TFlex.DOCs.Technology.References;
using TFlex.Model.Technology.References;
using TFlex.Model.Technology.References.TechnologyElements;
using TFlex.Model.Technology.References.TechnologyElements.TechnologicalParameters;
using TFlex.Technology;

namespace TechnologyMacros
{
    public class KTE : MacroProvider
    {
        /// <summary>
        /// Системные имена типов технологических элементов, которые нужно получить с чертежа
        /// </summary>
        private static readonly string[] _technologicalElementTypes =
        {
           TechnologicalElement.BaseType,    // базовый элемент
           TechnologicalElement.ComplexType  // составной элемент
        };

        public KTE(MacroContext context) : base(context)
        {
        }

        public override void Run()
        {
        }

        public void ПолучитьКТЭ()
        {
            var process = (ReferenceObject)ТекущийОбъект as StructuredTechnologicalProcess;
            if (process == null)
                Error("Объект '{0}' не является технологическим процессом", ТекущийОбъект);

            if (process.SketchFile == null)
                Error("Невозможно получить КТЭ: не задан эскиз техпроцесса.");

            GetStructureTechnologicalElements(process);

            Message("КТЭ", "Получение КТЭ завершено.");
        }

        public void ОбновитьКТЭ()
        {
            var kte = (ReferenceObject)ТекущийОбъект as StructureTechnologicalElement;
            if (kte == null)
                Error("Объект '{0}' не является КТЭ", ТекущийОбъект);

            CADStructureElement structuredElement = GetCADStructureElement(kte.Process.TechnologicalProcessCADObjectReceiver, kte.SearchString);
            if (structuredElement == null)
                Error("Технологический элемент '{0}' не найден на чертеже '{1}'.", kte, kte.Process.SketchFile);

            // обновление КТЭ
            UpdateStructureTechnologicalElement(kte, kte.Process, structuredElement, false, false);

            // Обновление связанных КТЭ для составного
            if (kte.IsComplex)
            {
                foreach (StructureTechnologicalElement te in kte.LinkedTechnologicalElements)
                {
                    structuredElement = GetCADStructureElement(kte.Process.TechnologicalProcessCADObjectReceiver, te.SearchString);
                    if (structuredElement != null)
                    {
                        UpdateStructureTechnologicalElement(te, kte.Process, structuredElement, false, false);
                    }
                }
            }

            Message("КТЭ", "Обновление КТЭ '{0}' завершено.", kte.ToString());
        }

        /// <summary>
        /// Получить структурные элементы
        /// </summary>
        /// <param name="process">Технологический процесс</param>
        private void GetStructureTechnologicalElements(StructuredTechnologicalProcess process)
        {
            bool supportStructureElements;

            // получаем структурные элементы документа с указанными типами
            var structureElements = process.TechnologicalProcessCADObjectReceiver.GetCADStructureElementCollection(out supportStructureElements, _technologicalElementTypes);
            if (structureElements != null && structureElements.Count > 0)
            {
                if (!Question(String.Format("Найдено {0} КТЭ. Продолжить?", structureElements.Count)))
                    return;
            }
            else
            {
                if (supportStructureElements)
                    Error("Не найдены технологические элементы на чертеже:\r\n'{0}'", process.SketchFile);
                else
                    Error("Текущая версия CAD не поддерживает структурные элементы.");
            }

            bool canAddType = false; // можно ли создавать новые типы ТЭ в справочнике
            bool canEditType = false; // можно ли редактировать существующие типы ТЭ в справочнике

            if (Question("Создавать новые типы технологических элементов в справочнике 'Технологические элементы'?"))
                canAddType = true;

            WaitingDialog.Show("Получение КТЭ", true);

            int i = 1;
            foreach (CADStructureElement structuredElement in structureElements)
            {
                if (!WaitingDialog.NextStep(String.Format("{0} из {1}: КТЭ '{2}'", i++, structureElements.Count, structuredElement.Name)))
                    return;

                // Ищем в данном техпроцессе ктэ с таким же гуидом, как у структурного элемента
                StructureTechnologicalElement KTE = null;
                var kteCollection = process.GetTechnologicalElements(true).Where(kte => kte.SearchString == structuredElement.SearchString).ToList();
                if (kteCollection.Count > 1)
                {
                    string messageKTE = string.Empty;
                    foreach (StructureTechnologicalElement element in kteCollection)
                    {
                        messageKTE += String.Format("\r\n{0}; ID = {1}", element.Name, element.SearchString);
                    }
                    Message("Получение КТЭ", "В техпроцессе найдены КТЭ с одинаковыми идентификаторами:\r\n {0}", messageKTE);
                    continue;
                }
                else
                    KTE = kteCollection.FirstOrDefault();

                UpdateStructureTechnologicalElement(KTE, process, structuredElement, canAddType, canEditType);
            }

            WaitingDialog.Hide();
        }

        /// <summary>
        /// Обновить информацию о КТЭ
        /// </summary>
        /// <param name="KTE">КТЭ</param>
        /// <param name="process">Технологический процесс</param>
        /// <param name="structureElement">Структурный элемент</param>
        /// <param name="canAddType">Флаг добавления новых типов структурных элементов</param>
        /// <param name="canEditType">Флаг редактирования типов структурных элементов</param>
        private void UpdateStructureTechnologicalElement(StructureTechnologicalElement KTE, StructuredTechnologicalProcess process, CADStructureElement structureElement, bool canAddType, bool canEditType)
        {
            try
            {
                if (process.IsAdded)
                {
                    process.ApplyChanges();
                }

                if (KTE == null)
                {
                    // Создаем КТЭ в техпроцессе
                    KTE = process.Reference.CreateReferenceObject(process, process.Reference.Classes.ConstructiveTechnologicalElement) as StructureTechnologicalElement;
                    KTE.SearchString.Value = structureElement.SearchString;
                }
                else
                    KTE.BeginChanges();

                // устанавливаем иконку КТЭ из структурного типа CAD
                if (structureElement.StructureType.Icon != null)
                    KTE.IconCAD.Value = new IconImage(structureElement.StructureType.Icon);

                // %%TODO
                // определение ориентации КТЭ
                var orientationElementValue = structureElement.CADObjectInfo.ValueCollection.FirstOrDefault(
                    elementValue => elementValue.ParameterInfo.SystemName == TechnologicalElement.OrientationParameter);

                if (orientationElementValue != null)
                {
                    if (orientationElementValue.EnumValue != null)
                    {
                        var valueName = KTE.Orientation.ParameterInfo.ValueList.GetName(orientationElementValue.EnumValue.DefaultValueAsString) ?? String.Empty;
                        KTE.Orientation.Value = KTE.Orientation.ParameterInfo.ValueList.GetValue(valueName) as String ?? KTE.Orientation.ParameterInfo.DefaultValue.ToString();
                    }
                }

                //Ищем в ОТП соответствующий технологический элемент (по системному имени типа ТЭ и ориентации), если он еще не подключен
                StructureTechnologicalElement commonTechnologicalElement = null;
                try
                {
                    commonTechnologicalElement = KTE.Process.CommonTechnologicalProcess != null
                     ? KTE.Process.CommonTechnologicalProcess.GetTechnologicalElements(true).SingleOrDefault(te =>
                     te.TypeSystemName == structureElement.StructureType.TypeSystemName &&
                     te.Orientation == KTE.Orientation.Value)
                     : null;
                }
                catch (InvalidOperationException)
                {
                    Message("КТЭ", "Для КТЭ '{0}' невозможно однозначно установить технологический элемент ОТП", structureElement.Name);
                }

                //Подключаем соответствующий технологический элемент из ОТП или справочника,
                //при изменении связи копируются параметры из справочника 'Технологические элементы': макрос "Изменение связей ТП"
                //устанавливается ориентация
                //флаг "Составной"
                if (commonTechnologicalElement != null) // если есть технологический элемент в общем техпроцессе, подключаем его
                {
                    StructureTechnologicalElement kteCommonTechnologicalElement = KTE.CommonTechnologicalElement;
                    if (kteCommonTechnologicalElement != commonTechnologicalElement)
                    {
                        KTE.CommonTechnologicalElement = commonTechnologicalElement;
                        // Флаг "Составной" устанавливается из справочного элемента ТЭ ОТП
                        //KTE.IsComplex.Value = commonTechnologicalElement.TechnologicalElementType != null ? commonTechnologicalElement.TechnologicalElementType.IsComplex : false;
                    }
                }
                else // иначе подключаем справочный технологический элемент
                {
                    KTE.TechnologicalElementType = GetTechnologicalElementType(structureElement.StructureType, canAddType, canEditType);

                    // Флаг "Составной" устанавливается из справочного элемента
                    //if (technologicalElementType != null)
                    //    KTE.IsComplex.Value = technologicalElementType.IsComplex;
                }

                // устанавливаем наименование КТЭ после подключения ТЭ из ОТП из справочника типов
                //if (String.IsNullOrEmpty(structureElement.DisplayName))
                KTE.Name.Value = structureElement.CADObjectInfo.ToString(); //.DisplayName; //.CADObjectInfo.ToString();

                if (KTE.IsAdded) // для нового объекта нужно применить изменения, иначе не сохраняется значение параметра (parameterKTE.Expression.Value)
                {
                    KTE.ApplyChanges();
                }

                // Создаем параметры КТЭ
                foreach (var cadElementValue in structureElement.CADObjectInfo.ValueCollection)
                {
                    if (String.IsNullOrEmpty(cadElementValue.ParameterInfo.SystemName))
                        continue;
                    if (cadElementValue.ParameterInfo.SystemName == TechnologicalElement.IsInnerParameter) // пропускаем флаг "внутренний"
                        continue;
                    if (cadElementValue.ParameterInfo.SystemName == TechnologicalElement.OrientationParameter) // ориентацию установили ранее
                        continue;
                    if (cadElementValue.ParameterInfo.IsAuxiliary || cadElementValue.ParameterInfo.IsStatic)
                        continue;

                    CreateKTEParameterFromCADValue(KTE, cadElementValue);
                }

                // связанные с КТЭ структурные элементы могут быть параметрами, если они типа "Технологический размер"
                var parameterAsLinkCollection = structureElement.CADObjectInfo.LinkedObjectInfoCollection.Where(cadLinkedObjectInfo =>
                    !String.IsNullOrEmpty(cadLinkedObjectInfo.LinkedObjectSearchString) &&
                    cadLinkedObjectInfo.LinkedObjectGroupType == CADObjectTypes.StructureElement &&
                    cadLinkedObjectInfo.LinkedObjectTypeSystemName == CADStructureDimension.SystemName_Type);

                // Получить описания размеров и создать по ним параметры КТЭ
                List<CADStructureDimension> dimensionCollection = new List<CADStructureDimension>();
                // сформировать список идентификаторов связанных технологических размеров
                List<string> parameterIDCollection = parameterAsLinkCollection.Select<CADLinkedObjectInfo, string>(cadLinkedObjectInfo => cadLinkedObjectInfo.LinkedObjectSearchString).ToList();
                // получить список структурных элементов типа "Технологический размер" по их идентификаторам
                List<CADStructureElement> parameterObjectCollection = structureElement.Receiver.GetCADObjectCollection(parameterIDCollection.ToList()).OfType<CADStructureElement>().ToList();

                foreach (CADLinkedObjectInfo cadLinkedObjectInfo in parameterAsLinkCollection)
                {
                    CADStructureElement structureObject = parameterObjectCollection.FirstOrDefault(parameterObject => parameterObject.SearchString == cadLinkedObjectInfo.LinkedObjectSearchString);
                    if (structureObject != null)
                    {
                        CADStructureDimension dimension = new CADStructureDimension(cadLinkedObjectInfo.LinkDescriptor.LinkSystemName, cadLinkedObjectInfo.LinkDescriptor.LinkName, structureObject);
                        dimensionCollection.Add(dimension);
                    }
                }

                foreach (CADStructureDimension dimensionObject in dimensionCollection)
                {
                    CreateKTEParameterFromStructureDimension(KTE, dimensionObject);
                }


                KTE.ClearLinks(KTE.Links.ToMany.LinkGroups.Find(StructureTechnologicalElement.TechnologicalElementLinks.LinkedTechnologicalElements));
                // подключение связанных КТЭ (кроме тех, которые типа "Технологический размер")
                foreach (CADLinkedObjectInfo cadLinkedObjectInfo in structureElement.CADObjectInfo.LinkedObjectInfoCollection.Except(parameterAsLinkCollection))
                {
                    if (!String.IsNullOrEmpty(cadLinkedObjectInfo.LinkedObjectSearchString) && cadLinkedObjectInfo.LinkedObjectGroupType == CADObjectTypes.StructureElement)
                    {
                        StructureTechnologicalElement linkedKTE = null;
                        try
                        {
                            linkedKTE = KTE.Process.GetTechnologicalElements(true).SingleOrDefault(kte => kte.SearchString == cadLinkedObjectInfo.LinkedObjectSearchString);
                        }
                        catch (InvalidOperationException)
                        {
                            Message("КТЭ", "Для КТЭ '{0}' невозможно подключить КТЭ с идентификатором '{1}' - найдено несколько объектов.",
                                structureElement.Name,
                                cadLinkedObjectInfo.LinkedObjectSearchString);
                            continue;
                        }

                        if (linkedKTE != null && !KTE.LinkedTechnologicalElements.Contains(linkedKTE))
                        {
                            KTE.AddLinkedObject(KTE.TechnologicalElementsLink.LinkGroup, linkedKTE);
                        }
                    }
                }

                KTE.EndChanges();
            }
            catch
            {
                if (KTE != null && KTE.Changing)
                    KTE.CancelChanges();
                throw;
            }
        }

        /// <summary>
        /// Создать параметр КТЭ из описания связанного структурного элемента типа "Технологический размер"
        /// </summary>
        /// <param name="KTE">Объект КТЭ</param>
        /// <param name="dimensionObject">Описание структурного элемента типа "Технологический размер"</param>
        private void CreateKTEParameterFromStructureDimension(StructureTechnologicalElement KTE, CADStructureDimension dimensionObject)
        {
            Filter filter = new Filter(KTE.Parameters.ParameterGroup);
            filter.Terms.AddTerm(KTE.Parameters.ParameterGroup[TechnologicalParameter.TechnologicalParameterParameters.Name], ComparisonOperator.Equal, dimensionObject.Name);
            KTEParameter parameterKTE = KTE.Parameters.Find(filter).FirstOrDefault() as KTEParameter;

            try
            {
                if (parameterKTE == null)
                {
                    parameterKTE = KTE.Parameters.CreateReferenceObject((KTE.Process.Parameters.Classes as TechnologicalParametersClassTree).KTEParameter) as KTEParameter;
                    parameterKTE.Name.Value = dimensionObject.Name;
                    parameterKTE.ApplyChanges();
                }
                else
                    parameterKTE.BeginChanges();

                parameterKTE.IsMain.Value = true;
                parameterKTE.Expression.Value = String.Empty; // очистить выражение

                if (String.IsNullOrEmpty(parameterKTE.Description))
                    parameterKTE.Description.Value = dimensionObject.Description;

                if (!String.IsNullOrEmpty(dimensionObject.CADDimensionSearchString)) // явно указан размер
                {
                    CADDim dimension = KTE.Process.TechnologicalProcessCADValuesReceiver.GetCADDimension(dimensionObject.CADDimensionSearchString);
                    if (dimension != null)
                    {
                        if (!String.IsNullOrEmpty(dimension.Tolerance)) // с указанным квалитетом
                            parameterKTE.Expression.Value = String.Format("{0} + Доп[\"{1}\"]", dimension.Nominal.ToString(System.Globalization.CultureInfo.InvariantCulture.NumberFormat), dimension.Tolerance);
                        //parameterKTE.Expression.Value = String.Format("{0} + Доп[\"{1}\"]", dimension.Nominal.ToString().Replace(',', '.'), dimension.Tolerance);
                        else // с указанными отклонениями
                            parameterKTE.Expression.Value = String.Format("{0} + Доп[{1}, {2}]", dimension.Nominal.ToString(System.Globalization.CultureInfo.InvariantCulture.NumberFormat), dimension.DevHigh.ToString("G", System.Globalization.CultureInfo.InvariantCulture.NumberFormat), dimension.DevLow.ToString("G", System.Globalization.CultureInfo.InvariantCulture.NumberFormat));
                        //parameterKTE.Expression.Value = String.Format("{0} + Доп[{1}, {2}]", dimension.Nominal.ToString().Replace(',', '.'), dimension.DevHigh.ToString().Replace(',', '.'), dimension.DevLow.ToString().Replace(',', '.'));
                    }
                }
                else
                {
                    string expression = string.Empty;
                    if (!String.IsNullOrEmpty(dimensionObject.VariableName))
                        expression = CreateParameterExpression(dimensionObject.VariableName);
                    else
                        expression = CreateParameterExpression(dimensionObject.Expression);

                    parameterKTE.Expression.Value = expression;

                    if (parameterKTE.Qualitet <= 0)
                    {
                        parameterKTE.Litera = dimensionObject.Litera;
                        parameterKTE.Qualitet = dimensionObject.Qualitet;
                    }

                    if (parameterKTE.Qualitet <= 0)
                    {
                        parameterKTE.LowerAllowance = dimensionObject.Lower;
                        parameterKTE.UpperAllowance = dimensionObject.Upper;
                    }
                }

                if (!String.IsNullOrEmpty(dimensionObject.CADRoughnessSearchString)) // явно указана шероховатость
                {
                    CADRoughness roughness = KTE.Process.TechnologicalProcessCADValuesReceiver.GetCADRoughness(dimensionObject.CADRoughnessSearchString);
                    if (roughness != null)
                        parameterKTE.RoughnessDescription = String.Format("{0}{1}", roughness.RoughnessHeightParameterType, roughness.MaximumOrNominal);
                }
                else
                    parameterKTE.RoughnessClass = dimensionObject.Class;

                parameterKTE.EndChanges();
            }
            catch
            {
                if (parameterKTE != null && parameterKTE.Changing)
                    parameterKTE.CancelChanges();
                throw;
            }
        }

        /// <summary>
        /// Создать параметр КТЭ из описания значения структурного элемента
        /// </summary>
        /// <param name="KTE">Объект КТЭ</param>
        /// <param name="cadElementValue">Описание значения (свойства) структурного элемента</param>
        private void CreateKTEParameterFromCADValue(StructureTechnologicalElement KTE, CADElementValue cadElementValue)
        {
            Filter filter = new Filter(KTE.Parameters.ParameterGroup);
            filter.Terms.AddTerm(KTE.Parameters.ParameterGroup[TechnologicalParameter.TechnologicalParameterParameters.Name], ComparisonOperator.Equal, cadElementValue.ParameterInfo.SystemName);
            KTEParameter parameterKTE = KTE.Parameters.Find(filter).FirstOrDefault() as KTEParameter;

            try
            {
                if (parameterKTE == null)
                {
                    parameterKTE = KTE.Parameters.CreateReferenceObject((KTE.Process.Parameters.Classes as TechnologicalParametersClassTree).KTEParameter) as KTEParameter;
                    parameterKTE.Name.Value = cadElementValue.ParameterInfo.SystemName;
                    parameterKTE.ApplyChanges();
                }
                else
                    parameterKTE.BeginChanges();

                if (String.IsNullOrEmpty(parameterKTE.Description))
                    parameterKTE.Description.Value = cadElementValue.ParameterInfo.DisplayName;

                string expression = string.Empty;
                if (cadElementValue.AttachedVariable != null)
                {
                    if (!String.IsNullOrEmpty(cadElementValue.AttachedVariable.Name))
                        expression = CreateParameterExpression(cadElementValue.AttachedVariable.Name);
                    else
                        expression = CreateParameterExpression(cadElementValue.AttachedVariable.Expression);
                }
                else
                {
                    if (cadElementValue.EnumValue == null) // %%TODO
                    {
                        expression = CreateParameterExpression(cadElementValue.ValueAsString);
                    }
                }
                if (parameterKTE.Expression.Value != expression)
                    parameterKTE.Expression.Value = expression;

                parameterKTE.EndChanges();
            }
            catch
            {
                if (parameterKTE != null && parameterKTE.Changing)
                    parameterKTE.CancelChanges();
                throw;
            }
        }

        private string CreateParameterExpression(string cadExpression)
        {
            string expression = cadExpression;

            if (!String.IsNullOrEmpty(cadExpression))
            {
                string[] expressionMembers = cadExpression.Split(new char[] { '+', '-', '*', '/' }, StringSplitOptions.RemoveEmptyEntries);
                foreach (string member in expressionMembers)
                {
                    double memberValue = 0;

                    if (Double.TryParse(member, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture.NumberFormat, out memberValue))
                    {
                        expression = expression.Replace(member, memberValue.ToString(System.Globalization.CultureInfo.InvariantCulture.NumberFormat));
                        continue;
                    }
                    expression = expression.Replace(member, String.Format("Перем[\"{0}\",\"ЭскизТП\"]", member));
                }
            }

            return expression;
        }

        /// <summary>
        /// Получить тип технологического элемента из справочника "Технологические элементы"
        /// </summary>
        /// <param name="structureType">Структурный тип CAD</param>
        /// <param name="canAddType">Флаг создания нового типа ТЭ, если он не найден</param>
        /// <param name="canEditType">Флаг редактирования существующего типа ТЭ</param>
        /// <returns>Созданный тип технологического элемента</returns>
        private TechnologicalElement GetTechnologicalElementType(CADStructureType structureType, bool canAddType, bool canEditType)
        {
            if (String.IsNullOrEmpty(structureType.TypeSystemName))
                return null;

            TechnologicalElementsReference technologicalElementsReference = null;

            var info = ServerGateway.Connection.ReferenceCatalog.Find(TechnologyReferences.TechnologicalElements);
            if (info != null && info.ActivityStatus == GroupActivityStatus.Normal)
            {
                technologicalElementsReference = info.CreateReference() as TechnologicalElementsReference;
            }

            if (technologicalElementsReference == null)
                return null;

            TechnologicalElement technologicalElementType = technologicalElementsReference.FindTechnologicalElementType(structureType.TypeSystemName);

            if (!canAddType && !canEditType) // нельзя ни добавлять, ни редактировать технологические элементы
            {
                return technologicalElementType;
            }

            // определение родительского технологического элемента
            TechnologicalElement parentTechnologicalElement = null;
            if (structureType.CADObjectInfo.ParentElementType != null)
            {
                parentTechnologicalElement = technologicalElementsReference.FindTechnologicalElementType(structureType.CADObjectInfo.ParentElementType.SystemName);
            }

            try
            {
                if (technologicalElementType == null) // тип технологического элемента не найден
                {
                    if (canAddType)
                    {
                        // создаем новый тип технологического элемента
                        technologicalElementType = technologicalElementsReference.CreateReferenceObject(parentTechnologicalElement, technologicalElementsReference.Classes.TechnologicalElement) as TechnologicalElement;
                        technologicalElementType.Name.Value = structureType.TypeName;
                        technologicalElementType.TypeSystemName.Value = structureType.TypeSystemName;
                    }
                }
                else
                {
                    if (canEditType)
                    {
                        technologicalElementType.BeginChanges();
                        if (parentTechnologicalElement != null)
                            technologicalElementType.SetParent(parentTechnologicalElement);
                    }
                }

                if (technologicalElementType != null && (technologicalElementType.Changing || technologicalElementType.IsAdded))
                {
                    technologicalElementType.Description.Value = structureType.TypeDescription;

                    // устанавливаем иконку технологического элемента из структурного типа CAD
                    if (structureType.Icon != null)
                        technologicalElementType.IconCAD.Value = new IconImage(structureType.Icon);

                    // флаг "Составной" устанавливается по системному имени типа "Составной элемент", заданного в редакторе структурных элементов в CAD
                    technologicalElementType.IsComplex.Value = structureType.CADObjectInfo.GetHierarchyTypes().Exists(elementType => elementType.SystemName == TechnologicalElement.ComplexType);

                    // флаг абстрактный
                    technologicalElementType.IsAbstract.Value = structureType.CADObjectInfo.IsAbstract;

                    // Создаем список параметров из описания типа структурного элемента
                    foreach (CADElementTypeValue cadTypeValue in structureType.TypeValueCollection)
                    {
                        if (String.IsNullOrEmpty(cadTypeValue.ParameterInfo.SystemName))
                            continue;

                        switch (cadTypeValue.ParameterInfo.SystemName)
                        {
                            case TechnologicalElement.OrientationParameter:
                                continue;

                            case TechnologicalElement.IsInnerParameter:
                                if (cadTypeValue.ParameterInfo.IsStatic)
                                    technologicalElementType.IsInner.Value = (bool)cadTypeValue.DefaultValueAsObject;
                                continue;

                            default:
                                if (cadTypeValue.ParameterInfo.IsAuxiliary)
                                    continue;

                                if (cadTypeValue.DefaultEnumValue != null)
                                    continue;

                                // создаем / обновляем параметр
                                technologicalElementType.CreateTEParameter(cadTypeValue.ParameterInfo.SystemName, cadTypeValue.ParameterInfo.DisplayName);

                                continue;
                        }
                    }

                    technologicalElementType.EndChanges();
                }
            }
            catch
            {
                if (technologicalElementType != null && technologicalElementType.Changing)
                    technologicalElementType.CancelChanges();
                throw;
            }

            return technologicalElementType;
        }

        /// <summary>
        /// Получить структурный элемент CAD'а по идентификатору  
        /// </summary>
        /// <param name="receiver"></param>
        /// <param name="searchString">Идентификатор структурного элемента</param>
        /// <returns>Структурный элемент</returns>
        private CADStructureElement GetCADStructureElement(CADObjectReceiver receiver, string searchString)
        {
            if (receiver == null)
                return null;

            if (String.IsNullOrEmpty(searchString))
                return null;

            TFlex.Model.Technology.References.CADObject cadObject = receiver.GetCADObjectBySearchString(searchString);
            CADStructureElement structureElement = null;

            try
            {
                structureElement = (CADStructureElement)cadObject;
            }
            catch (InvalidCastException)
            {
                Message("КТЭ", "Объект '{0}' (ID '{1}', тип '{2}') не является структурным элементом CAD.", cadObject.DisplayName, searchString, cadObject.CADObjectInfo.SubType);
            }

            return structureElement;
        }
    }
}


