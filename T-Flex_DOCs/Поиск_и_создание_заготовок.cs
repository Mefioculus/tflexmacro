using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using TFlex.DOCs.Model;
using TFlex.DOCs.Model.Classes;
using TFlex.DOCs.Model.FilePreview.CADInteraction;
using TFlex.DOCs.Model.Macros;
using TFlex.DOCs.Model.Macros.ObjectModel;
using TFlex.DOCs.Model.References;
using TFlex.DOCs.Model.References.Documents;
using TFlex.DOCs.Model.References.Files;
using TFlex.DOCs.Model.References.Links;
using TFlex.DOCs.Model.References.Materials;
using TFlex.DOCs.Model.References.Nomenclature;
using TFlex.DOCs.Model.Search;
using TFlex.DOCs.References.MaterialMark;
using TFlex.Model.Technology.References.Computations;
using TFlex.Model.Technology.References.ParametersProvider;
using TFlex.Model.Technology.References.TechnologyElements;
using TFlex.Model.Technology.References.TechnologyElements.TechnologicalParameters;

namespace Macros
{
    public class PieceMacroProvider : MacroProvider
    {
        // гуид списка объекта "Список условий"
        private Guid _termListGuid = new Guid("39c6e216-0d25-4e73-9acc-9564b810a073");
        //гуид параметра "Фильтр" (строка)
        private Guid _filterStringParameterGuid = new Guid("6b34b23a-4386-4cc2-82da-9fb64488b832");

        // гуид типа "Условие"
        private Guid _termTypeGuid = new Guid("35a91044-281c-4525-a023-26ffe6e42650");

        // гуид параметра "Параметр заготовки"
        private Guid _termNameParameterGuid = new Guid("9c60c612-d07e-412f-843b-a16245498d41");
        // гуид параметра "Оператор"
        private Guid _termOperatorParameterGuid = new Guid("a4400e00-0937-4651-bb7b-8071ddf63a58");
        // гуид параметра "Значение"
        private Guid _termValueParameterGuid = new Guid("b16eb22f-a4fa-417d-b010-38ca57db86be");

        // Расчёт "Поиск заготовок - Найти по параметрам"
        private Guid calcFindPieces = new Guid("6f874a08-3240-480b-888b-312035dadc4a");
        // Расчёт "Поиск заготовок - Найти все"
        private Guid calcFindAllPieces = new Guid("708af59c-dae2-425d-8661-1e99aef8265e");

        // связи Заготовка -> Типы заготовок
        private Guid DigitalStructureToPieceTypesRelation = new Guid("576c970b-21de-40d6-9b5a-442f473cad8b");
        
        // связь Справочный параметр заготовки номенклатуры
        private Guid DictionaryParameterOfPieceRelation = new Guid("e6995298-01ab-44f0-8cf0-89eff7635c13");
        
        public PieceMacroProvider(MacroContext context)
            : base(context)
        {
        }

        public override void Run()
        {
        }

        public void СозданиеВыбораЗаготовки()
        {
            PieceSelectionObject pieceSelectionObject = (PieceSelectionObject)ТекущийОбъект;
            LinkInfo linkInfo = pieceSelectionObject.Reference.LinkInfo;

            double dseMass = 0;
            var dse = linkInfo != null ? linkInfo.MasterObject as MaterialObject : null;
            if (dse != null)
            {
                dseMass = dse.Mass;

                MaterialObject materialDSE = dse.Children.FirstOrDefault(child => child.Class.IsInherit(NomenclatureTypes.Keys.Material)) as MaterialObject;
                if (materialDSE != null)
                {
                    TFlex.DOCs.Model.References.Materials.MaterialReferenceObject materialReferenceObject = materialDSE.LinkedObject as TFlex.DOCs.Model.References.Materials.MaterialReferenceObject;
                    if (materialReferenceObject != null)
                    {
                        pieceSelectionObject.PieceMaterialLink.SetLinkedObject(materialReferenceObject);

                        ReferenceObject materialAssortment = materialReferenceObject.Links.ToOne[new Guid("09aa0bd3-c7a1-4a94-8ee6-b911748cfd32")].LinkedObject;
                        if (materialAssortment != null)
                        {
                            Объект типЗаготовки = НайтиОбъект("91dd3751-4630-4a22-b008-9301f13befbf", "d9fa477b-da55-4b8a-b583-c98fffa92e6e", materialAssortment.Class.Guid);
                            if (типЗаготовки != null)
                            {
                                pieceSelectionObject.Links.ToOne[new Guid("56f9676a-5ca2-4f6c-b0d4-f1e06f823173")].SetLinkedObject((ReferenceObject)типЗаготовки);
                            }
                        }
                    }
                }
            }

            IEnumerable<TechnologicalParameter> technologicalParameters = pieceSelectionObject.Parameters.Objects.OfType<TechnologicalParameter>();
            if (!technologicalParameters.Any())
                return;

            if (pieceSelectionObject.SketchFile != null)
            {
                // Установить значения технологических параметров с эскиза детали

                CADVar[] cadVariableCollection = (pieceSelectionObject as ITechnologicalObjectSketchProvider).CADValuesReceiver.CADVariables;
                foreach (TechnologicalParameter parameter in technologicalParameters)
                {
                    Объект справочныйПараметр = НайтиОбъект("a787c3db-c9c5-4a53-8ce7-07e923d08481", "304f2d60-fa49-401e-b59d-8c98ffb8d21e", parameter.Name);
                    if (справочныйПараметр != null)
                    {
                        string наименованиеПеременнойCAD = справочныйПараметр["4587058a-866c-4c1d-81c4-ef2bba8b2f91"];
                        if (!String.IsNullOrEmpty(наименованиеПеременнойCAD))
                        {
                            CADVar variable = cadVariableCollection.FirstOrDefault(cadVariable => cadVariable.Name == наименованиеПеременнойCAD);
                            if (variable == null)
                                continue;

                            string expression = String.Format("Перем[\"{0}\",\"{1}\"]", variable.Name, "ЭскизВЗ");

                            if (parameter.Expression != expression)
                            {
                                parameter.BeginChanges();
                                parameter.Expression.Value = expression;
                                parameter.EndChanges();
                            }
                        }
                    }
                }
            }

            TechnologicalParameter massParameter = technologicalParameters.FirstOrDefault(parameter => parameter.Name == "Масса");
            if (dseMass > 0)
            {
                if (massParameter != null && massParameter.Nominal != dseMass)
                {
                    massParameter.BeginChanges();
                    massParameter.Expression.Value = dse.Mass.Value.ToString();
                    massParameter.EndChanges();
                }
            }
        }

        public void НайтиВсеЗаготовки()
        {
            ReferenceInfo referenceInfo = Context.Connection.ReferenceCatalog.Find(new Guid("853d0f07-9632-42dd-bc7a-d91eae4b8e83"));
            Filter filter = new Filter(referenceInfo);
            ClassObject nomenclaturePieceClass = referenceInfo.Classes.Find(new Guid("7a28bc48-2671-46f5-803f-e2f77840686c"));
            filter.Terms.AddTerm("[Тип]", ComparisonOperator.Equal, nomenclaturePieceClass);

            NomenclatureObject piece = ВыбратьЗаготовкуИзСписка(referenceInfo.CreateReference(), filter);
            if (piece != null)
                ПодключитьЗаготовкуКВыборуЗаготовки(piece);
        }

        public void НайтиЗаготовкуПоПараметрам()
        {
            // Расчёт поиска заготовок 
            ComputationObject computation = TFlex.Model.Technology.References.Computations.ComputationsReference.Instance.Find(calcFindPieces) as ComputationObject;
            if (computation == null)
                Ошибка("Не найден расчёт: {0}.", calcFindPieces);

            PieceSelectionObject pieceSelectionObject = (PieceSelectionObject)ТекущийОбъект;

            // тип заготовки в Выборе заготовки
            ReferenceObject pieceType = pieceSelectionObject != null ? pieceSelectionObject.PieceTypeObject : null;
            // материал заготовки в Выборе заготовки
            MaterialReferenceObject pieceMaterial = pieceSelectionObject != null ? pieceSelectionObject.PieceMaterial : null;

            ReferenceInfo referenceInfo = Context.Connection.ReferenceCatalog.Find(new Guid("853d0f07-9632-42dd-bc7a-d91eae4b8e83"));
            Filter filter = new Filter(referenceInfo);
            ClassObject nomenclaturePieceClass = referenceInfo.Classes.Find(new Guid("7a28bc48-2671-46f5-803f-e2f77840686c"));
            filter.Terms.AddTerm("[Тип]", ComparisonOperator.Equal, nomenclaturePieceClass);

            // по типу заготовки
            if ((bool)pieceSelectionObject[new Guid("36c2a889-b89b-47a4-8f88-c32cda091ad7")])
            {
                if (pieceType != null)
                    filter.Terms.AddTerm("[Тип заготовки]->[Guid]", ComparisonOperator.Equal, pieceType.SystemFields.Guid);
            }

            if (pieceMaterial != null)
            {
                // по марке
                if ((bool)pieceSelectionObject[new Guid("ec212aa7-8ffd-4294-9d80-0107df29e036")])
                {
                    AbstractMarkReferenceObject materialMark = pieceMaterial.MaterialMark;
                    if (materialMark != null)
                        filter.Terms.AddTerm("[Материал заготовки]->[Марка материала]->[Guid]", ComparisonOperator.Equal, materialMark.SystemFields.Guid);
                }
                // по точному сортаменту
                if ((bool)pieceSelectionObject[new Guid("e957349a-baff-4266-a1e5-00d9c4f19075")])
                {
                    ReferenceObject materialAssortment = pieceMaterial.Links.ToOne[new Guid("09aa0bd3-c7a1-4a94-8ee6-b911748cfd32")].LinkedObject;
                    if (materialAssortment != null)
                        filter.Terms.AddTerm("[Материал заготовки]->[Сортамент]->[Guid]", ComparisonOperator.Equal, materialAssortment.SystemFields.Guid);
                }
            }

            List<Условие> terms = new List<Условие>();
            IEnumerable<PieceSelectionParameter> pieceSelectionParameters = pieceSelectionObject.Parameters.Objects.OfType<PieceSelectionParameter>();
            if (pieceSelectionParameters.Any())
            {
                foreach (PieceSelectionParameter pieceSelectionParameter in pieceSelectionParameters)
                {
                    if (!pieceSelectionParameter.IsSearchParameter)
                        continue;

                    ReferenceObject linkedParameterDescription = pieceSelectionParameter.DescriptionPieceParameter;
                    if (linkedParameterDescription == null)
                        continue;

                    string name = linkedParameterDescription[new Guid("4587058a-866c-4c1d-81c4-ef2bba8b2f91")].Value.ToString(); // наименование переменной CAD
                    string comparsionOperator = pieceSelectionParameter[new Guid("6d67fda5-4a48-4712-bff7-f9feb43b1b29")].Value.ToString();
                    terms.Add(new TFlex.DOCs.Model.Macros.ObjectModel.Условие(name, comparsionOperator, pieceSelectionParameter.Value));
                }
            }

            ReferenceObject[] pieces = (computation.Run(this.Context, filter.ToString(), terms.ToArray()) as IEnumerable<ReferenceObject>).ToArray();

            NomenclatureObject piece = null;
            if (pieces.Any())
            {
                IInputDialog inputDialog = Context.CreateInputDialog();
                inputDialog.AddMultiselectFromList("Pieces", pieces, false);
                if (inputDialog.Show(Context)) 
                    piece = inputDialog.GetValue("Pieces");

            }
            else  // заготовки не найдены
            {
                Сообщение("Поиск заготовок", "Заготовки с указанными параметрами не найдены.");
                piece = СоздатьЗаготовкуВНоменклатуре();
            }

            if (piece is null)
                return;
            
            ПодключитьЗаготовкуКВыборуЗаготовки(piece);
        }

        public void CreateHierarchyLink()
        {
        	Debugger.Launch();
            Debugger.Break();
        	
            PieceSelectionObject pieceSelectionObject = (PieceSelectionObject)ТекущийОбъект;
            if (!pieceSelectionObject.IsNew)
                return;

            LinkInfo linkInfo = pieceSelectionObject.Reference.LinkInfo;
            var materialObject = linkInfo?.MasterObject as MaterialObject;
            var hierarchyLink = materialObject.CreateChildLink(pieceSelectionObject.PieceNomenclatureObject) as NomenclatureHierarchyLink;
            hierarchyLink.EndChanges();
        }

        private NomenclatureObject ВыбратьЗаготовкуИзСписка(Reference reference, Filter filter)
        {
            ISelectObjectDialog selectObjectDialog = Context.CreateSelectObjectDialog(reference);
            selectObjectDialog.Filter = filter;
            if (!selectObjectDialog.Show())
                return null;

            return selectObjectDialog.SelectedObjects.FirstOrDefault() as NomenclatureObject;
        }

        /// <summary>
        /// Создание заготовки в номенклатуре
        /// </summary>
        /// <returns>Созданная заготовка</returns>
        public NomenclatureObject СоздатьЗаготовкуВНоменклатуре()
        {
            if (!Вопрос("Создать заготовку?"))
            {
                return null;
            }

            PieceSelectionObject pieceSelectionObject = (PieceSelectionObject)ТекущийОбъект;

            // тип заготовки
            ReferenceObject pieceType = pieceSelectionObject.PieceTypeObject;
            if (pieceType == null)
            {
                Ошибка("Не указан тип заготовки.");
            }

            // Шаблон чертежа заготовки            
            FileObject pieceTemplateDrawing = pieceSelectionObject.PieceTemplateDrawing;
            if (pieceTemplateDrawing == null) //%%TODO создать новый файл чертежа CAD
            {
                Ошибка("У типа '{0}' отсутствует шаблон чертежа заготовки.", pieceType.ToString());
            }

            string pieceName = String.Format("Заготовка {0}", pieceType[new Guid("40b4422e-741e-4996-ae60-038d1c2fd933")]);
            MaterialReferenceObject material = pieceSelectionObject.PieceMaterial;
            if (material != null)
                pieceName = String.Format("{0} '{1}'", pieceName, material.ToString().Trim());

            FileObject pieceDrawing = null;

            // Ищем папку "Заготовки" для хранения чертежа заготовки
            FileReference fileReference = new FileReference(pieceSelectionObject.Reference.Connection);
            TFlex.DOCs.Model.References.Files.FolderObject fileBilletFolder = null;
            Filter filter = null;
            if (Filter.TryParse("[Тип] = 'Папка' И [Наименование] = 'Заготовки'", fileReference.ParameterGroup, out filter))
            {
                fileBilletFolder = fileReference.Find(filter, 1).OfType<TFlex.DOCs.Model.References.Files.FolderObject>().FirstOrDefault();
                if (fileBilletFolder == null)
                {
                    // Создать папку "Заготовки"
                    TFlex.DOCs.Model.References.Files.FolderObject ktdFolder = null;
                    filter = null;
                    if (Filter.TryParse("[Тип] = 'Папка' И [Наименование] = 'Конструкторско-технологические документы'", fileReference.ParameterGroup, out filter))
                    {
                        ktdFolder = fileReference.Find(filter, 1).OfType<TFlex.DOCs.Model.References.Files.FolderObject>().FirstOrDefault();
                    }

                    ClassObject classObject = fileReference.Classes.Find("Папка");
                    fileBilletFolder = fileReference.CreateReferenceObject(ktdFolder, classObject) as TFlex.DOCs.Model.References.Files.FolderObject;
                    fileBilletFolder.Name.Value = "Заготовки";
                    fileBilletFolder.EndChanges();
                }
            }

            // создаём чертёж заготовки
            var copySet = fileReference.CopyReferenceObject(pieceTemplateDrawing, fileBilletFolder);
            try
            {
                pieceDrawing = copySet.GetNewObject(pieceTemplateDrawing) as FileObject;
                pieceName = String.Format("{0} [{1}]", pieceName, pieceDrawing.SystemFields.CreationDate.ToString("yyyy.MM.dd HH.mm.ss"));
                char[] invalidChars = System.IO.Path.GetInvalidFileNameChars();
                foreach (char invalidChar in invalidChars)
                {
                    int charIndex = pieceName.IndexOf(invalidChar);
                    if (charIndex >= 0)
                        pieceName = pieceName.Remove(charIndex, 1);
                }
                pieceDrawing.Name.Value = String.Format("{0}.grb", pieceName);

                copySet.EndChanges();
            }
            catch
            {
                if (copySet != null && copySet.Changing)
                {
                    copySet.CancelChanges();
                }
                throw;
            }

            pieceSelectionObject.BeginChanges();
            pieceSelectionObject.DrawingFileLink.SetLinkedObject(pieceDrawing);
            pieceSelectionObject.EndChanges();

            // Создаем объект заготовки в номенклатуре
            IPieceObjectCreatorService blankCreatorService = new PieceObjectCreatorService(Context.Connection);

            NomenclatureObject piece = null;
            Debugger.Launch();
            Debugger.Break();
            IPieceObject blankObject = blankCreatorService.CreatePieceObject();
            piece = blankObject.Piece;
            piece.Name.Value = pieceName;
            // подключить тип заготовки
            piece.SetLinkedObject(DigitalStructureToPieceTypesRelation, pieceSelectionObject.PieceTypeObject);

            // подключить материал
            piece.SetLinkedObject(new Guid("92a14dee-4cf8-4080-b933-3d85265b1745"), pieceSelectionObject.PieceMaterial);

            // передаем основные параметры в заготовку
            IEnumerable<PieceSelectionParameter> pieceSelectionParameters = pieceSelectionObject.Parameters.Objects.OfType<PieceSelectionParameter>();
            foreach (PieceSelectionParameter pieceSelectionParameter in pieceSelectionParameters)
            {
                if (pieceSelectionParameter.Name == "Масса")
                {
                    // масса заготовки
                    piece[new Guid("ee3cbb2b-3c92-4fef-85e9-d5bc3c9ce206")].Value = pieceSelectionParameter.Value;
                }

                if (!pieceSelectionParameter.IsContextParameter)
                    continue;

                ReferenceObject pieceParameter = piece.CreateListObject(new Guid("25de86a3-49b7-4160-baa7-0f757474eb6a"), new Guid("1cbe12b3-71e3-459f-beca-90a9e34ebd52"));
                pieceParameter[new Guid("45bb20d3-c206-4f73-a8f6-b14f32c5dcf5")].Value = pieceSelectionParameter.Name; // наименование
                pieceParameter[new Guid("cc146711-7035-4876-a03d-90d37e4e024b")].Value = pieceSelectionParameter.Value; // значение
                pieceParameter[new Guid("16d50f2f-0b93-423e-b127-8e9a63a3be76")].Value = pieceSelectionParameter.IsMain; // основной
                ReferenceObject descriptionParameter = pieceSelectionParameter.DescriptionPieceParameter;
                if (descriptionParameter != null)
                {
                    try
                    {
                        pieceParameter.SetLinkedObject(DictionaryParameterOfPieceRelation, descriptionParameter);
                    }
                    catch
                    {
                        // пока проглотим
                    }
                }

                pieceParameter.EndChanges();
            }

            // подключить чертеж заготовки к связанному документу заготовки с передачей параметров в контекст
            ProductDocumentObject pieceDocument = piece.LinkedObject as ProductDocumentObject;
            if (pieceDocument != null)
            {
                pieceDocument.BasicMaterial = pieceSelectionObject.PieceMaterial; // установка материала документа
                ((IPieceSelectionParametersProvider) pieceSelectionObject).AttachDrawingToDocument(pieceDocument);
            }

            if (piece.SaveSet != null)
            {
                piece.SaveSet.EndChanges();
            }
            else
            {
                if (piece.Changing)
                    piece.EndChanges();
            }

            return piece;
        }

        /// <summary>
        ///Подключение заготовки номенклатуры к текущему выбору заготовки
        /// </summary>
        private void ПодключитьЗаготовкуКВыборуЗаготовки(NomenclatureObject piece)
        {
            if (piece == null)
                return;

            PieceSelectionObject pieceSelectionObject = (PieceSelectionObject)ТекущийОбъект;

            if (pieceSelectionObject.PieceNomenclatureObjectLink == null)
                return;

            pieceSelectionObject.BeginChanges();
            pieceSelectionObject.PieceNomenclatureObjectLink.SetLinkedObject(piece); // при изменении связи выполнение расчёта- Если материал заготовки другой, изменить его в выборе заготовки
            pieceSelectionObject.EndChanges();
        }

        public void ПолучитьМатериалИзДокумента()
        {
            NomenclatureObject piece = (NomenclatureObject)ТекущийОбъект;
            ProductDocumentObject pieceDocument = piece.LinkedObject as ProductDocumentObject;
            if (pieceDocument != null)
            {
                MaterialReferenceObject documentMaterial = pieceDocument.BasicMaterial;
                if (documentMaterial != null)
                {
                    // подключить материал
                    piece.SetLinkedObject(new Guid("92a14dee-4cf8-4080-b933-3d85265b1745"), documentMaterial);
                }
            }
        }

        public void ПолучитьПараметрыИзКонтекстаДокумента()
        {
            NomenclatureObject piece = (NomenclatureObject)ТекущийОбъект;
            ProductDocumentObject pieceDocument = piece.LinkedObject as ProductDocumentObject;
            if (pieceDocument != null)
            {
                FileObject documentFile = pieceDocument.GetFiles().FirstOrDefault();
                if (documentFile != null)
                {
                    TFlexCadContext context = null;
                    byte[] contextData = documentFile.FindFileContext(pieceDocument);
                    if (contextData != null && contextData.Length > 0)
                    {
                        context = new TFlexCadContext(contextData);
                    }
                    if (context == null)
                        return;

                    foreach (KeyValuePair<string, TFlexCadVariableValue> pair in context)
                    {
                        // ищем параметр заготовки с такой же переменной из справочного параметра
                        ReferenceObject pieceParameter = piece.Links.ToMany[new Guid("25de86a3-49b7-4160-baa7-0f757474eb6a")].Objects.
                              FirstOrDefault(ro => ro.Links.ToOne[new Guid("e6995298-01ab-44f0-8cf0-89eff7635c13")].LinkedObject != null
                              && ro.Links.ToOne[new Guid("e6995298-01ab-44f0-8cf0-89eff7635c13")].LinkedObject[new Guid("4587058a-866c-4c1d-81c4-ef2bba8b2f91")].Value.ToString() == pair.Key);
                        if (pieceParameter == null)
                        {
                            Объект справочныйПараметр = НайтиОбъект("Параметры заготовок", "4587058a-866c-4c1d-81c4-ef2bba8b2f91", pair.Key);
                            if (справочныйПараметр != null)
                            {
                                ReferenceObject descriptionParameter = (ReferenceObject)справочныйПараметр;
                                if (descriptionParameter != null)
                                {
                                    pieceParameter = piece.CreateListObject(new Guid("25de86a3-49b7-4160-baa7-0f757474eb6a"), new Guid("1cbe12b3-71e3-459f-beca-90a9e34ebd52"));
                                    pieceParameter.SetLinkedObject(new Guid("e6995298-01ab-44f0-8cf0-89eff7635c13"), descriptionParameter);
                                    pieceParameter[new Guid("45bb20d3-c206-4f73-a8f6-b14f32c5dcf5")].Value = descriptionParameter[new Guid("304f2d60-fa49-401e-b59d-8c98ffb8d21e")]; // наименование
                                }
                            }
                            else
                                continue;
                        }
                        else
                        {
                            pieceParameter.BeginChanges();
                        }

                        if (!String.IsNullOrEmpty(pair.Value.Text))
                        {
                            pieceParameter[new Guid("cc146711-7035-4876-a03d-90d37e4e024b")].Value = pair.Value.Text; // значение
                        }
                        else
                        {
                            pieceParameter[new Guid("cc146711-7035-4876-a03d-90d37e4e024b")].Value = pair.Value.Real; // значение
                            pieceParameter[new Guid("16d50f2f-0b93-423e-b127-8e9a63a3be76")].Value = true;
                        }

                        pieceParameter.EndChanges();
                    }
                }
            }
        }

        public void СоздатьУсловияПоискаПоСтрокеФильтра()
        {
            // очищаем список условий
            Context.ReferenceObject.ClearObjectList(_termListGuid);

            // текущий фильтр
            string filterString = Context.ReferenceObject.ParameterValues[_filterStringParameterGuid].Value as string;
            if (String.IsNullOrEmpty(filterString))
                return;

            // формируем новый список условий по условиям строки фильтра
            List<SearchTerm> termCollection = TFlex.DOCs.Model.Macros.ObjectModel.SearchTerm.Parse(filterString).ToList();
            foreach (SearchTerm term in termCollection)
            {
                ReferenceObject pieceTerm = Context.ReferenceObject.CreateListObject(_termListGuid, _termTypeGuid);
                pieceTerm[_termNameParameterGuid].Value = term.Parameter;
                pieceTerm[_termOperatorParameterGuid].Value = term.Operator;
                pieceTerm[_termValueParameterGuid].Value = term.Value;
                pieceTerm.EndChanges();
            }
        }

        public void СоздатьСтрокуФильтраПоУсловиямПоиска()
        {
            string newFilterString = String.Empty;

            // обновляем строку фильтра по списку условий
            List<ReferenceObject> pieceTermCollection = Context.ReferenceObject.GetObjects(_termListGuid);
            foreach (ReferenceObject pieceTerm in pieceTermCollection)
            {
                string currentTerm =
                    $"[{pieceTerm[_termNameParameterGuid].Value} {pieceTerm[_termOperatorParameterGuid].Value} '{(pieceTerm[_termValueParameterGuid].Value != null ? pieceTerm[_termValueParameterGuid].Value.ToString().Replace(' ', '_') : String.Empty)}']";

                if (String.IsNullOrEmpty(newFilterString))
                {
                    newFilterString = currentTerm;
                }
                else
                {
                    newFilterString += String.Format(" И {0}", currentTerm);
                }
            }

            Context.ReferenceObject.ParameterValues[_filterStringParameterGuid].Value = newFilterString;
        }
    }
    
    public interface IPieceObject
    {
        NomenclatureObject Piece { get; }
    }

    public class PieceReferenceObjectAdapter : IPieceObject
    {
        public PieceReferenceObjectAdapter(ReferenceObject piece)
        {
            Piece = (NomenclatureObject)piece;
        }

        public NomenclatureObject Piece { get; }
    }

    
    public interface IPieceObjectCreatorService
    {
        IPieceObject CreatePieceObject();
    }

    public class PieceObjectCreatorService : IPieceObjectCreatorService
    {
        private readonly ServerConnection _connection;

        public PieceObjectCreatorService(ServerConnection connection)
        {
            _connection = connection;
        }

        public IPieceObject CreatePieceObject()
        {
            var nomenclatureReference = new NomenclatureReference(_connection);
            NomenclatureType blankType = nomenclatureReference.Classes.Piece;
            var pieceObject = (NomenclatureReferenceObject)nomenclatureReference.CreateReferenceObject(blankType);
            pieceObject.CreateSaveSet();
            return new PieceReferenceObjectAdapter(pieceObject);
        }
    }
}

