using System;
using TFlex.DOCs.Model.Macros;
using TFlex.DOCs.Model.Macros.ObjectModel;
using TFlex.DOCs.Model.Plugins;
using TFlex.DOCs.Model.References.Links;
using TFlex.DOCs.Model.References;
using TFlex.DOCs.Model;
using TFlex.DOCs.UI.Common;
using TFlex.DOCs.UI.Objects.Managers;
using TFlex.DOCs.Model.Classes;
using System.Windows.Forms;


public class Macro : MacroProvider
{
    private readonly Guid Идентификатор_справочника_СТО = new Guid("904fc7da-77df-4763-94b9-1ada11aafa4a");
    private readonly Guid Идентификатор_поля_наименование_средства = new Guid("82042c52-5aef-4183-bdb7-2fdd5e69efdc");

    private class Идентификаторы
    {
        #region Получение экземпляра класса
        private static readonly Guid идентификаторСправочникаТехпроцессов = new Guid("353a49ac-569e-477c-8d65-e2c49e25dfeb");
        private static readonly Guid идентификаторСправочникаОпераций = new Guid("8f779dda-4f4f-498b-a91d-2f669a6e7a93");
        private static readonly Guid идентификаторСправочникаПереходов = new Guid("edc9f590-1b93-485e-b628-13dd41f01541");

        private static readonly Guid идентификаторТипаТехпроцесс = new Guid("376d88a3-60eb-409e-afcd-da75231c05f4");
        private static readonly Guid идентификаторТипаОперация = new Guid("553fde7e-beb8-4194-9126-f500057a1a81");
        private static readonly Guid идентификаторТипаПереход = new Guid("e5750071-e3ea-4a0f-8fd9-79c9ae8e7282");

        private static ReferenceInfo _справочникТехпроцессов;
        private static ReferenceInfo СправочникТехпроцессов
        {
            get
            {
                if (_справочникТехпроцессов == null)
                    _справочникТехпроцессов = ReferenceCatalog.FindReference(идентификаторСправочникаТехпроцессов);
                return _справочникТехпроцессов;
            }
        }
        
        private static ClassObject _типТехпроцесс;
        private static ClassObject ТипТехпроцесс
        {
            get
            {
                if (_типТехпроцесс == null)
                    _типТехпроцесс = СправочникТехпроцессов.Classes.AllClasses.Find(идентификаторТипаТехпроцесс);
                return _типТехпроцесс;
            }
        }

        private static ClassObject _типОперация;
        private static ClassObject ТипОперация
        {
            get
            {
                
                if (_типОперация == null)
                    _типОперация = ТипТехпроцесс.GetOneToManyRelations().Find(идентификаторСправочникаОпераций).Classes.AllClasses.Find(идентификаторТипаОперация);
                return _типОперация;
            }
        }

        private static ClassObject _типПереход;
        private static ClassObject ТипПереход
        {
            get
            {
                if (_типПереход == null)
                    _типПереход = ТипОперация.GetOneToManyRelations().Find(идентификаторСправочникаПереходов).Classes.AllClasses.Find(идентификаторТипаПереход);
                return _типПереход;
            }
        }

        private static readonly Идентификаторы идентификаторыДляПерехода = new Идентификаторы("3a71d45f-2f6b-4409-b60b-4e5309b2d153", "dcc61c9c-0f5c-4d75-9ad0-72ce26fb4bb7");
        private static readonly Идентификаторы идентификаторыДляОперации = new Идентификаторы("81e4363d-e363-422b-8c98-c15011654f4b", "bfb9b583-4adc-49fb-ae82-bc6af0e4d6b5");
        private static readonly Идентификаторы идентификаторыДляТехпроцесса = new Идентификаторы("63919d08-1c93-4298-a05f-28d214da6435", "4fe03ae6-98ca-41de-86c8-4bf7fb77577f");
        
        public static Идентификаторы ПолучитьИдентификаторыДляТипа(ClassObject тип)
        {
            if (тип == ТипТехпроцесс || тип.IsInherit(ТипТехпроцесс))
                return идентификаторыДляТехпроцесса;
            else if(тип == ТипОперация || тип.IsInherit(ТипОперация))
               return идентификаторыДляОперации; 
            else if(тип == ТипПереход || тип.IsInherit(ТипПереход))
                return идентификаторыДляПерехода;
            else
                return null;
        }
        #endregion

        public Guid Текстовое_Поле
        {
            get;
            private set;
        }
       
        public Guid Связь_к_СТО
        {
            get;
            private set;
        }

        private Идентификаторы(Guid текстовогоПоля, Guid связи_к_СТО)
        {
            Текстовое_Поле = текстовогоПоля;
            Связь_к_СТО = связи_к_СТО;
        }

        private Идентификаторы(string текстовогоПоля, string связи_к_СТО)
            : this(new Guid(текстовогоПоля), new Guid(связи_к_СТО))
        {
        }
    }

    private Идентификаторы _текущие_идентификаторы;
    private Идентификаторы ТекущиеИдентификаторы
    {
        get
        {
            if(_текущие_идентификаторы == null)
                _текущие_идентификаторы = Идентификаторы.ПолучитьИдентификаторыДляТипа(Context.ReferenceObject.MasterObject.Class);
            return _текущие_идентификаторы;
        }
    }

    private OneToOneLink _связь_к_СТО;
    private OneToOneLink Связь_к_СТО
    {
        get
        {
            if(_связь_к_СТО == null)
            {
                _связь_к_СТО = Context.ReferenceObject.Links.ToOne[ТекущиеИдентификаторы.Связь_к_СТО];
            }
            return _связь_к_СТО;
        }
    }
    
    private ReferenceInfo _справочник_СТО;
    private ReferenceInfo Справочник_СТО
    {
        get
        {
            if(_справочник_СТО == null)
                _справочник_СТО = ReferenceCatalog.FindReference(Идентификатор_справочника_СТО);
            return _справочник_СТО;
        }
    }

    public Macro(MacroContext context)
        : base(context)
    {
    }

    public void ВыбратьОборудование()
    {
        ВызватьДиалог("a0468938-0e70-414d-b6f3-db40244d1d30");
    }

    public void ВыбратьПриспособление()
    {
        ВызватьДиалог("722a4733-3819-41f5-a1be-67f8dd3f91e7");
    }

    public void ВыбратьСОЖ()
    {
        ВызватьДиалог("900d23f2-ae0d-4c49-a0d9-742b5cd6d462");
    }

    public void ВыбратьВспомогательныйИнструмент()
    {
        ВызватьДиалог("cbdce63b-c91a-4f28-97c4-36f50e1c265c");
    }

    public void ВыбратьРежущийИнструмент()
    {
        ВызватьДиалог("43c7c3c3-95bf-4934-9032-4a6c7a27bb2f");
    }

    public void ВыбратьСредстваИзмерения()
    {
        ВызватьДиалог("05573d4f-955a-4925-8ab8-9621d8b2e747");
    }

    public void ВыбратьСборочныйИнструмент()
    {
        ВызватьДиалог("a3de31c9-4b00-4209-b69f-9061fd4fdcc2");
    }

    public void ВыбратьСборочныеКомплектующие()
    {
        ВызватьДиалог("60a1c24d-35bd-41d0-be45-a2fbdcc41bd6");
    }

    public void ВыбратьМатериал()
    {
        ВызватьДиалог("2a90d021-1aa0-4cd6-9dc4-723ef05eddb4");
    }

    public void ВыбратьМодельнуюИОпочнуюОснастку()
    {
        ВызватьДиалог("b2de4e76-9ef4-4021-a0b9-e7b3e6a72f4f");
    }

    public void ВыбратьСредстваЗащиты()
    {
        ВызватьДиалог("f42303e0-27a3-4050-bbe8-de7ad23421dd");
    }

    private void ВызватьДиалог(string typeGuid)
    {
        if (Context != null && Context.ReferenceObject != null)
        {
            Guid guid;
            if (Guid.TryParse(typeGuid, out guid))
            {
                if (guid != Guid.Empty && Связь_к_СТО != null && Справочник_СТО != null)
                {
                    using (TFlex.DOCs.UI.Common.References.ISelectReferenceObjectDialog dialog = ObjectCreator.CreateObjectByFoundContext<TFlex.DOCs.UI.Common.References.ISelectReferenceObjectDialog>(Справочник_СТО.Guid))
                    {
                        dialog.Initialize(Справочник_СТО, false);
                        dialog.IsMultipleSelect = false;
                        ClassObject co = Справочник_СТО.Classes.Find(guid);
                        if (co != null)
                        {
                            dialog.ReferenceEnvironmentControl.VisualRepresentation.AddValidClass(Справочник_СТО.Classes.Find(guid));
                            dialog.ReferenceEnvironmentControl.VisualRepresentation.Refresh(false);
                        }

                        ReferenceObject currentObject = Связь_к_СТО != null ? Связь_к_СТО.LinkedObject : null;
                        if (currentObject != null && (co == currentObject.Class || co.GetAllChildClasses().Contains(currentObject.Class)))
                            dialog.FocusedObject = currentObject;

                        if (dialog.ShowDialog((Context as UIMacroContext).OwnerWindow) == DialogOpenResult.Ok)
                        {
                            try
                            {
                                Связь_к_СТО.MasterObject.BeginChanges();
                                УстановитьСвязанныйОбъект(dialog.FocusedObject);
                                УстановитьЗначениеТекстовогоПоля(dialog.FocusedObject);
                                Связь_к_СТО.MasterObject.EndChanges();
                            }
                            catch
                            {
                                Связь_к_СТО.MasterObject.CancelChanges();
                                throw;
                            }
                        }
                    }
                }
                (Context as UIMacroContext).ReferenceUIObject.RaiseObjectChanged();
            }
            else
                throw new ArgumentException("Невозможно преобразовать указаную строку в GUID");
        }
    }

    private void УстановитьСвязанныйОбъект(ReferenceObject linkedObject)
    {
        Связь_к_СТО.SetLinkedObject(linkedObject);
    }
    
    private void УстановитьЗначениеТекстовогоПоля(ReferenceObject linkedObject)
    {
        Связь_к_СТО.MasterObject[ТекущиеИдентификаторы.Текстовое_Поле].Value = linkedObject[Идентификатор_поля_наименование_средства].Value;
    }
    
}

