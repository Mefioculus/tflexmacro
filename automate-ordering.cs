using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using TFlex.DOCs.Model.Macros;
using TFlex.DOCs.Model.Macros.ObjectModel;
using TFlex.DOCs.Model.References;
using TFlex.DOCs.Model.Structure;

public class Macro : MacroProvider
{
    #region Guids

    private static class Guids {
        public static class Properties {
            public static Guid ПорядковыйНомерПротокола = new Guid("d662eed7-c2a2-41fc-9b35-05e527349cc7");
            public static Guid ПорядковыйНомерОбразца = new Guid("e3e591c4-662a-4b23-8315-d0eeed89fda6");

        }

        private static class References {
            public static Guid АрхивЦЗЛ = new Guid("f7f43d73-857c-41f9-b449-38ee72caa221");
        }
    }

    #endregion Guids


    public Macro(MacroContext context)
        : base(context)
    {
    }

    public override void Run()
    {
    }
    
    #region Нумерация в документах ОГТ
    public void НумерацияВДокументыОГТ()
    {
        // Для начала получаем текущий объект (как я понимаю, это должен быть объект, который в данный момент создается
        Объект запись = ТекущийОбъект;
    
        // Далее, получаем список всех объектов данного справочника
        Объекты существующиеЗаписи = НайтиОбъекты("500d4bcf-e02c-4b2e-8f09-29b64d4e7513", "Тип", "Документы ОГТ");
  
        //Проверка на то, что в данном справочнике уже существуют записи
        if (существующиеЗаписи != null)
        {
            int значениеПоследнегоНомера = 0;        
    	   // Мы получаем номер записи с самым большим значением
            foreach(Объект существующаяЗапись in существующиеЗаписи)
            {
        	   if (существующаяЗапись.Параметр["Номер"] > значениеПоследнегоНомера)
                {
                    значениеПоследнегоНомера = существующаяЗапись.Параметр["Номер"];
                }
            }
            ТекущийОбъект.Параметр["Номер"] = значениеПоследнегоНомера + 1;
        }
        else
        {
            ТекущийОбъект.Параметр["Номер"] = 1;
        }
    }
    #endregion нумерация в документах ОГТ
    
    #region Нумерация в документах ОГТ через ID
    public void НумерацияВДокументыОГТчерезID()
    {
        Объект record = ТекущийОбъект;
        
        int idRecord = record.Параметр["Id"];
        
        // Далее пытаемся найти предыдущий созданный объект
        Объект previousRecord = null;
        while (true)
        {
            idRecord--;
            previousRecord = НайтиОбъект("500d4bcf-e02c-4b2e-8f09-29b64d4e7513", "ID", idRecord.ToString());
            if (previousRecord != null)
            {
                break;
            }
        }
        
        // Присваиваем текущей записи новый порядковый номер
        record.Параметр["Номер"] = previousRecord.Параметр["Номер"] + 1;
        
    }
    #endregion Нумерация в документах ОГТ через ID
 
    #region Нумерация в документах ОГТ 2
    public void НумерацияВДокументыОГТ2()
    {
        // Для начала получаем текущий объект (как я понимаю, это должен быть объект, который в данный момент создается
        Объект запись = ТекущийОбъект;
        
        
         // Далее, получаем список всех объектов данного справочника
        Объекты существующиеЗаписи = НайтиОбъекты("500d4bcf-e02c-4b2e-8f09-29b64d4e7513", "Тип", "Документы ОГТ");   
        
         if (существующиеЗаписи != null)
        {
            
    	    List<int> list_num = new List<int>();
            // Мы получаем номер записи с самым большим значением
            foreach(Объект существующаяЗапись in существующиеЗаписи)
              	    list_num.Add(существующаяЗапись["Номер"]);
            //Сообщение("Информация",list_num.Max().ToString());
            ТекущийОбъект.Параметр["Номер"] = list_num.Max() + 1;
        }
    }
    #endregion нумерация в документах ОГТ 2

    #region Нумерация в документах ОГТ 3
    public void НумерацияВДокументыОГТ3_создание()
    {
        int номер = (int)ГлобальныйПараметр["ID_Документы ОГТ"] + 1;
        //Сообщение("Номер",номер.ToString());
        ТекущийОбъект["Номер"] = номер;
        ГлобальныйПараметр["ID_Документы ОГТ"] = номер;
        ТекущийОбъект.Изменить();    
        ТекущийОбъект.Сохранить();
    }
    #endregion Нумерация в документах ОГТ 3

    #region Нумерация в документах ОГТ 3 удаление
    public void НумерацияВДокументыОГТ3_удаление()
    {
        int номер = (int)ГлобальныйПараметр["ID_Документы ОГТ"];
        
        //Сообщение("Номер",номер.ToString());
        
        if (номер==ТекущийОбъект["Номер"])
            ГлобальныйПараметр["ID_Документы ОГТ"] = номер-1;        
    }
    #endregion Нумерация в документах ОГТ 3 удаление

    #region Нумерация документов в архиве ЦЗЛ

    #region НумероватьНовыйОбъектАрхивЦЗЛ

    public void НумероватьНовыйОбъектАрхивЦЗЛ() {
        SetCountNumber("id", Guids.Properties.ПорядковыйНомерПротокола);
    }

    #endregion НумероватьНовыйОбъектАрхивЦЗЛ

    #region НумероватьНовыйОбразец

    public void НумероватьНовыйОбразец() {
        SetCountNumber("count", Guids.Properties.ПорядковыйНомерОбразца);
    }

    #endregion НумероватьНовыйОбразец

    #region SetCountNumber

    private void SetCountNumber(string method, Guid propsForCountNumber) {
        // Получаем текущий объект, которому нужно присвоить номер
        ReferenceObject newRecord = Context.ReferenceObject;
        // Получаем справочник, к которому относится текущий объект
        Reference referenceOfRecord = newRecord.Reference;

        if (referenceOfRecord == null) {
            Error(string.Format("При попытке вычислить порядковый номер объекта {0} не удалось определить, к какому объекту этот справочник относится", newRecord.ToString()));
        }
        
        // Производим загрузку всех элементов справочника
        referenceOfRecord.Objects.Load();

        if (referenceOfRecord.Objects.GetCount() != 0) {
            // Обработка случая вычисления максимального номера по id
            if (method.ToLower() == "id") {
                int id = 0;
                foreach (ReferenceObject oldRecord in referenceOfRecord.Objects) {
                    id = (id < oldRecord.SystemFields.Id) ? oldRecord.SystemFields.Id : id;
                }
                newRecord[propsForCountNumber].Value = ((int)referenceOfRecord.Find(id)[propsForCountNumber].Value) + 1;
            }
            // Обработка случая вычисления максимального номера по номеру
            else if (method.ToLower() == "count") {
                int orderNumber = 0;
                foreach (ReferenceObject oldRecord in referenceOfRecord.Objects) {
                    orderNumber = (orderNumber < (int)oldRecord[propsForCountNumber].Value) ? (int)oldRecord[propsForCountNumber].Value : orderNumber;
                }
                newRecord[propsForCountNumber].Value = orderNumber + 1;
            }
            else {
                Error("Неизвестный метод определения нумерации");
            }
        }
        else {
            Message("Информация", "Создание первого объекта");
            newRecord[propsForCountNumber].Value = 1;
        }
    }

    #endregion SetCountNumber

    #endregion Нумерация документов в архиве ЦЗЛ

}
