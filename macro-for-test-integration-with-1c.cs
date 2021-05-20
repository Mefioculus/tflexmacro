using System;
using System.Net; // Пространство имен для отправления запросов и получения ответов
using System.Text; // Пространство имен для работы с кодировками
using System.Linq;
using Newtonsoft.Json;
using System.Collections;
using System.Collections.Generic;
using TFlex.DOCs.Model.Macros;
using TFlex.DOCs.Model; // Для ParameterInfo
using TFlex.DOCs.Model.References; // Для работы со справочниками
using TFlex.DOCs.Model.Structure; // Для использования ParameterInfo
using TFlex.DOCs.Model.Search; // Для использования ComparisonOperator
using TFlex.DOCs.Model.Classes; // Для использования ClassObject 


public class Macro : MacroProvider {
    public Macro(MacroContext context) : base(context) {
    }

    private static class Guids {
        public static class References {
            public static Guid Предприятия = new Guid("4d852d70-2377-451a-874a-050e14176e4a");
        }
        public static class Parameters {
            public static Guid Наименование = new Guid("345b220e-5964-4168-b849-b9132645067f");
            public static Guid Код = new Guid("0f0e179c-62dc-458a-aad2-edc21eeae54f");
        }
        public static class Types {
            public static Guid Предприятие = new Guid("0655cf5a-e11c-4ed2-ae6b-ec31b7b74957");
        }
    }

    #region Properties

    private string url { get; set; } = "http://192.168.2.15/uch-test/hs/Get1C";

    #endregion Properties

    public override void Run() {
        Сообщение("Полученные перекодированные данные", GetString());

        string jsonString = GetString();

        var reference = DeserializeData(jsonString);

        Сообщение("Информация", string.Format("Всего получено записей {0}\nПервая запись '{1} - {2}'", reference.Count, reference[2].Code, reference[2].Name));

        UpdateTable(reference);
        Сообщение("Информация", "Обновление справочника завершено");
    }

    #region Get and decode message from 1C

    private string GetString() {
        // Создаем клиент для подключения к сервису 1С
        WebClient wc = new WebClient();
        wc.Credentials = new NetworkCredential("Gukov", "123");

        byte[] response = wc.DownloadData(url);

        // Перекодируем


        Encoding srcEncoding = Encoding.UTF8;
        Encoding dstEncoding = Encoding.GetEncoding(1251);

        byte[] encodingResponse = Encoding.Convert(srcEncoding, dstEncoding, response);
        

        return Encoding.Default.GetString(encodingResponse);
    }

    #endregion Get and decode message from 1C

    private List<Record> DeserializeData(string jsonString) {
        return JsonConvert.DeserializeObject<List<Record>>(jsonString);
    }

    private void UpdateTable(List<Record> records) {
        // Получаем справочник для работы
        ReferenceInfo referenceInfo = Context.Connection.ReferenceCatalog.Find(Guids.References.Предприятия);
        Reference reference = referenceInfo.CreateReference();

        // Получаем параметр Code справочника
        ParameterInfo CodeParameter = reference.ParameterGroup[Guids.Parameters.Код];

        foreach (Record record in records) {
            if (reference.Find(CodeParameter, ComparisonOperator.Equal, record.Code).Count == 0) {
                
                // Создаем объект справочника
                ClassObject classObject = reference.Classes.Find(Guids.Types.Предприятие);
                ReferenceObject newRecord = reference.CreateReferenceObject(classObject);
                
                // Присваиваем полям пораметра значения
                newRecord.BeginChanges();
                newRecord[Guids.Parameters.Наименование].Value = record.Name;
                newRecord[Guids.Parameters.Код].Value = record.Code;
                newRecord.EndChanges();
            }
        }

        
    }


    private class Record {
        public string Name { get; set; }
        public int Code { get; set; }
    }
}


