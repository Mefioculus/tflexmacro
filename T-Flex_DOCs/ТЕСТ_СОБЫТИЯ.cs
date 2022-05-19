using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using TFlex.DOCs.Model.Macros;
using TFlex.DOCs.Model.Macros.ObjectModel;
using TFlex.DOCs.Model.References.Files;
using TFlex.DOCs.Model.References;

using System.Text;
using TFlex.DOCs.Model.References.Users;
using TFlex.DOCs.Model;
using TFlex.DOCs.Model.Structure;
using TFlex.DOCs.Model.Search;
using TFlex.DOCs.Model.Classes;
//using TFlex.DOCs.Common;

public class Macro : MacroProvider
{
    public Macro(MacroContext context)
        : base(context)
    {
    }

    public override void Run()
    {
    
    }
    
    
         
     public void Test1()
    {
        Сообщение("Номер1","событие 1");
    }   

     public void Test2()
    {       
        Сообщение("Номер2","событие 2");
            
    }         

     
          public void Test3()
    {       
        Сообщение("Номер3","событие 3");
            
    }       

          public void Test4()
    {       
        Сообщение("Номер4","событие 4");
            
    }            


          public void Test5()
    {       
        Сообщение("Номер5","событие 5");
            
    }         

       public void Test6()
    {       
        Сообщение("Номер6","событие 6");
            
    }                   

       public void Test7()
    {       
        Сообщение("Номер7","событие 7");
            
    } 



         public void testuser()
    {
        System.Diagnostics.Debugger.Launch();
        System.Diagnostics.Debugger.Break();
        User user = (User)ТекущийОбъект;
        user.BeginChanges();
        string MailSettings = $"<SmtpServerSettings ServerName=\"zimbra.corp.aeroem.ru\" Login=\"{user.Login}\" Name=\"{user.FullName}\" EMail=\"{user.Email}\" UseDOCsEMail=\"True\" UseDOCsName=\"True\" />";
        user.MailSettings.Value = MailSettings;
        var sentmailparam = user.MailSendType;
        //user.MailSendType.Value = 2;
        user.EndChanges();
        RefreshReferenceWindow();
    }





    public void connect()
    {

        // Соединение с сервером приложений

        Console.WriteLine("Test");
        // Console.ReadLine();


        ServerConnection serverConnection = null;

        try
        {
            //Context.Connection.ser
            // Подключаемся к серверу

            //   serverConnection = ServerConnection.Open("Администратор", "123", "TFLEX-DOCS:22321");
            //serverConnection = ServerConnection.Open("Администратор", "123", "TFLEX-DOCS:22321");
            //serverConnection = ServerConnection.Open("Администратор", "Aem1234", "TFLEX-DOCS:21421"); //Test
            serverConnection = ServerConnection.Open("markinaa", "123", "TFLEX-DOCS:21324"); //merge
            Console.WriteLine("Подключение к серверу...");

            if (!serverConnection.IsConnected)
            {
                Console.WriteLine("Ошибка: невозможно подключиться к серверу");
                return;
            }

            // Подписываемся на событие потери соединения с сервером
            serverConnection.ConnectionLost += ServerGatewayOnConnectionLost;
            Console.WriteLine("Подлючение установлено");




            /*
             Guid ref_user = new Guid("8ee861f3-a434-4c24-b969-0896730b93ea");
             ReferenceInfo info = serverConnection.ReferenceCatalog.Find(ref_user);
             Reference reference = info.CreateReference();
             ParameterInfo parameterInfo = reference.ParameterGroup[SystemParameterType.Class];
             List<ReferenceObject> result = reference.Find(parameterInfo, ComparisonOperator.Equal, "Сотрудник");
            */




            /*
             Reference reference=new Reference();
             reference.Classes.Find("22ee7b25 - 0744 - 405c - 986e-01d62bfed86d");
                     MaterialReference materialReference = new MaterialReference(serverConnection);
                 ReferenceObjectCollection reff = materialReference.Objects;
             */


            /*


            Reference reference;
            // reference.Connection = serverConnection;

            //MaterialReference materialReference = new MaterialReference(serverConnection);
            //ReferenceObjectCollection reff = materialReference.Objects;
            SpecReference specReference = new SpecReference(serverConnection);
            ReferenceObjectCollection ss = specReference.Objects;
            //S
            //ReferenceObjectCollection spec = specReference.Objects;
            //spec.GetType
            //var reff = materialReference; //.ParameterGroup; //.AuthorParameterInfo;
            //Console.WriteLine(reff.GetType());
            //Name    "Автор" string

            /*                +AuthorParameterInfo { Автор}
                            TFlex.DOCs.Model.Structure.ParameterInfo { TFlex.DOCs.Model.Internal.Structure.AuthorParameterInfo}
            */
            //reff
            //  Console.ReadLine();
            //  Console.WriteLine(specReference.GetType());
            //List<ReferenceObject> list = reff.arr;
            //Console.WriteLine(list[2].ToString());
            /*
            foreach (var r in reff)
            {
                Console.WriteLine(r.ToString());
            }
            */

            /*
            foreach (var r in spec)
            {
                Console.WriteLine(r.ToString());
            }
            */

            /*
            // Создаем экземпляра файлового справочника
              Console.WriteLine("Импорт данных...");
              FileReference fileReference = new FileReference(serverConnection);

              // Задаем целевую папку для импорта файлов
              FolderObject parentFolder = fileReference.FindByRelativePath("Служебные файлы") as FolderObject;
              // Вызываем функцию импорта файлов
              FolderObject result = fileReference.Import(FolderPath,
                  new ImportParameters
                  {
                      DestinationFolder = parentFolder,
                      Recursive = true,
                      CreateClasses = true,
                      AutoCheckIn = false
                  }
              );
              // Выводим результат операции
              ViewResult(result);

              */
            // Console.WriteLine("Импорт данных успешно завершен.");
            // Console.ReadLine();

            // */
            //   FileCopy(serverConnection);
            //Console.ReadLine();

           //roleref(serverConnection);
            usercomp(serverConnection);
        }
        catch (Exception e)
        {
            // Выводим сообщение об ошибке
            Console.WriteLine("Ошибка: " + e.Message);
        }
        finally
        {
            if (serverConnection != null)
            {
                if (serverConnection.IsConnected)
                    serverConnection.Close();

                serverConnection.Dispose();
            }
        }

    }


    private static void ServerGatewayOnConnectionLost(object sender, ConnectionLostEventArgs connectionLostEventArgs)
    {
        Console.WriteLine("Внимание: Разрыв связи с сервером");
    }




    /// <summary>
    /// Получает все объекты справочника
    /// </summary>
    /// 
    public List<ReferenceObject> GetAnalog(Guid guidref)
    {
        ReferenceInfo info = Context.Connection.ReferenceCatalog.Find(guidref);
        Reference reference = info.CreateReference();
        List<ReferenceObject> result = reference.Objects.ToList();
        return result;
    }


    public List<ReferenceObject> GetAnalog(String str, Guid guidref, Guid parametr)

    {
        /*
        ReferenceInfo info = Context.Connection.ReferenceCatalog.Find(guidref);
        Reference reference = info.CreateReference();
        //ParameterInfo parameterInfo = reference.ParameterGroup.Parameters(guidref);
        //var group = reference.ParameterGroup.Parameters;
        //var f = reference.
        //var parameterInfo = group.Find("login");
        TFlex.DOCs.Model.Classes.ClassObject parameterInfo = info.Classes.Find(guidref);
        */

        ReferenceInfo info = Context.Connection.ReferenceCatalog.Find(guidref);
        Reference reference = info.CreateReference();
        //Находим тип «Длина»
        ClassObject classObject = info.Classes.Find(new Guid("41ca5dde-4b17-4144-b12f-447eacefd7e8"));
        var parameterInfo = classObject.ParameterGroups[0];

        List<ReferenceObject> result = reference.Find(null, ComparisonOperator.Equal, str);
        //Console.WriteLine("GetAnalog");
        return result;
    }


    /// <summary>
    /// Получает все объекты справочника
    /// </summary>
    /// 
    public List<ReferenceObject> GetAnalog(ServerConnection connect, Guid guidref)
    {
        ReferenceInfo info = connect.ReferenceCatalog.Find(guidref);
        Reference reference = info.CreateReference();
        List<ReferenceObject> result = reference.Objects.ToList();
        return result;
    }


    /// <summary>
    /// Получает объекты справочника по условию если parametr содердит str
    /// </summary>
    public List<ReferenceObject> GetAnalog(ServerConnection connect, string param, Guid guidref, List<string> str)

    {
        Guid login2 = new Guid("1a78cc76-8dd0-4263-98e7-d6525813e89c");
        Guid login = new Guid("42c81c2b-7354-46aa-9547-0f1a93e9d4e1");
        List<ReferenceObject> result = new List<ReferenceObject>();

        ReferenceInfo info = connect.ReferenceCatalog.Find(guidref);
        Reference reference = info.CreateReference();
        //
        ParameterInfo parameterInfo = null;
        if (param.Equals("tip"))
            parameterInfo = reference.ParameterGroup[SystemParameterType.Class];
        //if (param.Equals("tip"))
        if (param.Equals("login"))
            parameterInfo = reference.ParameterGroup[login2];

        if (parameterInfo != null)
            result = reference.Find(parameterInfo, ComparisonOperator.IsOneOf, str);
        Console.WriteLine("GetAnalog");
        return result;
    }


    public void connect_test()
    {
        Guid ref_spec = new Guid("c587b002-3be2-46de-861b-c551ed92c4c1");
        Guid ref_spec_shifr = new Guid("2a855e3d-b00a-419f-bf6f-7f113c4d62a0");


        Guid ref_user = new Guid("8ee861f3-a434-4c24-b969-0896730b93ea");
        Guid ref_user_login = new Guid("42c81c2b-7354-46aa-9547-0f1a93e9d4e1");
        Guid ref_user_login2 = new Guid("1a78cc76-8dd0-4263-98e7-d6525813e89c");

        var find_izd = GetAnalog("markinaa", ref_user, ref_user_login);

        //Guid ref_user = new Guid("8ee861f3-a434-4c24-b969-0896730b93ea");
        //var ObjUser2 = GetAnalog(connect, "tip", ref_user, );

        //var find_izd = GetAnalog("8А8615967", ref_spec, ref_spec_shifr);
    }

    public void roleref(ServerConnection connect)
    {
        StringBuilder strbild = new StringBuilder();
        Guid ref_role = new Guid("212a5ec8-3f36-4501-bb46-082d200ba05f");
        var ObjUser2 = GetAnalog(Context.Connection, ref_role);
        var ObjUser = GetAnalog(connect, ref_role);

        foreach (var item in ObjUser)
        {
        //    if (item.Guid.Equals(new Guid("056cfff0-c86b-43e6-8571-f74621184adb")))
       //     {
                foreach (var item2 in ObjUser2)
                {
                    if (item2.Guid == item.Guid)
                    {
                        if (item2 == item)
                        {
                            strbild.AppendLine($"Совпадают {item.ToString()}");
                        }
                        else
                        {
                            strbild.AppendLine($"Не совпадают {item.ToString()}");
                        }
                    }

                }


                var t1 = from i1 in ObjUser2 where i1.Guid == item.Guid select i1;
                var t2 = from i2 in ObjUser where i2.Guid == item.Guid select i2;
              //  Message("", $"{t1.Count().ToString()} {t2.Count().ToString()}");
                if (t1.Count() == 0)
                    strbild.AppendLine($"Отсутствует {item.ToString()}");

                if (t2.Count() == 0)
                    strbild.AppendLine($"Отсутствует2 {item.ToString()}");
           // }
            //var t1 = from item in ObjUser2  select (item).Guid 


            /*            var t1 = from i1 in ObjUser2 where i1.Guid == item.Guid select i1;
                        if (t1 != null)
                        {
                            if (item.Equals(t1))
                            {
                                strbild.AppendLine($"Совпадают {item.ToString()}");
                            }
                            else
                            {
                                strbild.AppendLine($"Не совпадают {item.ToString()}");
                            }
                        }
                        else
                        {
                            strbild.AppendLine($"Отсутствует {item.ToString()}");
                        }*/


            //var test0 = item.ParameterValues;
            //var test = item.GetObjectValue("Name");
            //var test2 = item["Name"];
            /*if (ObjUser2.Contains(item))
            {
                strbild.AppendLine(item.GetObjectValue("Name"))
            }*/

        }
        Message("", strbild.ToString());
        Message("", ObjUser[0].Equals(ObjUser2[0]));

        Message("", ObjUser.Count.ToString());
        Message("", ObjUser2.Count.ToString());
    }

    public void usercomp(ServerConnection connect)
    {

        // ServerConnection serverConnection = null;


        StringBuilder strbild = new StringBuilder();

        Guid ref_user = new Guid("8ee861f3-a434-4c24-b969-0896730b93ea");

        Guid login = new Guid("42c81c2b-7354-46aa-9547-0f1a93e9d4e1");

        List<string> userlist = new List<string> { "Сотрудник", "Администратор", "Отключенный пользователь" };
        List<string> grouplist = new List<string> { "Группа пользователей" };
        var str = userlist;
        var ObjUser = GetAnalog(Context.Connection, "tip", ref_user, str);
        var ObjUser2 = GetAnalog(connect, "tip", ref_user, str);
        IEnumerable<string> users1 = new List<string>();
        IEnumerable<string> users2 = new List<string>();
        IEnumerable<User> usersref1 = new List<User>();
        IEnumerable<User> usersref2 = new List<User>();
        if (str.Contains("Сотрудник"))
        {
            users1 = from item in ObjUser select ((User)item).Login.ToString();
            users2 = from item in ObjUser2 select ((User)item).Login.ToString();
        }

        if (str.Contains("Группа пользователей"))
        {
            users1 = from item in ObjUser select ((UsersGroup)item).FullName.ToString();
            users2 = from item in ObjUser2 select ((UsersGroup)item).FullName.ToString();
        }

        strbild.AppendLine($"----База1--WORK---отсутствуют на макете--{ObjUser.Count.ToString()}");

        foreach (var item in users1)
        {
            if (!users2.Contains(item.ToString()))
            {
                strbild.AppendLine(item.ToString());
            }

        }

        strbild.AppendLine($"----------База2------{ObjUser2.Count.ToString()}");
        foreach (var item in users2)
        {
            if (!users1.Contains(item.ToString()))
                strbild.AppendLine(item.ToString());

            if (item.Equals("nagibinadi"))
            {
                Message("", item);
                var ObjUsercreate = GetAnalog(connect, "login", ref_user, new List<string> { "nagibinadi" });


                /*    var ObjUserCreate = GetAnalog(Context.Connection, ref_user, str);
                 Объект newuser = СоздатьОбъект("Группы и пользователи", "Сотрудник");
                 newuser["Логин"] = itemD.Value.login;
                 newuser["Фамилия"] = itemD.Value.surname;
                 newuser["Имя"] = itemD.Value.first_name;
                 newuser["Короткое имя"] = itemD.Value.short_names;
                 newuser["Учётная запись Windows"] = sidAD; //"CORP\\" + itemD.Value.login;
                 if (itemD.Value.tel != null)
                     newuser["Телефон внутренний"] = itemD.Value.tel;
                 if (itemD.Value.email != null)
                     newuser["Электронная почта"] = itemD.Value.email;
                 newuser.Сохранить();*/



            }


        }









        Message("Пользователи отсутствуют", strbild);

        //(User)ObjUser.Select(xs1 => xs1)
        //  var propertylist = user_item.GetType().GetProperties().Select(property => property.Name).ToList();
        //from item in listObj select item.GetObjectValue("Номер").ToString();


        //  Message("serv1", ObjUser.Count.ToString());
        //  Message("serv2", ObjUser2.Count.ToString());

        /*  foreach (var obj in ObjUser)
          {
              var type1 = obj.GetType();
              //var type1 = obj.Ty
          }*/

    }


       

    
}
