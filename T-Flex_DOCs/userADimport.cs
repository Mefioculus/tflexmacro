using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using TFlex.DOCs.Model.Macros;
using TFlex.DOCs.Model.Macros.ObjectModel;
using System.Text;
using System.DirectoryServices;
using System.DirectoryServices.AccountManagement;
using TFlex.DOCs.Model.References.Users;
using TFlex.DOCs.Model.Search;
using TFlex.DOCs.Model.Structure;
using TFlex.DOCs.Model.References.Macros;

public class Macro : MacroProvider
{

    public Macro(MacroContext context)
        : base(context)
    {
        /* var macro = Context[nameof(MacroKeys.MacroSource)] as CodeMacro;
         if (macro.DebugMode.Value == true)
         {
             System.Diagnostics.Debugger.Launch();
             System.Diagnostics.Debugger.Break();
         }*/


#if DEBUG
        System.Diagnostics.Debugger.Launch();
        System.Diagnostics.Debugger.Break();
#endif
    }

    public override void Run()
    {
        //GetActiveDirectory2();
    }

    public void ActiveDirectoryTest()
    {
        //"Группа Active Directory" = "e2d5b73b-36ee-4636-bb49-a5071bfcbd7e"
        Объект SelectGroupDocs = ТекущийОбъект;
        //Сообщение("",$"{SelectGroupDocs.ToString()} {SelectGroupDocs.GetType().ToString()} {SelectGroupDocs.Тип} 2");
        var p = SelectGroupDocs.Параметр["Наименование"].ToString();
        var sid = SelectGroupDocs.Параметр["Группа Active Directory"].ToString();
        //string sid = "S-1-5-21-789336058-507921405-854245398-9938";
        if (sid.Length > 0)
        {
            string namegroup = new System.Security.Principal.SecurityIdentifier(sid).Translate(typeof(System.Security.Principal.NTAccount)).ToString();
            //Сообщение("", p.ToString());
            //Сообщение("", namegroup.ToString() );
            var child_user = SelectGroupDocs.ВсеДочерниеОбъекты;
            //Сообщение("", child_user.Count.ToString());
            Dictionary<string, Объект> userDocsGroup = new Dictionary<string, Объект>();
            Dictionary<string, ADuser> userADGroup = new Dictionary<string, ADuser>();
            GetActiveDirectory2(namegroup.Replace("CORP\\", ""), userADGroup);

            StringBuilder struserAD = new StringBuilder();
            foreach (var item in userADGroup)
            {
                struserAD.Append(item.Key + "\r\n");
            }
            Message("", $"Группа {namegroup} содержит: {userADGroup.Count} пользователей {struserAD.ToString()}");
            //CORP\Сотрудники - Служба_PLM
            /*       ADuser user = new ADuser();
                   user.login = "PSS1";
                   user.first_name = "Имя";
                   user.surname = "Фамилия";
                   user.short_names = "Фамилия И.И";
                   user.email = "поята@mail.com";
                   userADGroup.Add("PSS1", user);
                   user = new ADuser();
                   user.login = "PSS2";
                   user.first_name = "Имя1";
                   user.surname = "Фамилия1";
                   user.short_names = "Фамилия1 И.И";
                   user.email = "почта1@mail.com";
                   userADGroup.Add("PSS2", user);*/

            /*user = new ADEser();
            user.login = "PSS";
            user.first_name = "Имя";
            user.surname = "Фамилия";
            user.short_names = "Фамилия И.И";
            user.email = "почта@mail.com";
            userADGroup.Add("PSS", user);*/

            /*        user = new ADuser();
                    user.login = "PSS4";
                    user.first_name = "Имя4";
                    user.surname = "Фамилия4";
                    user.short_names = "Фамилия4 И.И";
                    user.email = "почта4@mail.com";
                    userADGroup.Add("PSS4", user);*/

            /*user = new ADEser();
            user.login = "PSS5";
            user.first_name = "Имя5";
            user.surname = "Фамилия5";
            user.short_names = "Фамилия5 И.И";
            user.email = "почта5@mail.com";
            userADGroup.Add("PSS5", user);*/
            //Dictionary<string, Объект> test = (Dictionary<string, Объект>)(from t in child_user select (t.Параметр["Логин"].ToString(), t));

            foreach (var item in child_user)
            {
                string login = item.Параметр["Логин"].ToString();
                userDocsGroup.Add(login, item);
                if (!userADGroup.ContainsKey(login))
                {
                    //Сообщение("userref объекта нет в ADUser удалить из группы", login.ToString());
                    var connect = item.РодительскиеПодключения;
                    foreach (var iconnect in connect)
                    {
                        if (iconnect.РодительскийОбъект.Guid == SelectGroupDocs.Guid)
                            iconnect.Удалить();
                    }
                }
            }

            /*       foreach (var item in child_user)
                {
                    string login = item.Параметр["Логин"].ToString();
                    //userDocsGroup.Add(login, item);
                    Сообщение("login", login.ToString());*/

            foreach (var itemD in userADGroup)
            {
                Объект userref = НайтиОбъект("Группы и пользователи", Условие("Логин", "=", itemD.Key));
                // Создаем пользавателя если его нет и добавляем в группу
                if (userref is not null)
                {
                    if (userDocsGroup.ContainsKey(itemD.Key))
                    {
                        //Сообщение("userref объект уже есть в группе", userref.ToString());
                    }

                    else
                    {
                        // Сообщение("userref добавлен в группу", userref.ToString());
                        //userref.Изменить();
                        Подключение подключение = СоздатьПодключение(SelectGroupDocs, userref);
                        подключение.Сохранить();
                        //userref.ЗадатьПользователяВладельца(объект);
                        //userref.Сохранить();
                    }
                }
                else
                {
                    string sidAD = new System.Security.Principal.NTAccount(itemD.Value.login).Translate(typeof(System.Security.Principal.SecurityIdentifier)).ToString();

                    Объект GroupAllUser = НайтиОбъект("Группы и пользователи", Условие("ID", "=", 6)); //maket
                    //Объект GroupAllUser = НайтиОбъект("Группы и пользователи", Условие("ID", "=", 3)); //server tflex-docs
                    Объект newuser = СоздатьОбъект("Группы и пользователи", "Сотрудник", SelectGroupDocs);

                    newuser["Логин"] = itemD.Value.login;
                    newuser["Фамилия"] = itemD.Value.surname;
                    newuser["Имя"] = itemD.Value.first_name;
                    newuser["Короткое имя"] = itemD.Value.short_names;
                    newuser["Учётная запись Windows"] = sidAD; //"CORP\\" + itemD.Value.login;


                    if (itemD.Value.email != null)
                    {
                        string MailSettings = $"<SmtpServerSettings ServerName=\"zimbra.corp.aeroem.ru\" Login=\"{itemD.Value.login}\" Name=\"{itemD.Value.short_names}\" EMail=\"{itemD.Value.email}\" UseDOCsEMail=\"True\" UseDOCsName=\"True\" />";
                        newuser["Электронная почта"] = itemD.Value.email;
                        ((User)newuser).MailSettings.Value = MailSettings;
                        //var sentmailparam = newuser.MailSendType;                    
                    }



                    if (itemD.Value.tel != null)
                        newuser["Телефон внутренний"] = itemD.Value.tel;

                    newuser.Сохранить();
                    Подключение подключение = СоздатьПодключение(GroupAllUser, newuser);
                    подключение.Сохранить();
                }
            }

        }
    }
    public Dictionary<string, ADuser> GetActiveDirectory2(string groupname, Dictionary<string, ADuser> usersgroup)
    {
        //Dictionary<string, ADuser> usersgroup = new Dictionary<string, ADuser>();
        PrincipalContext pc = new PrincipalContext(ContextType.Domain);
        var name = System.Security.Principal.WindowsIdentity.GetCurrent().Name;
        //Console.WriteLine(name);
        //Сотрудники-Служба_PLM
        //Сотрудники-АСУП
        GroupPrincipal group = GroupPrincipal.FindByIdentity(pc, groupname);
        List<Principal> ListUsers = new List<Principal>(GetActiveDirectory3(group));

        StringBuilder stringBuilder = new StringBuilder();
        foreach (var g in ListUsers)
        {
            var user_item = (UserPrincipal)g;
            Console.WriteLine($"{g.ToString()}  {g.GetType()}");
            ADuser aduser = new ADuser();
            aduser.login = user_item.SamAccountName;
            aduser.short_names = user_item.Name;
            aduser.first_name = user_item.GivenName;
            aduser.surname = user_item.Surname;
            aduser.email = user_item.UserPrincipalName;
            aduser.tel = user_item.VoiceTelephoneNumber;
            if (!usersgroup.ContainsKey(aduser.login))
            {
                if (aduser.login != null && aduser.short_names != null && aduser.surname != null)
                    usersgroup.Add(aduser.login, aduser);
                else
                    Сообщение("", $"У пользователя {aduser.login} отсутсует обязательное поле");
            }
            //onsole.WriteLine(.ToString());
            stringBuilder.Append(g.ToString());
        }
        //Message("",stringBuilder.ToString());
        return usersgroup;
    }

    public List<Principal> GetActiveDirectory3(GroupPrincipal group)
    {
        List<Principal> ListUsers = new List<Principal>();

        PrincipalContext pc = new PrincipalContext(ContextType.Domain);
        //group = GroupPrincipal.FindByIdentity(pc, "Сотрудники-АСУП");
        //group.m;
        var MemberPrincipal = group.GetMembers();
        var filtergroup = from t in MemberPrincipal where (t is System.DirectoryServices.AccountManagement.GroupPrincipal) select t;
        var filteruser = from t in MemberPrincipal where (t is System.DirectoryServices.AccountManagement.UserPrincipal) select t;
        ListUsers.AddRange(filteruser);
        //Console.WriteLine($"{group.ToString()} всего {MemberPrincipal.Count()} группы {filtergroup.Count()} пользователи {filteruser.Count()}");
        if (filtergroup.Count() > 0)
        {
            foreach (var itemgroup in filtergroup)
            {
                GetActiveDirectory3((GroupPrincipal)itemgroup);
            }
        }
        return ListUsers;
    }

    private void AddToUsersGroup(User user, string UsersGroupName)
    {
        if (user is null)
            return;

        var classUserGroup = Context.Connection.References.Users.Classes.AllClasses.Find(Guids.ClassUserGroup);
        if (classUserGroup is null)
            return;

        //string UsersGroupName = "Сотрудник";

        // Находим группу пользователей "Отключенные пользователи"
        var userParameterGroup = Context.Connection.References.Users.ParameterGroup;
        using var filter = new Filter(userParameterGroup);
        filter.Terms.AddTerm(userParameterGroup[UserReferenceObject.Fields.FullName],
            ComparisonOperator.Equal, UsersGroupName);
        filter.Terms.AddTerm(userParameterGroup[SystemParameterType.Class],
            ComparisonOperator.IsInheritFrom, classUserGroup);

        var UsersGroup = Context.Connection.References.Users.Find(filter).FirstOrDefault();
        if (UsersGroup is null) // Если группа не существует, то создаем
        {
            UsersGroup = Context.Connection.References.Users.CreateReferenceObject(classUserGroup);
            UsersGroup[UserReferenceObject.Fields.FullName].Value = UsersGroupName;
            UsersGroup.EndChanges();
        }

        // Подключаем отключенного пользователя в группу
        var link = UsersGroup.CreateChildLink(user);
        link.EndChanges();
    }

    /// <summary>
    /// Добавляет текущего пользователя в группу Сотрудники
    /// </summary>
    public void AddUserGroupAllUdsers()
    {
        Объект сотрудник = ТекущийОбъект;
        //Объект GroupAllUser = НайтиОбъект("Группы и пользователи", Условие("ID", "=", 6)); //maket
        сотрудник.Изменить();
        //сотрудник.РодительскийОбъект = GroupAllUser;
        сотрудник.Сохранить();
        AddToUsersGroup((User)сотрудник, "Сотрудники");
    }

    public void SetLoginAD()
    {
        Объект user = ТекущийОбъект;
        string login = user["Логин"];
        var sidUser = user["Учётная запись Windows"].ToString();
        //System.Security.Principal.NTAccount f = new System.Security.Principal.NTAccount(login);
        //SecurityIdentifier s = (SecurityIdentifier)f.Translate(typeof(SecurityIdentifier));

        string sidAD = "";
        string getlogin = "";
        try
        {
            sidAD = new System.Security.Principal.NTAccount(login).Translate(typeof(System.Security.Principal.SecurityIdentifier)).ToString();

            try
            {
                getlogin = new System.Security.Principal.SecurityIdentifier(sidUser).Translate(typeof(System.Security.Principal.NTAccount)).ToString();
            }
            catch
            {
                Объекты users = НайтиОбъекты("Группы и пользователи", Условие("Логин", "=", $"{login}")); //maket
                if (users.Count == 1)
                {
                    user.Изменить();
                    user["Учётная запись Windows"] = sidAD;
                    user.Сохранить();
                }
                else
                    Message("", $"В базе несколько записей с логином {login}");
                //getlogin = "Неверный SID";
            }

        }
        catch
        {
            //Message("", "Пользователя не существует");
        }



        //Message("", sidUser);
        //Message("", sidAD);
        //Message("", getlogin);
    }


    public void SetLoginAD(Объект user)
    {
        //Объект user = ТекущийОбъект;
        string login = user["Логин"];
        var sidUser = user["Учётная запись Windows"].ToString();
        //System.Security.Principal.NTAccount f = new System.Security.Principal.NTAccount(login);
        //SecurityIdentifier s = (SecurityIdentifier)f.Translate(typeof(SecurityIdentifier));

        string sidAD = "";
        string getlogin = "";
        try
        {
            sidAD = new System.Security.Principal.NTAccount(login).Translate(typeof(System.Security.Principal.SecurityIdentifier)).ToString();

            try
            {
                getlogin = new System.Security.Principal.SecurityIdentifier(sidUser).Translate(typeof(System.Security.Principal.NTAccount)).ToString();
            }
            catch
            {
                Объекты users = НайтиОбъекты("Группы и пользователи", Условие("Логин", "=", $"{login}")); //maket
                if (users.Count == 1)
                {
                    user.Изменить();
                    user["Учётная запись Windows"] = sidAD;
                    user.Сохранить();
                }
                else
                    Message("", $"В базе несколько записей с логином {login}");
                //getlogin = "Неверный SID";
            }

        }
        catch
        {
            //Message("", "Пользователя не существует");
        }

    }



    public void SetMailSelectUser()
    {
        Объекты selectUsers = ВыбранныеОбъекты;
        foreach (var item in selectUsers)

        {
            SetLoginAD(item);
            SetMail(item);

        }
    }

        public void SetMail(Объект user)
    {
        //Объект user = ТекущийОбъект;
        //User user = (User)ТекущийОбъект;
        string login = user["Логин"];
        var sidUser = user["Учётная запись Windows"].ToString();
        var email = user["Электронная почта"];
        // System.Security.Principal.NTAccount f = new System.Security.Principal.NTAccount(login);
        //SecurityIdentifier s = (SecurityIdentifier)f.Translate(typeof(SecurityIdentifier));
        var userAD = GetAdUsers3(login);
    
        //user["Электронная почта"] = userAD.EmailAddress;
        //string sidAD = "";
        //string getlogin = "";
        /*try
        {*/
        //sidAD = new System.Security.Principal.NTAccount(login).Translate(typeof(System.Security.Principal.SecurityIdentifier)).ToString();
        //var user_nt = new System.Security.Principal.NTAccount(login);
        string query = @"И [Относительный путь] содержит 'Архив ОГТ\ogtMK' И [Относительный путь] не содержит 'Старые версии МК, КЭ.ТИ,СТО' И [Относительный путь] не содержит 'КЭ'";
        var пользователи = НайтиОбъекты("Группы и пользователи", "[Тип] Входит в список 'Администратор, Сотрудник, Отключенный пользователь'");

        /*foreach (var item in пользователи)
        { 
        
        }*/

        //User user = (User)ТекущийОбъект;
        /*        if (email == null)
                {*/
        if (userAD != null && (userAD.EmailAddress != null || userAD.UserPrincipalName!=null))
        {
            if (userAD.EmailAddress != null)
                user["Электронная почта"] = userAD.EmailAddress.ToString();
            else if (userAD.UserPrincipalName.ToString().Contains("@"))
                user["Электронная почта"] = userAD.UserPrincipalName.ToString();

            user.Сохранить();
      
        User user_cl = (User)user;
            user_cl.BeginChanges();
            var s = user_cl.GetObjectValue("Электронная почта");
            var m = user_cl.MailSettings;
            var m3 = user_cl.ParameterValues[23];
            var m4 = user_cl.ParameterValues;
            var m5 = user_cl.ParameterValues[27];
          //  user_cl.ParameterValues[27].Value  = userAD.EmailAddress.ToString();

        //user_cl.GetObjectValue("Электронная почта")= userAD.EmailAddress;
        //GetObjectValue("Name")
        //ParameterValues("Электронная почта")
        //em= userAD.EmailAddress;
        //user_cl.

        string MailSettings = $"<SmtpServerSettings ServerName=\"zimbra.corp.aeroem.ru\" Login=\"{user_cl.Login}\" Name=\"{user_cl.FullName}\" EMail=\"{user_cl.Email}\" UseDOCsEMail=\"True\" UseDOCsName=\"True\" />";

            user_cl.MailSettings.Value = MailSettings;
            var sentmailparam = user_cl.MailSendType;
            //user.MailSendType.Value = 2;
            user_cl.EndChanges();
        }
        //  }
        //RefreshReferenceWindow();
    }


    public void RuncopyLink()
    {
        var users = ВыбранныеОбъекты;

        foreach (var item in users)
                copyLink("null1" , item["Логин"]);
    }

    public void RuncopyLink(string login, string login_copy)
    {
        copyLink(login, login_copy);
    }


    public void copyLink(string login, string copytologin)
    {

        Объект LoginUserRef = НайтиОбъект("Группы и пользователи", Условие("Логин", "=", String.Format("{0}", login)));
        Объект CopyToLoginRef = НайтиОбъект("Группы и пользователи", Условие("Логин", "=", String.Format("{0}", copytologin)));


        Подключения linkorig = ((Объект)LoginUserRef).РодительскиеПодключения;
        Подключения linkcopy = ((Объект)CopyToLoginRef).РодительскиеПодключения;


        foreach (var item in linkcopy)
        {
            item.Удалить();
        }

        //link9.Add(link5);

        foreach (var item in linkorig)
        {
            Подключение link = СоздатьПодключение(item.РодительскийОбъект, CopyToLoginRef);
            link.Сохранить();
        }

        //link_chald3.Clear();

        //Message("",$"{link_chald.ToString}")
    }


    public void testuser()
    {
        System.Diagnostics.Debugger.Launch();
        System.Diagnostics.Debugger.Break();
        User user = (User)ТекущийОбъект;
        user.BeginChanges();
        string MailSettings = $"<SmtpServerSettings ServerName=\"zimbra.corp.aeroem.ru\" Login=\"{user.Login}\" Name=\"{user.FullName}\" EMail=\"{user.Email}\" UseDOCsEMail=\"True\" UseDOCsName=\"True\" />";
        user.MailSettings.Value = MailSettings;
        //ewuser["Электронная почта"]
        var sentmailparam = user.MailSendType;
        //user.MailSendType.Value = 2;
        user.EndChanges();
      //  RefreshReferenceWindow();
    }

    public UserPrincipal GetAdUsers3(string login)
    {
        PrincipalContext pc = new PrincipalContext(ContextType.Domain);
        UserPrincipal usr = UserPrincipal.FindByIdentity(pc,
                                                   IdentityType.SamAccountName,
                                                   login);
        return usr;
    }

}
public static class Guids
{
    internal static readonly Guid ClassUserGroup = new("8C34B00D-D26B-4F7D-88AB-4E4A07595669");
}

public class ADuser
{
    public string login;
    public string short_names;
    public string first_name;
    public string surname;
    public string email;
    public string tel;
}






