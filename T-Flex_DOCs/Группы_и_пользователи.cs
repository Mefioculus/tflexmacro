using System;
using System.Linq;
using TFlex.DOCs.Model.Macros;
using TFlex.DOCs.Model.References;
using TFlex.DOCs.Model.References.Users;
using TFlex.DOCs.Model.Search;
using TFlex.DOCs.Model.Structure;

public class Macro : MacroProvider
{
    public Macro(MacroContext context)
        : base(context)
    {
    }

    public override void Run()
    {
    }

    /// <summary>
    /// Обработчик события "Сохранение объекта"
    /// </summary>
    public void СохранениеОтключенногоПользователя()
    {
        if (Context.ReferenceObject is not User user || !user.Class.IsDisconnectedUser || user.IsPrototype)
            return;

        // Сбрасываем электронную почту, учетную запись Windows, служебный телефон
        user.Email.Value = String.Empty;
        user.Sid.Value = String.Empty;
        user.BusinessPhone.Value = String.Empty;

        // Удаляем все подключения пользователя
        var parentHierarchyLinks = user.Parents.GetHierarchyLinks();
        Reference.Delete(parentHierarchyLinks);

        // Подключаем отключенного пользователя в группу пользователей "Отключенные пользователи"
        AddToDisconnectedUsersGroup(user);

        RefreshReferenceWindow();
    }

    private void AddToDisconnectedUsersGroup(User user)
    {
        if (user is null)
            return;

        var classUserGroup = Context.Connection.References.Users.Classes.AllClasses.Find(Guids.ClassUserGroup);
        if (classUserGroup is null)
            return;

        string disconnectedUsersGroupName = "Отключенные пользователи";

        // Находим группу пользователей "Отключенные пользователи"
        var userParameterGroup = Context.Connection.References.Users.ParameterGroup;
        using var filter = new Filter(userParameterGroup);
        filter.Terms.AddTerm(userParameterGroup[UserReferenceObject.Fields.FullName],
            ComparisonOperator.Equal, disconnectedUsersGroupName);
        filter.Terms.AddTerm(userParameterGroup[SystemParameterType.Class],
            ComparisonOperator.IsInheritFrom, classUserGroup);

        var disconnectedUsersGroup = Context.Connection.References.Users.Find(filter).FirstOrDefault();
        if (disconnectedUsersGroup is null) // Если группа не существует, то создаем
        {
            disconnectedUsersGroup = Context.Connection.References.Users.CreateReferenceObject(classUserGroup);
            disconnectedUsersGroup[UserReferenceObject.Fields.FullName].Value = disconnectedUsersGroupName;
            disconnectedUsersGroup.EndChanges();
        }

        // Подключаем отключенного пользователя в группу
        var link = disconnectedUsersGroup.CreateChildLink(user);
        link.EndChanges();
    }

    private static class Guids
    {
        internal static readonly Guid ClassUserGroup = new("8C34B00D-D26B-4F7D-88AB-4E4A07595669");
    }
}
