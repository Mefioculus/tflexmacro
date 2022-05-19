//TFlex.DOCs.Common.dll

using System;
using TFlex.DOCs.Model.Macros;
using TFlex.DOCs.Model.Access;
using TFlex.DOCs.Model.References;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using TFlex.DOCs.Model.Macros.ObjectModel.Types.Extensions;
using TFlex.DOCs.Model.References.Users;
using TFlex.DOCs.Model.Macros.ObjectModel;

public class SetAccess : MacroProvider
{
    public SetAccess(MacroContext context)
        : base(context)
    {
    	
    #if DEBUG
        System.Diagnostics.Debugger.Launch();
        System.Diagnostics.Debugger.Break();
    #endif
    }

    public override void Run()
    {
        foreach (var ТП in ВыбранныеОбъекты)
        {
            ReferenceObject TP = (ReferenceObject)ТП;
            НазначитьДоступНаОбъектСправочника(TP);
        }
    }

    public void НазначитьДоступНаОбъектСправочника(ReferenceObject TP)
    {
        var accessManager = AccessManager.GetReferenceObjectAccess(TP);
        accessManager.IsInherit = true;
        accessManager.Save();
    }
    
    
    
    public void setStatusRight()
    {
          StringBuilder str = new StringBuilder();
        var allRef = Context.Connection.ReferenceCatalog.GetReferences();
        var allAccesses = AccessGroup.GetGroups(Context.Connection); // Получаем все доступы.
        var referencesAccesses = allAccesses.FindAll(a => a.Type.IsReference).ToArray();
        var ObjectAccesses = allAccesses.FindAll(a => a.Type.IsObject).ToArray();
        var RefAccesses = allAccesses.FindAll(a => a.Type.IsReference).ToArray();
        int countfor = allRef.Count;
        for (int i = 0; i < countfor; i++)
        {
            var referenceAccessManager2 = AccessManager.GetReferenceAccess(allRef[i]);
            //AccessManager.
            var objectsAccessManager2 = AccessManager.GetReferenceObjectAccess(allRef[i]);

            foreach (var item in referenceAccessManager2.ToList())
            {
                if (item.Owner == null && item.Access.Name == "Просмотр")
                {
                    SetAccessRef(referenceAccessManager2, RefAccesses[0], AccessType.Reference);
                    
                }
                if (item.Owner == null)
                    str.AppendLine($"Справочник {allRef[i].Name} Все пользователи {item.Access.Name}");
                else
                    str.AppendLine($"Справочник {allRef[i].Name} {item.Owner.ToString()} {item.Access.Name}");
            }

            foreach (var item in objectsAccessManager2.ToList())
            {
                if (item.Owner == null && item.Access.Name == "Редакторский")
                {
                    SetAccessRef(objectsAccessManager2, ObjectAccesses[2], AccessType.Object);                    
                }

                if (item.Owner == null)
                str.AppendLine($"Объекты Справочника {allRef[i].Name} Все пользователи {item.Access.Name}");
                else
                str.AppendLine($"Объекты Справочника {allRef[i].Name} {item.Owner.ToString()} {item.Access.Name}");

            }
            str.AppendLine("-----------------------");
        }

        Message("",str.ToString());

    }



    public void SetAccessRef(AccessManager accessManager, AccessGroup accessGroup, AccessType type)

    {
        var allUserAccess = accessManager.FirstOrDefault(a => a.Owner == null && a.Access != null && a.Access.Type == type);
        RefObj user = FindObject("Группы и пользователи", "[Наименование] = 'Администраторы'");
        if (allUserAccess != null)
        {
            if (type == AccessType.Reference)

            {
                AccessManagerExtensions.Clear(accessManager, type);
                accessManager.SetAccess((UsersGroup)user, accessGroup);
            }

            if (allUserAccess.Owner == null && type == AccessType.Object)
            {
                accessManager.SetAccess(0, allUserAccess.Owner, accessGroup, allUserAccess.Access, null,
                    allUserAccess.CommandType, allUserAccess.AccessTypeID, allUserAccess.AccessDirection, null);
            }
        }
        accessManager.Save();
    }
    
    
    
    public void setStatusRight5()
    {
        StringBuilder str = new StringBuilder();
        var allRef = Context.Connection.ReferenceCatalog.GetReferences();
        var allAccesses = AccessGroup.GetGroups(Context.Connection); // Получаем все доступы.
        var referencesAccesses = allAccesses.FindAll(a => a.Type.IsReference).ToArray();
        var ObjectAccesses = allAccesses.FindAll(a => a.Type.IsObject).ToArray();
        var RefAccesses = allAccesses.FindAll(a => a.Type.IsReference).ToArray();
        int countfor =  allRef.Count;
        for (int i = 0; i < countfor; i++)
        {
            var referenceAccessManager2 = AccessManager.GetReferenceAccess(allRef[i]);
            //AccessManager.
            var objectsAccessManager2 = AccessManager.GetReferenceObjectAccess(allRef[i]);

            foreach (var item in referenceAccessManager2.ToList())
            {

                SetAccessRef5(referenceAccessManager2, RefAccesses[0], AccessType.Reference);
                /*  if (item.Owner.ToString() == "Администраторы" && item.Access.Name == "Просмотр")
                  {


                  }*/
                if (item.Owner == null)
                    str.AppendLine($"Справочник {allRef[i].Name} Все пользователи {item.Access.Name}");
                else
                    str.AppendLine($"Справочник {allRef[i].Name} {item.Owner.ToString()} {item.Access.Name}");
            }

            foreach (var item in objectsAccessManager2.ToList())
            {
                if (item.Owner == null && item.Access.Name == "Редакторский")
                {
                    SetAccessRef5(objectsAccessManager2, ObjectAccesses[2], AccessType.Object);
                }

                if (item.Owner == null)
                    str.AppendLine($"Объекты Справочника {allRef[i].Name} Все пользователи {item.Access.Name}");
                else
                    str.AppendLine($"Объекты Справочника {allRef[i].Name} {item.Owner.ToString()} {item.Access.Name}");

            }
            str.AppendLine("-----------------------");
        }

        Message("", str.ToString());

    }



    public void SetAccessRef5(AccessManager accessManager, AccessGroup accessGroup, AccessType type)

    {
        var allUserAccess = accessManager.FirstOrDefault(a => a.Owner == null && a.Access != null && a.Access.Type == type);
        RefObj user = FindObject("Группы и пользователи", "[Наименование] = 'Администраторы'");
        if (allUserAccess == null)
        {
            if (type == AccessType.Reference)

            {
                AccessManagerExtensions.Clear(accessManager, type);
                accessManager.SetAccess(null, accessGroup);
                /*  accessManager.SetAccess(0, allUserAccess.Owner, accessGroup, allUserAccess.Access, null,
                  allUserAccess.CommandType, allUserAccess.AccessTypeID, allUserAccess.AccessDirection, null);*/
                // AccessManagerExtensions.Clear(accessManager, type);
                // accessManager.SetAccess((UsersGroup)user, accessGroup);
            }

          /*  if (allUserAccess.Owner == null && type == AccessType.Object)
            {
                accessManager.SetAccess(0, allUserAccess.Owner, accessGroup, allUserAccess.Access, null,
                    allUserAccess.CommandType, allUserAccess.AccessTypeID, allUserAccess.AccessDirection, null);
            }*/
        }
        accessManager.Save();
    }



}

