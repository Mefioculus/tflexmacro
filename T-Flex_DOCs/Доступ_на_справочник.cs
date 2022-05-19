using System;

using System.Linq;

using TFlex.DOCs.Model.Macros;

using TFlex.DOCs.Common;

using TFlex.DOCs.Model.Access;

using TFlex.DOCs.Client.ViewModels.References.SelectionDialogs;

using TFlex.DOCs.Client;

public class RoomReservationMacro : MacroProvider

{

    public RoomReservationMacro(MacroContext context)

        : base(context)

    {

    }

 

    public override void Run()

    {

    }

 

    public void НазначитьДоступНаСправочники()

    {

        var selectReferenceVM = new SelectReferenceViewModel(Context.Connection, true);

        if (!ApplicationManager.OpenDialog(selectReferenceVM))

            return;

 

        var allAccesses = AccessGroup.GetGroups(Context.Connection); // Получаем все доступы.

        var referencesAccesses = allAccesses.FindAll(a => a.Type.IsReference).ToArray();

        var objectsAccesses = allAccesses.FindAll(a => a.Type.IsObject).ToArray();

 

        var полеДоступНаСправочники = "Доступ на справочники";

        var полеДоступНаОбъекты = "Доступ на объекты";

        var диалогВвода = СоздатьДиалогВвода("Выбор устанавливаемых доступов");

        диалогВвода.Высота = 80;

        диалогВвода.ДобавитьВыборИзСписка(полеДоступНаСправочники, null, true, referencesAccesses);

        диалогВвода.ДобавитьВыборИзСписка(полеДоступНаОбъекты, null, true, objectsAccesses);

        if (!диалогВвода.Показать())

            return;

 

        var referenceAccess = (AccessGroup)диалогВвода.Значение(полеДоступНаСправочники);

        var objectsAccess = (AccessGroup)диалогВвода.Значение(полеДоступНаОбъекты);

 

        ДиалогОжидания.Показать("Установка доступов", true);

        var references = selectReferenceVM.SelectedObjects;

        foreach (var reference in references)

        {

            if (!ДиалогОжидания.СледующийШаг(string.Format("Установка доступов на справочник {0}", reference)))

                return;

 

            var referenceAccessManager = AccessManager.GetReferenceAccess(reference);

            SetAccess(referenceAccessManager, referenceAccess, AccessType.Reference);

 

            var objectsAccessManager = AccessManager.GetReferenceObjectAccess(reference);

            SetAccess(objectsAccessManager, objectsAccess, AccessType.Object);

        }

        ДиалогОжидания.Скрыть();

    }

 

    private void SetAccess(AccessManager accessManager, AccessGroup accessGroup, AccessType type)

    {

        var allUserAccess = accessManager.FirstOrDefault(a => a.Owner == null && a.Access != null && a.Access.Type == type);

        if (allUserAccess == null)

        {

            accessManager.SetAccess(null, accessGroup);

        }

        else

        {

            accessManager.SetAccess(0, allUserAccess.Owner, accessGroup, allUserAccess.Access, null,

                allUserAccess.CommandType, allUserAccess.AccessTypeID, allUserAccess.AccessDirection, null);

        }

 

        accessManager.Save();

    }

}
