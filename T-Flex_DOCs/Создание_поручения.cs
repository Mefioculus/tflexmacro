/* Дополнительные ссылки
TFlex.DOCs.Model.Office.dll
*/

using System;
using System.Collections.Generic;
using System.Linq;
using TFlex.DOCs.Model.Macros;
using TFlex.DOCs.Model.Macros.ObjectModel;
using TFlex.DOCs.Model.References.RecordControlCards;
using TFlex.DOCs.Model.References.Users;


public class CreateOrdersMacro : MacroProvider
{
    // Guid связи 'Подписал/Кому' справочника 'Регистрационно-контрольные карточки'
    private static readonly Guid RcUsersLinkGuid = new Guid("a3746fbd-7bda-4326-8e4b-f6bbefc2ecce");

    public CreateOrdersMacro(MacroContext context)
        : base(context)
    {
    }

    public override void Run()
    {
        var rccObject = Context.ReferenceObject as RCCReferenceObject;
        if (rccObject == null)
            return;        

        ЗакрытьДиалог(true, false);

        var задание = Канцелярия.СоздатьЗадание();
        задание.Тема = string.Format("{0} от {1} {2}",
            rccObject.RegistrationNumber,
            rccObject.RegistrationDate.Value.ToShortDateString(),
            rccObject.GetObject(RCCReference.RCCRegistrationJouranLink));

        задание.Текст = string.Format(
            "Вам назначена регистрационно-контрольная карточка с регистрационным номером \"{0}\" от {1}. Содержание: {2}",
            rccObject.RegistrationNumber,
            rccObject.RegistrationDate.Value.ToShortDateString(),
            rccObject.Text);

        задание.КонтрольнаяДата = null;
        задание.ДатаНачала = rccObject.RegistrationDate;
        задание.ДатаЗавершения = !rccObject.PlannedFinish.IsNull ? new DateTime?(rccObject.PlannedFinish.Value) : null;
        задание.ДобавитьОбъектВложения(Объект.CreateInstance(rccObject, Context));

        if (!rccObject.Class.IsInnerRecordControlCard && !rccObject.Class.IsOutgoingRecordControlCard)
        {
            var userReferenceObjects = rccObject.GetObjects(RcUsersLinkGuid).OfType<UserReferenceObject>().ToArray();
            var users = GetUserList(userReferenceObjects);

            foreach (var user in users)
                задание.ДобавитьИсполнителя(Объект.CreateInstance(user, Context));
        }

        ПоказатьДиалогСвойств(задание);
    }

    private List<User> GetUserList(UserReferenceObject[] userReferenceObjects)
    {
        var users = new List<User>();

        foreach (var item in userReferenceObjects)
        {
            if (!item.Class.IsGroup && !item.Class.IsProductionUnit)
            {
                if (item.Class.IsUser && !users.Contains((User) item))
                    users.Add((User) item);
            }
            else
            {
                var responsible = item.GetObject(UsersGroup.MailResponsible) as User;

                if (responsible != null)
                {
                    if (!users.Contains(responsible))
                        users.Add(responsible);
                }
                else
                {
                    foreach (var user in item.GetAllInternalUsers())
                    {
                        if (!users.Contains(user))
                            users.Add(user);
                    }
                }
            }
        }

        return users;
    }
}

