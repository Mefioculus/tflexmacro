/* Дополнительные ссылки:
System.Web.dll
*/

using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using TFlex.DOCs.Model.Macros;
using TFlex.DOCs.Model.Macros.ObjectModel;
using TFlex.DOCs.Model.References;
using TFlex.DOCs.Model.References.NormativeReferenceInfo.RequestsNsi;
using TFlex.DOCs.Model.References.Users;


/// <summary>
/// Макрос для обработки событий справочника "Заявки на НСИ"
/// </summary>
public class RequestsNsiMacro : MacroProvider
{
    /// <summary> Представляет уникальный идентификатор (GUID) параметра "Номер" справочника "Заявки на НСИ" </summary>
    public static readonly Guid Number = new Guid("000ee6c3-b040-4926-b0e6-7b934f21fe94");

    public RequestsNsiMacro(MacroContext context)
        : base(context)
    {
        if (Context.Connection.ClientView.HostName == "MOSINS")
            if (Вопрос("Хотите запустить в режиме отладки?"))
            {
                System.Diagnostics.Debugger.Launch();
                System.Diagnostics.Debugger.Break();
            }
    }

    /// <summary>
    /// Справочник "Заявки НСИ" - Событие "Изменение стадии объекта"
    /// </summary>
    public void СобытиеИзменениеСтадииЗаявки()
    {
        if (!(Context.ModelChangedArgs is ObjectStageChangedEventArgsBase changedEventArgs))
            return;

        var changedStageGuid = changedEventArgs.NewStage.Guid;
        var currentObj = Context.ReferenceObject;
        if (!StagesGuid.ObjectsEventStages.Contains(changedStageGuid) || currentObj == null || !currentObj.CanEdit)
            return;

        currentObj.BeginChanges();
        var dateNow = DateTime.Now;
        if (changedStageGuid == StagesGuid.ProcessedStageGuid)
        {
            currentObj[BidGuid.DateStartParameterGuid].Value = dateNow;
        }
        else if (changedStageGuid == StagesGuid.CompletedStageGuid)
        {
            currentObj[BidGuid.DateEndParameterGuid].Value = dateNow;
            var dateStart = currentObj[BidGuid.DateStartParameterGuid].GetDateTime();
            double count = GetWorkDaysCount(Context.Connection.ClientView.GetUser() as UserReferenceObject, dateStart, dateNow);
            currentObj[BidGuid.DurationParameterGuid].Value = count;
        }
        currentObj.EndChanges();
    }

    public void ВысчитатьДлительностьВыполненияЗаявки()
    {
        var bidReference = Context.Connection.ReferenceCatalog.Find(BidGuid.Reference)?.CreateReference();
        if (bidReference == null)
            return;

        bidReference.LoadSettings.AddParameters(BidGuid.DateStartParameterGuid);
        var objects = bidReference.Objects.ToHashSet();
        var saveset = new HashSet<ReferenceObject>();
        foreach (var bidObject in objects)
        {
            if (bidObject.SystemFields.Stage == null || !StagesGuid.ServerEventStages.Contains(bidObject.SystemFields.Stage.Guid))
                continue;

            var dateStart = bidObject[BidGuid.DateStartParameterGuid].GetDateTime();
            if (dateStart == null)
                continue;

            if (!bidObject.CanEdit)
                continue;

            bidObject.BeginChanges(false);
            bidObject[BidGuid.DurationParameterGuid].Value = GetWorkDaysCount(Context.Connection.ClientView.GetUser(), dateStart, DateTime.Now);
            saveset.Add(bidObject);
        }

        if (saveset.Count != 0)
            Reference.EndChanges(saveset);
    }

    /// <summary>
    /// Вернуть кол-во расчетных дней за указанный период.
    /// </summary>
    /// <param name="start">Дата начала периода.</param>
    /// <param name="end">Дата окончания периода.</param>
    /// <returns>Кол-во расчетных дней.</returns>
    private double GetWorkDaysCount(UserReferenceObject userObject, DateTime start, DateTime end)
    {
        var date = start;
        var spawn = end - start;
        var totalDays = spawn.TotalDays;
        while (date.Date <= end.Date)
        {
            if (!userObject.WorkTimeManager.GetWorkingIntervals(date).Any())
                totalDays -= 1;

            date = date.AddDays(1);
        }

        return totalDays;
    }

    /// <summary>
    /// Метод для добавления комментария о смене стадии в БП 'MDM. Управление Заявкой НСИ'
    /// </summary>
    public void ДобавитьКомментарийОсменеСтадии(Объект requestsNsiObj,
        Объект authorObj,
        string stageName,
        string userComment = null)
    {
        if (requestsNsiObj == null || authorObj == null || String.IsNullOrWhiteSpace(stageName))
            return;

        RequestNsiObject requestNsi = (ReferenceObject)requestsNsiObj;
        var author = (ReferenceObject)authorObj as UserReferenceObject;

        if (requestNsi == null || author == null)
            return;

        requestNsi.Main.Modify(o =>
        {
            var newComment = requestNsi.CreateComment();
            newComment.Theme.Value = $"Смена состояния на '{stageName}'";
            newComment.Content.Value = GetContentMessage(author, stageName, userComment);
            newComment.EndChanges();
        });
    }

    private static string GetContentMessage(UserReferenceObject author, string stageName, string userComment)
    {
        var dt = DateTime.Now;
        string dateString = $"{dt:dd-MM-yyyy} {dt:T}";

        string userAction =
            $"<div>{HttpUtility.HtmlEncode($"Пользователь '{author.FullName}' изменил состояние на '{stageName}' - '{dateString}'")}</div>";

        string comment = String.Empty;

        if (!String.IsNullOrWhiteSpace(userComment))
        {
            comment += "<br>";
            comment += "<div>";
            comment += "<div><i>Комментарий от пользователя:</i></div>";

            string[] separator = { Environment.NewLine, "\n" };
            var coomentsStrings = userComment.Split(separator, StringSplitOptions.RemoveEmptyEntries);

            foreach (string coomentsString in coomentsStrings)
                comment += $"<div>{HttpUtility.HtmlEncode(coomentsString)}</div>";

            comment += "</div>";
        }

        return $"<!DOCTYPE html><html><head></head><body>{userAction}{comment}</body></html>";
    }

    public void ПрисвоитьРегистрационныйНомер(Объект requestObj)
    {
        if (requestObj == null)
            return;

        var request = (ReferenceObject)requestObj;
        request.Modify(o => { request[Number].Value = $"НСИ-{request.SystemFields.Id}/{DateTime.Now.Year}"; }, true,
            false, true);
    }

    private class StagesGuid
    {
        /// <summary>
        /// Стадия обработка
        /// </summary>
        public static readonly Guid ProcessedStageGuid = new Guid("e294a1d1-bbc8-4230-b740-32d7b3e0a566");

        /// <summary>
        /// Стадия "Корректировка"
        /// </summary>
        public static readonly Guid AdjustmentStageGuid = new Guid("18df455a-0dc8-43a9-b256-c0fd6898df1b");

        /// <summary>
        /// Стадия выполнено
        /// </summary>
        public static readonly Guid CompletedStageGuid = new Guid("4e869516-ce04-4375-bd35-4a9795aa8428");

        /// <summary>
        /// Стадии которые будут учитыватьс на событии изменение стадии объекта
        /// </summary>
        public static readonly Guid[] ObjectsEventStages = new Guid[] { ProcessedStageGuid, CompletedStageGuid };

        /// <summary>
        /// Стадии которые будут учитываться при обработчике сервера
        /// </summary>
        public static readonly Guid[] ServerEventStages = new Guid[] { ProcessedStageGuid, AdjustmentStageGuid };
    }

    private class BidGuid
    {
        /// <summary>
        /// Справочник "Заявки НСИ"
        /// </summary>
        public static readonly Guid Reference = new Guid("9e21f099-73df-4948-b8c2-22ffd731a5e8");

        /// <summary>
        /// Параметр "Длительность обработки"
        /// </summary>
        public static readonly Guid DurationParameterGuid = new Guid("bf16795b-5fd3-4557-a2dd-bf5ae6d9b8d7");

        /// <summary>
        /// Параметр "Дата начала обработки"
        /// </summary>
        public static readonly Guid DateStartParameterGuid = new Guid("5ad9bc74-da34-4e6c-a981-8f6bca8c5b4f");

        /// <summary>
        /// Параметр "Дата начала обработки"
        /// </summary>
        public static readonly Guid DateEndParameterGuid = new Guid("2574d208-222c-47a6-8885-c11a34465ff6");
    }
}

