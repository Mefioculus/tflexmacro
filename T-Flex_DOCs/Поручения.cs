using System;
using System.Collections.Generic;
using System.Linq;
using TFlex.DOCs.Model.Macros;
using TFlex.DOCs.Model.Macros.ObjectModel;
using TFlex.DOCs.Model.References;
using TFlex.DOCs.Model.References.Assignments;
using TFlex.DOCs.Model.References.Links;
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

    private static List<int> Supervisors;
    private static List<int> Subordinates;

    //return ВыполнитьМакрос("24d95d52-b3c1-4ec2-98d0-15eec649d642", "ПолучитьРуководителей");
    public List<int> ПолучитьРуководителей()
    {
        if (Supervisors == null)
        {
            var пользователь = Context.Connection.ClientView.GetUser();
            var группы = GetSupervisorGroups(пользователь);
            var руководители = GetSupervisors(группы);

            Supervisors = руководители.Select(o => o.SystemFields.Id).ToList();
            if (Supervisors.Count == 0)
                Supervisors.Add(0);
        }

        return Supervisors;
    }

    //return ВыполнитьМакрос("24d95d52-b3c1-4ec2-98d0-15eec649d642", "ПолучитьПодчинённых");
    public List<int> ПолучитьПодчинённых()
    {
        if (Subordinates == null)
        {
            var пользователь = Context.Connection.ClientView.GetUser();
            var подчинённые_группы = GetSubordinateGroups(пользователь);
            var подчинённые_сотрудники = GetSubordinates(подчинённые_группы);

            Subordinates = подчинённые_сотрудники.Select(o => o.SystemFields.Id).ToList();
            if (Subordinates.Count == 0)
                Subordinates.Add(0);
        }

        return Subordinates;
    }

    //return ВыполнитьМакрос("24d95d52-b3c1-4ec2-98d0-15eec649d642", "СброситьКэш");
    public void СброситьКэш()
    {
        Supervisors = null;
        Subordinates = null;
    }

    //return ВыполнитьМакрос("24d95d52-b3c1-4ec2-98d0-15eec649d642", "ПослатьНаВнешнююПочту", ТекущийОбъект, false, false, false, false, false);
    public string[] ПослатьНаВнешнююПочту(Объект поручение, bool автор, bool исполнитель, bool руководитель, bool контролёр, bool рассылка)
    {
        var mailList = GetMailList((ReferenceObject)поручение, автор, исполнитель, руководитель, контролёр, рассылка);
        return mailList.Select(user => user.Email.ToString()).Where(email => !string.IsNullOrWhiteSpace(email)).Distinct().ToArray();
    }

    //return ВыполнитьМакрос("24d95d52-b3c1-4ec2-98d0-15eec649d642", "ПослатьНаВнутреннююПочту", ТекущийОбъект, false, false, false, false, false);
    public User[] ПослатьНаВнутреннююПочту(Объект поручение, bool автор, bool исполнитель, bool руководитель, bool контролёр, bool рассылка)
    {
        var mailList = GetMailList((ReferenceObject)поручение, автор, исполнитель, руководитель, контролёр, рассылка);
        return mailList.Distinct().ToArray();
    }

    //return ВыполнитьМакрос("24d95d52-b3c1-4ec2-98d0-15eec649d642", "НеобходимоОтправитьСообщениеИсполнителю", ТекущийОбъект);
    public bool НеобходимоОтправитьСообщениеИсполнителю(Объект поручение)
    {
        var assignment = (AssignmentReferenceObject)поручение;
        if (assignment.SystemFields.Author != assignment.Executor)
            return true;

        var eventHandlers = Context.Connection.ServerTaskManager.GetEventHandlers();
        // Поиск и проверка события "08. Изменение сроков поручения. Сообщение всем кроме исполнителя"
        var eventHandler = eventHandlers.FirstOrDefault(h => h.Guid == new Guid("1e4a4514-d057-42e3-9689-29ca2243dacf"));
        return (eventHandler == null || !eventHandler.Enabled);
    }

    //return ВыполнитьМакрос("24d95d52-b3c1-4ec2-98d0-15eec649d642", "ВернутьIdПорученийПоБпАннулированныхДляТекущегоИсполнителя");
    public ICollection<int> ВернутьIdПорученийПоБпАннулированныхДляТекущегоИсполнителя()
    {
        var assignmentReference = new AssignmentReference(Context.Connection);
        var parameterGroup = assignmentReference.ParameterGroup;
        using var filter = new Filter(parameterGroup);

        filter.Terms.AddTerm(parameterGroup[SystemParameterType.Class], ComparisonOperator.IsInheritFrom,
            assignmentReference.Classes.ProcessAssignment);

        filter.Terms.AddTerm(parameterGroup[AssignmentReferenceObject.FieldKeys.Status], ComparisonOperator.Equal,
            (int)AssignmentStatus.InProgress);

        var currentUser = Context.Connection.ClientView.GetUser();
        var assignmentExecutorsTable = parameterGroup.OneToManyTables.Find(AssignmentReferenceObject.RelationKeys.AssignmentExecutors);
        var executorLink = assignmentExecutorsTable.OneToOneLinks.Find(AssignmentExecutorReferenceObject.RelationKeys.Executor);
        if (executorLink != null)
        {
            var term = new ReferenceObjectTerm(filter.Terms);
            term.Path.AddGroup(assignmentExecutorsTable);
            term.Path.AddGroup(executorLink);
            term.Operator = ComparisonOperator.Equal;
            term.Value = currentUser;
        }

        var assignments = assignmentReference.Find(filter).OfType<ProcessAssignmentReferenceObject>().ToList();
        if (assignments.Count == 0)
            return Array.Empty<int>();

        var assignmentExecutorsRelation = assignmentReference.LoadSettings.AddRelation(AssignmentReferenceObject.RelationKeys.AssignmentExecutors);
        assignmentExecutorsRelation.AddRelation(AssignmentExecutorReferenceObject.RelationKeys.Executor);

        var relationTree = new RelationTree(assignments) { SetObjectsInLinks = true };
        relationTree.Fill(assignmentReference.LoadSettings);

        var cancelledForCurrentUserProcessAssignments = assignments.Where(a =>
        {
            var assignmentExecutors = a.Executors.AsList;
            return assignmentExecutors.FirstOrDefault(e => e.Executor == currentUser)?.StatusType == AssignmentStatus.Cancelled;
        });

        return cancelledForCurrentUserProcessAssignments.Select(a => a.Id).ToList();
    }

    //Получить связь "Руководитель" в справочнике "Группы и пользователи"
    private ParameterGroup GetManagerLink()
    {
        return Context.Connection.References.Users.ParameterGroup.OneToOneLinks.Find(new Guid("fdb41549-2adb-40b0-8bb5-d8be4a6a8cd2"));
    }

    //Получаем список групп и подразделений, в которые входит пользователь
    private List<UserReferenceObject> GetSupervisorGroups(UserReferenceObject userObject)
    {
        List<UserReferenceObject> groups = new List<UserReferenceObject>();

        //Добавляем группы пользователей, в который входит userObject
        groups.AddRange(userObject.Parents.OfType<UsersGroup>());

        //Добавляем производственные подразделения, в который входит userObject
        groups.AddRange(userObject.Parents.OfType<ProductionUnit>());

        return groups;
    }

    //Получаем список руководителей групп и подразделений
    private List<User> GetSupervisors(List<UserReferenceObject> groups)
    {
        List<User> chiefs = new List<User>();
        var link = Context.Connection.References.Users.ParameterGroup.OneToOneLinks.Find(new Guid("fdb41549-2adb-40b0-8bb5-d8be4a6a8cd2"));
        if (link == null)
            return chiefs;

        chiefs.AddRange(groups.Select(g => g.GetObject(link)).OfType<User>().Distinct());
        return chiefs;
    }

    //Получаем список групп, в которых userObject является руководителем
    private List<UserReferenceObject> GetSubordinateGroups(UserReferenceObject userObject)
    {
        UserReference userReference = Context.Connection.References.Users;

        //Фильтр [Руководитель].[ID] = Id userObject'а
        Filter filter = new Filter(userReference.ParameterGroup);
        ReferenceObjectTerm term = new ReferenceObjectTerm(filter.Terms);
        term.Path.AddGroup(GetManagerLink());
        term.Path.AddParameter(userReference.ParameterGroup[SystemParameterType.ObjectId]);
        term.Operator = ComparisonOperator.Equal;
        term.Value = userObject.SystemFields.Id;

        return userReference.Find(filter).OfType<UserReferenceObject>().ToList();
    }

    //Получаем список сотрудников групп
    private List<User> GetSubordinates(List<UserReferenceObject> groups)
    {
        List<User> subordinates = new List<User>();
        foreach (UserReferenceObject group in groups.Where(g => g.HasChildren))
            subordinates.AddRange(group.Children.OfType<User>());

        return subordinates.Distinct().ToList();
    }

    //Получает для поручения список пользователей, которым нужно отправить сообщение
    private List<User> GetMailList(ReferenceObject referenceObject, bool toAuthor, bool toExecutor, bool toChief, bool toController, bool toMailList)
    {
        List<User> result = new List<User>();
        if (!(referenceObject is AssignmentReferenceObject assignment))
            return result;

        if (toAuthor)
        {
            var author = GetAuthor(assignment);
            result.Add(author);
        }

        if (toExecutor || toChief)
            AddUsersToMailList(assignment, result, toExecutor, toChief);

        if (toController)
        {
            User controller = assignment.Controller;
            if (controller != null)
                result.Add(controller);
        }

        if (toMailList)
        {
            var mailList = assignment.MailingList;
            foreach (UserReferenceObject mailItem in mailList)
            {
                if (mailItem is User user)
                    result.Add(user);
                else if (mailItem.HasChildren)
                    result.AddRange(mailItem.Children.RecursiveLoad(false).OfType<User>());
            }
        }

        return result;
    }

    private void AddUsersToMailList(ReferenceObject assignment, List<User> result, bool toExecutor, bool toChief)
    {
        List<User> executors = new List<User>();

        if (toExecutor)
        {
            var assignmentExecutors = GetExecutors(assignment);
            if (assignmentExecutors.Count > 0)
                executors.AddRange(assignmentExecutors);
        }

        if (toChief && executors.Count > 0)
        {
            List<UserReferenceObject> allGroups = new List<UserReferenceObject>();
            foreach (var executor in executors)
            {
                var groups = GetSupervisorGroups(executor);
                if (groups.Count > 0)
                    allGroups.AddRange(groups);
            }

            if (allGroups.Count > 0)
            {
                var chiefs = GetSupervisors(allGroups);
                executors.AddRange(chiefs);
            }
        }

        result.AddRange(executors);
    }

    private User GetAuthor(ReferenceObject referenceObject)
    {
        var author = referenceObject.SystemFields.Author;
        var assignment = referenceObject as AssignmentReferenceObject;
        if (assignment == null)
            return author;

        if (assignment.Class.IsProcessAssignment)
        {
            var processAssignment = (ProcessAssignmentReferenceObject)assignment;
            var onBehalf = processAssignment.OnBehalf;
            if (onBehalf != null)
                author = onBehalf;
        }

        return author;
    }

    private ICollection<User> GetExecutors(ReferenceObject referenceObject)
    {
        if (!(referenceObject is AssignmentReferenceObject assignment))
            return Array.Empty<User>();

        if (assignment.Class.IsProcessAssignment)
        {
            var processAssignment = (ProcessAssignmentReferenceObject)assignment;
            var assignmentExecutors = processAssignment.Executors.AsList;
            if (assignmentExecutors.Count == 0)
                return Array.Empty<User>();

            return assignmentExecutors.Select(e => e.Executor).ToList();
        }

        var executor = assignment.Executor;
        // Если поручение является поручением со списком возможных исполнителей и исполнитель задан, то возможным исполнителям не посылаем сообщения.
        if (assignment is AssignmentWithExecutorsListObject assignmentWithEL && executor is null)
        {
            var possibleExecutors = assignmentWithEL.PossibleExecutors.OfType<User>().ToList();
            return possibleExecutors;
        }

        if (executor is null)
            return Array.Empty<User>();

        return new List<User> { executor };
    }

    //return ВыполнитьМакрос("24d95d52-b3c1-4ec2-98d0-15eec649d642", "ВернутьВозможныхИсполнителей", false);
    public List<ReferenceObject> ВернутьВозможныхИсполнителей(bool учитыватьПодчиненных)
    {
        var currentUser = Context.Connection.ClientView.GetUser();
        List<ReferenceObject> possibleExecutors = new List<ReferenceObject>();
        possibleExecutors.Add(currentUser);
        if (учитыватьПодчиненных)
        {
            var subordinateGroups = GetSubordinateGroups(currentUser);
            var subordinates = GetSubordinates(subordinateGroups);
            possibleExecutors.AddRange(subordinates);
        }

        return possibleExecutors;
    }

    // return ВыполнитьМакрос("24d95d52-b3c1-4ec2-98d0-15eec649d642", "ВернутьIdЕслиВышестоящееГрупповое");
    public int ВернутьIdЕслиВышестоящееГрупповое()
    	=> ВернутьВышестоящееГрупповое() is null ? -1 : Context.ReferenceObject.Id;
    
    // return ВыполнитьМакрос("24d95d52-b3c1-4ec2-98d0-15eec649d642", "ВернутьВышестоящееГрупповое");
    public GroupAssignmentReferenceObject ВернутьВышестоящееГрупповое()
        => (Context.ReferenceObject as AssignmentReferenceObject)?.GetGroupAssignment();

    // return ВыполнитьМакрос("24d95d52-b3c1-4ec2-98d0-15eec649d642", "ВернутьIdСоисполнителей");
    public List<int> ВернутьIdСоисполнителей()
    {
        List<int> usersId = new List<int>() { -1 };
        var assignment = Context.ReferenceObject as AssignmentReferenceObject;
        if (assignment != null)
        {
            var parentAssignment = assignment.GetParentAssignment() as GroupAssignmentReferenceObject;
            if (parentAssignment != null)
            {
                var dependentAssignments = parentAssignment.DependentAssignments.AsList;
                if (dependentAssignments.Count > 0)
                    usersId.AddRange(dependentAssignments.Where(da => da.Executor != null).Select(da => da.Executor.SystemFields.Id).Distinct());
            }
        }

        return usersId;
    }

    // Бизнес-процессы

    // return ВыполнитьМакрос("24d95d52-b3c1-4ec2-98d0-15eec649d642", "НеСодержитНормативныеДокументы");
    public bool НеСодержитНормативныеДокументы()
    {
        var assignment = Context.ReferenceObject as ProcessAssignmentReferenceObject;
        return assignment is null || assignment.Data.Documents.Length == 0;
    }
}
