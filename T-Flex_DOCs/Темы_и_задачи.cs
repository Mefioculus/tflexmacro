using System;
using System.Collections.Generic;
using System.Linq;
using TFlex.DOCs.Model.Macros;
using TFlex.DOCs.Model.References;
using TFlex.DOCs.Model.References.Assignments;
using TFlex.DOCs.Model.References.Tasks;
using TFlex.DOCs.Model.Search;
using TFlex.DOCs.Model.Structure;

public class CheckingTaskStates : MacroProvider
{
    private static readonly Dictionary<TaskState, string> _textStateMapping = new Dictionary<TaskState, string>()
    {
        { TaskState.Expired, "Просрочено" },
        { TaskState.AssignmentsOverdue, "Просрочены поручения" },
        { TaskState.DeviationFromHigherTask, "Не соответствует вышестоящей" },
        { TaskState.NotInWork, "Не выполняется" },
        { TaskState.Default, "По умолчанию" }
    };

    private static readonly TimeSpan _executeInterval = TimeSpan.FromSeconds(3);
    private static Dictionary<int, TaskState> _statusCache;
    private static DateTime _lastFillingTime = DateTime.MinValue;

    public CheckingTaskStates(MacroContext context)
        : base(context)
    {
    }

    public override void Run()
    {
    }

    // ВыполнитьМакрос("cf668b9b-ee86-4a28-a394-9326e06ecacb", "ВернутьIdВходящихЗадач");
    public IEnumerable<int> ВернутьIdВходящихЗадач()
    {
        var currentTask = Context.ReferenceObject;
        if (currentTask is null)
            return Array.Empty<int>();

        var currentTaskWithSubtasks = currentTask.Children.RecursiveLoad(true);
        currentTaskWithSubtasks.Add(currentTask);
        
        var tasksIds = currentTaskWithSubtasks.Select(task => task.SystemFields.Id).ToList();
        if (tasksIds.Count == 0)
            tasksIds.Add(0);

        return tasksIds;
    }

    // ВыполнитьМакрос("cf668b9b-ee86-4a28-a394-9326e06ecacb", "ВернутьАктуальность");
    public string ВернутьАктуальность()
    {
        var task = Context.ReferenceObject;
        if (task is null)
            return String.Empty;

        if (_statusCache is null || DateTime.Now - _lastFillingTime > _executeInterval)
            FillStatusCache();

        _lastFillingTime = DateTime.Now;
        string result = String.Empty;
        if (_statusCache.TryGetValue(task.SystemFields.Id, out TaskState val))
            _textStateMapping.TryGetValue(val, out result);

        return result ?? String.Empty;
    }

    #region Для выделения цветом в новом клиенте
    // ВыполнитьМакрос("cf668b9b-ee86-4a28-a394-9326e06ecacb", "ВернутьIdЕслиПросрочено");
    public int ВернутьIdЕслиПросрочено()
    {
        var task = Context.ReferenceObject;
        var taskState = GetTaskState();

        return taskState == TaskState.Expired || taskState == TaskState.AssignmentsOverdue ? task.SystemFields.Id : -1;
    }

    // ВыполнитьМакрос("cf668b9b-ee86-4a28-a394-9326e06ecacb", "ВернутьIdПриОтклоненииОтВышестоящей");
    public int ВернутьIdПриОтклоненииОтВышестоящей()
    {
        var task = Context.ReferenceObject;
        var taskState = GetTaskState();

        return taskState == TaskState.DeviationFromHigherTask ? task.SystemFields.Id : -1;
    }

    private TaskState GetTaskState()
    {
        var task = Context.ReferenceObject;
        if (task is null)
            return TaskState.Default;

        if (_statusCache is null || DateTime.Now - _lastFillingTime > _executeInterval)
            FillStatusCache();

        _lastFillingTime = DateTime.Now;

        return _statusCache.TryGetValue(task.SystemFields.Id, out TaskState val) ? val : TaskState.Default;
    }
    #endregion

    private void FillStatusCache()
    {
        if (_statusCache == null)
            _statusCache = new Dictionary<int, TaskState>();
        else
            _statusCache.Clear();

        TasksReference taskReference = Context.ReferenceObject.Reference as TasksReference;
        taskReference.LoadSettings.AddRelation(TasksReferenceObject.RelationKeys.Assignments);
        taskReference.LoadSettings.AddParameters(new Guid[] { TasksReferenceObject.FieldKeys.TargetDate, TasksReferenceObject.FieldKeys.Status });

        List<TasksReferenceObject> tasks;
        using (Filter filter = new Filter(taskReference.ParameterGroup))
        {
            filter.Terms.AddTerm(taskReference.ParameterGroup[SystemParameterType.ObjectId], ComparisonOperator.GreaterThan, 0);
            tasks = taskReference.Find(filter).OfType<TasksReferenceObject>().ToList();
        }

        foreach (var task in tasks)
        {
            if (_statusCache.ContainsKey(task.SystemFields.Id))
                continue;

            var state = TaskState.Default;
            if (task.StatusType != TaskStatus.InProgress)
                state = TaskState.NotInWork;
            else if (CheckOverdue(task))
                state = TaskState.Expired;
            else if (CheckDeviationFromHigherTask(task))
                state = TaskState.DeviationFromHigherTask;
            else if (IsContainsIncorrectAssignments(task))
                state = TaskState.AssignmentsOverdue;

            if (state != TaskState.Default)
                _statusCache.Add(task.SystemFields.Id, state);
        }
    }

    private bool CheckOverdue(TasksReferenceObject task)
    {
        var targetDateParameter = task[TasksReferenceObject.FieldKeys.TargetDate];
        if (targetDateParameter.IsEmpty)
            return false;

        DateTime targetDate = targetDateParameter.GetDateTime().Date;
        return targetDate < DateTime.Now.Date;
    }

    private bool CheckDeviationFromHigherTask(TasksReferenceObject task)
    {
        var targetDateParameter = task[TasksReferenceObject.FieldKeys.TargetDate];
        if (targetDateParameter.IsEmpty)
            return false;

        DateTime targetDate = targetDateParameter.GetDateTime().Date;

        var higherLvlTask = task;
        while (higherLvlTask.Parent != null)
        {
            higherLvlTask = higherLvlTask.Parent as TasksReferenceObject;
            if (_statusCache.ContainsKey(higherLvlTask.SystemFields.Id) && _statusCache[higherLvlTask.SystemFields.Id] == TaskState.Expired)
                return true;

            var parentTargetDateParameter = higherLvlTask[TasksReferenceObject.FieldKeys.TargetDate];
            if (parentTargetDateParameter.IsEmpty)
                continue;

            DateTime parentTargetDate = parentTargetDateParameter.GetDateTime().Date;
            if (parentTargetDate < targetDate)
                return true;
        }

        return false;
    }

    private bool IsContainsIncorrectAssignments(TasksReferenceObject task)
    {
        var targetDateParameter = task[TasksReferenceObject.FieldKeys.TargetDate];
        if (targetDateParameter.IsEmpty)
            return false;

        var assignments = task.Assignments;
        return assignments.AsList.Any(assignment =>
        {
            DateTime startDate = assignment[AssignmentReferenceObject.FieldKeys.StartDate].GetDateTime();
            DateTime endDate = assignment[AssignmentReferenceObject.FieldKeys.EndDate].GetDateTime();
            DateTime targetDate = targetDateParameter.GetDateTime().Date;

            return startDate.Date > targetDate || endDate.Date > targetDate;
        });
    }

    private enum TaskState
    {
        Default,
        Expired,
        DeviationFromHigherTask,
        AssignmentsOverdue,
        NotInWork
    }
}
