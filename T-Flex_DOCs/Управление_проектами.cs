/*Ссылки
TFlex.DOCs.ProjectManagement.dll*/

using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;

using TFlex.DOCs.Model.Macros;
using TFlex.DOCs.Model.Macros.ObjectModel;
using TFlex.DOCs.Model.References;
using TFlex.DOCs.Model.References.Calendar;

using TFlex.DOCs.References.ProjectManagement;
using TFlex.DOCs.References.Labels;
using TFlex.DOCs.References.ResourcesUsed;
using TFlex.DOCs.References.ProjectResources;
using TFlex.DOCs.Resources.ProjectManagement;

public class Macro : MacroProvider
{
    public Macro(MacroContext context)
        : base(context)
    {
    }

    public override void Run()
    {
    }
    
    //Меняет местами мастера и зависимого элемента
    public void ПоменятьМестами()
    {
    	if (ТекущийОбъект["ID"] != 0) //Менять мастера и зависимого элементов можно только для вновь создаваемой зависимости
    		return;
    	
    	var работа1 = Параметр["Работа 1"];
        var работа2 = Параметр["Работа 2"];
        var проект1 = Параметр["Проект 1"];
        var проект2 = Параметр["Проект 2"];
        
        var reference = new ProjectManagementReference(TFlex.DOCs.Model.ServerGateway.Connection);
        var task1 = reference.Find(new Guid(работа1.ToString())) as ProjectElement;
        var task2 = reference.Find(new Guid(работа2.ToString())) as ProjectElement;
        
        if (task1 == null || task2 == null)
        	return;
        
        if (task2.IsDependedFrom(task1))
        {
        	Сообщение("Исключение", "При создании зависимости от '" + task2.ToString() + "' к '" + task1.ToString() + "' появится цикл.");
        	return;
        }
        
        Параметр["Работа 1"] = работа2;
        Параметр["Работа 2"] = работа1;
        if (проект1 != проект2)
        {
            Параметр["Проект 1"] = проект2;
            Параметр["Проект 2"] = проект1;
        }
        
    }
    
    public List<int> ПолучитьСписокIDПредустанновленныхРесурсов(Объект текущий)
    {
    	var resource = (ResourcesUsedReferenceObject)текущий;
    	return resource.GetPredefinedResourcesIDs();
    }
    
    public List<int> ПолучитьСписокIDПроектовРесурса(Объект текущий)
    {
    	List<int> списокIDПроектов = new List<int>();
    	
    	ProjectResourcesReferenceObject ресурс = (ProjectResourcesReferenceObject)текущий;
    	if (ресурс != null)
    		списокIDПроектов = ресурс.GetProjectsIDs();
    	
    	return списокIDПроектов;
    }
    
    //Получить дату, на которой будет отображаться метка
    public string ДатаМетки()
    {
    	LabelObject labelObject = (LabelObject)ТекущийОбъект;
    	if (labelObject == null)
    		return "Дата не определена";
    	DateTime date = labelObject.ActualDate;
    	if (labelObject.ProjectElement.Project.PlanInDays)
    		return date.ToString("dd.MM.yyyy");
    	return date.ToString("dd.MM.yyyy HH:mm:ss");
    	//return ВыполнитьМакрос("Управление проектами", "ДатаМетки");
    }
    
    //Получает день недели, на который вносятся изменения в календаре
    public string ДеньНеделиИзменения()
    {
    	if (ТекущийОбъект == null)
    		return string.Empty;
    	
    	DateTime date = Параметр["e3af0c61-e6a8-4b53-ae52-385d5afed046"]; //Дата
    	return ДеньНедели(date);
    }
    
    private string ДеньНедели(DateTime date)
    {
        var formatInfo =  System.Globalization.CultureInfo.CurrentCulture.DateTimeFormat;
        string day = date.ToString("dddd", formatInfo);
        char[] arr = day.ToCharArray();
        arr[0] = Char.ToUpper(arr[0]);
        return new string(arr);
    }
    
    //Блокирует процент выполнения, если значение параметра "Определение процента" в проекте не Вручную 
    public bool Блокировка()
    {
    	ProjectElement projectElement = (ProjectElement)ТекущийОбъект;
    	return projectElement.Progress.IsReadOnly;
    	//return ВыполнитьМакрос("Управление проектами", "Блокировка");
    }
    
    public bool СкрытьДоступы()
    {
    	ProjectElement projectElement = (ProjectElement)ТекущийОбъект;
    	Project project = projectElement.Project;
    	return project == null || !project.GetObjectValue("106bf511-dd62-476c-9ac5-22fd1387d78a").ToBoolean(); //Использовать доступ
    	//return ВыполнитьМакрос("Управление проектами", "СкрытьДоступы");
    }
    
    public bool СкрытьЗадержку()
    {
    	//Найти в справочнике "Управление проектами" объект с уникальным идентификатором как в параметре "Проект 1"
    	Объект проект = НайтиОбъект("86ef25d4-f431-4607-be15-f52b551022ff", "[Guid]", Параметр["Проект 1"]);
    	if (проект != null)
    		return проект["Планировать в днях"];
    	return false;
    	//return ВыполнитьМакрос("Управление проектами", "СкрытьЗадержку");
    }
    
    public bool КалендарьПуст()
    {
    	Project project = (Project)ТекущийОбъект;
    	var calendar = project.Calendar as TFlex.DOCs.Model.References.Calendar.CalendarReferenceObject;
    	if (calendar != null)
    	   return !calendar.HasAnyWorkTimeInterval();
    	return true;
    	//return ВыполнитьМакрос("Управление проектами", "КалендарьПуст");
    }
    
    // Когда у используемого ресурса включен флаг "Фактическое значение", задаёт его начало и окончание
    // из фактических сроков выполнения работы
    public void УстановитьИнтервалИпользованияРесурса()
	{
	    Guid isActualResourceUsedParam = new Guid("4227e515-66a4-43ae-9418-346854748986");	    
        Guid startDateResourceUsedParam = new Guid("57695721-084c-48bf-8c39-667d27ee1aaf");
        Guid endDateResourceUsedParam = new Guid("1680be8c-8527-4243-85ab-b3ae6dc38140");
        Guid projectResourceUsedParam = new Guid("1a22ee46-5438-4caa-8b75-8a8a37b74b7e");
        Guid resourceResourceUsedParam = new Guid("7f882c52-bfae-4a93-a7a7-9f548215898f");
        
        Guid actualStartDateProjectParam = new Guid("2f457df1-246d-4d23-a332-53f387940ba9");
        Guid actualEndDateProjectParam = new Guid("a93f9644-3e24-4a59-9eb9-d200c83044ee");
    	
    	var args = Context.ModelChangedArgs as ObjectParameterChangedEventArgs;
        if (args.Parameter.ParameterInfo.Guid != isActualResourceUsedParam)
            return;

    	var referenceObject = Context.ReferenceObject;
    	var isActual = (bool) referenceObject.ParameterValues[isActualResourceUsedParam].Value;
    	if (!isActual)
    		return;
   	
    	var project = referenceObject.GetObject(projectResourceUsedParam);
    	var resource = referenceObject.GetObject(resourceResourceUsedParam) as ProjectResourcesReferenceObject;

    	if (project == null || resource == null)
    		return;

    	var projectActualStartDate = project.ParameterValues[actualStartDateProjectParam];
        var projectActualEndDate = project.ParameterValues[actualEndDateProjectParam];
    	var question = String.Format(Texts.UpdateUsedResourceIntervalQuestion,
    	                             projectActualStartDate.ParameterInfo.Name,
    	                             projectActualEndDate.ParameterInfo.Name);
    	
    	if (!Вопрос(question))
    		return;
    	
    	if (projectActualStartDate.IsEmpty)
    		return;
    	
        referenceObject.ParameterValues[startDateResourceUsedParam].Value = projectActualStartDate.Value;
        referenceObject.ParameterValues[endDateResourceUsedParam].Value = projectActualEndDate.IsEmpty ?
                resource.WorkTimeManager.GetWorkDayEnd((DateTime)projectActualStartDate.Value) :
                projectActualEndDate.Value;
	}
    
    public void СоздатьКалендарьНаОсновеСуществующего()
    {
        var project = Context.ReferenceObject as Project;
        if (project is null)
            return;

        var state = project.State;
        if (state == ProjectElementState.Aborted || state == ProjectElementState.Completed)
            Error(String.Format(Messages.CannotChangeCalendarCauseOfProjectState, ProjectElementStateExt.GetString(state)));

        string calendarName = String.Format(Texts.ProjectCalendar, project.Name.Value);
        var sourceCalendar = project.Calendar ?? project.Reference.Connection.References.GlobalParameters.GlobalCalendar;
        string calendarPrototypeText = "Прототип календаря", nameText = "Наименование";

        var dialog = CreateInputDialog("Создать календарь проекта");
        dialog.AddSelectFromReference(calendarPrototypeText, "Календари", null, sourceCalendar);
        dialog.AddString(nameText, calendarName);
        if (!dialog.Show())
            return;

        sourceCalendar = (CalendarReferenceObject)dialog[calendarPrototypeText];
        calendarName = dialog[nameText];
        if (sourceCalendar is null)
            return;

        WaitingDialog.Show(Texts.NewCalendar, false);
        WaitingDialog.NextStep("Обновление календаря у проекта");
        var calendarCopy = sourceCalendar.CreateCopy() as CalendarReferenceObject;
        calendarCopy.Name.Value = calendarName;
        calendarCopy.EndChanges();
        project.Calendar = calendarCopy;
        project.Reload();
        WaitingDialog.Hide();
    }
}
