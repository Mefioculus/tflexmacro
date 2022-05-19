using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using TFlex.DOCs.Model;
using TFlex.DOCs.Model.Macros;
using TFlex.DOCs.Model.Macros.ObjectModel;
using TFlex.DOCs.Model.Stages;
using TFlex.DOCs.Model.References;

public class Macro : MacroProvider
{
    public Macro(MacroContext context)
        : base(context)
    {
    }

    public override void Run()
    {
        // Команда "Выдать наряд".
        Объект Obj = ТекущийОбъект;

        if (Obj.Параметр["Стадия"] != "Разработка")
        {
            MessageBox.Show("Для выдачи наряда его стадия должна быть \"Разработка\".\r\nПроцесс прерван.", "Предупреждение",
                MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        Stage stage = Stage.Find("Выдано");
        ReferenceObject RObj = (ReferenceObject)Obj;
        List<ReferenceObject> RList = new List<ReferenceObject>();
        RList.Add(RObj);
        Stage.ChangeObjectsStage(stage, RList);

        Obj.Параметр["Время выдачи"] = DateTime.Now;
        Obj.СвязанныйОбъект["Выдавший"] = ТекущийПользователь;
        RObj.ApplyChanges();

        ReferenceInfo referenceInfo =
            ReferenceCatalog.FindReference(new Guid("f83f5f88-a2ec-471a-8221-7f16d85b151d")); // Справочник "План ТОиР".
        if (referenceInfo != null)
        {
	        Reference reference = referenceInfo.CreateReference();
	        if (reference != null)
	        {
	            Объект maintObject = Obj.СвязанныйОбъект["3511727b-55a3-47ef-be3e-84d6701a34df"]; // Обслуживания и работы.
	            if (maintObject != null)
	            {
	                if (maintObject.Параметр["Стадия"] == "Проект")
	                {
	                    DialogResult result =
	                        MessageBox.Show("Изменить стадию связанного обслуживания " + maintObject.ToString() +
	                                        " на \"Выполнение\"?", "Вопрос", MessageBoxButtons.YesNo);
	                    if (result == System.Windows.Forms.DialogResult.Yes)
	                    {
	                        ReferenceObject RObj1 = (ReferenceObject)maintObject;
	                        Stage stage1 = Stage.Find("Выполнение");
	                        List<ReferenceObject> RList1 = new List<ReferenceObject>();
	                        RList1.Add(RObj1);
	                        Stage.ChangeObjectsStage(stage1, RList1);
	
	                        RObj1.BeginChanges();
	                        RObj1[new Guid("{5cf6e014-d505-44a5-bc8b-9ba59a726590}")].Value = DateTime.Now; // Дата начала.
	                        RObj1.EndChanges();
	                    }
	                }
	            }
	        }
	    }
    }

    public void Close()
    {
        // Команда "Закрыть наряд".
        Объект Obj = ТекущийОбъект;

        if (Obj.Параметр["Стадия"] != "Выдано")
        {
            MessageBox.Show("Для закрытия наряда его стадия должна быть \"Выдано\".\r\nПроцесс прерван.", "Предупреждение", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        Stage stage = Stage.Find("Закрыто");
        ReferenceObject RObj = (ReferenceObject)Obj;
        List<ReferenceObject> RList = new List<ReferenceObject>();
        RList.Add(RObj);
        Stage.ChangeObjectsStage(stage, RList);

        Obj.Параметр["88a3ab52-4b31-4aa4-b79d-cd5c04bdd719"] = DateTime.Now; // Время закрытия
        Obj.СвязанныйОбъект["Закрывший"] = ТекущийПользователь;
        RObj.ApplyChanges();
    }

    public void Cancel()
    {
        // Команда "Аннулировать наряд".
        Объект Obj = ТекущийОбъект;

        if (Obj.Параметр["Стадия"] != "Выдано")
        {
            MessageBox.Show("Для аннулирования наряда его стадия должна быть \"Выдано\".\r\nПроцесс прерван.", "Предупреждение", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        Stage stage = Stage.Find("Аннулировано");
        ReferenceObject RObj = (ReferenceObject)Obj;
        List<ReferenceObject> RList = new List<ReferenceObject>();
        RList.Add(RObj);
        Stage.ChangeObjectsStage(stage, RList);

        RObj.ApplyChanges();
    }
    
    public int GetPlanID()
    {
		if(TFlex.DOCs.Model.ReferenceCatalog.FindReference("План ТОиР") == null)
		{
		    return ТекущийОбъект["ID"];
		}
		else
		{
			return -1;
		}
    }

    public int GetFailuresID()
    {
		if(TFlex.DOCs.Model.ReferenceCatalog.FindReference("Дефекты и отказы") == null)
		{
		    return ТекущийОбъект["ID"];
		}
		else
		{
			return -1;
		}
    }

    public int GetParametersID()
    {
		if(TFlex.DOCs.Model.ReferenceCatalog.FindReference("Параметры состояния объектов ТОиР") == null)
		{
		    return ТекущийОбъект["ID"];
		}
		else
		{
			return -1;
		}
    }
}
