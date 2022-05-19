using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using TFlex.DOCs.Model.Macros;
using TFlex.DOCs.Model.Macros.ObjectModel;
using System.Text;

//-----------

using TFlex.DOCs.Model.Access;
using TFlex.DOCs.Model.Classes;

using TFlex.DOCs.Model.References;
using TFlex.DOCs.Model.References.Users;
using TFlex.DOCs.Model.Desktop;
using TFlex.DOCs.Model.References.Files;
using TFlex.DOCs.Model.References.Nomenclature;
using TFlex.DOCs.Model.References.Documents;

public class Macro : MacroProvider
{
    public Macro(MacroContext context)
        : base(context)
    {
    	//System.Diagnostics.Debugger.Launch();
       // System.Diagnostics.Debugger.Break();
    	
    }

    public override void Run()
    {
    }
    
    public void AccessCZ (Объект текущий)
    {
        StringBuilder strBuilder = new StringBuilder();
        //  StringBuilder strBuilder2 = new StringBuilder();
        Объект пользовтель = ТекущийПользователь;
        Объекты группы = пользовтель.РодительскиеОбъекты;
        
        try
        {
            var суперпользователи = НайтиОбъект("Группы и пользователи", "[Наименование] = 'Суперпользователи'");
            НазначитьДоступНаОбъект(текущий, суперпользователи);
        }
        catch
        {}
        
        //Объект текущий = ТекущийОбъект;
                       
        //  if (текущий["Код подразделения"].Equals(""))
        //    Message("Введите код подразделения","Введите код подразделения");
        if (!текущий["Код подразделения"].Equals(""))
        {
        	Объект user = НайтиОбъект("Группы и пользователи","Номер",текущий["Код подразделения"].ToString());
        	if (user == null)
        		return;
        	
            UserReferenceObject user2=(ReferenceObject)user as UserReferenceObject;
            try
            {
            	НазначитьДоступНаОбъект(текущий, user);
            }
            catch
            {}
        }
    }
    
    /// <summary>
    /// Назначает на объект выбранный доступ, есть возможность очистить все доступы
    /// </summary>
    /// <param name="объектНазначенияДоступа">Объект для изменения доступа</param>
    /// <param name="пользователь">Пользователь на которого будет назначатся доступ</param>
    /// <param name="наименованиеДоступа">Наименование доступа, которое необходимо установить на объект</param>
    /// <param name="очищатьСтарыеДоступы">Значение позволяющее удалить все доступы(для исспользования личных папок)</param>
    private void НазначитьДоступНаОбъект(Объект объектНазначенияДоступа, Объект пользователь, bool очищатьСтарыеДоступы = false)
    {	 
      	UserReferenceObject user=(ReferenceObject)пользователь as UserReferenceObject;
        //User user = (ReferenceObject)пользователь as User;
        if (user == null)
            throw new Exception(string.Format("Объект: {0} не является пользователем", пользователь));
         AccessManager accessManager = AccessManager.GetReferenceObjectAccess((ReferenceObject)объектНазначенияДоступа);
         if (accessManager.IsInherit)
       	{
        //Получем группу доступов
        AccessGroup accessEdit = AccessGroup.GetGroups(Context.Connection).FirstOrDefault(ag => ag.Type.IsObject && ag.Name == "Редакторский");
        AccessGroup accessEdit2 = AccessGroup.GetGroups(Context.Connection).FirstOrDefault(ag => ag.Type.IsObject && ag.Name == "Просмотр");
        //Получаем менеджер доступа на объект
       
        //Убираем наследование доступа от родителя
       //--------
      
            accessManager.SetInherit(true, true);
            accessManager.Save();
            //-------- 
            accessManager.SetInherit(false, false);
            //Очищаем все доступы
            if (очищатьСтарыеДоступы)
            accessManager.ToList().Clear();
            //Устанавливаем доступ пользователю
            accessManager.SetAccess(user, accessEdit);
            accessManager.SetAccess(null, accessEdit2);
            //Сохраняем изменения
            accessManager.Save();
        }
    }
      
    public void ActiveCz ()
	{
    	Объект текущий = ТекущийОбъект;
    	AccessCZ(текущий);
    }
      
    public void ActiveTP ()
    {
       	 /*StringBuilder strBuilder = new StringBuilder();
       	Объект текущий = ТекущийОбъект;
       	 strBuilder.AppendLine("1 "+текущий.ToString());
       	 strBuilder.AppendLine("2 "+текущий.Тип.ToString());
       	Объекты цехозаходы = текущий.ДочерниеОбъекты;
       foreach (var c in цехозаходы)
       	{
          //  strBuilder.AppendLine("Цехозаход "+c.ToString());
             AccessCZ (c);
       }*/
       
          
                  //     Сообщение("Цехозаход", strBuilder.ToString());
    }
}
