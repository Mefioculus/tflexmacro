/*
TFlex.DOCs.UI.Objects.dll
TFlex.DOCs.UI.Common.dll
TFlex.DOCs.UI.Types.dll
TFlex.DOCs.Common.dll
*/

using System;
using System.Collections.Generic;
using System.Linq;
using TFlex.DOCs.Model.Macros;
using TFlex.DOCs.Model.Access;
using TFlex.DOCs.Model.Classes;
using TFlex.DOCs.Model.Macros.ObjectModel;
using TFlex.DOCs.Model.References;
using TFlex.DOCs.Model.References.Users;
using TFlex.DOCs.Model.Desktop;
using TFlex.DOCs.Model.References.Files;


public class MacroCreateFolder : MacroProvider
{
    public MacroCreateFolder(MacroContext context)
        : base(context)
    {
        //System.Diagnostics.Debugger.Launch();
        //System.Diagnostics.Debugger.Break();
    }
    private string _редакторский = "Редакторский"; //Назначаемый доступ на папку
    private string _просмотрИсоздание = "Просмотр и создание"; //Назначаемый доступ на папку
    private string _наименованиеПапкиДляХраненияЛичныхПапок = "Личные папки";
    //private string _наименованиеПапкиДляСогласования = "Согласования";

    public override void Run()
    {
        //User user = Context.ReferenceObject as User;
        var user = ТекущийОбъект;
        if (user == null)
            return;
        //Назначение наименования личной папки пользователя
        string наименованиеЛичнойПапки = ТекущийОбъект.Параметр["Логин"];
        Объект папка = НайтиОбъект("Файлы", "Наименование", наименованиеЛичнойПапки);
        //Проверка на то, что пользователь запустивший макрос, Администратор
        if ((ТекущийПользователь.Параметр["Тип"] == "Администратор") && (папка == null))
        {
            //Проверка и создание папки "Личные папки" в корне справочника
            Объект личныеПапки = НайтиОбъект("Файлы", "Наименование", _наименованиеПапкиДляХраненияЛичныхПапок);
            if (личныеПапки == null)
            {
                личныеПапки = СоздатьОбъект("Файлы", "Папка");
                личныеПапки["Наименование"] = _наименованиеПапкиДляХраненияЛичныхПапок;
                личныеПапки.Сохранить();
            }

            папка = СоздатьОбъект("Файлы", "Папка", личныеПапки);//Создаем личную папку пользователя
            папка.Параметр["Наименование"] = наименованиеЛичнойПапки;
            папка.Сохранить();
            НазначитьДоступНаОбъект(папка, user, _редакторский);
            //Desktop.CheckIn((ReferenceObject)папка, "Создание", false);//Применение изменений
        }
    }

    public void ОпубликоватьФайл()
    {
        //ID папки согласования
        int folderID = 396;
        //Получаем справочник файлов
        var fileReference = new FileReference(Context.Connection);
        //Находим папку
        var folderObj = fileReference.Find(folderID);
        if(folderObj == null)
            Error("Не найдена папка с ID = " + folderID.ToString());

        if (!Вопрос("Переместить файлы в папку \"Согласования\"?"))
            return;

        WaitingDialog.Show("Подождите, идет обработка", true);

        var currentObjs = Context.GetSelectedObjects();
        //Коллекции для сохранеия/изменения объектов
        var saveObjs = new List<ReferenceObject>();
        var deleteObjs = new List<ReferenceObject>();
        WaitingDialog.NextStep("Подождите идет копирование объектов");
        foreach (var currentObj in currentObjs)
        {
            ClassObject classObj = currentObj.Class;
            //Создает копию объекта
            ReferenceObject createObj = currentObj.CreateCopy(classObj, folderObj);
            saveObjs.Add(createObj);
            //Отменяет изменения объекта, добавляет в список для удаления
            if (currentObj.IsCheckedOut)
                createObj.UndoCheckOut();
            deleteObjs.Add(currentObj);

        }
        WaitingDialog.NextStep("Подождите идет сохранение новых объектов на сервер");
        ReferenceObjectSaveSet.RunWithSaveSet(set =>
        {
            //Пакетное сохранение объектов на сервер
            foreach (ReferenceObject saveObj in saveObjs)
                set.Add(saveObj);
            set.EndChanges();
        });
        WaitingDialog.NextStep("Подождите идет применение изменений новых объектов на сервер");
        Desktop.CheckIn(saveObjs, "Копирование объектов", false);

        WaitingDialog.NextStep("Подождите идет удаление исходных объектов с сервера");
        //Удаление старых файлов
        Desktop.CheckOut(deleteObjs, true);
        Desktop.CheckIn(deleteObjs, "Удаление объектов после копирования", false);
        WaitingDialog.Hide();
    }

    /// <summary>
    /// Назначает на объект выбранный доступ, есть возможность очистить все доступы
    /// </summary>
    /// <param name="объектНазначенияДоступа">Объект для изменения доступа</param>
    /// <param name="пользователь">Пользователь на которого будет назначатся доступ</param>
    /// <param name="наименованиеДоступа">Наименование доступа, которое необходимо установить на объект</param>
    /// <param name="очищатьСтарыеДоступы">Значение позволяющее удалить все доступы(для исспользования личных папок)</param>
    private void НазначитьДоступНаОбъект(Объект объектНазначенияДоступа, Объект пользователь, string наименованиеДоступа, bool очищатьСтарыеДоступы = false)
    {
        User user = (ReferenceObject)пользователь as User;
        if (user == null)
            throw new Exception(string.Format("Объект: {0} не является пользователем", пользователь));
        //Получем группу доступов
        AccessGroup accessEdit = AccessGroup.GetGroups(Context.Connection).FirstOrDefault(ag => ag.Type.IsObject && ag.Name == наименованиеДоступа);
        //Получаем менеджер доступа на объект
        AccessManager accessManager = AccessManager.GetReferenceObjectAccess((ReferenceObject)объектНазначенияДоступа);
        //Убираем наследование доступа от родителя
        accessManager.SetInherit(false, false);
        //Очищаем все доступы
        if (очищатьСтарыеДоступы)
            accessManager.ToList().Clear();
        //Устанавливаем доступ пользователю
        accessManager.SetAccess(user, accessEdit);
        //Сохраняем изменения
        accessManager.Save();
    }
}

