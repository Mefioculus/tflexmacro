using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using TFlex.DOCs.Model.Macros;
using TFlex.DOCs.Model.Macros.ObjectModel;
using TFlex.DOCs.Model.References.Files;

public class Macro : MacroProvider
{
    public Macro(MacroContext context)
        : base(context)
    {
    }

    public override void Run()
    {
    }
    public string GetNameFolder()
    { 
        string abstPathName = @"Личные папки\" + ТекущийПользователь["Логин"];
        FileReference fileRef = new FileReference(Context.Connection);
        var folderObj = fileRef.FindByPath(abstPathName);
        if (folderObj == null)
        {
            //Сообщение("Предупреждение",
            //    string.Format("Личная папка {0} не найдена \n\rФайл будет записан в папку 'Файлы документов'", abstPathName));
            return @"Файлы документов";
        }
        return abstPathName;
    }
}
