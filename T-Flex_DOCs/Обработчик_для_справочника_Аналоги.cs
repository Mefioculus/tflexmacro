using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using TFlex.DOCs.Model.Macros;
using TFlex.DOCs.Model.Macros.ObjectModel;

public class Macro : MacroProvider
{
    public Macro(MacroContext context)
        : base(context)
    {
    }

    public override void Run()
    {
        var объекты =ВыбранныеОбъекты;
                foreach (var объект in объекты)
                            CreateLogin(объект)	;
                	//Сообщение("",объект.ToString());
    	
    }
    
    public void CreateLogin(Объект документ)
    	{
    	
    	 // Объект документ = ТекущийОбъект;
          string row=документ["Сотрудник"].ToString();
          string[] rus = { "а", "б", "в", "г", "д", "е", "ё", "ж", "з", "и", "й", "к", "л", "м", "н", "о", "п", "р", "с", "т", "у", "ф", "х", "ц", "ч", "ш", "щ", "ъ", "ы", "ь", "э", "ю", "я" };  
          string[] eng = { "a", "b", "v", "g", "d", "e", "", "j", "z", "i", "", "k", "l", "m", "n", "o", "p", "r", "s", "t", "u", "f", "h", "c", "ch", "sh", "sc", "", "y", "", "e", "iu", "ia" };
          string result="";
          
          row= row.Split(' ')[0]+" "+ row.Split(' ')[1][0]+"." + row.Split(' ')[2][0]+".";
          документ.Изменить();
          //Сообщение("", row);
          документ["Короткое имя"]=row;
    	  row=документ["name_id"];
            
            foreach (char c in row)
                {
                for (int i=0; i<32; i++)
                    {
                    if ((c.ToString()).Equals(rus[i]))
                        result += eng[i];
                    }
                }

          документ["Логин"]=result;
          документ.Сохранить();
    	  
    	  
    	}
    
}
