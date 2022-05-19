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
	Объект DCE;
	List<string> str =new List<string>();
    public Macro(MacroContext context)
        : base(context)
    {
    }

    public override void Run()
    {
    
    }
    
    
    
    public void Test()
  {
    String s="";
    
    List<string> n = new List<string> ();
    //Объекты документы = НайтиОбъекты("SPEC","[SHIFR] = '"+DCE+"'");
    Объекты документы = НайтиОбъекты("SPEC","[NAIM] содержит 'МАГНИТ'");
    Message("__"," " +документы.Count.ToString());
    
    
    foreach (Объект документ in документы)
    	{
    //	str.Add(DCE+" "+документ["SHIFR"]+" "+документ["POS"]);
        DCE=документ;
    	RecursR(документ);
    	}
    
    foreach (string sr in str)
        s+=sr+"\r\n";    	
     
 //   Message("_0_",s);
    
    
    
    
    
   // Message("__0",документы["SHIFR"].ToList().ToString());
   // Message("__2",str.Count.ToString());
   // Message("__3",str.ToString());
    Message("__4",s);
   
   
   
     try
            {
                using (StreamWriter sw = new StreamWriter(@"C:\001\test.txt"))
                {
                    sw.WriteLine(s);
                }
 
            }
            catch (Exception e)
            {
                Сообщение("", e.Message);
            }
   
   
  }

    
    
     public  void RecursR(Объект parent)
    	{
    	Объекты документы = НайтиОбъекты("SPEC","[SHIFR] = '"+parent["IZD"]+"'");
    	//Message("__2","[IZD] = "+документ["SHIFR"]+"\"");
        	if (документы.Count>0)
        		{
            	    foreach (Объект документ in документы)
                	{
                    	//str.Add(DCE+" "+документ["SHIFR"]+" "+документ["POS"]);
                    	RecursR(документ);
                	}
    	         }
        	// && Int32.Parse(документы["POS"].ToString())==0
        	//string test = ;
        	//int nn=;
        	if (документы.Count==1)
        		{
        		     if (Int32.Parse(документы[0]["POS"].ToString())==0)
        		     	{
                        str.Add(DCE["SHIFR"]+"; "+DCE["NAIM"]+" >> "+документы[0]["SHIFR"]+" ; "+документы[0]["NAIM"]+" "+документы[0]["POS"]);
                        }
            	}
        	
        		
    	}
    
    
    
    public  void Recurs(Объект chald)
    	{
    	Объекты документы = НайтиОбъекты("SPEC","[IZD] = '"+chald["SHIFR"]+"'");
    	//Message("__2","[IZD] = "+документ["SHIFR"]+"\"");
        	if (документы.Count>0)
        		{
            	    foreach (Объект документ in документы)
                	{
                    	str.Add(документ["SHIFR"]);
                    	Recurs(документ);
                	}
    	         }
        	
        		
    	}
    
    
}
