using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using TFlex.DOCs.Model;
using TFlex.DOCs.Model.Macros;
using TFlex.DOCs.Model.Macros.ObjectModel;
using TFlex.DOCs.Model.References;


public class Oper
                {
                    public DateTime DATA_OP;
                    public string IZG;
                    public int K_OPER;
                    public string NAIM_ST;
                    public decimal NORM_M;
                    public float NORM_M_P;
                    public int NORM_T;
                    public int NORM_T_P;
                    public int NPER;
                    public string NUM_OP;
                    public string OP_OP;
                    public string OSN_TARA;
                    public string PROF;
                    public int RAZR;
                    public string Row_num;
                    public string SHIFR;
                    public string SHIFR_OP;
                    public string SK_TXT;
                    public float VREM;
                }
    
   public class Razm_out
            	{
                	public string EDIZM;
                    public string GOST;
                    public string MARKA;
                    public string MAS;
                    public string NAME;
                    public string NMAT;
                    public string NOTH;
                    public string NPOT;
                    public string NVES;
                    public string OKP;
                    public string POKR_H;
                    public string POKR_S;
                    public string RAZM;
                    public string ROW_NUMBER;
                    public string SHIFR;
                    public string VID;
            	}

    
   public class kat_izvM
                {
                    public string SHIFR;
                    public string NAIM;
                    public string IZD;
                    public string N_IZV;
                    public string SH_IZM;
                    public string DATA_IZ;
                    public string K_PIZM;
                    public string ZADEL;
                    public string VNEDR;
                    public string DATA_VV;
                }

   /*
   public class Razm_out
                {
   	               	public string SHIFR;
                    public string OKP;
                    public int EDIZM;
                    public float NMAT;
                    public float NVES;
                    public float NOTH;
                    public float NPOT;
                    public float POKR_S;
                    public int POKR_H;
                    public float MAS;
                    public string NAME;
                    public string GOST;
                    public string MARKA;
                    public string VID;
                    public string RAZM;     
   	            }
*/

public class Macro : MacroProvider
	
	
{
	
    public Объект  result2=null;
    string guidTrud_out="d1456977-6865-491e-90a8-34be8f84892e";
    string guidKat_razm_out="b24948f4-2f22-4249-bc73-94cf087b6480";
    string guid_kat_izvm="ab60b5b9-eb17-4ca7-9dd5-7d2ac3e082dc";
    //"b24948f4-2f22-4249-bc73-94cf087b6480";
	StringBuilder strBuilder = new StringBuilder();
    StringBuilder strBuilder2 = new StringBuilder();
    ArrayList Tplist = new ArrayList();
    String Oboz;
    string prev_izg="";
    int nper=1;

    public Macro(MacroContext context)
        : base(context)
    {
    System.Diagnostics.Debugger.Launch();
    System.Diagnostics.Debugger.Break();
    }

public override void Run()
    {
 
	Объекты select_tp =ВыбранныеОбъекты ;
    foreach (Объект TP in select_tp)	
    		GetCz(TP);
    	
    
    }
       
public void Children()
    {
    
        Объекты objects= ТекущийОбъект.ДочерниеОбъекты;
        foreach (var child in objects)
        	{
           // Сообщение("",String.Format(" {0} {1}",child.ToString(), child.Тип.ToString()));

        }
        
    
    }
       
public void GetCz(Объект TP)
        {
        
    	
    	//Dictionary<int, string> countries = new Dictionary<int, string>(5);
        //Объект TP = НайтиОбъект("Технологические процессы", "Обозначение = '8А8.603.299'");//8А8.942.348 '8А8.942.348' 8А8.231.556-22
        //8А8603299
        //Объект TP= ТекущийОбъект;                                                                     
        Объекты цехозаходы= TP.ДочерниеОбъекты;
     //   ReferenceObjectCollection ref1=(ReferenceObjectCollection) цехозаходы;
     //   var routes = цехозаходы.OrderBy(o=>o.SystemFields.Order).ToArray();
        //AddOper(цехозаходы[0]);
        /// TP["Обозначение"]=   

        foreach (var cexp in цехозаходы)
                {
                    //strBuilder.AppendLine("Код №"+cexp["Код подразделения"]+" --- "+cexp["Номер цехоперехода"]);    
                    GetOper(cexp, TP["Обозначение"].ToString());
                    
                 }
        
       // Сообщение("",TP.Тип.ToString());
        //save_kat_razm_out()
        save_trud_out();
        GetZag(TP);
        GetIzm(TP);
         
         	
      //   Export_out(guidTrud_out,"Выгрузка в базу данных Trud_out");
     //   Export_out(guidKat_azm_out,"Выгрузка из kat_razm_out");
          // Export_trud_out();
    }
    
    
    
public void GetIzm(Объект izg)
	{

kat_izvM kat_izvM=new kat_izvM();


 var ДСЕ = izg.СвязанныйОбъект["Изготавливаемая ДСЕ"];
        if (ДСЕ != null)            //Если список не пустой, т.е. такие объекты найдены
        {
             if (ДСЕ.РодительскиеОбъекты!=null)
             	{
               // string izdel = "Изделие   "+Parent(ДСЕ)["Обозначение"].ToString(); //+"  Обозначение изделия "+Parent(ДСЕ)["Обозначение"].ToString() ;       
                 }
        }



      kat_izvM.SHIFR = izg["Обозначение"].ToString(); 
        kat_izvM.NAIM = izg["Наименование"].ToString();
        /*  
        kat_izvM.IZD =;
        kat_izvM.N_IZV =;
        kat_izvM.SH_IZM =;
        kat_izvM.DATA_IZ =;
        kat_izvM.K_PIZM =;
        kat_izvM.ZADEL =;
        kat_izvM.VNEDR =;
        kat_izvM.DATA_VV =;*/

        Объекты izms=izg.СвязанныеОбъекты["Изменения"];
                           //= техпроцесс.СвязанныеОбъекты["Изменения"]["Внедрить"].ToString();
                         //   Сообщение("",izm[0]["№ изменения"].ToString());
                            if (izms != null && izms.Count>0)
                                {
                                
                                        foreach (Объект izm in izms)
                                            {
                                                int izmmax= 0;
                                                if  (izmmax<=izm["ID"] )
                                                    {
                                                    izmmax=izm["ID"];
                                                    //Переменная["$graph_04"]=izm["Внедрить"];
                                                    var date_iz= DateTime.Parse(izm["Дата создания"]);
                                                    
                                                    kat_izvM.N_IZV=izm["№ изменения"].ToString();
                                                    kat_izvM.VNEDR =izm["Внедрить"].ToString();
                                                    kat_izvM.DATA_IZ =date_iz.ToShortDateString();
                                                    Объект  izv=izm.СвязанныйОбъект["Извещение об изменении"];
                                                    
                                                    

                                                    if (izv != null ) 
                                                        {
                                                           kat_izvM.SH_IZM = izv["Выпущено на"].ToString();                                               
                                                           //    Переменная["$graph_izm"] = String.Format("изм № {0} \r\n {1} \r\n от {2}", izm["№ изменения"].ToString() , izv["Выпущено на"].ToString(),tip.ToShortDateString()  );
                                                            
                                                         }

                                                    }
                                                }

                                    }
                            else
                            	{

// 
kat_izvM.VNEDR =izg["Внедрить для МК"].ToString();
kat_izvM.DATA_IZ = izg["Дата изменения"].ToString();
kat_izvM.N_IZV=izg["Номер изменения"].ToString();

                            }
    Save_KAT_IZVM_OUT(kat_izvM) ;                       	
    }
    
    

public void CopyTempleteDbf(string[] file_arr)
	{  
	foreach (string file in file_arr)
		{
            string path = @"C:\AEMexport\template\"+file;
            string newPath = @"C:\AEMexport\"+file;
            FileInfo fileInf = new FileInfo(path);
            if (fileInf.Exists)
            {
               fileInf.CopyTo(newPath, true);      
               // альтернатива с помощью класса File
               // File.Copy(path, newPath, true);
            }
        }        
    }
 

public void MoveDbf(string[] file_arr,string shifr)
    	{
	        string dir= @"C:\AEMexport\";
	        string direxport=dir+@"\export\"+shifr+@"\";
        	foreach (string file in file_arr)
                {
                            string path = dir+file;
                            if (Directory.Exists(direxport) == false)
                            	Directory.CreateDirectory(direxport);
                            string newPath = direxport+file.Replace(".","_tef.");
                            FileInfo fileInf = new FileInfo(path);
                            if (fileInf.Exists)
                            {
                               fileInf.MoveTo(newPath);       
                               // альтернатива с помощью класса File
                               // File.Move(path, newPath);
                            }
                 }
        }
    
public void ExportOutDbf()
	{   
	    string [] file_arr={"kat_razm.DBF","trud.DBF","trud.fpt","kat_izvM.dbf"};
	    //CopyTempleteDbf(file_arr);
	    
	    ReferenceObjectCollection listObj= getRef(guidTrud_out);
	    
	     var list = from item in listObj select (item.GetObjectValue("SHIFR").ToString());
                   // List<string> list2 = list.ToList();
                    

                 //   var distinctList = list.Distinct();
	    
	    
	   // string shifr="ЭСКИЗ145020547";
	  //  Export_out(ReferenceObjectCollection listObj,string rolename, string shifr)
	    foreach (string shifr in list.Distinct())
	    	{
	    	    CopyTempleteDbf(file_arr);
        	    Export_out(guidTrud_out,"Выгрузка в базу данных Trud_out",shifr);
                Export_out(guidKat_razm_out,"Выгрузка из kat_razm_out",shifr);
                Export_out(guid_kat_izvm,"Выгрузка в Изменения в конструкторском составе KAT_IZVM",shifr);
                MoveDbf(file_arr,shifr);
            }
	    Сообщение("Обмен данными", "Выгрузка завершена");
    }
    
    
    

    
public void GetZag(Объект izg)
    	{
    	Razm_out zagrazm=new Razm_out();
    	
    	
    	
    	var списокматериалов = izg.СвязанныеОбъекты["Материалы/Заготовки"];
    	 
    	//Сообщение("{}",String.Format("{0}  {1} {2} {3} ",zag[0].ToString(),zag[0]["Обозначение"].ToString(),  zag[0]["Размеры"].ToString()  ));
    	
    	
    /*	 Сообщение("",String.Format("{0}  {1}  {2}",списокматериалов[0].ToString(),
                                            списокматериалов[0]["Обозначение"].ToString(),
                                            списокматериалов[0]["Размеры"].ToString() )
                                    );
    	*/
    	
    	   foreach (var zagitem in списокматериалов)
    	   	{
    	   	  // Сообщение("",String.Format("{0} {1}",zagitem.Тип,zagitem["Основной"]));
               if (zagitem.Тип=="Материал-Заготовка" && zagitem["Основной"])
               	    {
               	var edizDocs=zagitem.СвязанныйОбъект["ЕИ количества"];
               	var edizFox= edizDocs.СвязанныйОбъект["KAT_EDIZ"]["KOD"].ToString();
               	var material=zagitem.СвязанныеОбъекты["Материал"];
               	
             //  	Сообщение("ediz",edizFox);
       //        	Сообщение("mat",material.Count.ToString());
               	
               	        zagrazm.EDIZM=edizFox;
                        //zagrazm.GOST
                        //zagrazm.MARKA
                        //zagrazm.MAS
                        zagrazm.NAME=izg["Наименование"].ToString();
                        //zagrazm.NMAT
                        //zagrazm.NOTH
                        //zagrazm.NPOT
                        //zagrazm.NVES
                        zagrazm.OKP=zagitem["Обозначение"].ToString();
                        //zagrazm.POKR_H
                        //zagrazm.POKR_S
                        zagrazm.RAZM=zagitem["Размеры"].ToString();
                     
                        zagrazm.SHIFR=izg["Обозначение"].ToString();
                        //zagrazm.VID
                        //Сообщение("{}",String.Format("{0}  {1} {2}",zagitem.ToString(),zagitem["Обозначение"].ToString(), zagitem["Размеры"].ToString()  ));
                        Save_kat_razm_out(zagrazm);
               	    }
               }
    	   
    	   
    	  // Export_kat_razm_out();
    	  // Export_trud_out();
        }
   
public void GetOper(Объект cexp,string DCE)
    {
        
        
        
        Объекты операции = cexp.ДочерниеОбъекты;
        foreach (var ioper in операции)
            {

                Oper opertab = new Oper();
                (string osn, string obor) osnast= GetOsn(ioper);
                (string name, string razr) prof= GetIsp(ioper); 
                
                if (prev_izg=="")
                    {
                    prev_izg=cexp["Код подразделения"];
                    opertab.NPER = nper;
                    }
                else if (!prev_izg.Equals(cexp["Код подразделения"]))
                    {
                        prev_izg = cexp["Код подразделения"];
                        nper += 1;
                        opertab.NPER = nper;
                    }
                else if (prev_izg.Equals(cexp["Код подразделения"]))
                    {
                	    
                        prev_izg = cexp["Код подразделения"];
                        opertab.NPER = nper;
                    }

//Сообщение("",String.Format("{0} {1} {2} {3} {3}", cexp["Код подразделения"],prev_izg,prev_izg.Equals(cexp["Код подразделения"]).ToString(),nper ));
            strBuilder.AppendLine(String.Format("oper {0} {1} {2}", ioper["Номер"], ioper["Код"], ioper["Наименование"]));
                opertab.NUM_OP = ioper["Номер"];
                opertab.SHIFR_OP = ioper["Код операции"];
//                опер["Штучное время"]=trud["NORM_T"]*60+trud["NORM_M"];
                
                 if (ioper["Штучное время"]>=1)  
                 	{
                opertab.NORM_T = (ioper["Штучное время"]-ioper["Штучное время"]%60)/60;               
                opertab.NORM_M = ioper["Штучное время"]%60;
                    }
                 else
                 	{
                 	opertab.NORM_T = 0;               
                    opertab.NORM_M = ioper["Штучное время"];
                     }
            //decimal norm_m1 = ioper["Штучное время"];
                opertab.SHIFR = DCE;
                opertab.IZG = cexp["Код подразделения"];
                opertab.NAIM_ST = osnast.obor;
                opertab.OSN_TARA = osnast.osn;
                opertab.OP_OP = GetPer(ioper);
                if (prof!=("",""))
                	{
                        opertab.PROF = prof.name;
                        opertab.RAZR = int.Parse(prof.razr);
                    }

                Tplist.Add(opertab);
            }
    }
       
public (string, string) GetOsn(Объект oper)
    	{
            string strObor = "";
            string strOsn = "";
            

        var списокОснащения = oper.СвязанныеОбъекты["Оснащение"];
            foreach (var osn in списокОснащения)
                {
                    strBuilder.AppendLine(String.Format("osn {0}  {1}", osn["Строка оснащения"].ToString(), osn.Тип.ToString() ));
            if (osn.Тип.ToString() == "Оснащение" || osn.Тип.ToString() == "Инструмент")
                strOsn += "," + osn["Строка оснащения"].ToString();
            if (osn.Тип.ToString() == "Оборудование")
                strObor += "," + osn["Строка оснащения"].ToString();
                }

        
        return (strOsn.TrimStart(new char[] { ',' }), strObor.TrimStart(new char[] { ',' })) ;
        
        }
    
 public (string,string) GetIsp(Объект oper)
        {
            (string name,string razr) prof=("","");
            var списокИсполнителей = oper.СвязанныеОбъекты["Исполнители операции"];
            foreach (var isp in списокИсполнителей)
                {
                                prof.name += " " + isp["Наименование"].ToString();
                                prof.razr += " " + isp["Разряд работ"].ToString();
                }
        return prof;
        }
    
    
 public string GetPer(Объект oper)
        	{
                string strcnst = "";
                var perehod = oper.ДочерниеОбъекты;
                
                foreach (var per in perehod)
                	{
                                strcnst += " " + per["Текст перехода"].ToString();
                    }
            return strcnst;
            }

        	
   /* 	public void GetLinkObject(Объект obj,string linkname,params string[] nameparam)
            {
               
    		
    	
            var	allobj = obj.СвязанныеОбъекты[linkname];
           
            
            
                foreach (var itemobj in allobj)
                    {
                	string strconstr="";
                	   foreach (var itemnamep in nameparam)
                	   	   strconstr+=" "+itemobj[itemnamep];
                    	   strBuilder.AppendLine(String.Format(strconstr));	
                    }
            }
    	
    	
    	
    	        public void GetChaldObjects(Объект obj , params string[] nameparam)
            {
               
                        
              var  allobj = obj.ДочерниеОбъекты;
                      
            
            
                foreach (var itemobj in allobj)
                    {
                    string strconstr="";
                       foreach (var itemnamep in nameparam)
                              strconstr+=" "+itemobj[itemnamep];
                           strBuilder.AppendLine(String.Format(strconstr));    
                    }
            }*/
    	
       public void save_trud_out()
            {
               // Сообщение("Count1", Tplist.Count.ToString());
                foreach (Oper item in Tplist)
                {
                    //Сообщение("",String.Format("{0} {1} {2} {3} {4} {5} {6} ",item.SHIFR, item.SHIFR_OP,item.RAZR,item.IZG,item.NAIM_ST,item.NUM_OP,item.OP_OP));
                    strBuilder2.AppendLine(String.Format("{0} {1} {2} {3} {4} {5} {6} {7} {8}", item.SHIFR, item.SHIFR_OP, item.PROF, item.RAZR, item.IZG, item.NAIM_ST, item.NUM_OP, item.OP_OP, item.OSN_TARA));

                    Объект oper_trud = СоздатьОбъект("TRUD_OUT");
                    //oper_trud["DATA_OP"] = item.DATA_OP;
                    oper_trud["IZG"] = item.IZG;
                    oper_trud["K_OPER"] = item.K_OPER;
                    oper_trud["NORM_M"] = item.NORM_M;
                    oper_trud["NORM_T"] = item.NORM_T;
                    oper_trud["NAIM_ST"] = item.NAIM_ST;
                    
//                    oper_trud["NORM_M_P"] = item.NORM_M_P;

//опер["Штучное время"]=trud["NORM_T"]*60+trud["NORM_M"];
                    
                    
                    //oper_trud["NORM_T_P"] = item.NORM_T_P;
// */
                    oper_trud["NPER"] = item.NPER;
                    oper_trud["NUM_OP"] = item.NUM_OP;
                    oper_trud["OP_OP"] = item.OP_OP;
                    oper_trud["OSN_TARA"] = item.OSN_TARA;
                   // if (item.PROF!=null)
                    oper_trud["PROF"] = (item.PROF!=null) ? item.PROF:"" ;
                   
                    // if (item.RAZR!=null)
                    oper_trud["RAZR"] = (item.RAZR!=null) ? item.RAZR:"";
                    //oper_trud["Row_num"] = item.Row_num;
                    oper_trud["SHIFR"] = item.SHIFR.Replace(".","");
                    oper_trud["SHIFR_OP"] = item.SHIFR_OP;
                    //oper_trud["SK_TXT"] = item.SK_TXT;
                    //oper_trud["VREM"] = item.VREM;
                    oper_trud.Сохранить();
                    }
                
                
                
                Сообщение("", strBuilder2.ToString());


                
               
        
                }



    public void Save_kat_razm_out(Razm_out zagrazm)
               {
                
               

                   Объект kat_razm_out = СоздатьОбъект("KAT_RAZM_OUT");
                
                    kat_razm_out["EDIZM"]=zagrazm.EDIZM;
                  //  kat_razm_out["GOST"]=zagrazm.GOST;
                 //   kat_razm_out["MARKA"]=zagrazm.MARKA;
                //    kat_razm_out["MAS"]=zagrazm.MAS;
                    kat_razm_out["NAME"]=zagrazm.NAME;
               //  kat_razm_out["NMAT"]=zagrazm.NMAT;
              //   kat_razm_out["NOTH"]=zagrazm.NOTH;
             //    kat_razm_out["NPOT"]=zagrazm.NPOT;
            //     kat_razm_out["NVES"]=zagrazm.NVES;
                    kat_razm_out["OKP"]=zagrazm.OKP;
               //     kat_razm_out["POKR_H"]=zagrazm.POKR_H;
               //     kat_razm_out["POKR_S"]=zagrazm.POKR_S;
                    kat_razm_out["RAZM"]=zagrazm.RAZM;
              //      kat_razm_out["ROW_NUMBER"]=zagrazm.ROW_NUMBER;
                    kat_razm_out["SHIFR"]=zagrazm.SHIFR.Replace(".","");
              //      kat_razm_out["VID"]=zagrazm.VID;
                
                kat_razm_out.Сохранить();
               
           
            }
    
    
        public void Save_KAT_IZVM_OUT(kat_izvM kat_izvM)
               {
                
               
    
                        Объект KAT_IZVM_OUT = СоздатьОбъект("KAT_IZVM_OUT");
                        KAT_IZVM_OUT["SHIFR"]=kat_izvM.SHIFR.Replace(".","");
                        KAT_IZVM_OUT["NAIM"]=kat_izvM.NAIM;
                       // KAT_IZVM_OUT["IZD"]=kat_izvM.IZD;
                        KAT_IZVM_OUT["N_IZV"]=kat_izvM.N_IZV;
                        KAT_IZVM_OUT["SH_IZM"]= (kat_izvM.SH_IZM != null) ? kat_izvM.SH_IZM : "";
                        KAT_IZVM_OUT["DATA_IZ"]=(kat_izvM.DATA_IZ != null) ? kat_izvM.DATA_IZ : "";
                        KAT_IZVM_OUT["K_PIZM"]=(kat_izvM.K_PIZM != null) ? kat_izvM.K_PIZM : "" ;
                        KAT_IZVM_OUT["ZADEL"]=(kat_izvM.ZADEL != null) ? kat_izvM.ZADEL : "" ;
                        KAT_IZVM_OUT["VNEDR"]=(kat_izvM.VNEDR != null) ? kat_izvM.VNEDR : "" ;
                        KAT_IZVM_OUT["DATA_VV"]=(kat_izvM.DATA_VV != null) ? kat_izvM.DATA_VV : "" ;
                        KAT_IZVM_OUT.Сохранить();
                   
           
            }
    
    
    
    
      public void Export_kat_razm_out()
    {

        /*   	ReferenceInfo referenceInfo = serverConnection.ReferenceCatalog.Find(new Guid(SpecRefID_OUT));
             Reference reference = referenceInfo.CreateReference();
             ReferenceObjectCollection reff = reference.Objects;*/

        //Message("", "");
        //TFlex.DOCs.Model.References.ReferenceObjectCollection
        string referenceGuid = "b24948f4-2f22-4249-bc73-94cf087b6480";



        /*var objects = объекты.To<ReferenceObject>();
        var result = objects.ToList();

        foreach (var ro in objects)
            result.AddRange(ro.Children.RecursiveLoad());*/


        //"8d727772-d7e5-4058-b7e1-046c510e7f76";//заказы на оснастку

        ReferenceInfo info = Context.Connection.ReferenceCatalog.Find(referenceGuid);

        Reference reference = info.CreateReference();
        ReferenceObjectCollection listObj = reference.Objects;
        //IEnumerable<Object> refObj
        // var test = listObj.GetEnumerator();
        var test = listObj.ToList();
       // Message("", listObj.GetType().ToString());
      //  Message("", listObj.Count().ToString());


         //DataExchangeAccessor.ExportObjects("Выгрузка из kat_razm_out", test);
        //DataExchangeResultsAccessor dataExchangeResultsAccessor = DataExchangeAccessor.Export("Выгрузка из kat_razm_out");
       //var test2 = new DataExchangeAccessor();
      //modifynum(listObj);

        var test1 = listObj.Last<ReferenceObject>();

     //   Message("", test1.ToString());




        //int номер = ТекущийОбъект["Номер выгрузки"];
        //фильтр = null    


        /*ОбменДанными.Экспортировать(
            "Выгрузка из kat_razm_out",
         //            ,
           //"KAT_RAZM_OUT",
            //String.Format("[Номер журнала выгрузки] = '{0}'", 1111), 
            показыватьДиалог: false,
            фильтр : "[SHIFR] содержит текст"
            );*/


        //ОбменДанными.Экспортировать("Выгрузка из kat_razm_out",);
        ОбменДанными.ЭкспортироватьОбъекты("Выгрузка из kat_razm_out",test,показыватьДиалог : false);
        /*
        
            string наименованиеПравила,
            string справочник = null,
            string фильтр = null,
            string путьКФайлу = null,
            bool показыватьДиалог = true,
            Dictionary<string, Object> дополнительныеДанные = null,
            bool расширенноеЖурналирование = false,
            Объект корневойОбъект = null

        */
       string pred = "c:\\AEMexport\\template.bat";
       System.Diagnostics.Process.Start(pred);

        Сообщение("Обмен данными", "Выгрузка завершена");
    }
      
    

      public ReferenceObjectCollection getRef(string referenceGuid)
      	{
          	ReferenceInfo info = Context.Connection.ReferenceCatalog.Find(referenceGuid);
            Reference reference = info.CreateReference();
            ReferenceObjectCollection listObj = reference.Objects;
            return listObj;
        }
      
      
      
        public void Export_out(string referenceGuid,string rolename, string shifr)
    {

        
        //string referenceGuid = "d1456977-6865-491e-90a8-34be8f84892e";
 //      Сообщение("",rolename);
           var listObj =  getRef(referenceGuid);

        
        
        var routes = listObj.Where(o => o.GetObjectValue("SHIFR").ToString().Equals(shifr));
         var reflist = routes.ToList();
      //  var reflist = listObj.ToList();


      //  Message("", listObj.GetType().ToString());
     //   Message("", listObj.Count().ToString());


        

        //var test1 = listObj.Last<ReferenceObject>();

        //Message("", test1.ToString());




        //int номер = ТекущийОбъект["Номер выгрузки"];
        //фильтр = null    


        /*ОбменДанными.Экспортировать(
            "Выгрузка из kat_razm_out",
         //            ,
           //"KAT_RAZM_OUT",
            //String.Format("[Номер журнала выгрузки] = '{0}'", 1111), 
            показыватьДиалог: false,
            фильтр : "[SHIFR] содержит текст"
            );*/


        //ОбменДанными.Экспортировать("Выгрузка из kat_razm_out",);
     
        /*
        
            string наименованиеПравила,
            string справочник = null,
            string фильтр = null,
            string путьКФайлу = null,
            bool показыватьДиалог = true,
            Dictionary<string, Object> дополнительныеДанные = null,
            bool расширенноеЖурналирование = false,
            Объект корневойОбъект = null

        */
       
      ОбменДанными.ЭкспортироватьОбъекты(rolename,reflist,показыватьДиалог : false);
      //ОбменДанными.ЭкспортироватьОбъекты("Выгрузка в базу данных Trud_out",test,показыватьДиалог : false);
    //   string pred = "c:\\AEMexport\\template.bat";
    //   System.Diagnostics.Process.Start(pred);
        ///*    
       foreach (ReferenceObject refob in reflist)
                    {                       
                        refob.Delete();
                    }
     //  */
        Сообщение("Обмен данными", "Выгрузка завершена");
    }
        
        
        
        
        
        
          public void Export_out2(string referenceGuid,string rolename)
    {

        
        //string referenceGuid = "d1456977-6865-491e-90a8-34be8f84892e";
    //   Сообщение("",rolename);

        ReferenceInfo info = Context.Connection.ReferenceCatalog.Find(referenceGuid);

        Reference reference = info.CreateReference();
        ReferenceObjectCollection listObj = reference.Objects;
        
        var reflist = listObj.ToList();


     //   Message("", listObj.GetType().ToString());
    //    Message("", listObj.Count().ToString());


        

        //var test1 = listObj.Last<ReferenceObject>();

        //Message("", test1.ToString());




        //int номер = ТекущийОбъект["Номер выгрузки"];
        //фильтр = null    


        /*ОбменДанными.Экспортировать(
            "Выгрузка из kat_razm_out",
         //            ,
           //"KAT_RAZM_OUT",
            //String.Format("[Номер журнала выгрузки] = '{0}'", 1111), 
            показыватьДиалог: false,
            фильтр : "[SHIFR] содержит текст"
            );*/


        //ОбменДанными.Экспортировать("Выгрузка из kat_razm_out",);
     
        /*
        
            string наименованиеПравила,
            string справочник = null,
            string фильтр = null,
            string путьКФайлу = null,
            bool показыватьДиалог = true,
            Dictionary<string, Object> дополнительныеДанные = null,
            bool расширенноеЖурналирование = false,
            Объект корневойОбъект = null

        */
       
      ОбменДанными.ЭкспортироватьОбъекты(rolename,reflist,показыватьДиалог : false);
      //ОбменДанными.ЭкспортироватьОбъекты("Выгрузка в базу данных Trud_out",test,показыватьДиалог : false);
       string pred = "c:\\AEMexport\\template.bat";
       System.Diagnostics.Process.Start(pred);
       foreach (ReferenceObject refob in reflist)
                    {
                        
                        refob.Delete();
                    }
     //   Сообщение("Обмен данными", "Выгрузка завершена");
    }
        
        



                public Объект Parent(Объект объект)
                        {
                        
                        if (объект.РодительскиеОбъекты.Count>0)
                                      {
                                    foreach (var подкл in объект.РодительскиеОбъекты)
                                            {
                                                                                                      
                                                     Parent(подкл);
                                             }       
                                    }
                                    
                
               
                    
                
                        if (объект.РодительскиеОбъекты.Count==0)
                                result2 = объект;
                        
                                 return result2;
                        }


}
