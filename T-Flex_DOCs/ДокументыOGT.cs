using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using TFlex.DOCs.Model.Macros;
using TFlex.DOCs.Model.Macros.ObjectModel;
using FoxProShifrsNormalizer;
using TFlex.DOCs.Model.References.Files;
using TFlex.DOCs.Model;
using TFlex.DOCs.Model.References;
using TFlex.DOCs.Model.Desktop;

public class Macro : MacroProvider
{
    public Macro(MacroContext context)
        : base(context)
    {

    	
#if DEBUG
        System.Diagnostics.Debugger.Launch();
        System.Diagnostics.Debugger.Break();
#endif
    }

    public override void Run()
    {


    }

    public void EventCreateLink()
    {
        Объект документ = ТекущийОбъект;
        if (документ.СвязанныйОбъект["Скан документа1"] == null)
            LinkcopyOboz();
    }
    
    public  void Refresh()
    {
        Guid DocOGT = new Guid("500d4bcf-e02c-4b2e-8f09-29b64d4e7513");
        ReferenceInfo info = Context.Connection.ReferenceCatalog.Find(DocOGT);
        Reference reference = info.CreateReference();
        reference.Refresh();


        Guid fileguid = new Guid("a0fcd27d-e0f2-4c5a-bba3-8a452508e6b3");
        ReferenceInfo infofile = Context.Connection.ReferenceCatalog.Find(fileguid);
        Reference referencefile = infofile.CreateReference();
        referencefile.Refresh();
        /*List<ReferenceObject> result = reference.Objects.ToList();
            return result;*/



        var fileReference = new FileReference(Context.Connection);
        Объект документ = ТекущийОбъект;
        var skanfile = документ.СвязанныйОбъект["Скан документа1"];
        if (skanfile != null)
        {
            var filepath = skanfile["Относительный путь"].ToString();
           FileObject file = fileReference.FindByRelativePath($@"{filepath}") as FileObject;
            if (file is null)
                return;
            file.GetHeadRevision();
            fileReference.Refresh();
        }

    }


    public void Обновить_содержимое_папки(Объект папка)
    {
        if (папка != null)
        {
            string path = папка.Параметр["Относительный путь"];
            FileReference reference = new FileReference();
            FileReferenceObject file = reference.FindByRelativePath(path);
            if (file != null)
                file.GetHeadRevision();
        }
    }


    public  void LinkDel()
    {
        Объект документ = ТекущийОбъект;
        var skanfile = документ.СвязанныйОбъект["Скан документа1"];
        if (skanfile != null)
        {
            документ.СвязанныйОбъект["Скан документа1"] = null;
            документ.Сохранить();
        }

    }

    public void RemoveDotFileName(Объект файл)
    {
        string filename = файл["Наименование"].ToString();
        if (filename.EndsWith(".pdf"))
        {
            filename =  filename.Substring(0, filename.Length - 4).Replace(".","")+".pdf";
            файл["Наименование"] = filename;
            файл.Сохранить();
        }
    }


    public void RemoveDotSelectFileName()
    {
        Объекты файлы = ВыбранныеОбъекты;
        foreach (var файл1 in файлы)
        {
            RemoveDotFileName(файл1);
        }
    }


    public void Test()
    {
        var usr = ТекущийПользователь;
        var obj = ТекущийОбъект;
    }


    public void Create_change()
    {

        Объект ДокументОГТ = ТекущийОбъект;
        var izm = ДокументОГТ["№ изменения"];
        var skanfile = ДокументОГТ.СвязанныйОбъект["Скан документа1"];
        string tip_doc = ДокументОГТ["Тип документа"].ToString();
        ДокументОГТ.Изменить();
        if (izm.ToString().Equals(""))
        {
            ДокументОГТ["№ изменения"] = 1;
        }
        else
        {
            ДокументОГТ["№ изменения"] = int.Parse(izm.ToString()) + 1;
        }
        Объект измененияОГТ = СоздатьОбъект("Изменения ОГТ", "Изменения ОГТ");
        измененияОГТ["Номер"] = ДокументОГТ["Номер"];
        измененияОГТ["Тип документа"] = ДокументОГТ["Тип документа"].ToString();
        измененияОГТ["Номер изменения"] = izm;
        измененияОГТ["Шифр"] = ДокументОГТ["Шифр извещения"].ToString();

        if (skanfile != null && !tip_doc.ToString().Equals(""))
        { 
            измененияОГТ.СвязанныйОбъект["Файл"] = skanfile;
            ДокументОГТ.СвязанныйОбъект["Скан документа1"] = null;            
            var folder = НайтиОбъект("Файлы", @$"[Наименование] = 'Старые версии {tip_doc}' И [Относительный путь] содержит 'Старые версии МК, КЭ.ТИ,СТО'");
            skanfile.Изменить();
            skanfile.РодительскийОбъект = folder;
            skanfile.Сохранить();
            Desktop.CheckIn((ReferenceObject)skanfile , "Изменили актуальную версию объекта", false);
        }
        else
            Message("Ошибка", "Параметр Тип Документа не заполнен");
        измененияОГТ.Сохранить();
        ДокументОГТ.Подключить("Изменения", измененияОГТ);
        ДокументОГТ.Сохранить();
    }
    

    public  void LinkcopyOboz()
    {
        Normalizer norm = new Normalizer();
        norm.setprefix = false;
        Объект документ = ТекущийОбъект;
        var Oboz=документ["Обозначение детали, узла"].ToString().Split(' ').First();
        string link_skan = документ["Скан документа"].ToString();
        string tip_doc = документ["Тип документа"].ToString();
        string filename = link_skan.Split('\\').Last();
        if (!filename.EndsWith(".pdf"))
            filename = filename.Replace(".", "") + ".pdf";
        //filename = filename + ".pdf";

        var Oboz_dot = norm.NormalizeShifrsFromFox(Oboz);

        /*if (документ.СвязанныйОбъект["Скан документа1"] == null)
            Message("", "No Link");*/
        // if (документ["№ изменения"]!=null)
        //int izm = Int32.Parse(документ["№ изменения"].ToString());
        
        var izm = документ["№ изменения"];
        try
        {
            if (!izm.ToString().Equals(""))
            {
                Oboz += $"({izm})";
                Oboz_dot += $"({izm})";
            }
            //[Наименование] = '8А8053337.pdf' И [Относительный путь] содержит 'Архив ОГТ\ogtMK\КЭ\'
            //[Наименование] = '8А8053337.pdf' И [Относительный путь] содержит 'Архив ОГТ\ogtMK\КЭ\'
            //Архив ОГТ\ogtMK\КЭ
            //Старые версии МК, КЭ.ТИ,СТО
            //8А8316033(4).pdf
            //[Наименование] = '8А8316033(4).pdf' И[Относительный путь] содержит 'Архив ОГТ\ogtMK\КЭ' И [Относительный путь] не содержит 'Старые версии МК, КЭ.ТИ,СТО'
            Oboz += ".pdf";
            Oboz_dot += ".pdf";
            Объект файлы = null;
            Объект файлы_dot = null;
            Объект файлы_link = null;

            string query = "";
            if (!tip_doc.ToString().Equals(""))
            {
                
                if (tip_doc.Equals("МК+КЭ"))                	
                    query = @$"И [Относительный путь] содержит 'Архив ОГТ\ogtMK\МК' И [Относительный путь] не содержит 'Старые версии МК, КЭ.ТИ,СТО' И [Относительный путь] не содержит 'изменения'";
                else if (tip_doc.Equals("МКосн"))
                    query = @$"И [Относительный путь] содержит 'Архив ОГТ\ogtMK\МК-разовая' И [Относительный путь] не содержит 'Старые версии МК, КЭ.ТИ,СТО' И [Относительный путь] не содержит 'изменения'";
                else
                    query = @$"И [Относительный путь] содержит 'Архив ОГТ\ogtMK\{tip_doc}' И [Относительный путь] не содержит 'Старые версии МК, КЭ.ТИ,СТО' И [Относительный путь] не содержит 'изменения'";



                файлы = НайтиОбъект("Файлы", @$"[Наименование] = '{Oboz}' {query}");
                файлы_dot = НайтиОбъект("Файлы", @$"[Наименование] = '{Oboz_dot}' {query}");
                файлы_link = НайтиОбъект("Файлы", @$"[Наименование] = '{filename}' {query}");
            }
            else
                Message("Ошибка", "Параметр Тип Документа не заполнен");

    

            

            //Объект файлы_dot = НайтиОбъект("Файлы", "[Относительный путь] содержит '" + Oboz_dot + ".pdf'");
            //Объект файлы_link = НайтиОбъект("Файлы", "[Относительный путь] содержит '" + filename);
            if (файлы == null && файлы_dot== null && файлы_link==null)
            	{
                //Message("Ошибка" , @$"Файл {Oboz} в папке {query} не найден");
                Message("Ошибка" , @$"Файл {Oboz} в папке АрхивОГТ не найден");
                //     Сообщение("Ошибка", $"Файл {Oboz} в справочнике Docs не найден ");
               //   Сообщение("Ошибка", $"Файл {Oboz} /r/n ({Oboz_dot}) /r/n {filename} в справочнике Docs не найден ");
            }
            else
            {
                документ.Изменить();
                if (файлы != null)
                    документ.СвязанныйОбъект["Скан документа1"] = файлы;
                else if (файлы_dot != null)
                    документ.СвязанныйОбъект["Скан документа1"] = файлы_dot;
                else if (файлы_link != null)
                    документ.СвязанныйОбъект["Скан документа1"] = файлы_link;
                документ.Сохранить();
            }
        
        }
        catch(Exception e)
        {
            //Сообщение("Файл в справочнике Docs не найден ",String.Format("{e.ToString} {файлы.ToString()}"));
        }
}
}
