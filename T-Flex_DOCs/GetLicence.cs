using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net.NetworkInformation;
using System.Text;
using System.Windows.Forms;
using TFlex.DOCs.Model;
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
    
    }

    public void licence()
    {
        licence(Context.Connection);
    }

        public void licence(ServerConnection connect)
    {
        var activeConnects = connect.GetActiveConnections();
        StringBuilder str_licenses = new StringBuilder();
        var licenses = connect.GetAvailableLicenses();

        int TotalUsedLicence = 0;
        foreach (var license in licenses)
        {
            string NameLicence = license.Module.Name;
            var AccessibleLicence = license.AccessibleCount;
            var TotalLicence = license.TotalCount;
            var triger = TotalLicence/3;
            str_licenses.Append($"{NameLicence} используется {TotalLicence - AccessibleLicence} доступно {AccessibleLicence} всего {TotalLicence} порог {triger} \r\n");
            TotalUsedLicence = TotalUsedLicence + (TotalLicence - AccessibleLicence);
            bool restart = false;
            //if (AccessibleLicence < triger)
            if (restart)
            {
                str_licenses.Append($"RESTART");
                string pathToDirectory = @"C:\scripts\restart_net_service";
                //string pathToFileString = "restart_tflex-docs_merge.bat";
                string pathToFileString = "restart_tflex-docs_copy_work.bat";
                ProcessStartInfo infoStartProcess = new ProcessStartInfo();
                infoStartProcess.WorkingDirectory = pathToDirectory;
                infoStartProcess.FileName = pathToFileString;
                infoStartProcess.CreateNoWindow = false;
                infoStartProcess.WindowStyle = ProcessWindowStyle.Hidden;
                Process restartnet = Process.Start(infoStartProcess);
                restartnet.WaitForExit();

                break;
            }
        }


        str_licenses.Append($"Всего лицензий используется {TotalUsedLicence}\r\n");
        str_licenses.Append($"______\r\n");
        int countconnect = 0;

        foreach (var activeConnect in activeConnects)
        {
            var hostname = activeConnect.HostName;
            var username = activeConnect.UserName;
            string status = SendPing(activeConnect.HostName);
            countconnect++;
            str_licenses.Append($"{countconnect} {hostname} {username} {status} \r\n");
        }

        str_licenses.Append($"Подключено: {activeConnects.Count}\r\n");
        Message("", str_licenses.ToString());
        Save(str_licenses.ToString().Split(new string[] { "\r\n" }, StringSplitOptions.RemoveEmptyEntries).ToList(), "c:\\work\\licence","lic2.txt");
        //var sep = new string[] { "sss" } ;
        ;
    }

    public string SendPing(string hostname)
    {
        Ping pingSender = new System.Net.NetworkInformation.Ping();
        string status = "";
        try
        {
            var h1 = pingSender.Send(hostname);
            if (h1.Status.ToString().Equals("Success"))
            {
                status = h1.Status.ToString();
            }
            else status = "offline";
        }
        catch (Exception e)
        {
            // Console.WriteLine(e);
            status = "offline";
        }
        return status;
    }

    /// <summary>
    /// Сохранение List в файл
    /// </summary>
    public void Save(List<string> strList, string path,string filename)
    {
        DateTime datenow = DateTime.Now;
        if (Directory.Exists(path) == false)
        {
            DirectoryInfo dir = Directory.CreateDirectory(path);
            Сообщение("Информация.", "Папка для сохранения экспортированных данных не была найдена и была создана по пути " + dir);
        }

        //string text = "text";
        try
        {
            using (StreamWriter sw = new StreamWriter($@"{path}\{(datenow.ToShortDateString()).Replace(".", "")}_{(datenow.ToShortTimeString()).Replace(":", "")}_{filename}", false, System.Text.Encoding.Default))
            {
                foreach (var str in strList)
                    sw.WriteLine(str.ToString());
            }


        }
        catch (Exception e)
        {
            Console.WriteLine(e.Message);
        }
    }
}
