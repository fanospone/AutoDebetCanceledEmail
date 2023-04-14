using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net;
using System.Net.Mail;
using System.Globalization;
using static System.Windows.Forms.AxHost;
using System.Runtime.Remoting.Messaging;
using System.IO;

namespace AutoDebetCanceledEmail.Services
{
    public class AutoDebetCanceledService
    {
        string smtpHost = System.Configuration.ConfigurationManager.AppSettings["smtp"].ToString();
        string smtpMailFrom = System.Configuration.ConfigurationManager.AppSettings["mailFrom"].ToString();
        string directory = Environment.CurrentDirectory;
        System.Net.Mail.Attachment attachment;
        Program x = new Program();

        public SmtpClient MailConfigure()
        {
            var smtp = new SmtpClient();

            smtp.Host = smtpHost;
            smtp.UseDefaultCredentials = false;
            smtp.Credentials = new NetworkCredential(smtpMailFrom, "");
            smtp.DeliveryMethod = SmtpDeliveryMethod.Network;

            return smtp;
        }
        public MailMessage MessageConfigure(string msgTo, string msgBody)
        {
            var msg = new MailMessage();
            int day = DateTime.Now.Day;
            int month = DateTime.Now.Month;
            int year = DateTime.Now.Year;

            string printDay = string.Empty;
            string printMonth = string.Empty;
            string printYear = string.Empty;


            if (day == 1)
            {
                if (month == 1)
                {

                    printDay = DateTime.DaysInMonth((year - 1), 12).ToString();
                    printMonth = CultureInfo.CurrentCulture.DateTimeFormat.GetMonthName(12);
                    printYear = (year - 1).ToString();
                }
                else
                {
                    printDay = DateTime.DaysInMonth(year, (month - 1)).ToString();
                    printMonth = CultureInfo.CurrentCulture.DateTimeFormat.GetMonthName(month - 1);
                    printYear = year.ToString();
                }
            }
            else
            {
                printDay = (day).ToString();
                printMonth = CultureInfo.CurrentCulture.DateTimeFormat.GetMonthName(month);
                printYear = year.ToString();
            }
            foreach (var mailTo in msgTo.Split(new[] { ";" }, StringSplitOptions.RemoveEmptyEntries))
            {
                msg.To.Add(mailTo);
            }

            msg.From = new MailAddress(smtpMailFrom);
            msg.Subject = $"Tolakan SKAD {printDay} {printMonth} {printYear} ";//"Auto Debet Monitoring";
            msg.IsBodyHtml = true;
            msg.Body = msgBody;

            attachment = new System.Net.Mail.Attachment(directory + @"\FileAttachment\" + x.filename + ".xlsx");
            msg.Attachments.Add(attachment);

            return msg;
        }
        public void Run(SmtpClient smtp, MailMessage msg)
        {
            //string subPath = @"E:\TAFS_APP\AutoDebetMonitoringEmail\FileAttachment";
            //string subPath = @"D:\Rifano\Work\AutoDebetCanceledEmail\AutoDebetCanceledEmail\bin\Debug\FileAttachment";
            //System.IO.DirectoryInfo di = new DirectoryInfo(subPath);
            //string[] xlsxList = Directory.GetFiles(subPath, "*.xlsx");
            
            //foreach (string f in xlsxList)
            //{
            //    System.IO.StreamReader file = new System.IO.StreamReader(f);
            //    file.Close();
            //    file.Dispose();
            //    File.Delete(f);
                
            //    //System.IO.File.Delete(subPath);
            //    //File.Delete(f);
            //}
            //foreach (FileInfo file in di.GetFiles())
            //{
            //    File.Delete(file.FullName);
            //    //file.Delete();
            //}
            smtp.Send(msg);
            //string[] files = Directory.GetFiles(subPath);
            //string[] dirs = Directory.GetDirectories(subPath);
            //foreach (string file in files)
            //{
            //    File.SetAttributes(file, FileAttributes.Normal);
            //    File.Delete(file);
            //}
            //System.IO.File.Delete(attachment);
        }
    }
}
