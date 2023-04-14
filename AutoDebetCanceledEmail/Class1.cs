using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using ClosedXML.Excel;
using OfficeOpenXml;

namespace IAMScheduler
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnStart_Click(object sender, EventArgs e)
        {
            try
            {
                p_startProcessMonthly();
                Application.Exit();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            //p_startProcessDaily();

        }
        private void p_startProcessMonthly()
        {
            //string startupPath = System.IO.Directory.GetCurrentDirectory();

            //string x = Environment.CurrentDirectory;
            //string filename = "Monthly" + DateTime.Now.ToString("yyyy-MM-dd");
            //DataTable dtMonthly = Database.GetDataMonthly();
            //XLWorkbook wb = new XLWorkbook();

            //wb.Worksheets.Add(dtMonthly, "Test");
            //wb.SaveAs(x + @"\File\Monthly\" + filename + ".xlsx");
            p_sendEmailMonthly();
            Thread.Sleep(5000);
            Application.Exit();
        }
        private void p_startProcessDaily()
        {
            DataTable dtInsco = Database.GetDataInscoMonthlyLoop();
            foreach (DataRow dr in dtInsco.Rows)
            {
                string inscocode = dr["INSCOCODE"].ToString();
                int id = int.Parse(dr["ESCROWINSCOID"].ToString());
                DataTable dtEmailDailyLoop = Database.GenerateDataDaily(inscocode, id);
                p_sendEmailDailyLoop(dtEmailDailyLoop);
            }
            p_sendEmailDaily();
        }
        private void p_sendEmailDailyLoop(DataTable dtEmail)
        {
            MailMessage message = new MailMessage();
            SmtpClient smtp = new SmtpClient();
            if (dtEmail.Rows[0]["MAKER_EMAIL"].ToString() != "")
                message.To.Add(new MailAddress(dtEmail.Rows[0]["MAKER_EMAIL"].ToString()));

            for (int i = 0; i < dtEmail.Rows.Count; i++)
            {
                message.To.Add(new MailAddress(dtEmail.Rows[i]["INSCO_PICEMAIL"].ToString()));
            }
            if (dtEmail.Rows[0]["HEAD_EMAIL"].ToString() != "")
                message.To.Add(new MailAddress(dtEmail.Rows[0]["HEAD_EMAIL"].ToString()));

            message.Subject = "IAMS Scheduler Generate Data Daily";
            message.IsBodyHtml = true; //to make message body as html  
            message.Body = dtEmail.Rows[0]["EMAIL_BODY"].ToString();
            smtp.Host = "172.30.50.135"; //for gmail host  
                                         //smtp.EnableSsl = true;
            smtp.UseDefaultCredentials = false;
            smtp.Credentials = new NetworkCredential("iams.notification@taf.co.id", "");
            smtp.DeliveryMethod = SmtpDeliveryMethod.Network;
            smtp.Send(message);
        }
        private void p_sendEmailDaily()
        {
            DataTable dtEmail = Database.GetUserEmailDaily();
            MailMessage message = new MailMessage();
            SmtpClient smtp = new SmtpClient();
            message.From = new MailAddress("iams.notification@taf.co.id");
            for (int i = 0; i < dtEmail.Rows.Count; i++)
            {
                message.To.Add(new MailAddress(dtEmail.Rows[i]["EMAIL"].ToString()));
            }
            string htmlString = @"<html>
                      <body>
                      <p>Dear Tim Marketing,</p></br>
                      <p>Data Insentif Monthly sudah tersedia di Dashboard IAMS.</p></br>
                      <p>Akses https://sister-uat.taf.co.id/iamsreborn/ </p>
                      <p>pada Menu : Report - Report Data Monthly</p></br>

                      <p>Terima Kasih</p>
                      </body>
                      </html>
                     ";
            message.Subject = " IAMS Daily Generate - " + DateTime.Now.ToShortDateString();
            message.Body = htmlString;
            message.IsBodyHtml = true;
            smtp.Host = "172.30.50.135"; //for gmail host  
                                         //smtp.EnableSsl = true;
            smtp.UseDefaultCredentials = false;
            smtp.Credentials = new NetworkCredential("iams.notification@taf.co.id", "");
            smtp.DeliveryMethod = SmtpDeliveryMethod.Network;
            smtp.Send(message);
            //Application.Exit();
        }
        private void p_sendEmailMonthly()
        {
            string x = Environment.CurrentDirectory;
            string result = Database.GenerateDataMonthly();
            DataTable dtEmail = Database.GetUserEmailMonthly();
            MailMessage message = new MailMessage();
            SmtpClient smtp = new SmtpClient();
            message.From = new MailAddress("iams.notification@taf.co.id");
            for (int i = 0; i < dtEmail.Rows.Count; i++)
            {
                message.To.Add(new MailAddress(dtEmail.Rows[i]["EMAIL"].ToString()));
            }
            message.Subject = "IAMS Monthly Generate - " + DateTime.Now.ToShortDateString();
            string htmlString = @"<html>
                      <body>
                      <p>Dear Tim Marketing,</p></br>
                      <p>Data Insentif Monthly sudah tersedia di Dashboard IAMS.</p></br>
                      <p>Akses https://sister-uat.taf.co.id/iamsreborn/ </p>
                      <p>pada Menu : Report - Report Data Monthly</p></br>

                      <p>Terima Kasih</p>
                      </body>
                      </html>
                     ";
            message.IsBodyHtml = true; //to make message body as html  
            message.Body = htmlString;
            System.Net.Mail.Attachment attachment;
            attachment = new System.Net.Mail.Attachment(x + @"\File\Monthly\" + filename + ".xlsx");
            message.Attachments.Add(attachment);
            smtp.Port = 587;
            smtp.Host = "172.30.50.135"; //for gmail host  
                                         //smtp.EnableSsl = true;
            smtp.UseDefaultCredentials = false;
            smtp.Credentials = new NetworkCredential("iams.notification@taf.co.id", "");
            smtp.DeliveryMethod = SmtpDeliveryMethod.Network;
            smtp.Send(message);
        }
        public async Task StartBgProses()
        {
            Task longRunningTask = ProcessMonthly();
            await longRunningTask;
        }
        public Task ProcessMonthly() // assume we return an int from this long running operation 
        {
            return Task.Run(() =>
            {
                p_startProcessMonthly();
            });

        }
        private void Form1_Load(object sender, EventArgs e)
        {
            try
            {

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void Form1_Shown(object sender, EventArgs e)
        {
            Thread.Sleep(5000);
            p_startProcessMonthly();
        }
    }
}