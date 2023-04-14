using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using AutoDebetCanceledEmail.Repositories;
using AutoDebetCanceledEmail.Services;
using System.Windows.Forms;
using ClosedXML.Excel;
using OfficeOpenXml;
using System.Data.Entity;

namespace AutoDebetCanceledEmail
{
    public class Program
    {
        public string filename = "AutoDebetMonitoring " + DateTime.Now.ToString("yyyy-MM-dd");
        static void Main(string[] args)
        {
            Console.WriteLine("Start Sending Email..");
            GenerateAttachment();
            var _AutoDebetCanceledEmailRepository = new AutoDebetCanceledEmailRepository();
            var data = _AutoDebetCanceledEmailRepository.Get();

            var msgTo = System.Configuration.ConfigurationManager.AppSettings["mailTo"].ToString();

            var _AutoDebetCanceledService = new AutoDebetCanceledService();
            var mailMsg = _AutoDebetCanceledService.MessageConfigure(msgTo, CreateBody(data));
            var mailConfigure = _AutoDebetCanceledService.MailConfigure();

            _AutoDebetCanceledService.Run(mailConfigure, mailMsg);
            Console.WriteLine("Sending email has been successfully!");
        }
        static void GenerateAttachment()
        {
            var AutoRepo = new AutoDebetCanceledEmailRepository();
            string startupPath = System.IO.Directory.GetCurrentDirectory();

            string x = Environment.CurrentDirectory;
            string filename = "AutoDebetMonitoring " + DateTime.Now.ToString("yyyy-MM-dd");
            DataTable dt = AutoRepo.Get();
            XLWorkbook wb = new XLWorkbook();

            wb.Worksheets.Add(dt, "AutoDebet");
            wb.SaveAs(x + @"\FileAttachment\" + filename + ".xlsx");
        }
        static string CreateBody(DataTable data)
        {
            string body = "";
            int day = DateTime.Now.Day;
            int month = DateTime.Now.Month;
            int year = DateTime.Now.Year;

            string printDay = (day).ToString();
            string printMonth = CultureInfo.CurrentCulture.DateTimeFormat.GetMonthName(month);
            string printYear = year.ToString();


            body += "<p>Dear All,</p>";
            //body += $"<p>Berikut merupakan list Tolakan SKAD per tanggal {printDay} {printMonth} {printYear} :<br /></ p > ";
            body += $"<p>Berikut kami lampirkan kontrak-kontrak pengajuan SKAD yang berstatus <span style='background-color:#f1c40f'> Reject Bank dan Finalisasi Return </span> pertanggal {printDay} {printMonth} {printYear} <br/></p>";
            body += "Mohon dicek kembali & dapat diajukan ulang";
            body += "<br /><br />";
            //body += @"<table border=1 cellpadding=1 cellspacing=1 style='width: 537px'>
            body += @"<table border=1>
                        <tr>
                            <td style='background-color:#9bbcfd; text-align:center'>No Kontrak</td>
                            <td style='background-color:#9bbcfd; text-align:center'>Nama Pemilik Rek</td>
                            <td style='background-color:#9bbcfd; text-align:center'>No Rek</td>
                            <td style='background-color:#9bbcfd; text-align:center'>Nama Bank</td>
                            <td style='background-color:#9bbcfd; text-align:center'>Keterangan Reject</td>
                            <td style='background-color:#9bbcfd; text-align:center'>Tanggal Reject</td>
                            <td style='background-color:#9bbcfd; text-align:center'>Cabang</td>
                        </tr>";
            foreach (DataRow row in data.Rows)
            {
                body += "<tr>";
                body += $"<td>{row["AGRMNT_NO"].ToString()}</td>";
                body += $"<td>{row["BANK_ACC_NAME"].ToString()}</td>";
                body += $"<td>{row["BANK_ACC_NO"].ToString()}</td>";
                body += $"<td>{row["BANK_NAME"].ToString()}</td>";
                body += $"<td>{row["BANK_REG_NOTES"].ToString()}</td>";
                body += $"<td>{row["BANK_RESULT_DT"].ToString()}</td>";
                body += $"<td>{row["OFFICE_NAME"].ToString()}</td>";
                body += "</tr>";
            }
            body += "</table>";
            body += "<br /><br />";
            body += "Terima kasih";
            return body;
        }

    }
}
