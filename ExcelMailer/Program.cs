using System;
using System.Data;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Collections.Generic;
using ClosedXML.Excel;
using System.Windows.Forms;

public class ExcelMailExport
{
    public static void Main()
    {
        try
        {
            DataSet ds = DALUtility.clsSQLExecute.Exec_Dataset_sp("[dbo].[GetWHLockDetailsByDate_Test_v6]");
            DataTable dt = ds.Tables[0];

            string excelPath = "data.xlsx";
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Data");
                for (int col = 0; col < dt.Columns.Count; col++)
                {
                    worksheet.Cell(1, col + 1).Value = dt.Columns[col].ColumnName;
                }
                for (int row = 0; row < dt.Rows.Count; row++)
                {
                    for (int col = 0; col < dt.Columns.Count; col++)
                    {
                        worksheet.Cell(row + 2, col + 1).Value = dt.Rows[row][col]?.ToString();
                    }
                }

                workbook.SaveAs(excelPath);
                Console.WriteLine("Excel file saved.");
            }

            string recipientEmail = "mr.ashish199sddgdsb6@gmail.com";
            string senderEmail = "ashish.chaudhardfgy@slc-india.com";
            string senderPassword = "jbxj sdgagbarryr vgrgeblk wfsva";

            var smtpClient = new SmtpClient("smtp.gmail.com")
            {
                Port = 587,
                Credentials = new NetworkCredential(senderEmail, senderPassword),
                EnableSsl = true
            };

            var message = new MailMessage
            {
                From = new MailAddress(senderEmail),
                Subject = "Stored Procedure Excel Report",
                Body = "Please find the attached Excel file.",
            };
            message.To.Add(recipientEmail);
            message.Attachments.Add(new Attachment(excelPath));

            smtpClient.Send(message);
            Console.WriteLine("Email sent.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error: " + ex.Message);
        }
    }
}
