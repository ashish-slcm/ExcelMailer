using System;
using System.Data;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Collections.Generic;
using ClosedXML.Excel;
using DALUtility;
using ExcelMailer.ELockMail;

public class ExcelMailExport
{
    static async Task Main(String[] args)
    {
        try
        {
            Console.WriteLine("Calling SendDailyReport");
            await ELockReportMail.SendDailyReportAsync();
            Console.WriteLine("Closing SendDailyReport");
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error: " + ex.Message);
        }
    }
}
