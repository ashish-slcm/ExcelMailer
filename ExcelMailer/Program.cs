using ExcelMailer.ELockMail;

public class ExcelMailExport
{
    static async Task Main(String[] args)
    {
        try
        {
            Console.WriteLine("Calling SendDailyReport Method");
            await ELockReportMail.SendDailyReportAsync();
            Console.WriteLine("Closing SendDailyReport Method");
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error: " + ex.Message);
        }
    }
}
