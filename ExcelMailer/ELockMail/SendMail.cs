using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Mail;
using System.Net;
using System.Text;
using System.Threading.Tasks;

namespace ExcelMailer.ELockMail
{
    public static class SendMail
    {
        private const string RECIPIENT_EMAIL = "ashish.chaudhary@slc-india.com";

        // Additional Recipients (comma-separated for multiple recipients)
        private const string CC_EMAILS = ""; // Optional CC recipients
        private const string BCC_EMAILS = "ashutosh.s@slc-india.com"; // BCC recipients

        // Email Configuration - Update these with your SMTP settings
        private const string SMTP_SERVER = "smtp.gmail.com"; // Gmail SMTP server
        private const string SMTP_PORT = "587"; // TLS port for Gmail
        private const string SENDER_EMAIL = "workwithaashuu@gmail.com"; // Your Gmail account
        private const string SENDER_PASSWORD = "yrvh wiey tkby lkvt"; // Gmail App Password (NOT regular password)
        private const string SENDER_NAME = "AgriSuraksha E-Lock System"; // Display name

        // Email Settings
        private const bool ENABLE_SSL = true; // Always true for Gmail
        private const int EMAIL_TIMEOUT = 30000; // 30 seconds timeout

        // Error Notification Recipients (for system errors)
        private const string ERROR_NOTIFICATION_EMAILS = "ashutosh.s@slc-india.com";


        public static async Task SendEmailWithAttachmentAsync(byte[] fileData, SummaryData summaryData, DateTime reportDate)
        {
            try
            {
                using (var client = new SmtpClient(SMTP_SERVER))
                {
                    client.Port = int.Parse(SMTP_PORT);
                    client.Credentials = new NetworkCredential(SENDER_EMAIL, SENDER_PASSWORD);
                    client.EnableSsl = ENABLE_SSL;
                    client.Timeout = EMAIL_TIMEOUT;

                    using (var mail = new MailMessage())
                    {
                        mail.From = new MailAddress(SENDER_EMAIL, SENDER_NAME);
                        mail.To.Add(RECIPIENT_EMAIL);
                        mail.Subject = $"Daily E-Lock Monitoring Report - {reportDate:dd-MM-yyyy}";
                        mail.IsBodyHtml = true;
                        mail.Priority = MailPriority.Normal;
                        mail.Body = CreateEmailBody(summaryData, reportDate);

                        // Add CCs
                        if (!string.IsNullOrWhiteSpace(CC_EMAILS))
                        {
                            foreach (var cc in CC_EMAILS.Split(',').Select(cc => cc.Trim()).Where(cc => !string.IsNullOrWhiteSpace(cc)))
                            {
                                mail.CC.Add(cc);
                            }
                        }

                        // Add BCCs
                        if (!string.IsNullOrWhiteSpace(BCC_EMAILS))
                        {
                            foreach (var bcc in BCC_EMAILS.Split(',').Select(bcc => bcc.Trim()).Where(bcc => !string.IsNullOrWhiteSpace(bcc)))
                            {
                                mail.Bcc.Add(bcc);
                            }
                        }

                        // Attach file
                        using (var stream = new MemoryStream(fileData))
                        {
                            string attachmentName = $"ELock_Daily_Report_{reportDate:yyyyMMdd}";
                            string mimeType;
                            string extension;

                            if (fileData.Length > 1000 && !IsLikelyCsv(fileData))
                            {
                                mimeType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                                extension = ".xlsx";
                            }
                            else
                            {
                                mimeType = "text/csv";
                                extension = ".csv";
                            }

                            var attachment = new Attachment(stream, $"{attachmentName}{extension}", mimeType);
                            mail.Attachments.Add(attachment);

                            await client.SendMailAsync(mail);
                            Console.WriteLine("Email sent successfully");
                        }
                    }
                }
            }
            catch (SmtpException smtpEx)
            {
                Console.WriteLine($"SMTP Error: {smtpEx.StatusCode} - {smtpEx.Message}");
                throw;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"General Error: {ex.Message}");
                throw;
            }
        }

        private static bool IsLikelyCsv(byte[] fileData)
        {
            try
            {
                string headerSample = Encoding.UTF8.GetString(fileData, 0, Math.Min(fileData.Length, 200));
                return headerSample.Contains("DAILY E-LOCK MONITORING REPORT") || headerSample.Contains(","); // crude CSV check
            }
            catch
            {
                return false;
            }
        }

        public static async Task SendErrorNotificationEmailAsync(Exception ex)
        {
            try
            {
                using (var client = new SmtpClient(SMTP_SERVER))
                {
                    client.Port = int.Parse(SMTP_PORT);
                    client.Credentials = new NetworkCredential(SENDER_EMAIL, SENDER_PASSWORD);
                    client.EnableSsl = ENABLE_SSL;
                    client.Timeout = EMAIL_TIMEOUT;

                    var mail = new MailMessage();
                    mail.From = new MailAddress(SENDER_EMAIL, SENDER_NAME);
                    mail.To.Add(RECIPIENT_EMAIL);

                    if (!string.IsNullOrWhiteSpace(ERROR_NOTIFICATION_EMAILS))
                    {
                        var errorEmailList = ERROR_NOTIFICATION_EMAILS.Split(',');
                        foreach (var errorEmail in errorEmailList)
                        {
                            if (!string.IsNullOrWhiteSpace(errorEmail.Trim()))
                            {
                                mail.CC.Add(errorEmail.Trim());
                            }
                        }
                    }

                    mail.Subject = $"❌ Daily E-Lock Report Generation Failed - {DateTime.Now:dd-MM-yyyy}";
                    mail.IsBodyHtml = true;
                    mail.Priority = MailPriority.High;

                    mail.Body = $@"
                    <html><body style='font-family: Arial, sans-serif;'>
                    <div style='background-color: #e74c3c; color: white; padding: 20px; text-align: center;'>
                        <h1>⚠️ Report Generation Failed</h1>
                    </div>
                    <div style='padding: 20px;'>
                        <h2>Error Details:</h2>
                        <p><strong>Time:</strong> {DateTime.Now:dd-MM-yyyy HH:mm:ss}</p>
                        <p><strong>Error Message:</strong> {ex.Message}</p>
                        <p><strong>Stack Trace:</strong></p>
                        <pre style='background-color: #f8f9fa; padding: 10px; border-radius: 5px;'>{ex.StackTrace}</pre>
                        
                        <h3>Recommended Actions:</h3>
                        <ul>
                            <li>Check database connectivity</li>
                            <li>Verify API endpoints are accessible</li>
                            <li>Review application logs</li>
                            <li>Contact system administrator if issue persists</li>
                        </ul>
                    </div>
                    </body></html>";

                    await client.SendMailAsync(mail);
                }
            }
            catch (Exception emailEx)
            {
                Console.WriteLine($"Failed to send error notification email: {emailEx.Message}");
            }
        }
        public static string CreateEmailBody(SummaryData summaryData, DateTime reportDate)
        {
            var sb = new StringBuilder();

            sb.AppendLine("<!DOCTYPE html>");
            sb.AppendLine("<html><head><meta charset='utf-8'>");
            sb.AppendLine("<style>");
            sb.AppendLine("body { font-family: Arial, sans-serif; margin: 20px; }");
            sb.AppendLine(".header { background-color: #2c3e50; color: white; padding: 20px; text-align: center; }");
            sb.AppendLine(".summary-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(250px, 1fr)); gap: 15px; margin: 20px 0; }");
            sb.AppendLine(".summary-card { border: 1px solid #ddd; border-radius: 8px; padding: 15px; background-color: #f8f9fa; }");
            sb.AppendLine(".summary-card h3 { margin-top: 0; color: #2c3e50; }");
            sb.AppendLine(".summary-value { font-size: 24px; font-weight: bold; color: #e74c3c; }");
            sb.AppendLine(".status-ongoing { color: #e74c3c; }");
            sb.AppendLine(".status-closed { color: #27ae60; }");
            sb.AppendLine(".status-opening { color: #f39c12; }");
            sb.AppendLine(".footer { margin-top: 30px; padding: 15px; background-color: #ecf0f1; border-radius: 5px; }");
            sb.AppendLine("</style></head><body>");

            // Header
            sb.AppendLine("<div class='header'>");
            sb.AppendLine("<h1>🔒 Daily E-Lock Monitoring Report</h1>");
            sb.AppendLine($"<h2>{reportDate:dd MMMM yyyy}</h2>");
            sb.AppendLine("</div>");

            // Executive Summary
            sb.AppendLine("<h2>📊 Executive Summary</h2>");
            sb.AppendLine("<div class='summary-grid'>");

            sb.AppendLine("<div class='summary-card'>");
            sb.AppendLine("<h3>Total E-Locks</h3>");
            sb.AppendLine($"<div class='summary-value'>{summaryData.TotalELocks}</div>");
            sb.AppendLine("</div>");

            sb.AppendLine("<div class='summary-card'>");
            sb.AppendLine("<h3>Assigned E-Locks</h3>");
            sb.AppendLine($"<div class='summary-value'>{summaryData.AssignedELocks}</div>");
            sb.AppendLine("</div>");

            sb.AppendLine("<div class='summary-card'>");
            sb.AppendLine("<h3>Currently Open E-Locks</h3>");
            sb.AppendLine($"<div class='summary-value status-ongoing'>{summaryData.CurrentlyOpenELock}</div>");
            sb.AppendLine("</div>");

            sb.AppendLine("<div class='summary-card'>");
            sb.AppendLine("<h3>Total Opened Today</h3>");
            sb.AppendLine($"<div class='summary-value'>{summaryData.TotalOpenELock}</div>");
            sb.AppendLine("</div>");

            sb.AppendLine("</div>");

            // Status Breakdown
            sb.AppendLine("<h2>📈 Status Breakdown</h2>");
            sb.AppendLine("<div class='summary-grid'>");

            sb.AppendLine("<div class='summary-card'>");
            sb.AppendLine("<h3>🔄 Ongoing Operations</h3>");
            sb.AppendLine($"<div class='summary-value status-ongoing'>{summaryData.StatusOngoingCount}</div>");
            sb.AppendLine("</div>");

            sb.AppendLine("<div class='summary-card'>");
            sb.AppendLine("<h3>✅ Completed Operations</h3>");
            sb.AppendLine($"<div class='summary-value status-closed'>{summaryData.StatusClosedCount}</div>");
            sb.AppendLine("</div>");

            sb.AppendLine("<div class='summary-card'>");
            sb.AppendLine("<h3>📤 Opening Operations</h3>");
            sb.AppendLine($"<div class='summary-value status-opening'>{summaryData.StatusOpeningCount}</div>");
            sb.AppendLine("</div>");

            sb.AppendLine("</div>");

            // Important Notes
            sb.AppendLine("<div class='footer'>");
            sb.AppendLine("<h3>📌 Important Notes</h3>");
            sb.AppendLine("<ul>");
            sb.AppendLine("<li><strong>Real-time Data:</strong> Battery levels and shackle status are updated with live API data</li>");
            sb.AppendLine("<li><strong>Ongoing Operations:</strong> Warehouses currently open and operational</li>");
            sb.AppendLine("<li><strong>Completed Operations:</strong> Warehouses that were opened and closed today</li>");
            sb.AppendLine("<li><strong>Alert:</strong> Monitor ongoing operations for extended open durations</li>");
            sb.AppendLine("</ul>");

            sb.AppendLine($"<p><strong>Report Generated:</strong> {DateTime.Now:dd-MM-yyyy HH:mm:ss} IST</p>");
            sb.AppendLine("<p><strong>System:</strong> AgriSuraksha E-Lock Monitoring System</p>");
            sb.AppendLine("</div>");

            sb.AppendLine("</body></html>");

            return sb.ToString();
        }
        public static async Task SendNoDataNotificationEmailAsync(DateTime reportDate)
        {
            try
            {
                using (var client = new SmtpClient(SMTP_SERVER))
                {
                    client.Port = int.Parse(SMTP_PORT);
                    client.Credentials = new NetworkCredential(SENDER_EMAIL, SENDER_PASSWORD);
                    client.EnableSsl = ENABLE_SSL;
                    client.Timeout = EMAIL_TIMEOUT;

                    var mail = new MailMessage();
                    mail.From = new MailAddress(SENDER_EMAIL, SENDER_NAME);
                    mail.To.Add(RECIPIENT_EMAIL);

                    if (!string.IsNullOrWhiteSpace(BCC_EMAILS))
                    {
                        var bccList = BCC_EMAILS.Split(',');
                        foreach (var bcc in bccList)
                        {
                            if (!string.IsNullOrWhiteSpace(bcc.Trim()))
                            {
                                mail.Bcc.Add(bcc.Trim());
                            }
                        }
                    }

                    mail.Subject = $"📊 Daily E-Lock Report - No Data Found - {reportDate:dd-MM-yyyy}";
                    mail.IsBodyHtml = true;
                    mail.Priority = MailPriority.Normal;

                    mail.Body = $@"
                    <html><body style='font-family: Arial, sans-serif;'>
                    <div style='background-color: #f39c12; color: white; padding: 20px; text-align: center;'>
                        <h1>📊 Daily E-Lock Report</h1>
                        <h2>{reportDate:dd MMMM yyyy}</h2>
                    </div>
                    <div style='padding: 20px;'>
                        <h2>ℹ️ No Data Available</h2>
                        <p>The daily report for <strong>{reportDate:dd-MM-yyyy}</strong> could not be generated because no data was found for this date.</p>
                        
                        <h3>Possible Reasons:</h3>
                        <ul>
                            <li>No E-Lock operations occurred on this date</li>
                            <li>Database connectivity issues</li>
                            <li>Data filtering parameters excluded all records</li>
                        </ul>
                        
                        <p><strong>Report Generated:</strong> {DateTime.Now:dd-MM-yyyy HH:mm:ss} IST</p>
                        <p><strong>System:</strong> AgriSuraksha E-Lock Monitoring System</p>
                    </div>
                    </body></html>";

                    await client.SendMailAsync(mail);
                    Console.WriteLine("No data notification email sent successfully");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to send no data notification email: {ex.Message}");
            }
        }
    }
}
