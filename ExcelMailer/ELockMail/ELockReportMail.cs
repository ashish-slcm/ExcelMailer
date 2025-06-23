using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Net.Mail;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using DALUtility;
using Newtonsoft.Json;
using System.Net.Http.Headers;
using ClosedXML.Excel;

namespace ExcelMailer.ELockMail
{
    public static class ELockReportMail
    {
        private const string SMTP_SERVER = "smtp.gmail.com";
        private const string SMTP_PORT = "587"; 
        private const string SENDER_EMAIL = "workwithaashuu@gmail.com";
        private const string SENDER_PASSWORD = "yrvh wiey tkby lkvt";
        private const string SENDER_NAME = "AgriSuraksha E-Lock System";

        public static async Task SendDailyReportAsync()
        {
            try
            {
                Console.WriteLine($"Starting daily email report generation at {DateTime.Now:dd-MM-yyyy HH:mm:ss}");

                // Get today's date for filtering
                DateTime reportDate = DateTime.Today;

                // STEP 1: Get raw data from database
                Console.WriteLine("STEP 1: Getting raw data from database...");
                var (reportData, summaryData) = GetReportDataForDate(reportDate);

                if (reportData.Rows.Count == 0)
                {
                    Console.WriteLine("No data found for daily report");
                    await SendMail.SendNoDataNotificationEmailAsync(reportDate);
                    return;
                }

                Console.WriteLine($"Found {reportData.Rows.Count} records in database");
                foreach (DataRow row in reportData.Rows)
                {
                    try
                    {
                        string company = row["Company"]?.ToString()?.ToUpper();
                        string imei = row["ElockIMEI"]?.ToString();
                        string assetId = row["AssetId"]?.ToString();
                        string whCode = row["WH_Code"]?.ToString();



                        if (string.IsNullOrWhiteSpace(whCode))
                        {
                            row["Battery %"] = "N/A";
                            row["Actual Shackle Status"] = "N/A";
                        }
                        else if (company == "TRACOLOGIC")
                        {
                            if (!string.IsNullOrWhiteSpace(imei))
                            {
                                Console.WriteLine($"Fetching TracologicData data for IMEI Number: {imei}");
                                var (battery, status, message) = await GetTracologicDataSyncAsync(imei);

                                row["Battery %"] = !string.IsNullOrEmpty(battery) ? battery : "N/A";
                                row["Actual Shackle Status"] = !string.IsNullOrEmpty(status) ? status : "N/A";
                            }
                            else
                            {
                                row["Battery %"] = "N/A";
                                row["Actual Shackle Status"] = "N/A";
                            }
                        }
                        else if (company == "IMZ")
                        {
                            if (!string.IsNullOrWhiteSpace(assetId))
                            {
                                Console.WriteLine($"Fetching IMZ data for asset ID: {assetId}");
                                var (battery, status, message) = await GetIMZDataSyncAsync(assetId);

                                row["Battery %"] = !string.IsNullOrEmpty(battery) ? battery : "N/A";
                                row["Actual Shackle Status"] = !string.IsNullOrEmpty(status) ? status : "N/A";
                            }
                            else
                            {
                                row["Battery %"] = "N/A";
                                row["Actual Shackle Status"] = "N/A";
                            }
                        }
                        else
                        {
                            // Handle unknown companies
                            row["Battery %"] = "N/A nnn";
                            row["Actual Shackle Status"] = "N/A nnnn";
                        }
                    }
                    catch (Exception rowEx)
                    {
                        Console.WriteLine($"Error updating row with real-time data: {rowEx.Message}");

                        // Set error values for this row and continue
                        try
                        {
                            row["Battery %"] = "Error";
                            row["Actual Shackle Status"] = "Error";
                        }
                        catch (Exception setEx)
                        {
                            Console.WriteLine($"Error setting error values: {setEx.Message}");
                        }
                    }
                }

                // STEP 5: Generate Excel file with enhanced data
                Console.WriteLine("STEP 5: Generating Excel file with enhanced data...");
                byte[] excelData = GenerateExcelFile(reportData, summaryData, reportDate);

                if (excelData == null || excelData.Length == 0)
                {
                    Console.WriteLine("Failed to generate Excel file");
                    return;
                }

                Console.WriteLine($"Generated Excel file: {excelData.Length} bytes");

                //STEP 6: Send email with Excel attachment
                Console.WriteLine("STEP 6: Sending email with attachment...");
                await SendMail.SendEmailWithAttachmentAsync(excelData, summaryData, reportDate);

                Console.WriteLine($"Daily email report sent successfully at {DateTime.Now:dd-MM-yyyy HH:mm:ss}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error sending daily email report: {ex.Message}");
                Console.WriteLine($"Stack trace: {ex.StackTrace}");

                // Send error notification email
                try
                {
                    await SendMail.SendErrorNotificationEmailAsync(ex);
                }
                catch (Exception emailEx)
                {
                    Console.WriteLine($"Error sending error notification email: {emailEx.Message}");
                }
            }
        }
        private static (DataTable reportData, SummaryData summaryData) GetReportDataForDate(DateTime reportDate)
        {
            try
            {
                SqlParameter[] prm = {
                    new SqlParameter("@search", DBNull.Value),
                    new SqlParameter("@filterDate", reportDate),
                    new SqlParameter("@statusFilter", DBNull.Value),
                    new SqlParameter("@sortOrder", "DESC"),
                    new SqlParameter("@sortColumn", "OpeningRequestTime"),
                    new SqlParameter("@pageNumber", 1),
                    new SqlParameter("@pageSize", 50000) // Get all records
                };

                DataSet ds = clsSQLExecute.Exec_Dataset_sp("[dbo].[GetWHLockDetailsByDate_Test_v6]", prm);

                DataTable dt = ds.Tables[0];
                DataTable summaryTable = ds.Tables.Count > 2 ? ds.Tables[2] : null;

                // Remove existing battery and status columns if they exist
                if (dt.Columns.Contains("Actual Shackle Status"))
                    dt.Columns.Remove("Actual Shackle Status");
                if (dt.Columns.Contains("Battery %"))
                    dt.Columns.Remove("Battery %");

                // Add new columns for enhanced data
                dt.Columns.Add("Actual Shackle Status", typeof(string));
                dt.Columns.Add("Battery %", typeof(string));

                // Initialize with default values
                foreach (DataRow row in dt.Rows)
                {
                    row["Battery %"] = "Loading...";
                    row["Actual Shackle Status"] = "Loading...";
                }

                // Extract summary data
                var summary = new SummaryData();
                if (summaryTable != null && summaryTable.Rows.Count > 0)
                {
                    summary = ExtractSummaryData(summaryTable.Rows[0]);
                }

                Console.WriteLine($"Retrieved {dt.Rows.Count} records from database");
                return (dt, summary);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error getting report data: {ex.Message}");
                throw;
            }
        }

        private static async Task<(string battery, string status, string message)> GetTracologicDataSyncAsync(string imei)
        {
            try
            {
                if (string.IsNullOrEmpty(imei))
                {
                    return ("N/A", "N/A", "Invalid IMEI");
                }

                using (var httpClient = new HttpClient())
                {
                    httpClient.Timeout = TimeSpan.FromSeconds(15);

                    var requestContent = new
                    {
                        curPage = 1,
                        pageSize = 1,
                        deviceCodes = new[] { imei },
                        dataType = 0,
                        gpsStartTime = "",
                        gpsEndTime = ""
                    };

                    string jsonContent = JsonConvert.SerializeObject(requestContent);
                    var content = new StringContent(jsonContent, System.Text.Encoding.UTF8, "application/json");

                    var request = new HttpRequestMessage(HttpMethod.Post, "http://api.hhdlink.top/hhdApi/public/api/iotDeviceData/getHistoryData")
                    {
                        Content = content
                    };

                    request.Headers.Add("accessKeyId", "iL1kI4vHKeGONhoa");
                    request.Headers.Add("accessSecret", "1PzEjQoOcfL43yi4mo0Vh4wNY7EPkinW");
                    request.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

                    var response = httpClient.SendAsync(request).Result;

                    if (response.IsSuccessStatusCode)
                    {
                        string responseContent = await response.Content.ReadAsStringAsync();
                        dynamic json = JsonConvert.DeserializeObject(responseContent);

                        if (json?.returnCode != null && json.returnCode?.ToString() == "200")
                        {
                            var records = json.data?.records;

                            var rec = records[0];

                            if (rec.battery != null && rec.deviceStatus != null)
                            {
                                int battery = (int)rec.battery;
                                string deviceStatus = rec.deviceStatus?.ToString() ?? "";
                                string shackleStatus = ParseShackleStatus(deviceStatus);

                                return ($"{battery}%", shackleStatus, "Data retrieved successfully");
                            }
                            else
                            {
                                return ("N/A", "Unknown", "Battery or status data not available");
                            }
                        }
                    }

                    return ("N/A", "N/A", "API Error");
                }
            }
            catch (Exception ex)
            {
                return ("N/A", "N/A", $"Error: {ex.Message}");
            }
        }

        private static async Task<(string battery, string status, string message)> GetIMZDataSyncAsync(string assetId)
        {
            try
            {
                if (string.IsNullOrEmpty(assetId) || !int.TryParse(assetId, out int assetIdInt))
                {
                    return ("N/A", "N/A", "Invalid Asset ID");
                }

                using (var httpClient = new HttpClient())
                {
                    httpClient.Timeout = TimeSpan.FromSeconds(15);

                    var requestContent = new
                    {
                        requesttype = "LIVETRACK",
                        vendorcode = "EX",
                        request = new
                        {
                            username = "sohan",
                            pin = "0dc53edbe6cee7ab26b2bc5f2ea28d92",
                            ipaddress = "49.249.65.178",
                            clienttype = "web",
                            accid = 2856,
                            assetid = assetIdInt
                        }
                    };

                    var content = new StringContent(JsonConvert.SerializeObject(requestContent), Encoding.UTF8, "application/json");
                    var request = new HttpRequestMessage(HttpMethod.Post, "https://api-ns1.imztech.io/excise/allrequest/")
                    {
                        Content = content
                    };

                    request.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

                    var response = await httpClient.SendAsync(request);

                    if (!response.IsSuccessStatusCode)
                    {
                        return ("N/A", "N/A", $"HTTP {response.StatusCode}");
                    }

                    string responseContent = await response.Content.ReadAsStringAsync();
                    dynamic json = JsonConvert.DeserializeObject(responseContent);

                    if (json?.resultcode == 0)
                    {
                        var data = json.response?.data;
                        if (data?.battery != null && data?.@lock != null)
                        {
                            int battery = Convert.ToInt32(data.battery);
                            int lockStatus = Convert.ToInt32(data.@lock);
                            return ($"{battery}%", lockStatus == 0 ? "Open" : "Closed", "Success");
                        }
                    }

                    return ("N/A", "N/A", "API Error");
                }
            }
            catch (Exception ex)
            {
                return ("N/A", "N/A", $"Error: {ex.Message}");
            }
        }


        public static string ParseShackleStatus(string deviceStatus)
        {
            if (string.IsNullOrEmpty(deviceStatus))
                return "Unknown";

            deviceStatus = deviceStatus.ToLower();

            if (deviceStatus.Contains("status_151") || deviceStatus.Contains("closed") || deviceStatus.Contains("lock"))
                return "Closed";
            else if (deviceStatus.Contains("status_150") || deviceStatus.Contains("open") || deviceStatus.Contains("unlock"))
                return "Open";
            else if (deviceStatus.Contains("error") || deviceStatus.Contains("fault"))
                return "Error";
            else if (deviceStatus.Contains("offline") || deviceStatus.Contains("disconnect"))
                return "Offline";
            else
                return "Unknown";
        }
        public static byte[] GenerateExcelFile(DataTable reportData, SummaryData summaryData, DateTime reportDate)
        {
            try
            {
                // Try ClosedXML first, fall back to CSV if it fails
                return GenerateExcelFileWithClosedXL(reportData, summaryData, reportDate);
            }
            catch (Exception closedXlEx)
            {
                Console.WriteLine($"ClosedXL failed: {closedXlEx.Message}. Falling back to CSV.");
                return GenerateCSVFile(reportData, summaryData, reportDate);
            }
        }

        public static SummaryData ExtractSummaryData(DataRow summaryRow)
        {
            var summary = new SummaryData();

            if (summaryRow != null)
            {
                summary.TotalELocks = GetSafeInt(summaryRow, "TotalELocks");
                summary.AssignedELocks = GetSafeInt(summaryRow, "AssignedELocks");
                summary.TotalOpenELock = GetSafeInt(summaryRow, "TotalOpenELock");
                summary.CurrentlyOpenELock = GetSafeInt(summaryRow, "CurrentlyOpenELock");
                summary.ClosedELock = GetSafeInt(summaryRow, "ClosedELock");
                summary.CurrentlyOpenWarehouse = GetSafeInt(summaryRow, "CurrentlyOpenWarehouse");
                summary.TotalOpenedWarehouse = GetSafeInt(summaryRow, "TotalOpenedWarehouse");
                summary.StatusClosedCount = GetSafeInt(summaryRow, "StatusClosedCount");
                summary.StatusOngoingCount = GetSafeInt(summaryRow, "StatusOngoingCount");
                summary.StatusOpeningCount = GetSafeInt(summaryRow, "StatusOpeningCount");
                summary.TracologicDevices = GetSafeInt(summaryRow, "TracologicDevices");
                summary.ImzDevices = GetSafeInt(summaryRow, "ImzDevices");
            }

            return summary;
        }

        public static int GetSafeInt(DataRow row, string columnName)
        {
            try
            {
                if (row.Table.Columns.Contains(columnName) && row[columnName] != DBNull.Value)
                {
                    var value = row[columnName];
                    if (value is int intValue) return intValue;
                    if (value is decimal decimalValue) return (int)decimalValue;
                    if (value is double doubleValue) return (int)doubleValue;
                    if (value is string stringValue && int.TryParse(stringValue, out int parsedValue))
                        return parsedValue;
                    return Convert.ToInt32(value);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error extracting int from column {columnName}: {ex.Message}");
            }
            return 0;
        }
        public static byte[] GenerateExcelFileWithClosedXL(DataTable reportData, SummaryData summaryData, DateTime reportDate)
        {
            using (var workbook = new XLWorkbook())
            {
                // SHEET 1: SUMMARY
                var summarySheet = workbook.Worksheets.Add("Daily Summary");
                CreateEnhancedSummarySheet(summarySheet, summaryData, reportDate);

                // SHEET 2: DETAILED REPORT
                var reportSheet = workbook.Worksheets.Add("E-Lock Details");
                CreateDetailedReportSheet(reportSheet, reportData);

                using (var stream = new MemoryStream())
                {
                    workbook.SaveAs(stream);
                    return stream.ToArray();
                }
            }
        }

        public static byte[] GenerateCSVFile(DataTable reportData, SummaryData summaryData, DateTime reportDate)
        {
            var csvContent = new StringBuilder();

            // Add summary section
            csvContent.AppendLine("DAILY E-LOCK MONITORING REPORT");
            csvContent.AppendLine($"Report Date,{reportDate:dd-MM-yyyy}");
            csvContent.AppendLine($"Generated On,{DateTime.Now:dd-MM-yyyy HH:mm:ss}");
            csvContent.AppendLine("");

            // Add summary statistics
            csvContent.AppendLine("MAIN STATISTICS");
            csvContent.AppendLine($"Total E-Locks,{summaryData.TotalELocks}");
            csvContent.AppendLine($"Assigned E-Locks,{summaryData.AssignedELocks}");
            csvContent.AppendLine($"Total Open E-Lock,{summaryData.TotalOpenELock}");
            csvContent.AppendLine($"Currently Open E-Lock,{summaryData.CurrentlyOpenELock}");
            csvContent.AppendLine($"Closed E-Lock,{summaryData.ClosedELock}");
            csvContent.AppendLine($"Currently Open Warehouse,{summaryData.CurrentlyOpenWarehouse}");
            csvContent.AppendLine($"Total Opened Warehouse,{summaryData.TotalOpenedWarehouse}");
            csvContent.AppendLine("");

            // Add status breakdown
            csvContent.AppendLine("STATUS BREAKDOWN");
            csvContent.AppendLine($"Closed Status Count,{summaryData.StatusClosedCount}");
            csvContent.AppendLine($"Ongoing Status Count,{summaryData.StatusOngoingCount}");
            csvContent.AppendLine($"Opening Status Count,{summaryData.StatusOpeningCount}");
            csvContent.AppendLine("");
            csvContent.AppendLine("");

            // Add detailed data header
            csvContent.AppendLine("DETAILED E-LOCK REPORT");

            // Define export columns
            var exportColumns = new Dictionary<string, string>
            {
                {"ElockIMEI", "E-lock IMEI"},
                {"Company", "Company"},
                {"State", "State"},
                {"WH_Code", "WH Code"},
                {"WH_Name", "WH Name"},
                {"WH Open Request By Name", "Open Request By"},
                {"WH Opening Request Time", "Opening Request Time"},
                {"WH Open Approved By Name", "Open Approved By"},
                {"WH Opening Approval Time", "Opening Approval Time"},
                {"WH Close Request By Name", "Close Request By"},
                {"WH Closing Request Time", "Closing Request Time"},
                {"WH Close Approved By Name", "Close Approved By"},
                {"WH Closing Approval Time", "Closing Approval Time"},
                {"WH Open Duration", "Duration"},
                {"StatusType", "Status"},
                {"Actual Shackle Status", "Current Shackle Status"},
                {"Battery %", "Battery Level"}
            };

            // Add headers
            var headers = exportColumns.Values.ToArray();
            csvContent.AppendLine(string.Join(",", headers.Select(h => $"\"{h}\"")));

            // Add data rows
            foreach (DataRow dataRow in reportData.Rows)
            {
                var values = new List<string>();
                foreach (var columnPair in exportColumns)
                {
                    string columnName = columnPair.Key;
                    object cellValue = "";

                    if (reportData.Columns.Contains(columnName))
                    {
                        cellValue = dataRow[columnName] ?? "";
                    }

                    // Escape CSV values
                    string value = cellValue.ToString().Replace("\"", "\"\"");
                    values.Add($"\"{value}\"");
                }
                csvContent.AppendLine(string.Join(",", values));
            }

            return Encoding.UTF8.GetBytes(csvContent.ToString());
        }


        public static void CreateEnhancedSummarySheet(IXLWorksheet summarySheet, SummaryData summaryData, DateTime reportDate)
        {
            // Title
            summarySheet.Cell("A1").Value = "DAILY E-LOCK MONITORING REPORT";
            summarySheet.Cell("A1").Style.Font.Bold = true;
            summarySheet.Cell("A1").Style.Font.FontSize = 18;
            summarySheet.Cell("A1").Style.Fill.BackgroundColor = XLColor.DarkBlue;
            summarySheet.Cell("A1").Style.Font.FontColor = XLColor.White;
            summarySheet.Range("A1:D1").Merge();
            summarySheet.Range("A1:D1").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

            // Report Information
            int row = 3;
            summarySheet.Cell(row, 1).Value = "Report Date:";
            summarySheet.Cell(row, 2).Value = reportDate.ToString("dd-MM-yyyy");
            summarySheet.Cell(row, 1).Style.Font.Bold = true;
            row++;

            summarySheet.Cell(row, 1).Value = "Generated On:";
            summarySheet.Cell(row, 2).Value = DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss");
            summarySheet.Cell(row, 1).Style.Font.Bold = true;
            row += 2;

            // Main Summary Statistics
            summarySheet.Cell(row, 1).Value = "MAIN STATISTICS";
            summarySheet.Cell(row, 1).Style.Font.Bold = true;
            summarySheet.Cell(row, 1).Style.Fill.BackgroundColor = XLColor.LightBlue;
            summarySheet.Range(row, 1, row, 3).Merge();
            row++;

            AddSummaryRow(summarySheet, ref row, "Total E-Locks", summaryData.TotalELocks);
            AddSummaryRow(summarySheet, ref row, "Assigned E-Locks", summaryData.AssignedELocks);
            AddSummaryRow(summarySheet, ref row, "Total Open E-Lock", summaryData.TotalOpenELock);
            AddSummaryRow(summarySheet, ref row, "Currently Open E-Lock", summaryData.CurrentlyOpenELock);
            AddSummaryRow(summarySheet, ref row, "Closed E-Lock", summaryData.ClosedELock);
            AddSummaryRow(summarySheet, ref row, "Currently Open Warehouse", summaryData.CurrentlyOpenWarehouse);
            AddSummaryRow(summarySheet, ref row, "Total Opened Warehouse", summaryData.TotalOpenedWarehouse);
            row++;

            // Status Breakdown
            summarySheet.Cell(row, 1).Value = "STATUS BREAKDOWN";
            summarySheet.Cell(row, 1).Style.Font.Bold = true;
            summarySheet.Cell(row, 1).Style.Fill.BackgroundColor = XLColor.LightGreen;
            summarySheet.Range(row, 1, row, 3).Merge();
            row++;

            AddSummaryRow(summarySheet, ref row, "Closed Status Count", summaryData.StatusClosedCount);
            AddSummaryRow(summarySheet, ref row, "Ongoing Status Count", summaryData.StatusOngoingCount);
            AddSummaryRow(summarySheet, ref row, "Opening Status Count", summaryData.StatusOpeningCount);

            // Format columns
            summarySheet.Column(1).Width = 30;
            summarySheet.Column(2).Width = 15;
            summarySheet.Column(3).Width = 20;
        }

        public static void AddSummaryRow(IXLWorksheet sheet, ref int row, string label, object value)
        {
            sheet.Cell(row, 1).Value = label;
            sheet.Cell(row, 1).Style.Font.Bold = true;
            sheet.Cell(row, 2).Value = value.ToString();
            sheet.Cell(row, 2).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            row++;
        }

        public static void CreateDetailedReportSheet(IXLWorksheet reportSheet, DataTable reportData)
        {
            // Define export columns
            var exportColumns = new Dictionary<string, string>
            {
                {"ElockIMEI", "E-lock IMEI"},
                {"Company", "Company"},
                {"State", "State"},
                {"WH_Code", "WH Code"},
                {"WH_Name", "WH Name"},
                {"WH Open Request By Name", "Open Request By"},
                {"WH Opening Request Time", "Opening Request Time"},
                {"WH Open Approved By Name", "Open Approved By"},
                {"WH Opening Approval Time", "Opening Approval Time"},
                {"WH Close Request By Name", "Close Request By"},
                {"WH Closing Request Time", "Closing Request Time"},
                {"WH Close Approved By Name", "Close Approved By"},
                {"WH Closing Approval Time", "Closing Approval Time"},
                {"WH Open Duration", "Duration"},
                {"StatusType", "Status"},
                {"Actual Shackle Status", "Current Shackle Status"},
                {"Battery %", "Battery Level"}
            };

            // Add headers
            int col = 1;
            foreach (var columnPair in exportColumns)
            {
                var headerCell = reportSheet.Cell(1, col);
                headerCell.Value = columnPair.Value;
                headerCell.Style.Font.Bold = true;
                headerCell.Style.Fill.BackgroundColor = XLColor.LightGray;
                col++;
            }

            // Add data rows
            int row = 2;
            foreach (DataRow dataRow in reportData.Rows)
            {
                col = 1;
                foreach (var columnPair in exportColumns)
                {
                    string columnName = columnPair.Key;
                    object cellValue = "";

                    if (reportData.Columns.Contains(columnName))
                    {
                        cellValue = dataRow[columnName] ?? "";
                    }

                    // Handle IMEI formatting
                    if (columnName == "ElockIMEI" && !string.IsNullOrEmpty(cellValue.ToString()))
                    {
                        reportSheet.Cell(row, col).SetValue(cellValue.ToString());
                        reportSheet.Cell(row, col).Style.NumberFormat.Format = "@";
                    }
                    else
                    {
                        reportSheet.Cell(row, col).Value = cellValue.ToString();
                    }

                    col++;
                }
                row++;
            }

            // Format the sheet
            reportSheet.ColumnsUsed().AdjustToContents();

            // Apply conditional formatting
            ApplyConditionalFormatting(reportSheet, exportColumns, row);
        }
        public static void ApplyConditionalFormatting(IXLWorksheet sheet, Dictionary<string, string> columns, int maxRow)
        {
            // Find battery and status columns
            int batteryCol = 0, statusCol = 0, shackleCol = 0;
            int col = 1;

            foreach (var columnPair in columns)
            {
                if (columnPair.Key == "Battery %") batteryCol = col;
                else if (columnPair.Key == "StatusType") statusCol = col;
                else if (columnPair.Key == "Actual Shackle Status") shackleCol = col;
                col++;
            }

            // Apply formatting
            for (int i = 2; i < maxRow; i++)
            {
                // Battery level formatting
                if (batteryCol > 0)
                {
                    var batteryCell = sheet.Cell(i, batteryCol);
                    var batteryValue = batteryCell.Value.ToString();

                    if (batteryValue.Contains("%"))
                    {
                        if (int.TryParse(batteryValue.Replace("%", ""), out int batteryLevel))
                        {
                            if (batteryLevel < 20)
                                batteryCell.Style.Font.FontColor = XLColor.Red;
                            else if (batteryLevel < 50)
                                batteryCell.Style.Font.FontColor = XLColor.Orange;
                            else
                                batteryCell.Style.Font.FontColor = XLColor.Green;
                        }
                    }
                }

                // Status formatting
                if (statusCol > 0)
                {
                    var statusCell = sheet.Cell(i, statusCol);
                    var statusValue = statusCell.Value.ToString().ToLower();

                    switch (statusValue)
                    {
                        case "ongoing":
                            statusCell.Style.Font.FontColor = XLColor.Red;
                            statusCell.Style.Font.Bold = true;
                            break;
                        case "closed":
                            statusCell.Style.Font.FontColor = XLColor.Green;
                            statusCell.Style.Font.Bold = true;
                            break;
                        case "opening":
                            statusCell.Style.Font.FontColor = XLColor.Orange;
                            statusCell.Style.Font.Bold = true;
                            break;
                    }
                }

                // Shackle status formatting
                if (shackleCol > 0)
                {
                    var shackleCell = sheet.Cell(i, shackleCol);
                    var shackleValue = shackleCell.Value.ToString().ToLower();

                    if (shackleValue == "open")
                        shackleCell.Style.Font.FontColor = XLColor.Red;
                    else if (shackleValue == "closed")
                        shackleCell.Style.Font.FontColor = XLColor.Green;
                }
            }
        }

    }
}
