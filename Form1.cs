using System;
using System.Windows.Forms;
using System.IO;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Sheets.v4;
using Google.Apis.Gmail.v1;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using System.Threading;
using System.Collections.Generic;
using MailKit.Net.Imap;
using MailKit.Security;
using MailKit;
using MailKit.Search;
using System.Configuration;
using System.ComponentModel;
using System.Linq;
using MimeKit;
using System.Data.SqlClient;
using MAPPOnBoardingStats.Models;
using MAPPOnBoardingStats.Utils;
using System.Threading.Tasks;

namespace MAPPOnBoardingStats
{


    public partial class Form1 : Form
    {
        List<PRIInfo> ThisPRIInfo = new List<PRIInfo>();
        List<HSMCInfo> ThisHSMCInfo = new List<HSMCInfo>();
        private static string[] Scopes = {
            GmailService.Scope.GmailReadonly
        };
        private static string ApplicationName = "Gmail API .NET Quickstart";
        public Form1()
        {
            InitializeComponent();
        }
        private void GetPaceRptIArray()
        {
            try
            {
                var filePace_Report_Instructions = "1YIQejjcFCcQsAIqzMNr7zKhsTD75ULUmBEMo-0_aP88";
                {
                    string[] Scopes = { SheetsService.Scope.SpreadsheetsReadonly };
                    string ApplicationName = "Google Sheets API .NET Quickstart";
                    UserCredential credential;

                    using (var stream =
                        new FileStream("PRI-credentials.json", FileMode.Open, FileAccess.Read))
                    {
                        // The file token.json stores the user's access and refresh tokens, and is created
                        // automatically when the authorization flow completes for the first time.
                        //string credPath = "PaceRptI-token.json";
                        string credPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "PaceRptI-token.json");
                        credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                            GoogleClientSecrets.Load(stream).Secrets,
                            Scopes,
                            "user",
                            CancellationToken.None,
                            new FileDataStore(credPath, true)).Result;
                        Console.WriteLine("Credential file saved to: " + credPath);
                    }

                    // Create Google Sheets API service.
                    var service = new SheetsService(new BaseClientService.Initializer()
                    {
                        HttpClientInitializer = credential,
                        ApplicationName = ApplicationName,
                    });

                    // Define request parameters.

                    String spreadsheetId = filePace_Report_Instructions;
                    String range = "A1:G";
                    Google.Apis.Sheets.v4.SpreadsheetsResource.ValuesResource.GetRequest request =
                            service.Spreadsheets.Values.Get(spreadsheetId, range);

                    // Prints the names and other details of Hotels not Submitting the Data for the MAPP Reports:
                    ValueRange response = request.Execute();
                    IList<IList<Object>> values = response.Values;
                    if (values != null && values.Count > 0)
                    {
                        var data = response.Values;

                        // Get the headers and get the index of the ldap and the approval status
                        // using the names you use in the headers
                        var headers = values[0];
                        var PMSIndex = headers.IndexOf("PMS");
                        var RequiredReportsIndex = headers.IndexOf("Required Reports");

                        //foreach (var row in values)
                        for (int i = 0; i < values.Count; i++)
                        {
                            // Print columns A and E, which correspond to indices 0 and 4.
                            var row = values[i];
                            var c1 = new PRIInfo();
                            c1.PMSName = row[PMSIndex].ToString().ToUpper();
                            c1.Report1 = row[RequiredReportsIndex].ToString();
                            ThisPRIInfo.Add(c1);
                        }
                    }
                    else
                    {
                        MessageBox.Show("No data found.");
                    }
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show($"Unable to Get the Data for PMS ReportNames.\n\nError message: {ex.Message}\n\n" +
                $"Details:\n\n{ex.StackTrace}");
            }
        }

        private void GetGroupEmailCredsIArray()
        {
            try
            {
                var fileHSM_Companion = "1Kw0SjpOwPeXR7tV_a76kYUVoGGMjFzKrfdeGWm6HLJs";
                {
                    string[] Scopes = { SheetsService.Scope.SpreadsheetsReadonly };
                    string ApplicationName = "Google Sheets API .NET Quickstart";
                    UserCredential credential;

                    using (var stream =
                        new FileStream("HSMC-credentials.json", FileMode.Open, FileAccess.Read))
                    {
                        // The file token.json stores the user's access and refresh tokens, and is created
                        // automatically when the authorization flow completes for the first time.
                        //string credPath = "HSM_Companion-token.json";
                        string credPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "HSM_Companion-token.json");
                        credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                            GoogleClientSecrets.Load(stream).Secrets,
                            Scopes,
                            "user",
                            CancellationToken.None,
                            new FileDataStore(credPath, true)).Result;
                        Console.WriteLine("Credential file saved to: " + credPath);
                    }

                    // Create Google Sheets API service.
                    var service = new SheetsService(new BaseClientService.Initializer()
                    {
                        HttpClientInitializer = credential,
                        ApplicationName = ApplicationName,
                    });

                    // Define request parameters.

                    String spreadsheetId = fileHSM_Companion;
                    String range = "A1:M";
                    Google.Apis.Sheets.v4.SpreadsheetsResource.ValuesResource.GetRequest request =
                            service.Spreadsheets.Values.Get(spreadsheetId, range);

                    // Prints the names and other details of Hotels not Submitting the Data for the MAPP Reports:
                    ValueRange response = request.Execute();
                    System.Collections.Generic.IList<System.Collections.Generic.IList<Object>> values = response.Values;
                    if (values != null && values.Count > 0)
                    {
                        var data = response.Values;

                        // Get the headers and get the index of the ldap and the approval status
                        // using the names you use in the headers
                        var headers = values[0];
                        var GroupIndex = headers.IndexOf("Group");
                        var GmAdressIndex = headers.IndexOf("MDO Gmail address");
                        var GmPwdIndex = headers.IndexOf("Gmail Password");


                        //foreach (var row in values)
                        for (int i = 0; i < values.Count; i++)
                        {
                            // Print columns A and E, which correspond to indices 0 and 4.
                            var row = values[i];
                            var c1 = new HSMCInfo();
                            c1.GroupName = row[GroupIndex].ToString().ToUpper();
                            c1.GroupEmail = row[GmAdressIndex].ToString();
                            c1.GroupEmailPwd = row[GmPwdIndex].ToString();
                            ThisHSMCInfo.Add(c1);
                        }
                    }
                    else
                    {
                        MessageBox.Show("No data found.");
                    }
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show($"Unable to Get the Data for PMS ReportNames.\n\nError message: {ex.Message}\n\n" +
                $"Details:\n\n{ex.StackTrace}");
            }
        }

        private string GetPMSReportName(string strPMSName)
        {
            var strReportNames = "Report";
            var strAllPMSNames = "ReportName";

            foreach (PRIInfo PRI in ThisPRIInfo)
            {
                strAllPMSNames = strAllPMSNames + PRI.PMSName + "\\n";
                if (PRI.PMSName.ToUpper() == strPMSName.ToUpper())
                {
                    strReportNames = PRI.Report1;
                    break;
                }
            }
            return strReportNames;
        }

        private string GetGroupEmailCreds(string strGroupEmail)
        {
            var strEmailPwd = "Report";
            var strAllGroupNames = "GroupName";

            foreach (HSMCInfo PRI in ThisHSMCInfo)
            {
                strAllGroupNames = strAllGroupNames + PRI.GroupName + "\\n";
                if (PRI.GroupEmail == strGroupEmail)
                {
                    strEmailPwd = PRI.GroupEmailPwd;
                }
            }
            //MessageBox.Show("All Group Names: \n" + strAllGroupNames);
            return strEmailPwd;
        }
       
        private void BtnGetNonSubmittals_Click(object sender, EventArgs e)
        {
            // Preload the Report Names for reference later
            GetPaceRptIArray();
            // Preload the Email Credentials for reference later
            GetGroupEmailCredsIArray();
            try
            {
                var fileMAPP_HSM = "1EkkGy5pZjFQnnSOdL9z4ELc4oiJ9ovG-edT3UsA9TUA";
                {
                    string[] Scopes = { SheetsService.Scope.SpreadsheetsReadonly };
                    string ApplicationName = "Google Sheets API .NET Quickstart";
                    UserCredential credential;

                    using (var stream =
                        new FileStream("credentials.json", FileMode.Open, FileAccess.Read))
                    {
                        // The file token.json stores the user's access and refresh tokens, and is created
                        // automatically when the authorization flow completes for the first time.
                        //string credPath = "token.json";
                        string credPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "token.json");
                        credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                            GoogleClientSecrets.Load(stream).Secrets,
                            Scopes,
                            "user",
                            CancellationToken.None,
                            new FileDataStore(credPath, true)).Result;
                        Console.WriteLine("Credential file saved to: " + credPath);
                    }

                    // Create Google Sheets API service.
                    var service = new SheetsService(new BaseClientService.Initializer()
                    {
                        HttpClientInitializer = credential,
                        ApplicationName = ApplicationName,
                    });

                    // Define request parameters.

                    String spreadsheetId = fileMAPP_HSM;
                    String range = "A1:Y";
                    Google.Apis.Sheets.v4.SpreadsheetsResource.ValuesResource.GetRequest request =
                            service.Spreadsheets.Values.Get(spreadsheetId, range);

                    // Prints the names and other details of Hotels not Submitting the Data for the MAPP Reports:
                    ValueRange response = request.Execute();
                    System.Collections.Generic.IList<System.Collections.Generic.IList<Object>> values = response.Values;
                    if (values != null && values.Count > 0)
                    {
                        var data = response.Values;

                        // Get the headers and get the index of the ldap and the approval status
                        // using the names you use in the headers
                        var GridViewRowCount = 0;
                        var headers = values[0];
                        var GroupNameIndex = headers.IndexOf("Hotel Management Group");
                        var GroupEmailIndex = headers.IndexOf("@mdo GROUP email address");
                        var PMSIndex = headers.IndexOf("PMS");
                        var SubmissionMethodIndex = headers.IndexOf("Submission Method");
                        var SubmittingPaceIndex = headers.IndexOf("Receiving Pace Reports?");
                        var HotelMDOEmailIndex = headers.IndexOf("MDO Email Address");
                        var strEmailPwd = "";
                        var strReportName1 = "";
                        var strReportName2 = "";
                        var strReportName3 = "";
                        var strReportName = "";


                        //foreach (var row in values)
                        for (int i = 0; i < values.Count; i++)
                        {
                            // Print columns A and E, which correspond to indices 0 and 4.
                            var row = values[i];
                            if (i == 0)
                            {
                                dataGridView1.Columns.Add("Group Name", row[GroupNameIndex].ToString());
                                dataGridView1.Columns.Add("Group Email", row[GroupEmailIndex].ToString());
                                dataGridView1.Columns.Add("Group MDO Email", row[HotelMDOEmailIndex].ToString());
                                dataGridView1.Columns.Add("Group Email PD", "Group Email PD");
                                dataGridView1.Columns.Add("Group PMS", row[PMSIndex].ToString());
                                dataGridView1.Columns.Add("PMS Report Name1", "PMS Report Name1");
                                dataGridView1.Columns.Add("PMS Report Name2", "PMS Report Name2");
                                dataGridView1.Columns.Add("PMS Report Name3", "PMS Report Name3");
                                dataGridView1.Columns.Add("Group PMS", row[SubmissionMethodIndex].ToString());
                                dataGridView1.Columns.Add("Group Submitting?", row[SubmittingPaceIndex].ToString());
                                dataGridView1.Columns.Add("Found Reports", "Found Reports");
                                dataGridView1.Columns.Add("SheetRowIndex", i.ToString());
                                dataGridView1.Columns.Add("EmailSearch", "Email Search Result");
                                //var GridViewRowCount = 1;
                                //AddTestSampleRow();
                            }
                            else
                            {
                                if (row.Count > SubmittingPaceIndex && row[SubmittingPaceIndex].ToString().ToUpper() == "NO")
                                {
                                    strReportName = "";
                                    strReportName1 = "";
                                    strReportName2 = "";
                                    strReportName3 = "";
                                    dataGridView1.Rows.Add();
                                    //GridViewRowCount = GridViewRowCount + 1;
                                    dataGridView1.Rows[GridViewRowCount].Cells[0].Value = row[GroupNameIndex].ToString();
                                    dataGridView1.Rows[GridViewRowCount].Cells[1].Value = row[GroupEmailIndex].ToString();
                                    strEmailPwd = GetGroupEmailCreds(row[GroupEmailIndex].ToString());
                                    dataGridView1.Rows[GridViewRowCount].Cells[2].Value = row[HotelMDOEmailIndex].ToString();
                                    dataGridView1.Rows[GridViewRowCount].Cells[3].Value = strEmailPwd;
                                    dataGridView1.Rows[GridViewRowCount].Cells[4].Value = row[PMSIndex].ToString();
                                    strReportName = GetPMSReportName(row[PMSIndex].ToString());
                                    string[] splittedReports = strReportName.Split('\n');
                                    if (splittedReports[0] != "") { strReportName1 = splittedReports[0]; } else { strReportName1 = ""; }
                                    if (splittedReports.Length > 1) { if (splittedReports[1] != "") { strReportName2 = splittedReports[1]; } else { strReportName2 = ""; } }
                                    if (splittedReports.Length > 2) { if (splittedReports[2] != "") { strReportName3 = splittedReports[2]; } else { strReportName3 = ""; } }


                                    dataGridView1.Rows[GridViewRowCount].Cells[5].Value = strReportName1;
                                    dataGridView1.Rows[GridViewRowCount].Cells[6].Value = strReportName2;
                                    dataGridView1.Rows[GridViewRowCount].Cells[7].Value = strReportName3;
                                    dataGridView1.Rows[GridViewRowCount].Cells[8].Value = row[SubmissionMethodIndex].ToString();
                                    dataGridView1.Rows[GridViewRowCount].Cells[9].Value = row[SubmittingPaceIndex].ToString();
                                    dataGridView1.Rows[GridViewRowCount].Cells[10].Value = "FALSE";
                                    dataGridView1.Rows[GridViewRowCount].Cells["SheetRowIndex"].Value = (i+1).ToString();
                                    GridViewRowCount = GridViewRowCount + 1;
                                }
                            }
                        }
                        textBoxNonSubmittalCount.Text = "Total Count of Hotels NOT Submitting is: " + GridViewRowCount.ToString();
                        foreach (DataGridViewRow row in dataGridView1.Rows)
                        {
                            row.HeaderCell.Value = (row.Index + 1).ToString();
                        }

                        dataGridView1.Sort(dataGridView1.Columns[1], ListSortDirection.Ascending);
                        foreach (DataGridViewRow row in dataGridView1.Rows)
                        {
                            //dataGridView1.Rows[GridViewRowCount].Cells[10].Value = ProcessEmailInbox3(row[1].ToString(), strEmailPwd, row[HotelMDOEmailIndex].ToString(), "Revenue Report", "", "").ToString();
                        }
                    }
                    else
                    {
                        MessageBox.Show("No data found.");
                    }
                }


            }
            catch (Exception ex)
            {
                MessageBox.Show($"Unable to Get the Data for Non Submittals.\n\nError message: {ex.Message}\n\n" +
                $"Details:\n\n{ex.StackTrace}");
            }

        }

       
        
        private void BtnProcessThisEmailAdr_Click(object sender, EventArgs e)
        {
            if (TxtEmail.Text == "" || TxtEmailPwd.Text == "")
            {
                MessageBox.Show("Both Email Address and the Password are required !");
            }
            else
            {
                try
                {
                    var SearchCount = 0;
                    var ProcessedFolderExists = false;


                    using (var client = new ImapClient())
                    {
                        client.Connect("imap.gmail.com", 993, SecureSocketOptions.SslOnConnect);

                        client.Authenticate(TxtEmail.Text, TxtEmailPwd.Text);
                        //var emailFolder = client.GetFolder(ProcessedFolder);
                        var emailFolder = client.GetFolder("Inbox");
                        emailFolder.Open(FolderAccess.ReadOnly);

                        // Find out if there is a Folder Name Processed or PROCESSED or Process
                        var topLevelFolder = client.GetFolder(client.PersonalNamespaces[0]);
                        var topFolders = topLevelFolder.GetSubfolders();
                        var strTargetFolder = "";

                        for (int idx1 = 0; idx1 < topFolders.Count; idx1++)
                        {
                            //MessageBox.Show("This Folder's Name: " + topFolders[idx1].Name);
                            var targetFolder1 =  "Processed";
                            var targetFolder2 = "processed";
                            var targetFolder3 = "process";
                            var targetFolder4 = "PROCESSED";
                            if (topFolders[idx1].Name == targetFolder1 || topFolders[idx1].Name == targetFolder2
                                || topFolders[idx1].Name == targetFolder3 || topFolders[idx1].Name == targetFolder4)
                            {
                                ProcessedFolderExists = true;
                                strTargetFolder = topFolders[idx1].Name;
                            }
                        }
                        

                        var dateFilter = SearchQuery.DeliveredAfter(DateTime.Today.AddDays(-2))
                                            .And(SearchQuery.FromContains(TxtFromEmail.Text)
                                            .Or(SearchQuery.ToContains(TxtFromEmail.Text)))
                                            .And(SearchQuery.BodyContains(TxtContainsText.Text));


                        if (ProcessedFolderExists)
                        {
                            emailFolder.Close(false);

                            emailFolder = client.GetFolder(strTargetFolder);
                            emailFolder.Open(FolderAccess.ReadOnly);

                            var searchResultsPFolder = emailFolder.Search(dateFilter);

                            for (int idx = 0; idx < searchResultsPFolder.Count; idx++)
                            {
                                var message = emailFolder.GetMessage(searchResultsPFolder[idx]);
                                SearchCount = SearchCount + 1;

                                var strAttachments = "";
                                foreach (MimeEntity attachment in message.Attachments)
                                {
                                    var fileName = attachment.ContentDisposition?.FileName ?? attachment.ContentType.Name;
                                    strAttachments = strAttachments + fileName + "\n";
                                }

                                if (SearchCount < 10) MessageBox.Show(message.From.ToString() + "::\n" + message.Subject + "\n" + strAttachments);
                            }
                        }
                        else
                        {
                            var searchResultsPFolder = emailFolder.Search(dateFilter);

                            for (int idx = 0; idx < searchResultsPFolder.Count; idx++)
                            {
                                var message = emailFolder.GetMessage(searchResultsPFolder[idx]);
                                SearchCount = SearchCount + 1;


                                var strAttachments = "";
                                foreach (MimeEntity attachment in message.Attachments)
                                {
                                    var fileName = attachment.ContentDisposition?.FileName ?? attachment.ContentType.Name;
                                    strAttachments = strAttachments + fileName + "\n";
                                }

                                if (SearchCount < 10) MessageBox.Show(message.From.ToString() + "::\n" + message.Subject + "\n" + strAttachments);

                            }
                        }
                        MessageBox.Show("Found a total of " + SearchCount.ToString());
                    }

                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Error in checking the email: ${ex.Message}");
                }
            }


        }
        private void BtnProcessSelectedEmailAdr_Click(object sender, EventArgs e)
        {
            if (dataGridView1.SelectedRows.Count > 0)
            {
                var selectedRow = dataGridView1.SelectedRows[0];
                var selectedGroupId = selectedRow.Cells["Group Email"].Value.ToString();
                var selectedGroupPwd = selectedRow.Cells["Group Email PD"].Value.ToString();
                if (string.IsNullOrEmpty(selectedGroupId) || string.IsNullOrWhiteSpace(selectedGroupPwd))
                {
                    MessageBox.Show("No group email or password on selected row");
                    return;
                }                   
                var selectedGroupEmail = new Dictionary<string, string>();
                selectedGroupEmail.Add(selectedGroupId, selectedGroupPwd);
                ProcessEmails(selectedGroupEmail);
            }
            else
            {
                MessageBox.Show("Please Select the Group Account that you would like to process !");
            }
        }

        private void ProcessAllRowsBtn_Click(object sender, EventArgs e)
        {
            ProcessEmails();
        }

        private void ProcessEmails(Dictionary<string, string> groupEmailSelected = null)
        {
            try
            {

                var ProcessedFolder = "Processed";
                var groupEmails = new Dictionary<string, string>();
                if (groupEmailSelected == null)
                {
                    for (int i = dataGridView1.CurrentCell.RowIndex; i < dataGridView1.Rows.Count; i++)
                    {
                        var row = dataGridView1.Rows[i];
                        var groupEmailInRow = row.Cells["Group Email"]?.Value?.ToString().ToLower();
                        if (string.IsNullOrWhiteSpace(groupEmailInRow))
                        {
                            row.Cells["EmailSearch"].Value = "No group email";
                            continue;
                        }
                        var groupEmailPwd = row.Cells["Group Email PD"]?.Value?.ToString();

                        if (string.IsNullOrWhiteSpace(groupEmailPwd))
                        {
                            row.Cells["EmailSearch"].Value = "No Password";
                        }
                        if (groupEmails.ContainsKey(groupEmailInRow))
                        {
                            continue;
                        }
                        groupEmails.Add(groupEmailInRow, groupEmailPwd);
                    }
                }
                else
                {
                    groupEmails = groupEmailSelected;
                }

                foreach (var gEmailKeyValuePair in groupEmails)
                {
                    var gEmailId = gEmailKeyValuePair.Key;
                    try
                    {                       
                        string GroupEmailId = gEmailId.Substring(0, gEmailId.IndexOf("@")).ToUpper();
                        var processedFolderSettingValue = ConfigurationManager.AppSettings.Get(GroupEmailId);
                        if (!string.IsNullOrWhiteSpace(processedFolderSettingValue))
                        {
                            ProcessedFolder = processedFolderSettingValue;
                        }
                        var pwd = gEmailKeyValuePair.Value;
                        using (var client = new ImapClient())
                        {
                            client.Connect("imap.gmail.com", 993, SecureSocketOptions.SslOnConnect);
                            client.Authenticate(gEmailId, pwd);
                            for (int i = 0; i < dataGridView1.Rows.Count; i++)
                            {
                                var row = dataGridView1.Rows[i];
                                try
                                {
                                    var grupEmailInRow = row.Cells["Group Email"].Value?.ToString();
                                    if (string.IsNullOrWhiteSpace(grupEmailInRow))
                                    {
                                        row.Cells["EmailSearch"].Value = "Processed; Group email is empty";
                                        continue;
                                    }
                                    if (row.Cells["Group Email"].Value.ToString().Equals(gEmailId, StringComparison.OrdinalIgnoreCase) && string.IsNullOrWhiteSpace(row.Cells["EmailSearch"].Value?.ToString()))
                                    {
                                        row.Cells["EmailSearch"].Value = "Processed; ";
                                        var strHotelEmailAddress = row.Cells["Group MDO Email"].Value.ToString();
                                        var strReportName1 = row.Cells["PMS Report Name1"].Value?.ToString();
                                        var strReportName2 = row.Cells["PMS Report Name2"].Value?.ToString();
                                        var strReportName3 = row.Cells["PMS Report Name3"].Value?.ToString();
                                        var reportsList = new List<string>() { strReportName2, strReportName2, strReportName3 };
                                        var strAttachments = strReportName1;                                                                           
                                        CheckEmailResult(row, reportsList, client.Inbox, strHotelEmailAddress);                                        
                                        var emailFolder = client.GetFolder(ProcessedFolder);
                                        CheckEmailResult(row, reportsList, emailFolder, strHotelEmailAddress);                                        
                                    }                                   
                                } 
                                catch(Exception ex)
                                {
                                    row.Cells["EmailSearch"].Value += ex.Message;
                                }
                            }
                        }
                    }
                    catch(Exception ex)
                    {
                        for (int i = 0; i < dataGridView1.Rows.Count; i++)
                        {
                            var row = dataGridView1.Rows[i];
                            if (row.Cells["Group Email"].Value?.ToString().ToUpper() == gEmailId.ToUpper())
                            {
                                row.Cells["EmailSearch"].Value += ex.Message;
                            }
                        }
                    }
                }

                MessageBox.Show("Processing completed. Check EmailSearch column for results");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error in checking the email: ${ex.Message}");
            }


        }

        private static void CheckEmailResult(DataGridViewRow row, List<string> reportsList, IMailFolder emailFolder, string strHotelEmailAddress)
        {
            try
            {
                string folderToProcess = emailFolder.Name;
                emailFolder.Open(FolderAccess.ReadOnly);
                var dateFilter = SearchQuery.DeliveredAfter(DateTime.Today.AddDays(-3))
                                               .And(SearchQuery.FromContains(strHotelEmailAddress));
                var searchResultsPFolder = emailFolder.Search(dateFilter);

                if (searchResultsPFolder.Count > 0)
                {
                    // row.Cells["Found Reports"].Value = "TRUE";  // HINT : Set found to true on just find the email 
                    row.Cells["EmailSearch"].Value += $" | Email Found ({folderToProcess}) from " + strHotelEmailAddress;
                }

                for (int idx = 0; idx < searchResultsPFolder.Count; idx++)
                {
                    var message = emailFolder.GetMessage(searchResultsPFolder[idx]);
                    bool foundReport = false;
                    foreach (MimeEntity attachment in message.Attachments)
                    {
                        var fileName = attachment.ContentDisposition?.FileName ?? attachment.ContentType.Name;
                        if (!string.IsNullOrWhiteSpace(fileName))
                        {
                            if (reportsList.Any(report => report != null && report.Equals(fileName, StringComparison.OrdinalIgnoreCase)))
                            {
                                row.Cells["Found Reports"].Value = "TRUE";
                                row.Cells["EmailSearch"].Value += $" | Report Found({folderToProcess}: {fileName})";
                                foundReport = true;
                                break;
                            }
                            else if (reportsList.Any(report => report != null &&
                                Path.GetFileNameWithoutExtension(report).Equals(Path.GetFileNameWithoutExtension(fileName), StringComparison.OrdinalIgnoreCase)))
                            {
                                row.Cells["EmailSearch"].Value += $" | Report Found But Different Extension ({folderToProcess}: {fileName})";
                            }
                        }
                    }
                    if (foundReport)
                    {
                        break;
                    }
                }
                emailFolder.Close();
            } 
            catch(Exception ex)
            {
                row.Cells["EmailSearch"].Value +=$"; {ex.Message}";
            }
        }

        private void BtnUpdateSheets(object sender, EventArgs e)
        {
            try
            {
                var fileMAPP_HSM = "1EkkGy5pZjFQnnSOdL9z4ELc4oiJ9ovG-edT3UsA9TUA";
                {
                    string[] Scopes = { SheetsService.Scope.Spreadsheets };
                    string ApplicationName = "Google Sheets API .NET Quickstart";
                    UserCredential credential;

                    using (var stream =
                        new FileStream("credentials.json", FileMode.Open, FileAccess.Read))
                    {
                        credential = GoogleAuthUtils.GetUserCredential(stream, Scopes, "tokenwrite.json");                        
                    }

                    // Create Google Sheets API service.
                    var service = new SheetsService(new BaseClientService.Initializer()
                    {
                        HttpClientInitializer = credential,
                        ApplicationName = ApplicationName,
                    });

                    // Define request parameters.

                    String spreadsheetId = fileMAPP_HSM;


                    for (int i = 0; i < dataGridView1.Rows.Count; i++)
                    {
                        var row = dataGridView1.Rows[i];
                        try
                        {
                            var foundReports = row.Cells["Found Reports"].Value?.ToString();
                            if (foundReports == "TRUE")
                            {
                                var sheetRowIndex = row.Cells["SheetRowIndex"].Value?.ToString();
                                if (!string.IsNullOrWhiteSpace(sheetRowIndex))
                                {
                                    String range2 = "V" + sheetRowIndex;
                                    ValueRange valueRange = new ValueRange();
                                    valueRange.MajorDimension = "COLUMNS";

                                    var oblist = new List<object>() { "Yes" };
                                    valueRange.Values = new List<IList<object>> { oblist };

                                    SpreadsheetsResource.ValuesResource.UpdateRequest update = service.Spreadsheets.Values.Update(valueRange, spreadsheetId, range2);
                                    update.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.RAW;
                                    UpdateValuesResponse result2 = update.Execute();
                                }
                            }
                        }
                        catch (Exception ex)
                        {

                        }
                    }
                }
                MessageBox.Show("Sheet update operation done");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private void BtnClearSelection_Click(object sender, EventArgs e)
        {
            dataGridView1.ClearSelection();
        }


        private void btnExit_Click(object sender, EventArgs e)
        {
            //cleanup  
            GC.Collect();
            GC.WaitForPendingFinalizers();
            Application.Exit();
        }

        private void SelectedRowEmailAuditBtn_Click(object sender, EventArgs e)
        {
            try
            {
                if (dataGridView1.SelectedRows.Count > 0)
                {
                    var selectedRow = dataGridView1.SelectedRows[0];
                    var selectedGroupId = selectedRow.Cells["Group Email"].Value.ToString();
                    var selectedGroupPwd = selectedRow.Cells["Group Email PD"].Value.ToString();
                    if (string.IsNullOrEmpty(selectedGroupId) || string.IsNullOrWhiteSpace(selectedGroupPwd))
                    {
                        MessageBox.Show("No group email or password on selected row");
                        return;
                    }
                    var strHotelEmailAddress = selectedRow.Cells["Group MDO Email"].Value.ToString();
                    if (string.IsNullOrEmpty(strHotelEmailAddress))
                    {
                        MessageBox.Show("Group MDO Email is empty");
                        return;
                    }


                    // TEST Record Start

                    //selectedGroupId = "affiliate@mydigitaloffice.ca";
                    //selectedGroupPwd = "dIgItAL0FFICe";
                    //strHotelEmailAddress = "crwfc@mydigitaloffice.ca";

                    // Test Record ends

                    using (var client = new ImapClient())
                    {
                        client.Connect("imap.gmail.com", 993, SecureSocketOptions.SslOnConnect);
                        client.Authenticate(selectedGroupId, selectedGroupPwd);
                        var ProcessedFolder = "Processed";
                        string GroupEmailId = selectedGroupId.Substring(0, selectedGroupId.IndexOf("@")).ToUpper();
                        var processedFolderSettingValue = ConfigurationManager.AppSettings.Get(GroupEmailId);
                        if (!string.IsNullOrWhiteSpace(processedFolderSettingValue))
                        {
                            ProcessedFolder = processedFolderSettingValue;
                        }

                        var emailAuditInfoInbox = GetEmailAuditResult(client.Inbox, strHotelEmailAddress);
                        var emailAuditInfoProcessed = new List<EmailAuditInfo>();
                        try
                        {
                            var processedFolder = client.GetFolder(ProcessedFolder);
                            emailAuditInfoProcessed = GetEmailAuditResult(processedFolder, strHotelEmailAddress);
                        }
                        catch(Exception ex)
                        {

                        }
                        var combinedResults = emailAuditInfoInbox.Concat(emailAuditInfoProcessed).ToList();

                        if (combinedResults.Count() > 0)
                        {
                            UpdateEmailAudit(combinedResults);
                        }
                        else
                        {
                            MessageBox.Show($"No email found in {selectedGroupId} from {strHotelEmailAddress} for today");
                            return;
                        }
                    }
                }
                else
                {
                    MessageBox.Show("No row selected. Please selecte any row.");
                    return;
                }
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private static void UpdateEmailAudit(List<EmailAuditInfo> emailAuditInfos)
        {
            try
            {
                var fileMAPP_HSM = "1wS0LUz46fF64UOkL_YqgdGGNiURgf4kyt6qOCLKImk0";
                {
                    string[] Scopes = { SheetsService.Scope.Spreadsheets };
                    string ApplicationName = "Google Sheets API .NET Quickstart";
                    UserCredential credential;

                    using (var stream =
                        new FileStream("credentials.json", FileMode.Open, FileAccess.Read))
                    {
                      
                        credential = GoogleAuthUtils.GetUserCredential(stream, Scopes, "tokenemailauditwrite.json");
                    }

                    // Create Google Sheets API service.
                    var service = new SheetsService(new BaseClientService.Initializer()
                    {
                        HttpClientInitializer = credential,
                        ApplicationName = ApplicationName,
                    });

                    // Define request parameters.

                    String spreadsheetId = fileMAPP_HSM;
                    // How the input data should be interpreted.
                    string valueInputOption = "Raw";

                    // The new values to apply to the spreadsheet.
                    List<ValueRange> data = new List<ValueRange>();  // TODO: Update placeholder value.

                    for (int i = 0; i < emailAuditInfos.Count; i++)
                    {
                        var row = emailAuditInfos[i];
                        var rowIndex = i + 2;
                        var dateReceivedRange = GetValueRange("A", rowIndex, row.DateReceived);
                        var fromAddressRange = GetValueRange("B", rowIndex, row.FromEmailAddress);
                        var toAddressRange = GetValueRange("C", rowIndex, row.ToEmailAddress);
                        var subjectRange = GetValueRange("J", rowIndex, row.Subject);
                        var bodyRange = GetValueRange("K", rowIndex, row.Body);
                        var attachmentsRange = GetValueRange("O", rowIndex, string.Join(",", row.Attachments));
                        data.Add(dateReceivedRange);
                        data.Add(fromAddressRange);
                        data.Add(toAddressRange);
                        data.Add(subjectRange);
                        data.Add(bodyRange);
                        data.Add(attachmentsRange);
                    }

                    if(data.Count > 0)
                    {
                        BatchUpdateValuesRequest requestBody = new BatchUpdateValuesRequest();
                        requestBody.ValueInputOption = valueInputOption;
                        requestBody.Data = data;
                        SpreadsheetsResource.ValuesResource.BatchUpdateRequest request = service.Spreadsheets.Values.BatchUpdate(requestBody, spreadsheetId);
                        var response = request.Execute();
                    }
                    else
                    {
                        MessageBox.Show("No data to update");
                    }
                }
                MessageBox.Show("Sheet update operation done");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private static ValueRange GetValueRange(string columnName, int rowNumber, string value)
        {
            var dateReceivedRange = columnName + rowNumber;
            ValueRange valueRange = new ValueRange();
            valueRange.MajorDimension = "COLUMNS";
            valueRange.Range = dateReceivedRange;
            var oblist = new List<object>() { value };
            valueRange.Values = new List<IList<object>> { oblist };
            return valueRange;
        }
        private static List<EmailAuditInfo> GetEmailAuditResult(IMailFolder emailFolder, string strHotelEmailAddress)
        {
            var emailAuditList = new List<EmailAuditInfo>();
            try
            {
               
                string folderToProcess = emailFolder.Name;
                emailFolder.Open(FolderAccess.ReadOnly);
                var dateFilter = SearchQuery.DeliveredAfter(DateTime.Today)
                                               .And(SearchQuery.FromContains(strHotelEmailAddress));
                var searchResultsPFolder = emailFolder.Search(dateFilter);          

                for (int idx = 0; idx < searchResultsPFolder.Count; idx++)
                {
                    var emailAuditInfo = new EmailAuditInfo();
                    emailAuditInfo.Attachments = new List<string>();
                    var message = emailFolder.GetMessage(searchResultsPFolder[idx]);
                    emailAuditInfo.DateReceived = message.Date.ToString();
                    emailAuditInfo.FromEmailAddress = message.From.ToString();
                    emailAuditInfo.ToEmailAddress = message.To.ToString();
                    emailAuditInfo.Subject = message.Subject.ToString();
                    emailAuditInfo.Body = message.Body.ToString();
                    foreach (MimeEntity attachment in message.Attachments)
                    {
                        var fileName = attachment.ContentDisposition?.FileName ?? attachment.ContentType.Name;
                        if (!string.IsNullOrWhiteSpace(fileName))
                        {
                            emailAuditInfo.Attachments.Add(fileName);
                        }
                    }
                    
                }
                emailFolder.Close();
            }
            catch (Exception ex)
            {
                
            }
            return emailAuditList;
        }
               

        private async void UpdateAccountButton_Click(object sender, EventArgs e)
        {
            try
            {
                var progress = new Progress<string>(s => progressLabel.Text = s);
                var result = await Task.Run(() =>
                {
                    string strGConnectionString = ConfigurationManager.AppSettings.Get("SourceDBConnectionString");
                    IProgress<string> prg = progress;
                    HotelDataLookupUtils hotelDataLookupUtils = new HotelDataLookupUtils();
                    prg.Report("Reading Org Id Lookup");
                    var groupIdDict = hotelDataLookupUtils.GetOrgIdLookUp();
                    prg.Report("Reading Hotel Id Lookup");
                    var hotelIdDict = hotelDataLookupUtils.GetHotelIdLookUp();

                    SheetUtils sheetUtils = new SheetUtils();
                    prg.Report("Reading HSM Sheet data");
                    var sheetDataList = sheetUtils.LoadHSMSheetData();
                    var customerList = new List<Customer>();
                    foreach (var sD in sheetDataList)
                    {
                        var customer = new Customer();
                        customer.GroupId = groupIdDict.ContainsKey(sD.HotelManagementGroup) ? groupIdDict[sD.HotelManagementGroup] : 0;
                        customer.HotelId = hotelIdDict.ContainsKey(sD.HotelCode) ? hotelIdDict[sD.HotelCode] : 0;
                        customer.HotelName = sD.HotelName;
                        customer.GroupName = sD.HotelManagementGroup;
                        customer.HotelCode = sD.HotelCode;
                        customer.MDOGroupEmail = sD.MdoGroupEmailAddres;
                        customer.MDOHotelEmail = sD.MdoEmailAddress;
                        customer.MDOSolutions = sD.MdoSolutions;
                        customer.PMSType = sD.PMS;
                        customer.PaceReport = sD.PaceReport;
                        customer.RevenueReport = sD.RevenueReport;
                        if (!string.IsNullOrWhiteSpace(sD.PaceReport))
                        {
                            var attachments = sD.PaceReport.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.None);
                            var attachmentTrimmed = attachments.Where(att => !string.IsNullOrWhiteSpace(att)).ToList();
                           
                            var att1 = attachmentTrimmed.ElementAtOrDefault(0);
                            var att2 = attachmentTrimmed.ElementAtOrDefault(1);
                            var att3 = attachmentTrimmed.ElementAtOrDefault(2);
                            var att4 = attachmentTrimmed.ElementAtOrDefault(3);
                            var ext1 = string.Empty;
                            var ext2 = string.Empty;
                            var ext3 = string.Empty;
                            var ext4 = string.Empty;
                            if (att1 != null)
                            {
                                var ext = att1.Contains(".") ? att1.Substring(att1.LastIndexOf(".")) : string.Empty;
                                ext1 = string.Concat(ext.Take(4)).Trim();
                            }
                            if (att2 != null)
                            {
                                var ext = att2.Contains(".") ? att2.Substring(att2.LastIndexOf(".")) : string.Empty;                               
                                ext2 = string.Concat(ext.Take(4)).Trim();
                            }
                            if (att3 != null)
                            {
                                var ext = att3.Contains(".") ? att3.Substring(att3.LastIndexOf(".")) : string.Empty;                               
                                ext3 = string.Concat(ext.Take(4)).Trim();
                            }
                            if (att4 != null)
                            {
                                var ext = att4.Contains(".") ? att4.Substring(att4.LastIndexOf(".")) : string.Empty;                                
                                ext4 = string.Concat(ext.Take(4)).Trim();
                            }
                            var extList = (new List<string>() { ext1, ext2, ext3, ext4 }).Where(ex => !string.IsNullOrWhiteSpace(ex)).ToList();
                            var uniqList = extList.Distinct(StringComparer.OrdinalIgnoreCase).ToList();
                            var uext1 = uniqList.ElementAtOrDefault(0);
                            var uext2 = uniqList.ElementAtOrDefault(1);
                            var uext3 = uniqList.ElementAtOrDefault(2);
                            var uext4 = uniqList.ElementAtOrDefault(3);
                            if(!string.IsNullOrWhiteSpace(uext1))
                            {
                                customer.AttachmentExtension1 = uext1;
                            }
                            if (!string.IsNullOrWhiteSpace(uext2))
                            {
                                customer.AttachmentExtension2 = uext2;
                            }
                            if(!string.IsNullOrWhiteSpace(uext3))
                            {
                                customer.AttachmentExtension3 = uext3;
                            }
                            if (!string.IsNullOrWhiteSpace(uext4))
                            {
                                customer.AttachmentExtension4 = uext4;
                            }
                            customer.NumOfAttachments = uniqList.Count;
                        }

                        customerList.Add(customer);
                    }
                        CustomerCreationHelper customerCreationHelper = new CustomerCreationHelper(strGConnectionString);
                        var returnMessage = customerCreationHelper.CreateCustomers(customerList, null, prg);

                        return returnMessage;
                    
                });
                MessageBox.Show(result);
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message);
                
            }
            progressLabel.Text = "";
        }       
    }

    internal class EmailAuditInfo
    {
        public string DateReceived { get; set; }
        public string FromEmailAddress { get; set; }
        public string ToEmailAddress { get; set; }
        public string Subject { get; set; }
        public string Body { get; set; }
        public List<string> Attachments { get; set; }
    }

    internal class PRIInfo
    {
        public string PMSName { get; set; }
        public string Report1 { get; set; }
        public string Report2 { get; set; }
        public string Report3 { get; set; }
        public string Report4 { get; set; }
        public string Report5 { get; set; }
    }

    internal class HSMCInfo
    {
        public string GroupName { get; set; }
        public string GroupEmail { get; set; }
        public string GroupEmailPwd { get; set; }
    }
}

