using System.Windows.Forms;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace ppData
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        /// 
        static string[] Scopes = { SheetsService.Scope.SpreadsheetsReadonly };
        static string ApplicationName = "Google Sheets API .NET Quickstart";
        static SheetsService service;

        [STAThread]
        static void Main()
        {
            UserCredential credential = doAccess();

            // Create Google Sheets API service.
            service = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });

            doOutput();
            Console.Read();

            //Application.EnableVisualStyles();
            //Application.SetCompatibleTextRenderingDefault(false);
            //Application.Run(new Form1());

        }

        private static void doOutput()
        {
            // Define request parameters.
            String spreadsheetId = "1Y9YeG-P5pfwqDfozgeVYiJJyxBOK2G7F15dedowvrCU";
            String range = "Subsheets!A1:AP50";
            SpreadsheetsResource.ValuesResource.GetRequest request =
                    service.Spreadsheets.Values.Get(spreadsheetId, range);
            request.ValueRenderOption = SpreadsheetsResource.ValuesResource.GetRequest.ValueRenderOptionEnum.FORMATTEDVALUE;
            request.MajorDimension = SpreadsheetsResource.ValuesResource.GetRequest.MajorDimensionEnum.COLUMNS;
            string error = "none";

            try
            {
                ValueRange response = request.Execute();
                IList<IList<Object>> sites = response.Values;
                foreach (var site in sites)
                {
                    Queue<string> data = new Queue<string>(site.Cast<String>());

                    var siteName = data.Dequeue();
                    if (siteName == "Site")
                        continue;
                    var siteURL = getKey(data.Dequeue());
                    if (String.IsNullOrEmpty(siteURL))
                        Console.WriteLine("No datasheet for {0}", siteName);
                    else
                    {
                        pullSiteData(siteName, siteURL, data);
                    }
                }
            }
            catch (Exception e)
            {
                error = e.ToString();
                Console.WriteLine("No global data ERROR={0}", error);
            };
        }

        private static string getKey(object urlobj)
        {
            string url = urlobj.ToString();
            if (String.IsNullOrEmpty(url))
                return "";

            if ('/'.Equals(url[url.Length - 1]))
                url = url.Substring(0, url.Length - 1);

            int i = url.LastIndexOf('/');
            return url.Substring(i + 1);
        }

        private static void pullSiteData(string siteName, String sheetKey, Queue<string> sheets)
        {
            SpreadsheetsResource.GetRequest request = service.Spreadsheets.Get(sheetKey);
            request.IncludeGridData = true;
            try
            {
                Google.Apis.Sheets.v4.Data.Spreadsheet response = request.Execute();

                foreach(var sheetName in sheets)
                {
                    try
                    {
                        Sheet sheet = response.Sheets.FirstOrDefault<Sheet>(s => s.Properties.Title == sheetName);
                        if (sheet == null)
                        {
                            Console.WriteLine("Couldn't find sheet {0} in spreadsheet {1} {2}", sheetName, siteName, sheetKey);
                            continue;
                        }

                        foreach (var row in sheet.Data.First<GridData>().RowData)
                        {
                            Console.Write("{0}\t{1}\t", siteName, sheetName);
                            foreach (var cell in row.Values)
                            {
                                Console.Write("{0}\t", (cell.FormattedValue ?? "").Trim());
                            }
                            Console.WriteLine();
                        }
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine("Error processing sheet data for {0}:{1}:{2}={3}", siteName, sheetKey, sheetName, e.Message);
                    };
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("Couldn't get sheet data for {0} on {1} = {2}", siteName, sheetKey, e.GetType());
            };
        }

        private static UserCredential doAccess()
        {
            // Check/Referesh Access Token
            UserCredential credential;

            using (var stream =
                new FileStream("client_id.json", FileMode.Open, FileAccess.Read))
            {
                string credPath = System.Environment.GetFolderPath(
                    System.Environment.SpecialFolder.Personal);
                credPath = Path.Combine(credPath, ".credentials/sheets.googleapis.com-dotnet-quickstart.json");

                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    Scopes,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;
                Console.WriteLine("Credential file saved to: " + credPath);
            }

            return credential;
        }
    }
}
