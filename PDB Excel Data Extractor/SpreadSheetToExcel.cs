using System;
using System.Collections.Generic;
using System.Data;
using System.IO;

using System.Text;
using System.Text.RegularExpressions;
using System.Threading;

using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Util.Store;


namespace PDB_Excel_Data_Extractor
{
    public class SpreadSheetToExcel : Common
    {
        static string[] Scopes = {SheetsService.Scope.SpreadsheetsReadonly};
        static string ApplicationName = "Google Sheets API .NET Quickstart";

        public void Transformer(string sourcePath, DataTable tableSctructure)
        {
            ExcelReader excel = new ExcelReader();
            var service = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = SetCredentials(),
                ApplicationName = ApplicationName,
            });
            string[] files = Directory.GetFiles(sourcePath);
            string text = "";
            for (int i = 0; i < files.Length; i++)
            {
                if (files[i].Contains("gsheet"))
                {
                    text = File.ReadAllText(files[i],
                        Encoding.UTF8);
                    DataTable table =
                        SpreadSheetToExcelDataTable(text, service, tableSctructure);

                    excel.ExportToExcel(table, files[i].Substring(0, files[i].Length - 7), "Sheet1", numberOfLastRow: 3,
                        startingCellIndex: 2);
                }
            }
           
        }

       
        private  DataTable SpreadSheetToExcelDataTable(string text, SheetsService service, DataTable table)
        {
            
            Regex regex = new Regex(@"=([0-9A-Za-z_-]+)");
            String spreadsheetId = "";
            Match match = regex.Match(text);
            if (match.Success)
            {
                spreadsheetId = match.Value.Substring(1);
            }
            String range = "Sheet1!B3:H";
            SpreadsheetsResource.ValuesResource.GetRequest request =
                service.Spreadsheets.Values.Get(spreadsheetId, range);


            ValueRange response = request.Execute();
            IList<IList<Object>> values = response.Values;
            if (values != null && values.Count > 0)
            {
                foreach (var row in values)
                {
                    switch (row.Count)
                    {
                        case 4:
                            table.Rows.Add(row[0], row[1], row[2], row[3], "", "");
                            break;

                        case 6:
                            table.Rows.Add(row[0], row[1], row[2], row[3], row[4], row[5]);
                            break;
                        case 7:
                            table.Rows.Add(row[0], row[1], row[2], row[3], row[4], row[5], row[6]);
                            break;
                        default:
                            table.Rows.Add("", "", "", "", "", "");
                            break;
                    }
                }
            }
            return table;
        }
        private  UserCredential SetCredentials()
        {
            UserCredential credential;
            using (var stream =
                new FileStream("client_secret.json", FileMode.Open, FileAccess.Read))
            {
                string credPath = Environment.GetFolderPath(
                    Environment.SpecialFolder.Personal);
                string secondPath =
                    AssemblyDirectory;
                credPath = Path.Combine(secondPath, ".credentials/sheets.googleapis.com-dotnet-quickstart.json");

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
