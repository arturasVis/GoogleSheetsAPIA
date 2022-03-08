using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Util.Store;
using LumenWorks.Framework.IO.Csv;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Threading;

namespace ConsoleApp1
{
    internal class Program
    {
        static string[] Scopes = { SheetsService.Scope.Spreadsheets };
        static string ApplicationName = "Desktop client 1";
        static string spreadsheetId = "1fEX7dsW_gGZFugrq9IvH18IuGLP8c02rJRuzIFtYtWI";
        static SheetsService service;
        static string path = "C:\\Users\\PC\\Dropbox\\PC\\Desktop\\csvFile\\";
        //static string sheet = "openorder";
        static void Main(string[] args)
        {
            Authentication();
            DataTable table = ReadSheet("openorder.csv");
            string[] col = { "A","B","C","D","E","F"};
            for (int i = 0; i < table.Rows.Count; i++)
            {
                for (int j=0;j< table.Columns.Count;j++)
                {
                    int k = i + 1;
                    string range = col[j] + k + ":" + col[j] + k;
                    UpdateSheet(table.Rows[i][j].ToString(), "openorder",range);
                }
                //UpdateSheet(table.Rows[i][0].ToString(), table.Rows[i][1].ToString(), table.Rows[i][2].ToString(), "openorder");
            }
           
        }
        static void Authentication()
        {
            UserCredential credential;

            using (var stream =
                new FileStream("cred.json", FileMode.Open, FileAccess.Read))
            {
                // The file token.json stores the user's access and refresh tokens, and is created
                // automatically when the authorization flow completes for the first time.

                string credPath = "token.json";
                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    Scopes,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;
                Console.WriteLine("Credential file saved to: " + credPath);
            }

            // Create Google Sheets API service.
            service = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });
            Console.WriteLine();
        }
        static void UpdateSheet(string s1,string sheet,string col)
        {

            var range = $"{sheet}!{col}";
            var valueRange = new ValueRange();
            var oblist = new List<object>() {s1};
            valueRange.Values = new List<IList<object>> { oblist };

            var appendRequest = service.Spreadsheets.Values.Append(valueRange, spreadsheetId, range);
            appendRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.USERENTERED;
            var appendReponse = appendRequest.Execute();


        }
        static DataTable ReadSheet(string name)
        {
            var csvTable = new DataTable();
            using (var csvReader = new CsvReader(new StreamReader(System.IO.File.OpenRead(path + name)), true))
            {
                csvTable.Load(csvReader);
            }
            return csvTable;
        }
    }
}
