﻿using Google.Apis.Auth.OAuth2;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using Google.Apis.HangoutsChat.v1;
using Google.Apis.HangoutsChat.v1.Data;
using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using System.Data;
using System.Net;

namespace GoogleApiTestForms
{

    public class GoogleApi
    {
        string[] Scopes = { SheetsService.Scope.Spreadsheets};
        private string ApplicationName;
        private string spreadsheetId;
        SheetsService service;
        private string path;
        UserCredential credential;
        HangoutsChatService hangoutsChatService;

        public GoogleApi(string ApplicationName, string spreadsheetId, string path)
        {
            this.spreadsheetId = spreadsheetId;
            this.ApplicationName = ApplicationName;
            this.path = path;
        }
        public void Authentication()
        {
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

            }

            // Create Google Sheets API service.
            service = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });

            hangoutsChatService = new HangoutsChatService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });
        }
        public DataTable ReadEntries(string start, string end, string sheet)
        {

            var range = $"{sheet}{start}:{end}";
            var request = service.Spreadsheets.Values.Get(spreadsheetId, range);
            var dataTable = new DataTable();
            var response = request.Execute();
            var values = response.Values;
            dataTable.Columns.Add(values[0][0].ToString());
            dataTable.Columns.Add(values[0][1].ToString());
            dataTable.Columns.Add(values[0][2].ToString());
            dataTable.Columns.Add(values[0][3].ToString());
            dataTable.Columns.Add(values[0][4].ToString());
            if (values.Count>0)
            {
                for (int i = 1; i < values.Count; i++)
                {
                    if (values[i].Count>4)
                    {
                        dataTable.Rows.Add(values[i][0], values[i][1], values[i][2], values[i][3], values[i][4]);
                    }
                }
            }
            
            return dataTable;
        }
        public void Testing()
        {
            var resource = new SpreadsheetsResource(service);
            
        }
        public void UpdateSheet(List<object> s1, string sheet, string start,string end)
        {
            var range = $"{sheet}{start}:{end}";
            var valueRange = new ValueRange();
            
            valueRange.Values = new List<IList<object>> { s1 };
            
            
            
            var appendRequest = service.Spreadsheets.Values.Append(valueRange, spreadsheetId, range);
            appendRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.USERENTERED;
            var appendReponse = appendRequest.Execute();
        }

        
      



    }
}

