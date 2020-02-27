using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Util.Store;
using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using System.Threading.Tasks;

namespace GSP.Core.Controller
{
    public class GoogleSheetsAPI
    {        
        private string[] Scopes = { SheetsService.Scope.SpreadsheetsReadonly };
        private static string ApplicationName = "Google Sheets Parser";
        private UserCredential Credential { get; }
        private SheetsService SheetsService { get; }

        public GoogleSheetsAPI(string jsonFile)
        {
            using (var stream = new FileStream(jsonFile, FileMode.Open, FileAccess.Read))
            {
                string credPath = Environment.GetFolderPath(Environment.SpecialFolder.Personal);
                credPath = Path.Combine(credPath, ".credentials/sheets.googleapis.com-dotnet-quickstart.json");

                Credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    Scopes,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;
            }

            SheetsService = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = Credential,
                ApplicationName = ApplicationName,
            });
        }

        public IList<IList<Object>> GetValue(string spreadsheetId, string range)
        {
            try
            {
                var request = SheetsService.Spreadsheets.Values.Get(spreadsheetId, range);
                return GetResponse(request).Result;
            }
            catch (Exception)
            {
                return default;
            }  
            
        }

        private static async Task<IList<IList<Object>>> GetResponse(SpreadsheetsResource.ValuesResource.GetRequest request)
        {
            ValueRange response = await request.ExecuteAsync();
            return response.Values;
        }
    }
}
