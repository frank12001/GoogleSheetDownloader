using System;
using System.IO;
using System.Threading;
using System.Collections.Generic;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Util.Store;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Sheets.v4.Data;
using Newtonsoft.Json;

namespace GoogleSheetDownload
{
    public class SheetDownloader
    {
        private const string SheetRange = "A1:Z";
        /// <summary>
        /// 將指定的 Google Sheet 下載，並輸出成 JSON
        /// </summary>
        /// <param name="sheetId"> Google 對每個 Sheet 定義的 Id 。 ex : https://docs.google.com/spreadsheets/d/11T_AEhwHLoOI-Shb7KYFIatsRxsguX_4iKTr1t79CdQ/edit#gid=1209401376  Id = 11T_AEhwHLoOI-Shb7KYFIatsRxsguX_4iKTr1t79CdQ </param>
        /// <param name="sheetName"> 指定的 Sheet 名稱 </param>
        /// <param name="exportPath"> 輸出 Json 的絕對路徑 。 ex : C:\Data </param>
        public void ExportSheet(string sheetId, string sheetName,string exportPath)
        {
            File.WriteAllText(exportPath, GetTargetSheet(sheetName, sheetId));
        }

       
        private string GetTargetSheet(string tableName, string spreadsheetId)
        {
            return ConvertToJSONString(_GetTargetSheet(tableName, spreadsheetId));
        }

        /// <summary>
        /// 將指定的表格取下來 。取完後會是 Google Api 定義的格式
        /// </summary>
        /// <param name="tableName"></param>
        /// <param name="sheetId"></param>
        /// <returns></returns>
        private IList<IList<Object>> _GetTargetSheet(string sheetId, string tableName)
        {
            String range = string.Format("{0}!{1}", tableName, SheetRange);
            SpreadsheetsResource.ValuesResource.GetRequest request = UserCredential().Spreadsheets.Values.Get(sheetId, range);
            request.ValueRenderOption = SpreadsheetsResource.ValuesResource.GetRequest.ValueRenderOptionEnum.FORMULA;

            ValueRange response = request.Execute();
            IList<IList<Object>> values = response.Values;
            return values;
        }

        /// <summary>
        /// 將載下來的表格，轉成 JSON
        /// </summary>
        /// <param name="values"></param>
        /// <returns></returns>
        private string ConvertToJSONString(IList<IList<Object>> values)
        {
            IList<Object> property = values[0];
            List<Dictionary<Object, Object>> result = new List<Dictionary<Object, Object>>();
            for (int i = 1; i < values.Count; i++)
            {
                Dictionary<Object, Object> raw = new Dictionary<object, object>();
                for (int j = 0; j < values[i].Count; j++)
                {
                    if (values[i][j] != null)
                        raw.Add(property[j], values[i][j]);
                }
                result.Add(raw);
            }
            return JsonConvert.SerializeObject(result);
        }


        /// <summary>
        /// 讀取認證
        /// </summary>
        /// <returns></returns>
        private static SheetsService UserCredential()
        {
            UserCredential credential;

            using (var stream = new FileStream("client_secret.json", FileMode.Open, FileAccess.Read))
            {
                string credPath = System.Environment.GetFolderPath(System.Environment.SpecialFolder.Personal);
                credPath = Path.Combine(credPath, ".credentials/sheets.googleapis.com-dotnet-quickstart.json");
                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    new string[] { SheetsService.Scope.SpreadsheetsReadonly },
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;
                Console.WriteLine("Credential file saved to: " + credPath);
            }

            // Create Google Sheets API service.
            var service = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = "DownLoad Sheet",
            });
            return service;
        }
    }
}
