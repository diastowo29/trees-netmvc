using System;
using System.IO;
using System.Web;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Net.Http;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using fif_api.Models;

/*EPPLUS - READXLSX*/
using OfficeOpenXml;
/*JSONCONVERT SERIALIZE*/
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
/*RESTREQUEST - CALL API*/
using RestSharp;
using RestSharp.Authenticators;

namespace fif_api.Controllers
{
    public class FifController : Controller
    {
        string zendeskDomain = "https://developmenttesting.zendesk.com";
        string zendeskUsername = "rahdityoluhung89@gmail.com";

        // static string zendeskDomain = "https://treesdemo1.zendesk.com";
        // static string zendeskUsername = "eldien.hasmanto@treessolutions.com";

        string zendeskPassword = "W3lcome123";

        public IEnumerable<string> Index()
        {
            // string users = callApi(zendeskDomain + "/api/v2/users.json");
            // Console.WriteLine(users);
            return new string[] { "fifController1", "fifController2" };
        }

        public IActionResult About()
        {
            ViewData["Message"] = "Your application description page.";

            return View();
        }

        public IActionResult Contact()
        {
            ViewData["Message"] = "Your contact page.";

            return View();
        }

        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }

        [HttpPost]
        public IEnumerable<string> DoMapping([FromBody] TeamViewModel teamViewModel){

            // Console.WriteLine(teamViewModel.ticket_id);
            // Console.WriteLine(teamViewModel.dealer_id);
            // Console.WriteLine(teamViewModel.ticket_level);
            

            string fileName = @"/home/diastowo/Documents/DOT NET/excel/Dealer Hirarki v3 (Include Region).xlsx";
            List<Dictionary<string, string>> excelContent = readXlsx(fileName);
            for (int i=0; i<excelContent.Count; i++) {
                if (excelContent[i]["Dealer ID"] == teamViewModel.dealer_id) {
                    Console.WriteLine(excelContent[i]["Dealer Name"]);
                }
            }

            return new string[] { "domapping" };
        }

        public IEnumerable<string> ReadExcel() {
            

            return new string[] { "fif1", "fif2" };
        }
        
        public List<Dictionary<string, string>> readXlsx (string filePath) {
            Dictionary<string, string> mappingList = new Dictionary<string, string>();
            Console.WriteLine("===== DO XLSX =====");
            List<String> keys = new List<String>();
            List<Dictionary<string, string>> mappingArray = new List<Dictionary<string, string>>();
            int skipIndex = 0;

            Console.WriteLine(filePath);
            var package = new ExcelPackage(new FileInfo(filePath));
            ExcelWorksheet sheet = package.Workbook.Worksheets[0];
            var rowCount = sheet.Dimension.End.Row;
            var colCount = sheet.Dimension.End.Column;
            for (int i=1; i<=rowCount; i++) {
                mappingList = new Dictionary<string, string>();
                // Console.WriteLine("===== NEW ROW =====");
                int myColCounter = 0;

                for (int j=1; j<=colCount; j++) {
                    if (i == 1) {
                        // bool keyExist = false;
                        for (int k=0; k<keys.Count; k++) {
                            if (keys[k] == sheet.Cells[i,j].Value.ToString()) {
                                skipIndex = j;
                            }
                        }
                        keys.Add(sheet.Cells[i,j].Value.ToString());
                    } else {
                        if (j != skipIndex) {
                            string values = "";
                            if (sheet.Cells[i,j].Value == null) {
                                values = "";
                            } else {
                                values = sheet.Cells[i,j].Value.ToString();
                            }
                            mappingList.Add(keys[j-1], values);
                            myColCounter++;
                        }
                    }
                }
                if (i != 1) {
                    mappingArray.Add(mappingList);
                }
            }

            // string jsonString = JsonConvert.SerializeObject(mappingArray);
            // Console.WriteLine(jsonString);
            return mappingArray;
        }

        public string callApi (String urls) {
            Console.WriteLine("CALL GET: " + urls);

            var client = new RestClient(urls);
            client.Authenticator = new HttpBasicAuthenticator(zendeskUsername, zendeskPassword);

            var request = new RestRequest("", Method.GET);
            // easily add HTTP Headers
            // request.AddHeader("Authorization", "Basic " + zendeskToken);

            IRestResponse response = client.Execute(request);
            var content = response.Content;
            return content;
        }
    }



    public class TeamViewModel {
        public string ticket_id { get; set; }
        public string dealer_id { get; set; }
        public string ticket_level { get; set; }
    }
}
