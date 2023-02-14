using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.Net.Http.Headers;
using Newtonsoft.Json;
using System.Net;
using IntakeBase=Xom.Gci.Addin.LvMake.Intake.Base;
using IIntakeBase = Xom.Gci.Addin.LvMake.IIntake;

namespace ConsoleApp1
{
    public class Program  :IntakeBase.IntakeLIMSDataFetch(null)
    {

        public static Excel.Workbook workbook ;

        
           public  void Main(string[] args)
        {
            //NewExcelFile();
            IntakeFormExcelFile();
            Login();
        }

        public static void NewExcelFile()
        {
            string fileName = "TestExcel.xlsx";
            string sourcePath = @"C:\Users\naveene\Downloads";
            string targetPath = @"C:\Users\naveene\Downloads\New folder";

            string sourceFile = System.IO.Path.Combine(sourcePath, fileName);
            string destFile = System.IO.Path.Combine(targetPath, fileName);
            Excel.Application excelApp = null;
            Excel.Range range = null;
            Excel.Workbook workbook = null;
            Excel.Worksheet worksheet = null;
            System.IO.Directory.CreateDirectory(targetPath);

            if (System.IO.Directory.Exists(sourcePath))
            {
                DirectoryInfo d = new DirectoryInfo(sourcePath);

                // Copy the files and overwrite destination files if they already exist.
                foreach (var file in d.GetFiles("*.xlsx"))
                {
                    if (file.Name == fileName)
                    {
                        // Use static Path methods to extract only the file name from the path.
                        fileName = System.IO.Path.GetFileName(file.Name);
                        destFile = System.IO.Path.Combine(targetPath, fileName);
                        if (File.Exists(destFile))
                        {
                            File.Delete(destFile);
                        }
                        System.IO.File.Copy(sourceFile, destFile, true);
                        excelApp = new Excel.Application();
                        Console.WriteLine("Loading the Excel File ");

                        workbook = excelApp.Workbooks.Open(destFile);
                        worksheet = workbook.Worksheets[1];
                        worksheet.Cells[1, 2].Value2 = "Creating a cell value.!!!";
                        workbook.Save();
                        workbook.Close();
                        Console.WriteLine("Value has return Successfully!!!!! ");
                    }

                }
            }
            else
            {
                Console.WriteLine("Source path does not exist!");
            }
        }

        public static void IntakeFormExcelFile()
        {
            string fileName = "R2301-009984-001_Intake.xlsx";
            string sourcePath = @"C:\Users\naveene\Downloads";
            string targetPath = @"C:\Users\naveene\Downloads\New folder";

            string sourceFile = System.IO.Path.Combine(sourcePath, fileName);
            string destFile = System.IO.Path.Combine(targetPath, fileName);
            Excel.Application excelApp = null;
            Excel.Range range = null;
            
            Excel.Worksheet worksheet = null;
            System.IO.Directory.CreateDirectory(targetPath);

            if (System.IO.Directory.Exists(sourcePath))
            {
                DirectoryInfo d = new DirectoryInfo(sourcePath);

                // Copy the files and overwrite destination files if they already exist.
                foreach (var file in d.GetFiles("*.xlsx"))
                {
                    if (file.Name == fileName)
                    {
                        // Use static Path methods to extract only the file name from the path.
                        fileName = System.IO.Path.GetFileName(file.Name);
                        destFile = System.IO.Path.Combine(targetPath, fileName);
                        if (File.Exists(destFile))
                        {
                            File.Delete(destFile);
                        }
                        System.IO.File.Copy(sourceFile, destFile, true);
                        excelApp = new Excel.Application();
                        Console.WriteLine("Loading the Excel File ");
                      
                        workbook = excelApp.Workbooks.Open(destFile);
                        IntakeBase.IntakeValidate validate = new IntakeBase.IntakeValidate();
                        IntakeBase.OpenWorkbook openWorkbook = new IntakeBase.OpenWorkbook();
                        openWorkbook.LIMSDataFetch.FetchAndLoadSheetData();
                                    worksheet = workbook.Worksheets[2];
                        worksheet.Cells[5, 6].Value2 = "BTEC_10035";
                        worksheet.Cells[4, 6].Value2 = "BTEC";
                
                        workbook.Save();
                        workbook.Close();
                        Console.WriteLine("Value has return Successfully!!!!! ");
                    }

                }
            }
            else
            {
                Console.WriteLine("Source path does not exist!");
            }
        }

        public static void Login()
        {
            // To adhere to security guidelines - LIMS portal upgraded to TSL level 1.2 & 1.3 
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
            HttpClient httpClient;
            LvConnection connection;
            string lvConnectionID = string.Empty;
            string jSessionID = string.Empty;
            string userName = "naveene";
            string password = "exxonmobilinfosys";

            string lvurl = "https://hoeapp910.na.xom.com/labvantage/rest";
            string absoluteurl = "https://hoeapp910.na.xom.com/labvantage";
            string databaseid = "LabVantageNADEV";
            connection = new LvConnection();
            connection.DatabaseId = databaseid;
            connection.Username = userName;
            connection.Password = password;
            HttpClientHandler httpHandler = new HttpClientHandler();
            httpClient = new HttpClient(httpHandler);
            var connObj = new StringContent(JsonConvert.SerializeObject(connection));
            httpClient.DefaultRequestHeaders.Add("User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/57.0.2987.133 Safari/537.36");
            var response = httpClient.PostAsync($"{lvurl}/connections", connObj).Result;
            bool IsUnauthorized = response.StatusCode != HttpStatusCode.Created ? true : false;

            if (!IsUnauthorized)
            {
                foreach (var headerItem in response.Headers)
                {
                    IEnumerable<string> values;
                    values = response.Headers.GetValues("Set-Cookie");
                    foreach (var valueItem in values)
                    {
                        if (valueItem.ToLower().Contains("connectionid="))
                        {
                            lvConnectionID = WebUtility.UrlDecode(valueItem.Split(';').ToList().FirstOrDefault(id => id.Contains("connectionid=")).Replace("connectionid=", ""));
                            break;
                        }
                    }
                    foreach (var valueItem in values)
                    {
                        if (valueItem.ToLower().Contains("JSESSIONID=".ToLower()))
                        {
                            jSessionID = valueItem.Split(';').ToList().FirstOrDefault(id => id.Contains("JSESSIONID=")).ToString();
                            break;
                        }
                    }
                    if (!string.IsNullOrEmpty(lvConnectionID))
                    {
                        break;
                    }
                }
            }

            if (!string.IsNullOrEmpty(lvConnectionID))
            {
                httpClient.DefaultRequestHeaders.Remove("Authorization");
                httpClient.DefaultRequestHeaders.Add("Authorization", "Token " + lvConnectionID);
                HttpResponseMessage reqResponse = httpClient.GetAsync($"{absoluteurl}/files/IntakeForm/{"R2212-009913-001"}").Result;

                if (reqResponse.StatusCode == HttpStatusCode.OK)
                {
                    string limsDataString = reqResponse.Content.ReadAsStringAsync().Result;
                }
                else if (reqResponse.StatusCode == HttpStatusCode.Unauthorized)
                {
                    IsUnauthorized = true;
                    //MessageBox.Show("Connection timed out. Kindly re-login to LabVantage");
                }
            }
        }
    }
}
