using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http;
using Microsoft.Office.Core;
using Newtonsoft.Json;

using System.Net;

using Excel = Microsoft.Office.Interop.Excel;
using System.Net.Http.Headers;

using System.Net;
using IntakeBase=Xom.Gci.Addin.LvMake.Intake.Base;
using IIntakeBase = Xom.Gci.Addin.LvMake.IIntake;
using Xom.Gci.Addin.LvMake.SimpleIntake;
using Xom.Gci.Addin.LvMake.Helpers;
using Xom.Gci.Addin.LvMake.Common.Models;
using Microsoft.Office.Interop.Excel;
using System.Reflection;

namespace ConsoleApp1
{
    public class Program 
    {

        public static Excel.Workbook workbook ;
        public static string userName = "naveene";
        public static string password = "exxonmobilinfosys";

        public static string lvurl = "https://hoeapp910.na.xom.com/labvantage/rest";
        public static string absoluteurl = "https://hoeapp910.na.xom.com/labvantage";
        public static string databaseid = "LabVantageNADEV";
        public LvHelper LvHelper { get; set; }
        public  static void Main(string[] args)
        {
            //NewExcelFile();
   
            IntakeFormExcelFile();
           // Login();
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
            string fileName = "R2302-010016-001_Intake.xlsx";
            string sourcePath = @"C:\Users\naveene\Downloads\New folder (2)";
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
                        workbook = excelApp.ActiveWorkbook;

                        //DocumentProperties prop = workbook.CustomDocumentProperties as DocumentProperties;

                        // New code to get document properties
                        Dictionary<string, string> docProps = new Dictionary<string, string>();
                        if (workbook != null)
                        {
                            dynamic properties = workbook.CustomDocumentProperties;
                            foreach (dynamic p in properties)
                            {
                                docProps.Add(Convert.ToString(p.Name), Convert.ToString(p.Value));
                            }
                        }

                        //IntakeBase.IntakeValidate validate = new IntakeBase.IntakeValidate();
                        IntakeBase.OpenWorkbook openWorkbook = new IntakeBase.OpenWorkbook();
                      
                        var configWorksheet = (Excel.Worksheet)workbook.Sheets[GlobalValue.ConfigSheet];

                        openWorkbook.ConfigHelper = new ConfigHelper(configWorksheet);


                        openWorkbook.ConfigHelper.LoadConfigValues();

                        Excel.Worksheet mainSheet = (Excel.Worksheet)workbook.Worksheets[GlobalValue.MainSheet];
                   
                        openWorkbook.AddinHelper = new SimpleIntakeAddinHelper(workbook, false);
                        openWorkbook.AddinHelper.Unprotect();
                        openWorkbook.Ribbon = new SimpleIntakeRibbon(workbook);
                        openWorkbook.Inventory = new SimpleIntakeInventory(workbook);
                        openWorkbook.ActualIngredient = new SimpleIntakeActualIng(workbook);
                        openWorkbook.Formulation = new SimpleIntakeFormulation(workbook);
                        openWorkbook.Tests = new SimpleIntakeTests(workbook);
                        openWorkbook.LIMSDataFetch = new SimpleIntakeLIMSDataFetch(workbook);
                        openWorkbook.DynamicConfigHelper = new DynamicConfigHelper(workbook);
                        openWorkbook.DropdownHelper = new DropdownHelper(workbook);
                        openWorkbook.SyncHelper = new SyncHelper(workbook);
                        openWorkbook.ColumnWidthHelper = new ColumnWidthHelper(workbook);
                        openWorkbook.LockCellsHelper = new LockCellsHelper(workbook);
                        openWorkbook.CopyCellsHelper = new CopyCellsHelper(workbook);
                        openWorkbook.CfHelper = new ConditionalFormattingHelper(workbook);
                        openWorkbook.RibbonConfigHelper = new RibbonConfigHelper(workbook);
                        openWorkbook.PrintConfigHelper = new PrintConfigHelper(workbook);
                        openWorkbook.ChangeLogHelper = new ChangeLogHelper(workbook);
                        openWorkbook.Ingredient = new SimpleIntakeIngredient(workbook, false);
                        openWorkbook.TreatedIngredient = new SimpleIntakeTreatedIngredient(workbook, false);
                        openWorkbook.TreatedBatches = new SimpleIntakeTreatedBatches(workbook);
                        string userName = "naveene";
                        string password = "exxonmobilinfosys";
                        string lvurl = "https://hoeapp910.na.xom.com/labvantage/rest";
                        string absoluteurl = "https://hoeapp910.na.xom.com/labvantage";
                        string databaseid = "LabVantageNADEV";
                        openWorkbook.LvHelper = new LvHelper(userName, password, lvurl, databaseid);

                        openWorkbook.LvHelper.Login();
                        openWorkbook.ProcessVariables = new SimpleIntakeProcessVariables(workbook);
                        openWorkbook.VariableSetting = new SimpleIntakeVariableSetting(workbook);
                        openWorkbook.Blends = new SimpleIntakeBlends(workbook, false);
                        openWorkbook.BatchContainer = new SimpleIntakeBatchContainer(workbook);
                        openWorkbook.Context = new SimpleIntakeContext(workbook, false);
                        openWorkbook.ReviewComplete = new SimpleIntakeReviewComplete(workbook, false);
                        openWorkbook.Validation = new SimpleIntakeValidation(workbook, false);
                     
                        LIMSDataModel limsDataModel = openWorkbook.LIMSDataFetch.FetchAndLoadMasterData(openWorkbook.LvHelper, "https://hoeapp910.na.xom.com/labvantage", "R2302-010016-001");
                        if (limsDataModel != null)
                        {
                            //Reinstantiating Validation object citing Master Data Refresh
                            //if (IsSimpleIntakeFile(workbook) || IsRubberIntakeFile(workbook))
                            //{
                                openWorkbook.Validation = new SimpleIntakeValidation(workbook, false);
                                openWorkbook.Ribbon = new SimpleIntakeRibbon(workbook);
                                openWorkbook.ReviewComplete = new SimpleIntakeReviewComplete(workbook, false);
                                openWorkbook.ProcessVariblesBatchContainers = new SimpleIntakeContainerProcessVariables(workbook);
                                Excel.Worksheet sheetName = (Excel.Worksheet)workbook.Worksheets[GlobalValue.FormulationSheet];
                                openWorkbook.Validation.ApplyCellFormats(sheetName);
                                openWorkbook.LIMSDataFetch.FetchAndLoadSheetData(openWorkbook.LvHelper, limsDataModel, openWorkbook.Validation, openWorkbook.Context, openWorkbook.Ingredient, openWorkbook.Formulation,
                                openWorkbook.Blends, openWorkbook.ProcessVariables, openWorkbook.ActualIngredient, openWorkbook.BatchContainer,
                                openWorkbook.Inventory, openWorkbook.Ribbon, openWorkbook.CfHelper, openWorkbook.ColumnWidthHelper, openWorkbook.TreatedIngredient, openWorkbook.TreatedBatches, openWorkbook.Tests, openWorkbook.ProcessVariblesBatchContainers, openWorkbook.ReviewComplete, true);
                            //}
                            //else if (IsStructuredIntakeFile(workbook))
                            //{

                            //    //openWorkbook.Validation = new FilmIntakeValidation(workbook);
                            //    openWorkbook.Validation = new SimpleIntakeValidation(workbook, false);
                            //    openWorkbook.Ribbon = new SimpleIntakeRibbon(workbook);
                            //    openWorkbook.ReviewComplete = new SimpleIntakeReviewComplete(workbook, false);
                            //    openWorkbook.ProcessVariblesBatchContainers = new SimpleIntakeContainerProcessVariables(workbook);
                            //    Excel.Worksheet sheetName = (Excel.Worksheet)workbook.Worksheets[GlobalValue.StructuredFormulationSheet];
                            //    openWorkbook.Validation.ApplyCellFormats(sheetName);
                            //    openWorkbook.LIMSDataFetch.FetchAndLoadSheetData(openWorkbook.LvHelper, limsDataModel, openWorkbook.Validation, openWorkbook.Context, openWorkbook.Ingredient, openWorkbook.Formulation,
                            //    openWorkbook.Blends, openWorkbook.ProcessVariables, openWorkbook.ActualIngredient, openWorkbook.BatchContainer,
                            //    openWorkbook.Inventory, openWorkbook.Ribbon, openWorkbook.CfHelper, openWorkbook.ColumnWidthHelper, openWorkbook.TreatedIngredient, openWorkbook.TreatedBatches, openWorkbook.Tests, openWorkbook.ProcessVariblesBatchContainers, openWorkbook.ReviewComplete, false);
                            //    openWorkbook.Formulation.UpdateShortNamesInDropdown();
                            //    //openWorkbook.ProcessVariblesBatchContainers = new FilmIntakeContainerProcessVariables(workbook);
                            //}
                        }
          
                        openWorkbook.AddinHelper.Protect();
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


            public static bool IsStructuredIntakeFile(Excel.Workbook workbook)
            {
                try
                {
                    DocumentProperties properties = workbook.BuiltinDocumentProperties as DocumentProperties;
                    string keywords = properties["Keywords"].Value;

                    return (!string.IsNullOrEmpty(keywords) && keywords.Contains("LV Make") && keywords.Contains("Structured"));
                }
                catch
                {
                    return false;
                }
            }

            public static bool IsRubberIntakeFile(Excel.Workbook workbook)
            {
                try
                {
                    DocumentProperties properties = workbook.BuiltinDocumentProperties as DocumentProperties;
                    string keywords = properties["Keywords"].Value;

                    return (!string.IsNullOrEmpty(keywords) && keywords.Contains("LV Make") && keywords.Contains("Rubber"));
                }
                catch
                {
                    return false;
                }
            }
            public static bool IsMonoIntakeFile(Excel.Workbook workbook)
            {
                try
                {
                    DocumentProperties properties = workbook.BuiltinDocumentProperties as DocumentProperties;
                    string keywords = properties["Keywords"].Value;

                    return (!string.IsNullOrEmpty(keywords) && keywords.Contains("LV Make") && keywords.Contains("Mono-Layer"));
                }
                catch
                {
                    return false;
                }
            }

            public static bool IsSimpleIntakeFile(Excel.Workbook workbook)
            {
                try
                {
                    DocumentProperties properties = workbook.BuiltinDocumentProperties as DocumentProperties;
                    string keywords = properties["Keywords"].Value;

                    return (!string.IsNullOrEmpty(keywords) && keywords.Contains("LV Make") && (keywords.Contains("Simple") || keywords.Contains("Mono-Layer")));
                }
                catch
                {
                    return false;
                }
            }
        }
}
