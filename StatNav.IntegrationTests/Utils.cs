using System;
using System.Collections.Generic;
using System.Linq;
using System.Configuration;
using System.IO;
using System.Data;
using OpenQA.Selenium;
using System.Net.Mail;
using OpenQA.Selenium.Support.UI;
using System.Diagnostics;
using Excel;
using System.Net;
using AventStack.ExtentReports.Reporter;
using AventStack.ExtentReports;

namespace TestAutomationFramework
{
    public static class Utils

    {
        public  static DataTable ExcelToDataTable(string fileName,string strSheetname)
        {
            //open file and returns as Stream
            FileStream stream = File.Open(fileName, FileMode.Open, FileAccess.Read);
            //Createopenxmlreader via ExcelReaderFactory
            IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream); //.xlsx                            
            //Set the First Row as Column Name
            excelReader.IsFirstRowAsColumnNames = true;
            //Return as DataSet
            DataSet result = excelReader.AsDataSet();
            //Get all the Tables
            DataTableCollection table = result.Tables;
            //Store it in DataTable
            DataTable resultTable = table[strSheetname];

            //return
            return resultTable;
        }
        

        public static string ReadData(int rowNumber, string columnName)
        {
            try
            {
                //Retriving Data using LINQ to reduce much of iterations
                string data = (from colData in dataCol
                               where colData.colName == columnName && colData.rowNumber == rowNumber
                               select colData.colValue).SingleOrDefault();

                //var datas = dataCol.Where(x => x.colName == columnName && x.rowNumber == rowNumber).SingleOrDefault().colValue;
                return data.ToString();
                //colData = null;
            }
            catch (Exception )
            {
                return null;
            }
        }


       public static List<Datacollection> dataCol = null;
        

        public static void PopulateInCollection(string fileName,string strSheetname)
           
        {

            dataCol = new List<Datacollection>();
            DataTable table = ExcelToDataTable(fileName,strSheetname);
            
            //Iterate through the rows and columns of the Table
            for (int row = 1; row <= table.Rows.Count; row++)
            {
                for (int col = 0; col < table.Columns.Count; col++)
                {
                    Datacollection dtTable = new Datacollection()
                    {
                        rowNumber = row,
                        colName = table.Columns[col].ColumnName,
                        colValue = table.Rows[row - 1][col].ToString()
                    };
                    //Add all the details for each row
                    dataCol.Add(dtTable);
                }
            }
            
        }
       
        public  class Datacollection
        {
            public int rowNumber { get; set; }
            public string colName { get; set; }
            public string colValue { get; set; }
        }



        public static string Capture(string ScreenshotName)
        {
            Screenshot file = ((ITakesScreenshot)AppDriver.driver).GetScreenshot();
            string path = System.Reflection.Assembly.GetCallingAssembly().CodeBase;
            string uptobinpath = path.Substring(0, path.LastIndexOf("bin")) + ScreenshotName + ".png";
            string localpath = new Uri(uptobinpath).LocalPath;
            file.SaveAsFile(localpath, ScreenshotImageFormat.Png);
            return localpath;
        }

        public static void SendMail(String Subject, String contentBody)
        {
            MailMessage mail = new MailMessage();
            string[] toAddrs = ConfigurationManager.AppSettings["ToMail"].Split(';');
            foreach (string addr in toAddrs)
            {
                if (!string.IsNullOrEmpty(addr))
                {
                    mail.To.Add(addr);
                }
            }

            mail.From = new MailAddress(ConfigurationManager.AppSettings["FromMail"]);
            mail.Subject = ConfigurationManager.AppSettings["MailSubject"];
            mail.IsBodyHtml = true;
            mail.Attachments.Add(new Attachment(ConfigurationManager.AppSettings["ReportsPath"] + "Execution_Report.html"));
            mail.Body = ConfigurationManager.AppSettings["MailBody"];

            SmtpClient smtp = new SmtpClient();
            smtp.Host = ConfigurationManager.AppSettings["SmtpHost"];
            smtp.Port = Int32.Parse(ConfigurationManager.AppSettings["SmtpPort"]);
            smtp.Credentials = new System.Net.NetworkCredential(ConfigurationManager.AppSettings["SmtpUserId"], ConfigurationManager.AppSettings["SmtpPassword"]);
            smtp.EnableSsl = ConfigurationManager.AppSettings["EnableSSL"] == "Y";
            smtp.Send(mail);
        }

        public static string AppendTimeStamp(this string filename)
        {
                return string.Concat(
                Path.GetFileNameWithoutExtension(filename),
                DateTime.Now.ToString("yyyyMMddHHmmssfff"),
                Path.GetExtension(filename));
        }

        public static void Contentdelete(string path)
        {
            System.IO.DirectoryInfo di = new DirectoryInfo(path);
            foreach(FileInfo file in di.GetFiles())
            {
                file.Delete();
            }
            
            foreach (DirectoryInfo file1 in di.GetDirectories())
            {
                file1.Delete(true);
            }
        }

        public static void InitiateDelete()
        {
            Utils.Contentdelete(ConfigurationManager.AppSettings["ReportsPath"]);
        }

        public static void CreateFileOrFolder(string subfolder)
        {
            string folderName = ConfigurationManager.AppSettings["ReportsPath"]; 
            string pathString = System.IO.Path.Combine(folderName, subfolder);
            System.IO.Directory.CreateDirectory(pathString);
        }

        public static void DriverSetup()
        {

            foreach (var process in Process.GetProcessesByName("chromedriver"))
            { process.Kill(); }
            foreach (var process in Process.GetProcessesByName("geckodriver"))
            { process.Kill(); }
            foreach (var process in Process.GetProcessesByName("IEDriverServer"))
            { process.Kill(); }

        }

        public static void checkAlert()
        {
            try
            {
                WebDriverWait wait = new WebDriverWait(AppDriver.driver, TimeSpan.FromMilliseconds(100));
                wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.AlertIsPresent());
                IAlert alert = AppDriver.driver.SwitchTo().Alert();
                alert.Accept();
            }
            catch (Exception )
            {
                
            }
        }

        public static void Extent_Test(string htmlFilepath)
        {
            //Html Report Initialization
            AppDriver.htmlReporter = new ExtentV3HtmlReporter(htmlFilepath);
            AppDriver.htmlReporter.Config.Theme = AventStack.ExtentReports.Reporter.Configuration.Theme.Standard;
            AppDriver.extent = new ExtentReports();
            AppDriver.extent.AttachReporter(AppDriver.htmlReporter);
            string hostname = Dns.GetHostName();
            OperatingSystem OS = Environment.OSVersion;
            AppDriver.extent.AddSystemInfo("Operating System", OS.ToString());
            AppDriver.extent.AddSystemInfo("HostName", hostname);
            //AppDriver.extent.AddSystemInfo("Browser", AppDriver.driver.GetType().ToString());
            AppDriver.extent.AddSystemInfo("Browser", "Chrome");
        }
    }
}

