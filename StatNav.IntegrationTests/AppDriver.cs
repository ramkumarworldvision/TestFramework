using AventStack.ExtentReports;
using AventStack.ExtentReports.Reporter;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;

namespace TestAutomationFramework
{
     enum ProperType
    {
        id,
        name,
        linkText,
        xpath
    }
    public class AppDriver
    {
        public static IWebDriver driver { get; set; }
        public static ExtentV3HtmlReporter htmlReporter { get; set; }
        public static ExtentReports extent { get; set; }
        public static ExtentTest test { get; set; }
        public static WebDriverWait wait { get; set; }

        public static Screenshot file;

    }
      
}
