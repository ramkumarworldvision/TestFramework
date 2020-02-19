using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;

namespace TestAutomationFramework
{
    public static class SeleniumSetMethods
    {
        
        public static void entertext(this IWebElement element, string value)
        {
            element.SendKeys(value);

        }

        public static void clicks(this IWebElement element)
        {
            element.Click();
        }

        public static void selectdropdown(this IWebElement element, int value)
        {

            new SelectElement(element).SelectByIndex(value);
        }

        public static void selectdropdowntext(this IWebElement element, string value)
        {
            new SelectElement(element).SelectByText(value);
        }

        public static void selectdropdownvalue(this IWebElement element, string value)
        {
            new SelectElement(element).SelectByValue(value);
        }
    }
}
