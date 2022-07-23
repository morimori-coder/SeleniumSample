using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;

namespace SeleniumSample
{
    internal class SelenimProcess
    {
        private IWebDriver _WebDriver;
        public void DoSelenium()
        {
            ChromeDriverService service = ChromeDriverService.CreateDefaultService();
            service.HideCommandPromptWindow = true;

            ChromeOptions options = new ChromeOptions();
            options.AddArgument("--start-maximized");

            using (_WebDriver = new ChromeDriver(service, options)) 
            {
                _WebDriver.Navigate().GoToUrl("https://www.google.co.jp");
            }
        }
    }
}
