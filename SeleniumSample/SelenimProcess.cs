using System;
using System.Collections.Generic;
using System.Linq;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;

namespace SeleniumSample
{
    internal class SelenimProcess
    {
        private IWebDriver _WebDriver;

        public List<IWebElement> SearchResult { get; set; }

        public void SearchChrome()
        {
            ChromeDriverService service = ChromeDriverService.CreateDefaultService();
            service.HideCommandPromptWindow = true;

            ChromeOptions options = new ChromeOptions();
            options.AddArgument("--start-maximized");

            using (_WebDriver = new ChromeDriver(service, options)) 
            {
                _WebDriver.Navigate().GoToUrl("https://www.google.co.jp");
                var element = _WebDriver.FindElement(By.CssSelector("検索バー"));
                element.SendKeys("検索したいキーワード");
                element.SendKeys(Keys.Enter);

                var h3List = _WebDriver.FindElements(By.ClassName("各検索結果")).ToList();
                var nextLink = _WebDriver.FindElement(By.Id("つぎへ"));
                int counter = 0;
                try 
                {
                    while (nextLink != null)
                    {
                        nextLink.Click();
                        h3List.AddRange(_WebDriver.FindElements(By.ClassName("各検索結果")).ToList());
                        nextLink = _WebDriver.FindElement(By.Id("つぎへ"));
                        counter++;
                    }
                    this.SearchResult = h3List;
                } catch (Exception e)
                {
                    Console.WriteLine(e.Message);
                }
            }
        }
    }
}