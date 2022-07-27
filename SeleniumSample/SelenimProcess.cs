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

        public List<string> SearchResult { get; set; }

        public void SearchChrome()
        {
            ChromeDriverService service = ChromeDriverService.CreateDefaultService();
            service.HideCommandPromptWindow = true;

            ChromeOptions options = new ChromeOptions();
            options.AddArgument("--start-maximized");

            using (_WebDriver = new ChromeDriver(service, options))
            {
                _WebDriver.Navigate().GoToUrl("https://www.google.co.jp");
                var element = _WebDriver.FindElement(By.CssSelector("検索テキストボックス"));
                element.SendKeys("検索キーワード");
                element.SendKeys(Keys.Enter);

                var h3List = new List<string>();
                List<IWebElement> currentElements = new List<IWebElement>();
                IWebElement nextLink;
                try
                {
                    while (true)
                    {
                        currentElements = _WebDriver.FindElements(By.ClassName("URLなどの要素")).ToList();
                        currentElements.ForEach(e => { h3List.Add(e.Text.Replace("\r\n","\t")); });

                        if (_WebDriver.FindElements(By.Id("つぎへ")).Count > 0)
                        {
                            nextLink = _WebDriver.FindElement(By.Id("つぎへ"));
                            nextLink.Click();
                        }
                        else
                            break;
                    }
                    this.SearchResult = h3List;
                }
                catch (Exception e)
                {
                    Console.WriteLine(e.Message);
                }

            }
        }
    }
}