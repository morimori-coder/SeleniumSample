using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;


namespace SeleniumSample
{
    internal class Program
    {
        

        static void Main(string[] args)
        {
            var seleniumProcess = new SelenimProcess();
            seleniumProcess.SearchChrome();
            var excel = new ExcelOperation();
            excel.Excel_OutPutEx(seleniumProcess.SearchResult);
        }

        
    }
}
