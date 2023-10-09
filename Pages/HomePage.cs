using Newtonsoft.Json;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.UI;
using PathologyLabManagementAutomation.Models;
using Spire.Xls;
using System;
using System.Threading;

namespace PathologyLabManagementAutomation.Pages
{
    class HomePage
    {
        private IWebDriver _driver;

        public HomePage(IWebDriver driver)
        {
            _driver = driver;

            // Login to the application
            _driver.FindElement(By.Name("email")).SendKeys("test@kennect.io");
            _driver.FindElement(By.Name("password")).SendKeys("Qwerty@1234");
            _driver.FindElement(By.XPath("//span[@class='MuiButton-label'][text()='Login']")).Click();
            Thread.Sleep(2000); // Wait for login to complete
        }

        public void Home(string cost, string dis)
        {
            Actions actions = new Actions(_driver);

            // Find the element or location where you want to move the cursor
            IWebElement targetElement = _driver.FindElement(By.XPath("//p[text()='Discount for customer']"));
            ((IJavaScriptExecutor)_driver).ExecuteScript("arguments[0].scrollIntoView(true);", targetElement);
            Thread.Sleep(2000);

            // Enter cost value
            _driver.FindElement(By.Id("patient-test")).Click();
            _driver.FindElement(By.Id("patient-test")).SendKeys(cost);
            //Thread.Sleep(2000); // Wait for input to populate

            // Move the cursor 12 pixels downward from the current position
            IWebElement combobox = _driver.FindElement(By.XPath("//div[@role = 'combobox']"));
            actions.MoveToElement(combobox, 0, 60).Perform();
            //Thread.Sleep(2000); // Wait for cursor movement

            // Click at the current cursor position
            actions.Click().Perform();

            // Click the dropdown
            _driver.FindElement(By.XPath("//div[@aria-haspopup = 'listbox']")).Click();

            // Select the specified 'dis' value from the dropdown
            string disamt = $"//li[@data-value = '{dis}']";
            _driver.FindElement(By.XPath(disamt)).Click();
        }

        public void ReadExcel_Home(string path, string Sheetname)
        {
            Workbook workbook = new Workbook();
            workbook.LoadFromFile(path);
            Worksheet sheet = workbook.Worksheets[Sheetname];
            int columnCount = sheet.Columns.Count();
            int rowCount = sheet.Rows.Count();

            if (columnCount <= 6)
            {
                sheet.InsertColumn(6, 1);
            }

            sheet.Rows[1].ToList()[6].Text = "Status";

            for (int i = 2; i < rowCount; i++)
            {
                string data = sheet.Rows[i].ToList()[5].Value;
                CalculatorModel cm = JsonConvert.DeserializeObject<CalculatorModel>(data);

                sheet.Rows[i].ToList()[6].Text = GetResult(cm.Cost, cm.Discount);
            }

            workbook.SaveToFile(path, ExcelVersion.Version2013);
        }

        public string GetResult(string cost, string discount)
        {
            Home(cost, discount);

            try
            {
                IWebElement check = _driver.FindElement(By.XPath("//div[@class = 'MuiBox-root jss93']"));
                IWebElement todo = _driver.FindElement(By.XPath("//div[@class='MuiDrawer-root MuiDrawer-docked']"));

                if (check.Displayed && todo.Displayed)
                {
                    // Clear the input field
                    IWebElement inputField = _driver.FindElement(By.XPath("//input[@id = 'patient-test']"));
                    inputField.Click();
                    inputField.Clear();

                    // Click the clear button
                    IWebElement clearButton = _driver.FindElement(By.XPath("//button[@title='Clear']//span[@class='MuiIconButton-label']//*[name()='svg']"));
                    clearButton.Click();

                    return "Pass";
                }

                return "Fail";
            }
            catch (Exception ex)
            {
                Console.WriteLine("An exception occurred: " + ex.Message);
            }

            return "Code Fail";
        }
    }
}
