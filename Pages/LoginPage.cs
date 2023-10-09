using Newtonsoft.Json;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;
using PathologyLabManagementAutomation.Drivers;
using PathologyLabManagementAutomation.Models;
using SeleniumExtras.PageObjects;
using SeleniumExtras.WaitHelpers;
using Spire.Xls;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PathologyLabManagementAutomation.Pages
{
    class LoginPage : Driver
    {
        [FindsBy(How = How.Name, Using = "email")]
        private IWebElement Email { get; set; }

        [FindsBy(How = How.Name, Using = "password")]
        private IWebElement Password { get; set; }

        [FindsBy(How = How.XPath, Using = "//span[@class='MuiButton-label'][text()='Login']")]
        private IWebElement Loginbtn { get; set; }

        [FindsBy(How = How.XPath, Using = "//div[@class = 'MuiAlert-message']")]
        private IWebElement AlertPopup { get; set; }

        [FindsBy(How = How.XPath, Using = "//button[@aria-label='account of current user']")]
        private IWebElement Logoutbtn { get; set; }

        [FindsBy(How = How.XPath, Using = "//span[@class='MuiButton-label'][text() = 'Sign out of Lab']")]
        private IWebElement Signoutbtn { get; set; }

        





        //public IWebDriver _driver;
        public LoginPage(IWebDriver driver)
        {
            _driver = driver;
            PageFactory.InitElements(_driver, this);
        }
        public void Login(string empid, string password)
        {
            _driver.Navigate().Refresh();
            Email.Click();
            Email.SendKeys(Keys.Control + "a");
            Email.SendKeys(Keys.Delete);
            Email.Clear();
            Email.SendKeys(empid);
            Password.Click();
            Password.SendKeys(Keys.Control + "a");
            Password.SendKeys(Keys.Delete);
            Password.Clear();
            Password.SendKeys(password);
            Loginbtn.Click();
        }

        Worksheet sheet;
        int columnCount;
        int rowCount;
        public void ReadExcel_Login(string path, string Sheetname)
        {
            Workbook workbook = new Workbook();
            workbook.LoadFromFile(path);
            sheet = workbook.Worksheets[Sheetname];
            columnCount = sheet.Columns.Count();
            rowCount = sheet.Rows.Count();
            if (columnCount <= 6)
            {
                sheet.InsertColumn(6, 1);
            }
            sheet.Rows[1].ToList()[6].Text = "Status";
 

            for (int i = 2; i < rowCount; i++)
            {
                try
                {
                    string data = sheet.Rows[i].ToList()[5].Value;
                    LoginModel lm = JsonConvert.DeserializeObject<LoginModel>(data);

                    sheet.Rows[i].ToList()[6].Text = GetResult(lm.email,lm.password);
                }
                catch (Exception ex)
                {

                    throw;
                }
               
            }
            workbook.SaveToFile(path, ExcelVersion.Version2013);
        }
        // runner
        public string GetResult(string _user, string _pass)
        {

            Login(_user, _pass);
            try
            {
                WebDriverWait wait = new WebDriverWait(_driver, TimeSpan.FromSeconds(10));

                // Wait for the alert to be present
                //IWebElement alert = wait.Until(AlertPopup.Displayed);

                // If an alert is present, handle it
                var ele = wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("//div[@class = 'MuiAlert-message']")));

                //var el=_driver.FindElement(By.XPath("//div[@class = 'MuiAlert-message']"));

                if (ele != null)
                {
                    return "Fail";
                }
                
            }

            catch (Exception ex)
            {
                //Console.WriteLine("An exception occurred: " + ex.Message);
                Logoutbtn.Click();
                Signoutbtn.Click();
            }
            return "Pass";
        }
    }
}
