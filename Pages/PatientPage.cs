using Newtonsoft.Json;
using OpenQA.Selenium;
using PathologyLabManagementAutomation.Drivers;
using PathologyLabManagementAutomation.Models;
using OpenQA.Selenium.Interactions;
using Spire.Xls;
using System;
using System.Linq;
using System.Threading;
using System.Xml.Linq;

namespace PathologyLabManagementAutomation.Pages
{
    class PatientPage : Driver
    {
        private IWebDriver _driver;

        public PatientPage(IWebDriver driver)
        {
            _driver = driver;
            // Login to the application
            Login();
            InitializePatientCreation();
        }

        private void Login()
        {
            _driver.FindElement(By.Name("email")).SendKeys("test@kennect.io");
            _driver.FindElement(By.Name("password")).SendKeys("Qwerty@1234");
            _driver.FindElement(By.XPath("//span[@class='MuiButton-label'][text()='Login']")).Click();
            Thread.Sleep(2000); // Wait for login to complete
        }

        private void InitializePatientCreation()
        {
            _driver.FindElement(By.XPath("//span[normalize-space()='Patients']")).Click();
            _driver.FindElement(By.XPath("//span[text() = 'Add Patient']")).Click();
        }

        public void PatientContact(string name, string email, string phone)
        {
            _driver.FindElement(By.Name("name")).SendKeys(name);
            _driver.FindElement(By.Name("email")).SendKeys(email);
            _driver.FindElement(By.Name("phone")).SendKeys(phone);

            ScrollToBottom();
            _driver.FindElement(By.XPath("//img[@alt = 'add']")).Click();
        }

        public void GeneralDetails(string height, string weight, string gender, string age, string bp1, string bp2)
        {
            _driver.FindElement(By.Name("height")).SendKeys(height);
            _driver.FindElement(By.Name("weight")).SendKeys(weight);

            SelectGender(gender);

            _driver.FindElement(By.Name("age")).SendKeys(age);

            ScrollToElement(_driver.FindElement(By.Name("systolic")));
            _driver.FindElement(By.Name("systolic")).SendKeys(bp1);
            _driver.FindElement(By.Name("diastolic")).SendKeys(bp2);

            ScrollToBottom();
            _driver.FindElement(By.XPath("//img[@alt = 'add']")).Click();
        }

        public void SelectGender(string gender)
        {
            _driver.FindElement(By.Id("mui-component-select-gender")).Click();
            string genderOptionXPath = $"//li[@data-value='{gender}']";
            _driver.FindElement(By.XPath(genderOptionXPath)).Click();
        }

        public void AddTest(string test, string dis, string sellab, string dname)
        {
            // Initialize Actions for performing advanced interactions
            Actions actions = new Actions(_driver);

            // Locate and click the patient test input field
            _driver.FindElement(By.Id("patient-test")).Click();
            // Enter the test value
            _driver.FindElement(By.Id("patient-test")).SendKeys(test);

            // Click the dropdown
            _driver.FindElement(By.XPath("//div[@aria-haspopup = 'listbox']")).Click();

            // Select the specified 'dis' value from the dropdown
            string disamt = $"//li[@data-value = '{dis}']";
            _driver.FindElement(By.XPath(disamt)).Click();

            // Locate and click the combobox
            IWebElement combobox = _driver.FindElement(By.XPath("//div[@role = 'combobox']"));
            // Move the cursor 60 pixels downward from the current position within the combobox and click
            actions.MoveToElement(combobox, 0, 60).Click().Perform();

            // Locate and send keys to the patient-tests-labs-label input field
            IWebElement SelLab = _driver.FindElement(By.Id("patient-tests-labs-label"));
            SelLab.SendKeys(sellab);
            // Move the cursor 60 pixels downward from the current position within the SelLab input field and click
            actions.MoveToElement(SelLab, 0, 60).Click().Perform();

            // Locate and send keys to the doctor_name input field
            _driver.FindElement(By.Name("doctor_name")).SendKeys(dname);
            // Locate and click the doctor_name dropdown
            IWebElement ddrp = _driver.FindElement(By.Name("doctor_name"));
            // Move the cursor 60 pixels downward from the current position within the dropdown and click
            actions.MoveToElement(ddrp, 0, 60).Click().Perform();

            // Scroll to the bottom of the page using JavaScript
            IJavaScriptExecutor js = (IJavaScriptExecutor)_driver;
            js.ExecuteScript("window.scrollTo(0, document.body.scrollHeight);");

            // Locate and click the 'add_box' element
            _driver.FindElement(By.XPath("//span[text()='add_box']")).Click();

            // Locate and click the 'Equipment Name' div
            _driver.FindElement(By.XPath("//div[@aria-label = 'Equipment Name']")).Click();
            // Locate and click a specific equipment option
            _driver.FindElement(By.XPath("//li[@data-value = 'UMLmlWLxrpDLyfzipFRY']")).Click();

            // Locate and click the input field for quantity
            IWebElement qty = _driver.FindElement(By.XPath("//input[@aria-label = 'Required']"));
            qty.Click();
            // Clear the input field and enter the quantity value
            qty.Clear();
            qty.SendKeys("5");

            // Locate and click the 'check' element
            _driver.FindElement(By.XPath("//span[text() = 'check']")).Click();

            // Locate the target element to scroll into view
            IWebElement targetElement = _driver.FindElement(By.XPath("//img[@alt = 'add']"));
            // Scroll the target element into view
            ((IJavaScriptExecutor)_driver).ExecuteScript("arguments[0].scrollIntoView(true);", targetElement);
            // Click the target element
            targetElement.Click();

        }

        public void ReadExcel_Patient(string path, string Sheetname)
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
                PatientModel pm = JsonConvert.DeserializeObject<PatientModel>(data);

                sheet.Rows[i].ToList()[6].Text = GetResult(pm.name, pm.email, pm.phone, pm.height, pm.weight, pm.gender, pm.age, pm.bp1, pm.bp2, pm.test, pm.dis, pm.sellab, pm.dname);
            }

            workbook.SaveToFile(path, ExcelVersion.Version2013);
        }

        public string GetResult(string name, string email, string phone, string height, string weight, string gender, string age, string bp1, string bp2, string test, string dis, string sellab, string dname)
        {
            PatientContact(name, email, phone);
            GeneralDetails(height, weight, gender, age, bp1, bp2);
            AddTest(test, dis, sellab, dname);

            try
            {
                _driver.FindElement(By.XPath("//span[normalize-space()='Patients']")).Click();
                IWebElement searchbox = _driver.FindElement(By.XPath("//input[@placeholder = 'Search']"));
                searchbox.SendKeys(name);
                string namechk = $"//td[@value = '{name}']";
                IWebElement namelist = _driver.FindElement(By.XPath(namechk));

                return namelist != null ? "Pass" : "Fail";
            }
            catch (Exception ex)
            {
                Console.WriteLine("An exception occurred: " + ex.Message);
                return "Code Fail";
            }
        }

        // Helper methods for scrolling and interacting with elements
        private void ScrollToBottom()
        {
            IJavaScriptExecutor js = (IJavaScriptExecutor)_driver;
            js.ExecuteScript("window.scrollTo(0, document.body.scrollHeight);");
        }

        private void ScrollToElement(IWebElement element)
        {
            ((IJavaScriptExecutor)_driver).ExecuteScript("arguments[0].scrollIntoView(true);", element);
        }

        private void MoveToElementAndClick(By by)
        {
            Actions actions = new Actions(_driver);
            IWebElement element = _driver.FindElement(by);
            actions.MoveToElement(element, 0, 60).Click().Perform();
        }
    }
}
