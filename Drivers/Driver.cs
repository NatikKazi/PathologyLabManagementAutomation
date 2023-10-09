using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using SeleniumExtras.PageObjects;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WebDriverManager.DriverConfigs.Impl;
using WebDriverManager;

namespace PathologyLabManagementAutomation.Drivers
{
     class Driver
    {
        public IWebDriver _driver;

        [SetUp]
        public void Init()
        {
            new DriverManager().SetUpDriver(new ChromeConfig());
            _driver = new ChromeDriver();
            _driver.Manage().Window.Maximize();
            _driver.Navigate().GoToUrl("https://gor-pathology.web.app/");
            //Implicit wait
            _driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(10);
            PageFactory.InitElements(_driver, this);
        }

        [TearDown]
        public void Cleanup()
        {
            _driver.Quit();
        }
    }
}
