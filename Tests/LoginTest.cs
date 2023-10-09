using PathologyLabManagementAutomation.Drivers;
using PathologyLabManagementAutomation.Pages;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PathologyLabManagementAutomation.Tests
{
    [TestFixture]
     class LoginTest : Driver
    {
        [Test]
        public void LoginFunc()
        {
            LoginPage page = new LoginPage(_driver);
            page.ReadExcel_Login("D:\\VisualStudioMain\\PathologyLabManagementAutomation\\Data Test\\PathologyLabManagement.xlsx", "Login Page");
        }
    }
}
