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
    class PatientTest : Driver
    {
        [Test]
        public void LoginFunc()
        {
            PatientPage page = new PatientPage(_driver);
            page.ReadExcel_Patient("D:\\VisualStudioMain\\PathologyLabManagementAutomation\\Data Test\\PathologyLabManagement.xlsx", "Patient Page");
        }
    }
}
