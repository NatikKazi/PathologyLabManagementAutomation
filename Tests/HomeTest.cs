using PathologyLabManagementAutomation.Pages;
using PathologyLabManagementAutomation.Drivers;

namespace PathologyLabManagementAutomation.Tests
{
    [TestFixture]
    class HomeTest : Driver
    {
        [Test]
        public void Homeverification()
        {
            HomePage homepage = new HomePage(_driver);
            homepage.ReadExcel_Home("D:\\VisualStudioMain\\PathologyLabManagementAutomation\\Data Test\\PathologyLabManagement.xlsx", "Home Page");
        }
    }
}
