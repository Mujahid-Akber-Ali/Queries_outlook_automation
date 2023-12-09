

using AventStack.ExtentReports;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenQA.Selenium;
using OpenQA.Selenium.Appium;
using OpenQA.Selenium.Appium.Windows;
using System;
using System.Threading;

namespace Queries
{
    [TestClass]
    public class UnitTest1
    {
        public TestContext instance;
        public TestContext TestContext
        {
            set { instance = value; }
            get { return instance; }
        }

         [TestMethod]
        [DataSource("Microsoft.VisualStudio.TestTools.DataSource.CSV", "data.csv", "data#csv", DataAccessMethod.Sequential)]
        public void TestMethod1()
        {
            string URL = TestContext.DataRow["url"].ToString();

            System.Diagnostics.Process.Start
         (@"C:\Program Files (x86)\Windows Application Driver\WinAppDriver.exe");

            AppiumOptions options = new AppiumOptions();

            options.AddAdditionalCapability("app", @"C:\Program Files (x86)\Microsoft Office\root\Office16\OUTLOOK.EXE");
            options.AddAdditionalCapability("deviceName", "WindowsPC");
            var driver = new WindowsDriver<WindowsElement>
                (new Uri("http://127.0.0.1:4723"), options);
           
            driver.Manage().Window.Size = new System.Drawing.Size(1540, 900);
            driver.FindElementByName("Search").SendKeys(URL);
            Thread.Sleep(2000);
            driver.FindElementByName("Submit Search").Click();
            driver.FindElementByName("Ribbon Tabs").Click();
            driver.FindElementByName("Table View").Click();
            Thread.Sleep(3000);
            TakeScreenshot(driver, Status.Pass, "Screenshot");
            driver.FindElementByName("Search").Clear();
            driver.FindElementByName("Close").Click();
        }

        public void TakeScreenshot(IWebDriver driver, Status status, string stepDetail)
        {

            string path = @"C:\Users\DELL\source\repos\Queries\Screenshot\" + DateTime.Now.ToString("yyyyMMddHHmmss");
            Screenshot image_username = ((ITakesScreenshot)driver).GetScreenshot();
            image_username.SaveAsFile(path + ".png", ScreenshotImageFormat.Png);
            ExtentReport.exChildTest.Log(status, stepDetail, MediaEntityBuilder
                .CreateScreenCaptureFromPath(path + ".png").Build());


        }
    }
}
