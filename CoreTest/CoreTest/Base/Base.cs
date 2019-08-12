using AventStack.ExtentReports;
using AventStack.ExtentReports.Reporter;
using NUnit.Framework;
using NUnit.Framework.Interfaces;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Mail;
using System.Reflection;
using System.Text;
using System.Threading;

namespace CoreTest.Base
{
    class Base
    {
        public IWebDriver driver = new ChromeDriver(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location));

        public AventStack.ExtentReports.ExtentReports _extent = new AventStack.ExtentReports.ExtentReports();

        //protected ExtentReports _extent;
        protected ExtentTest _test;
        static string email_pref = "a";
        static string email_suff = "@test.com";

        ///For report directory creation and HTML report template creation
        ///For driver instantiation
        [OneTimeSetUp]
        public void BeforeClass()
        {
            try
            {
                //To create report directory and add HTML report into it

                _extent = new AventStack.ExtentReports.ExtentReports();
                var dir = AppDomain.CurrentDomain.BaseDirectory.Replace("\\bin\\Debug", "");
                DirectoryInfo di = Directory.CreateDirectory(dir + "\\Report");
                var htmlReporter = new ExtentHtmlReporter(dir + "\\Report" + "\\Automation_Report" + ".html");
                _extent.AddSystemInfo("Environment", " Netboost Test");
                _extent.AddSystemInfo("User Name", "Stano");
                _extent.AttachReporter(htmlReporter);
            }
            catch (Exception e)
            {
                throw (e);
            }


        }

        ///Getting the name of current running test to extent report
        [SetUp]
        public void BeforeTest()
        {
            try
            {
                _test = _extent.CreateTest(TestContext.CurrentContext.Test.Name);
            }
            catch (Exception e)
            {
                throw (e);
            }
        }


        /// Finish the execution and logging the detials into HTML report
        [TearDown]
        public void AfterTest()
        {
            try
            {
                var status = TestContext.CurrentContext.Result.Outcome.Status;
                var stacktrace = "" + TestContext.CurrentContext.Result.StackTrace + "";
                var errorMessage = TestContext.CurrentContext.Result.Message;
                Status logstatus;
                switch (status)
                {
                    case TestStatus.Failed:
                        logstatus = Status.Fail;
                        string screenShotPath = Capture(driver, TestContext.CurrentContext.Test.Name);
                        _test.Log(logstatus, "Test ended with " + logstatus + " – " + errorMessage);
                        _test.Log(logstatus, "Snapshot below: " + _test.AddScreenCaptureFromPath(screenShotPath));
                        break;
                    case TestStatus.Skipped:
                        logstatus = Status.Skip;
                        _test.Log(logstatus, "Test ended with " + logstatus);
                        break;
                    default:
                        logstatus = Status.Pass;
                        _test.Log(logstatus, "Test ended with " + logstatus);
                        break;
                }
            }
            catch (Exception e)
            {
                throw (e);
            }
        }

        ///To flush extent report
        ///To quit driver instance
        [OneTimeTearDown]
        public void AfterClass()
        {
            try
            {
                _extent.Flush();
            }
            catch (Exception e)
            {
                throw (e);
            }
            driver.Quit();
        }

        /// To capture the screenshot for extent report and return actual file path
        private string Capture(IWebDriver driver, string screenShotName)
        {
            string localpath = "";
            try
            {
                Thread.Sleep(4000);
                ITakesScreenshot ts = (ITakesScreenshot)driver;
                Screenshot screenshot = ts.GetScreenshot();
                string pth = System.Reflection.Assembly.GetCallingAssembly().CodeBase;
                var dir = AppDomain.CurrentDomain.BaseDirectory.Replace("\\bin\\Debug", "");
                DirectoryInfo di = Directory.CreateDirectory(dir + "\\Defect_Screenshots\\");
                string finalpth = pth.Substring(0, pth.LastIndexOf("bin")) + "\\Defect_Screenshots\\" + screenShotName + ".png";
                localpath = new Uri(finalpth).LocalPath;
                screenshot.SaveAsFile(localpath);
            }
            catch (Exception e)
            {
                throw (e);
            }
            return localpath;
        }
        public void Send_Report_In_Mail()
        {
            MailMessage mail = new MailMessage();
            SmtpClient SmtpServer = new SmtpClient("smtp.gmail.com");
            mail.From = new MailAddress("chinedugodson@gmail.com");
            mail.To.Add("stanosrg@gmail.com");

            StringBuilder TimeAndDate = new StringBuilder(DateTime.Now.ToString());
            TimeAndDate.Replace("/", "_");
            TimeAndDate.Replace(":", "_");

            mail.Subject = "Automation Test Report_" + TimeAndDate;

            mail.Body = "Please find the attached report to get details.";

            string actualPath = AppDomain.CurrentDomain.BaseDirectory.Replace("\\bin\\Debug", "\\Report"); //Reports should store in Test_Execution_Reports folder

            var mostRecentlyModified = Directory.GetFiles(actualPath, "*.html")
            .Select(f => new FileInfo(f))
            .OrderByDescending(fi => fi.LastWriteTime)
            .First()
            .FullName;

            Attachment attachment;
            attachment = new Attachment(mostRecentlyModified);
            mail.Attachments.Add(attachment);

            SmtpServer.Port = 587;
            SmtpServer.Credentials = new System.Net.NetworkCredential("chinedugodson@gmail.com", "faithimex");
            SmtpServer.EnableSsl = true;

            SmtpServer.Send(mail);

        }

    }
}
