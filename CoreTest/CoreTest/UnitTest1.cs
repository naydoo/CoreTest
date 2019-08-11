using NUnit.Framework;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System.IO;
using System.Reflection;

namespace Tests
{
    public class Tests
    {
        public IWebDriver driver = new ChromeDriver(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location));
        [SetUp]
        public void Setup()
        {
            driver.Navigate().GoToUrl("https://google.com");
        }

        [Test]
        public void Test1()
        {
            driver.FindElement(By.Name("q")).SendKeys("test");
            driver.Close();
            driver.Quit();
        }
    }
}