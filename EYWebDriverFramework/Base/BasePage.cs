using EYWebDriverFramework.Config;
using EYWebDriverFramework.Selenium;
using EYWebDriverFramework.Utils;
using NUnit.Framework;
using NUnit.Framework.Interfaces;
using OpenQA.Selenium;
using System;
using System.Configuration;
using System.Globalization;
using System.IO;
using System.Threading;

namespace EYWebDriverFramework.Base
{
    public abstract class BasePage<T> where T : WebDriverBaseAbstractPageObject
    {
        public T _initialPage;

        protected T InitialPage
        {
            get
            {
                return _initialPage;
            }
        }

        protected virtual void TakeScreenshot(string testResultFile = "")
        {            
            Screenshot screenshot = ((ITakesScreenshot)Driver.Instance).GetScreenshot();
            string filePath = string.Format(ConfigReader.TestResultsPath + "{0}.png", LogHelpers.NowFileName());

            screenshot.SaveAsFile(filePath, ScreenshotImageFormat.Png);
            TestContext.AddTestAttachment(filePath, "Screenshot on failure or error");
            if (testResultFile != null)
            {
                ConfigReader.reportGenerated = null;
            }
        }

        //With this method a log file is created and InlinePowershell task fails if it's found
        protected virtual void CreateFailedFlagFile()
        {
            StreamWriter flagFile = File.AppendText(Settings.TestResultsPath + @"\FailedFlag.log");
            flagFile.Close();
        }

        [SetUp]
        protected virtual void SetUp()
        {
            InitializeSettings();
            CultureInfo currentCulture = new CultureInfo("en-Us");
            Thread.CurrentThread.CurrentCulture = currentCulture;
            
            Uri DefaultBaseUrl = new Uri(ConfigurationManager.AppSettings["Url"]);
            string CurrentTestName = TestContext.CurrentContext.Test.Name;
            LogHelpers.Write(string.Format("Executing Test: {0}", CurrentTestName));
            LogHelpers.Write(string.Format("Navigated to site: {0}", DefaultBaseUrl));
            Driver.Get();
            Driver.Instance.Navigate().GoToUrl(DefaultBaseUrl);
            _initialPage = (T)Activator.CreateInstance(typeof(T), Driver.Instance);
        }

        [TearDown]
        protected virtual void TearDown()
        {
            if (TestContext.CurrentContext.Result.Outcome.Status == TestStatus.Failed)
            {
                CreateFailedFlagFile();
                TakeScreenshot(ConfigReader.reportGenerated);
                LogHelpers.Write(string.Format("FAILED execution for test: {0}. Error Message: {1}", TestContext.CurrentContext.Test.Name, TestContext.CurrentContext.Result.Message));
            }
            else
            {
                LogHelpers.Write(string.Format("SUCCESS execution for test: {0}.", TestContext.CurrentContext.Test.Name));
            }
            LogHelpers.CloseLogFile();
            InitialPage.Quit();
        }

        public static void InitializeSettings()
        {
            // Set all the settings for the frame work
            ConfigReader.SetFrameworkSettings(TestContext.CurrentContext.Test.Name);
            // Set log
            LogHelpers.CreateLogFile(TestContext.CurrentContext.Test.Name);
        }
    }
}
