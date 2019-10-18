using System.Collections.Generic;
using AllokateCJV.Pages;
using EYWebDriverFramework.Config;
using AllokateCJV.Tests;
using EYWebDriverFramework.Base;
using EYWebDriverFramework.Utils;
using NUnit.Framework;
using System;
using System.Runtime.InteropServices;
using System.IO;
using System.Globalization;
using System.Threading;
using NUnit.Framework.Interfaces;

namespace AllokateCJV
{
    /// <summary>
    /// This class was made to test comparison without accessing AlloKate and without having to edit
    /// the BasePage class in order to make it browserless.
    /// It uses a fixed golden copy file's path to fake a bad comparison so adding more files was not
    /// necessary, it also copies it to a new in case the test whose file is originally from is being
    /// run and the test will pass instead of failing because of file being used.
    /// </summary>
    [TestFixture]
    public class BrowserLessTests
    {
        private void CreateFailedFlagFile()
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
            string CurrentTestName = TestContext.CurrentContext.Test.Name;
            LogHelpers.Write(string.Format("Executing Test: {0}", CurrentTestName));
        }

        [TearDown]
        protected virtual void TearDown()
        {
            if (TestContext.CurrentContext.Result.Outcome.Status == TestStatus.Failed)
            {
                CreateFailedFlagFile();
                LogHelpers.Write(string.Format("FAILED execution for test: {0}. Error Message: {1}", TestContext.CurrentContext.Test.Name, TestContext.CurrentContext.Result.Message));
            }
            else
            {
                LogHelpers.Write(string.Format("SUCCESS execution for test: {0}.", TestContext.CurrentContext.Test.Name));
            }
            LogHelpers.CloseLogFile();
        }

        private static void InitializeSettings()
        {
            ConfigReader.SetFrameworkSettings(TestContext.CurrentContext.Test.Name);
            LogHelpers.CreateLogFile(TestContext.CurrentContext.Test.Name);
        }

        private static readonly TestGenerator TestsList = new TestGenerator(Paths.TestCasesNew(@"InputData/"));
        private static readonly IEnumerable<TestCaseData> BrowserlessTestData = TestsList.RetrieveReportTests();

        /// <summary>
        /// It takes the same gonlden file for good comparison but makes a copy of it before doing it
        /// to avoid "File being used by another proccess" error.
        /// </summary>
        /// <param name="testData">It uses the same variable as the normal CompareTest method.</param>
        [Test, TestCaseSource(nameof(BrowserlessTestData))]
        [Category("BrowserlessTestCompareReports")]
        public void BrowserlessCompareReport(CompareTestCase testData)
        {
            //GOOD PATH
            string goldenCopyPath = Paths.GetGoldenCopyPath(testData.ProjectName,
                                            testData.GoldenCopyPath);

            TestContext.WriteLine("1st spreadsheet is the golden copy of report");
            TestContext.AddTestAttachment(goldenCopyPath, "Golden copy of report");

            string fakeDownloadedReport = "C:\\Users\\GK122LX\\Desktop\\FrameworkEY_cloned\\InputData\\R1.11.3\\SS_R1.11.3_Smoketest_Remedial_FIFO\\Golden Copies\\Set_1_6H_CostRecoveryPartnersBSI_GOLDEN-Bad.xlsm";

            //string fakeDownloadedReport = Path.Combine(Settings.DownloadPath, badCopy.Substring(0, badCopy.Length - 5) + " - FakeDownloadedReport.xlsm");
                
            //File.Copy(goldenCopyPath, fakeDownloadedReport);

            TestContext.WriteLine("2nd spreadsheet is the downloaded report");
            TestContext.AddTestAttachment(fakeDownloadedReport, "Downloaded report");

            var reporteCompartor = new ExcelReportComparator(goldenCopyPath);

            Assert.That(reporteCompartor.DevopsCompareExcelReport(fakeDownloadedReport),
                "Downloaded report and golden copy should have equal data");
        }

    }
}