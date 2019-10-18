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

namespace AllokateCJV
{
    [TestFixture]
    public class ReportComparatorTests : BasePage<HomePage>
    {
        public static bool DayA = true;
        private static readonly TestGenerator Tests = new TestGenerator(Paths.TestCasesNew(@"InputData/"));
        private static readonly TestGenerator EydsTests = new TestGenerator(Paths.TestCasesNew(@"Input-EYDS/"));

        private static readonly IEnumerable<TestCaseData> ReviewCalcTestsData = Tests.RetrieveReviewCalcsTests();
        [Test, TestCaseSource(nameof(ReviewCalcTestsData))]
        [Category("ProdReview&Calcs")]
        public void ReviewAndCalculation(ReviewTestCase testData)
        {
            var ReviewPage = InitialPage.GoToProjectLink(testData.ProjectLink)
                .GoToCalculationPage().CancelActiveJob()
                .GoToReviewPage().ExecuteReviewProcess();

            Assert.That(ReviewPage.IsNoDiscrepanciesFoundMessageDisplayed());

            var CalculationPage = ReviewPage.GoToCalculationPage()
                .ExecuteCalculationProcess(testData.StartPeriod, testData.StartAtAsset, testData.StopPeriod, testData.StopAtAsset);
            Assert.That(CalculationPage.GetCalculationResultMessage().Equals("Calculation Completed Successfully"));
        }

        private static readonly IEnumerable<TestCaseData> CompareTestsData = Tests.RetrieveReportTests();
        [Test, TestCaseSource(nameof(CompareTestsData))]
        [Category("ProdCompareReports")]
        public void CompareReport(CompareTestCase testData)
        {
            string goldenCopyPath = Paths.GetGoldenCopyPath(testData.ProjectName,
                                                            testData.GoldenCopyPath);

            TestContext.WriteLine("1st spreadsheet is the golden copy of report");
            TestContext.AddTestAttachment(goldenCopyPath, "Golden copy of report");

            var DownloadReportModal = InitialPage.GoToProjectLink(testData.ProjectLink)
                .GoToReportsPage().SelectReport(testData.ReportGroup, testData.ReportName);
            string downloadedReport = DownloadReportModal
                .GenerateReport(testData.FullYear, testData.YpeNUM, testData.BooksetType, testData.LayerType)
                .DownloadReport();

            TestContext.WriteLine("2nd spreadsheet is the downloaded report");
            TestContext.AddTestAttachment(downloadedReport, "Downloaded report");

            var reporteCompartor = new ExcelReportComparator(goldenCopyPath);
            Assert.That(reporteCompartor.DevopsCompareExcelReport(downloadedReport),
                "Downloaded report and golden copy should have equal data");
        }

        private static readonly IEnumerable<TestCaseData> EYDSTestsData = EydsTests.RetrieveEYDSReviewCalcsTests();
        [Test, TestCaseSource(nameof(EYDSTestsData))]
        [Category("ProdEYDS")]
        public void EYDSDailyRun(EydsReviewTestCase EYDSTestsData)
        {
            Settings.Timeouts.Process = Settings.Eyds.ProcessTime;

            var purchasesPage = InitialPage.GoToProjectLink(EYDSTestsData.ProjectLink)
                .GoToCalculationPage().CancelActiveJob().GoToAssetsPurchasePage();
            DayA = purchasesPage.UpdateValues();
            TestContext.WriteLine("Set of execution day.{1}True means DayA files will be used, and False means DayB files.{1}" +
                "Current day value is: {0}", DayA.ToString(), System.Environment.NewLine);

            var reviewPage = purchasesPage.GoToProjectMenu().GoToReviewPage().ExecuteReReviewProcess().GetReviewProcessTime();
            string reviewTime = TimeHelper.FormatTimeForEmail(reviewPage.ReviewProcessTime);

            Assert.That(reviewPage.IsNoDiscrepanciesFoundMessageDisplayed());

            var calculationPage = reviewPage.GoToCalculationPage().ExecuteCalculationForEYDS().GetCalculationProcessTime();
            string calculationTime = TimeHelper.FormatTimeForEmail(calculationPage.CalculationProcessTime);
            string downloadedLog = calculationPage.DownloadProcessLog(Settings.DownloadPath);

            string flowResultString = string.Empty;

            if (TimeHelper.IsProcessSlow(reviewPage.ReviewProcessTime, Settings.Eyds.ReviewMaxProcessTime) ||
            TimeHelper.IsProcessSlow(calculationPage.CalculationProcessTime, Settings.Eyds.CalculationMaxProcessTime))
            {
                flowResultString += "Slow";
                TxtCreator.SlowProcessMsg(EYDSTestsData.ProjectLink, reviewTime, calculationTime);
                TestContext.AddTestAttachment(TxtCreator.txtFullPath, "Text file for email body");
            }

            TestContext.WriteLine("Log file of the Calculation process");
            TestContext.AddTestAttachment(downloadedLog, "Log file of calculation");

            Assert.That(calculationPage.GetCalculationResultMessage().Equals("Calculation Completed Successfully"));

            IEnumerable<EydsCompareTestCase> CompareEYDSTests = EydsTests.RetrieveEYDSReportTests(EYDSTestsData.ProjectName);

            var reportsPage = calculationPage.GoToReportsPage();

            Exception exception = null;

            foreach (EydsCompareTestCase testData in CompareEYDSTests)
            {
                LogHelpers.Write(string.Format("Work with {0} report.",testData.ReportName));
                string goldenCopyPath = Paths.GetGoldenCopyPath(testData.ProjectName,
                                                    testData.GoldenCopyPath);
                TestContext.WriteLine("Running comparison for " + testData.BooksetType + " Book Set.");
                TestContext.WriteLine("1st spreadsheet is the golden copy of report");
                TestContext.AddTestAttachment(goldenCopyPath, "Golden copy of report");

                var DownloadReportModal = reportsPage.SelectReport(testData.ReportGroup, testData.ReportName);
                string downloadedReport = DownloadReportModal
                    .GenerateReport(testData.BooksetType).DownloadReport();

                TestContext.WriteLine("2nd spreadsheet is the downloaded report");
                TestContext.AddTestAttachment(downloadedReport, "Downloaded report");

                var reportComparator = new ExcelReportComparator(goldenCopyPath);

                try
                {
                    flowResultString += reportComparator.DevopsCompareExcelReport(downloadedReport).ToString();
                }
                catch (COMException ex)
                {
                    flowResultString += "Exception";
                    exception = ex;
                }
                catch (FormatException ex)
                {
                    flowResultString += "Exception";
                    exception = ex;
                }
            }

            string errorMessage = string.Empty;
            bool result = true;

            if (flowResultString.Contains("Slow"))
            {
                result = false;
                errorMessage = "One of the processes exceeded the allowed time.";
            }
            else if (flowResultString.Contains("Exception"))
            {
                result = false;
                errorMessage = exception.Message;
            }
            else if (flowResultString.Contains(false.ToString()))
            {
                result = false;
                errorMessage = "Downloaded report and golden copy should have equal data";
            }

            //TODO: File is always created. Is a validation needed if test failed for exceptions
            //or even false comparation?
            TxtCreator.EndEydsRun(EYDSTestsData.ProjectLink, reviewTime, calculationTime);
            TestContext.AddTestAttachment(TxtCreator.txtFullPath, "Text file for email body");

            Assert.That(result, errorMessage);
        }


    }
}