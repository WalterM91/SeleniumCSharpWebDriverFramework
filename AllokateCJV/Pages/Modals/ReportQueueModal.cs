using EYWebDriverFramework.Selenium;
using EYWebDriverFramework.Config;
using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using EYWebDriverFramework.Utils;

namespace AllokateCJV.Pages.Modals
{
    public class ReportQueueModal : WebDriverBaseAbstractPageObject
    {
        #region Locators
        IWebElement UserName => WebDriver.FindElement(By.ClassName("user-name"));
        IWebElement DownloadReportButton => WebDriver.FindElement(By.XPath("(//div[contains(text(), '" + UserName.Text + "')]/../../../../../div[contains(@class,'report-job-detail')])[last()]/div[contains(@class,'download-reports-button')]"));
        IWebElement CloseQueueButton => WebDriver.FindElement(By.XPath("(//div[contains(text(), '" + UserName.Text + "')]/../../../../../div[contains(@class,'report-job-detail')])[last()]//div[contains(@class,'report-job-close-icon')]"));
        IWebElement ProgressLabel => WebDriver.FindElement(By.XPath("(//div[contains(text(), '" + UserName.Text + "')]/../../../../../div[contains(@class,'report-job-detail')])[last()]//div[contains(@class,'progress-text')]"));
        #endregion
        public ReportQueueModal(IWebDriver driver)
            : base(driver)
        {
        }

        public string DownloadReport()
        {
            WaitForReportGeneration();
            LogHelpers.Write(string.Format("Click \"Download Report\" button."));
            DownloadReportButton.ClickJS();
            Waits.WaitDownloadDocument();
            LogHelpers.Write(string.Format("Click \"Close Queue\" button."));
            Waits.WaitUntilElementPresent(drv => CloseQueueButton).ClickAndWaitForAjax();
            var directory = new DirectoryInfo(Settings.DownloadPath);
            var myFile = directory.GetFiles().Where(f => !f.Attributes.HasFlag(FileAttributes.Hidden)).OrderByDescending(f => f.LastWriteTime).First();
            return myFile.FullName;
        }

        private void WaitForReportGeneration()
        {
            Waits.GetWait().Until(drv => {
                try
                {
                    return ProgressLabel == null;
                }
                catch (NotFoundException)
                {
                    return true;
                }
            });

        }
    }
}
