using EYWebDriverFramework.Selenium;
using EYWebDriverFramework.Utils;
using OpenQA.Selenium;
using System;
using System.IO;
using System.Linq;

namespace AllokateCJV.Pages
{
    public class CalculationPage : TopMenu
    {
        #region Locators
        IWebElement StartPeriodSelect => WebDriver.FindElement(By.ClassName("start-at-period-select"));
        IWebElement StartAssetSelect => WebDriver.FindElement(By.ClassName("start-at-asset-select"));
        IWebElement StopPeriodSelect => WebDriver.FindElement(By.ClassName("stop-at-period-select"));
        IWebElement StopAssetSelect => WebDriver.FindElement(By.ClassName("stop-at-asset-select"));
        IWebElement BeginCalculationButton => WebDriver.FindElement(By.ClassName("begin-calculation-button"));
        IWebElement CalculationProgress => WebDriver.FindElement(By.ClassName("calculating-progress"));
        IWebElement CalculationResult => WebDriver.FindElement(By.ClassName("calculating-completed-progress"));
        IWebElement CalculationTimeElement =>
            WebDriver.FindElement(By.CssSelector(".calculate-summary-section.float-left .time-elapsed-heading.float-left"));
        IWebElement CancelJob => WebDriver.FindElements(By.ClassName("cancel-job-button")).FirstOrDefault();
        IWebElement DownloadLog => WebDriver.FindElement(By.CssSelector(".float-left.label-cursor"));
        IWebElement ForceReRunContinueButton => WebDriver.FindElement(By.ClassName("ype-button-modal"));
        IWebElement RegularTaxCalcDoneCheck =>
            WebDriver.FindElement(By.CssSelector(".float-left.regulartax-bar .ype-calc-status-line:nth-last-child(1) .ype-ending-circle .calculated-tick"));
        IWebElement AmtCalcDoneCheck => 
            WebDriver.FindElement(By.CssSelector(".float-left.amt-bar .ype-calc-status-line:nth-last-child(1) .ype-ending-circle .calculated-tick"));
        IWebElement EandPCalcDoneCheck => 
            WebDriver.FindElement(By.CssSelector(".float-left.eandp-bar .ype-calc-status-line:nth-last-child(1) .ype-ending-circle .calculated-tick"));
        IWebElement StateCalcDoneCheck =>
            WebDriver.FindElement(By.CssSelector(".float-left.state-bar .ype-calc-status-line:nth-last-child(1) .ype-ending-circle .calculated-tick"));
        
        #endregion
        #region Constructor
        public CalculationPage(IWebDriver driver)
            : base(driver)
        {


        }
        #endregion
        #region Properties
            public DateTime CalculationProcessTime { get; set; }
        #endregion
        #region Actions

        public CalculationPage ExecuteCalculationProcess(string startPeriodValue, string startAssetValue, string stopPeriodValue, string stopAssetValue)
        {
            SetStartPeriodValue(startPeriodValue);
            SetStopPeriodValue(stopPeriodValue);
            SetStartAssetValue(startAssetValue);
            SetStopAssetValue(stopAssetValue);
            return ClickBeginCalculationsButton();
        }

        public CalculationPage ExecuteCalculationForEYDS()
        {
            LogHelpers.Write(string.Format("Click \"Begin calculation\" button."));
            BeginCalculationButton.ClickJS();
            Waits.WaitUntilElementPresent(drv => ForceReRunContinueButton);

            LogHelpers.Write(string.Format("Click \"Force Re-Run\" button."));
            ForceReRunContinueButton.ClickAndWaitForAjax();
            Waits.WaitForProcessToFinish(drv => CalculationProgress);

            WaitForBackendCalcEnd();
            return this;
        }

        public CalculationPage GetCalculationProcessTime()
        {
            LogHelpers.Write(string.Format("Read \"Calculation process\" time."));
            Waits.WaitUntilElementPresent(drv => CalculationTimeElement);
            string elementValue = CalculationTimeElement.Read();
            string replacedValue = elementValue.Replace("Time Elapsed : ", "");
            string[] timeValues = replacedValue.Split(':');
            if (timeValues.Length == 1)
            {
                CalculationProcessTime = new DateTime(1, 1, 1, 0, 0, int.Parse(timeValues[0]));
            }
            else
            {
                CalculationProcessTime = new DateTime(1, 1, 1, 1, int.Parse(timeValues[0]), int.Parse(timeValues[1]));
            }
            return this;
        }

        public void WaitForBackendCalcEnd()
        {
            /*The elapsed time is updated with a Backend process, this occurs after
             *a Process completed" label appeared. Four elements with a green check
             *icon are also searched for since it's one of the visual changes that
             *ocurr in the page prior to the time being updated.*/

            LogHelpers.Write(string.Format("Wait for \"Calculation\" process to finish."));
            int backEndSuggestedTime = 60;

            Waits.WaitForProcessToFinish(drv => RegularTaxCalcDoneCheck);
            Waits.WaitForProcessToFinish(drv => AmtCalcDoneCheck);
            Waits.WaitForProcessToFinish(drv => EandPCalcDoneCheck);
            Waits.WaitForProcessToFinish(drv => StateCalcDoneCheck);
            try
            {
                Waits.WaitForElementIfChanges(drv => CalculationTimeElement, backEndSuggestedTime);
            }
            catch (WebDriverTimeoutException)
            {
               /*It's not sure if the data will always change after "Process completed"
                *label appeared, how many seconds after that it will occur or even if it
                *can be updated before because of a fast Calculation and a slow recognition
                *of the check elements. An until method is used for it to exit the wait,
                *but since that causes an exception I setted an empty catch.*/
            }
        }

        public void SetStartPeriodValue(string value)
        {
            LogHelpers.Write(string.Format("Select \"{0}\" option as Start period.", value));
            StartPeriodSelect.SelectByText(value);
        }
        public void SetStartAssetValue(string value)
        {
            LogHelpers.Write(string.Format("Select \"{0}\" option as Start asset.", value));
            StartAssetSelect.SelectByText(value);
        }
        public void SetStopPeriodValue(string value)
        {
            LogHelpers.Write(string.Format("Select \"{0}\" option as Stop period.", value));
            StopPeriodSelect.SelectByText(value);
        }
        public void SetStopAssetValue(string value)
        {
            LogHelpers.Write(string.Format("Select \"{0}\" option as Stop asset.", value));
            StopAssetSelect.SelectByText(value);
        }
        public string DownloadProcessLog(string downloadPath)
        {
            LogHelpers.Write(string.Format("Click \"Download log\" button."));
            DownloadLog.ClickJS();
            Waits.WaitDownloadDocument();
            //check if there is some window or something
            var directory = new DirectoryInfo(downloadPath);
            var myFile = directory.GetFiles().Where(f => !f.Attributes.HasFlag(FileAttributes.Hidden)).OrderByDescending(f => f.LastWriteTime).First();
            return myFile.FullName;

        }



        public CalculationPage ClickBeginCalculationsButton()
        {
            LogHelpers.Write(string.Format("Click \"Begin calculation\" button."));
            BeginCalculationButton.ClickJS();
            Waits.WaitForProcessToFinish(drv => CalculationProgress);
            return this;
        }

        public bool IsNoDiscrepanciesFoundMessageDisplayed()
        {
            return CalculationResult.IsDisplayed();
        }
        public string GetCalculationResultMessage()
        {
            return CalculationResult.Text;
        }

        public CalculationPage CancelActiveJob()
        {
            if (CancelJob != null)
            {
                LogHelpers.Write(string.Format("Cancel active job."));
                CancelJob.ClickAndWaitForAjax();
                Waits.WaitUntilElementPresent(drv => BeginCalculationButton);
            }
            return this;
        }

        #endregion
    }
}
