using EYWebDriverFramework.Selenium;
using EYWebDriverFramework.Utils;
using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;

namespace AllokateCJV.Pages
{
    public class ReviewPage : TopMenu
    {
        #region Locators
        private IWebElement ReviewButton => WebDriver.FindElement(By.XPath("//div[@class='asset-review-button float-left' and contains(text(),'Review')]"));
        private IWebElement ReReviewButton => WebDriver.FindElement(By.CssSelector(".rereview-button-section.float-left .asset-re-review-button.float-left"));
        private IWebElement DiagnosticResult => WebDriver.FindElement(By.ClassName("diagnostics-message-main"));
        private IWebElement ReviewProgress => WebDriver.FindElement(By.ClassName("review-progress-container"));
        private ReadOnlyCollection<IWebElement> ReviewTimeElements => 
            WebDriver.FindElements(By.CssSelector(".review-status-msg:last-child .initiated-by-value"));
        #endregion
        #region Properties
        public DateTime ReviewProcessTime { get; set; }
        #endregion

        public ReviewPage(IWebDriver driver)
            : base(driver)
        {
        }

        public ReviewPage ExecuteReReviewProcess()
        {
            Waits.WaitUntilElementPresent(drv => ReReviewButton);
            LogHelpers.Write(string.Format("Click \"Re-Review\" button."));
            ReReviewButton.ClickJS();
            Waits.WaitUntilElementPresent(drv => ReviewProgress);
            LogHelpers.Write(string.Format("Wait for \"Review process\" to finish..."));
            Waits.WaitForProcessToFinish(drv => ReviewProgress);
            return this;
        }



        public ReviewPage GetReviewProcessTime()
        {
            LogHelpers.Write(string.Format("Get \"Review process\" time."));
            /*'02 min' '17 sec' there is a WebElement for each part of the time*/

            #region TODO: Need to check if Waits.WaitUntilElementPresent is enough
            //ISSUE: Random behavior => ReviewTimeElements are read by the framework but Read()
            //returns empty and "Data is not on valid a string format" error is displayed. 
            //POSSIBLE REASON: When the browser is not maximized it returns empty indeed.
            //Not sure if that's the only reason.
            #endregion
            Waits.WaitUntilElementPresent(drv => ReviewTimeElements[0]);

            List<int> timeData = new List<int>();
            string numberPartData;
            foreach (IWebElement item in ReviewTimeElements)
            {
                numberPartData = string.Empty;
                foreach (char chr in item.Read().ToCharArray())
                {
                    if (chr.IsANumber())
                    {
                        numberPartData += chr;
                    }
                }
                timeData.Add(Convert.ToInt32(numberPartData));
            }
            if (timeData.Count == 1)
            {
                ReviewProcessTime = new DateTime(1, 1, 1, 0, 0, timeData[0]);
            }
            else
            {
                ReviewProcessTime = new DateTime(1, 1, 1, 1, timeData[0], timeData[1]);
            }
            return this;
        }

        public ReviewPage ExecuteReviewProcess()
        {
            
            Waits.WaitUntilElementPresent(drv => ReviewButton);
            LogHelpers.Write(string.Format("Click \"Review\" button."));
            ReviewButton.ClickJS();
            LogHelpers.Write(string.Format("Wait for \"Review process\" to finish..."));
            Waits.WaitUntilElementPresent(drv => ReviewProgress);
            Waits.WaitForProcessToFinish(drv => ReviewProgress);
            return this;
        }
        public bool IsNoDiscrepanciesFoundMessageDisplayed()
        {
            LogHelpers.Write(string.Format("Check if \"No discrepancies found\" message is displayed."));
            return DiagnosticResult.IsDisplayed();
        }
    }
}
