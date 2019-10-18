using EYWebDriverFramework.Selenium;
using EYWebDriverFramework.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;
using System;
using System.Linq;

namespace AllokateCJV.Pages.Modals
{
    public class DownloadReportModal: WebDriverBaseAbstractPageObject
    {
        #region Locators
        IWebElement BookSetSelect => WebDriver.FindElement(By.XPath("//div[@class='bookset-option']//select"));
        IWebElement LayerTypeSelect => WebDriver.FindElement(By.XPath("//div[@class='layer-types']")).FindElement(By.XPath("./..")).FindElement(By.TagName("select"));
        IWebElement EventOrPeriodButton => WebDriver.FindElement(By.ClassName("duration-scope-event-period"));
        IWebElement FullTaxYearButton => WebDriver.FindElement(By.ClassName("duration-scope-full-year"));
        IWebElement ScopeSelector => WebDriver.FindElements(By.ClassName("selected-duration-scope")).FirstOrDefault();
        IWebElement SelectEventsorPeriodForReportButton(int period) => WebDriver.FindElements(By.ClassName("event-period-row"))[period];
        IWebElement GenerateAndDownloadReportButton => WebDriver.FindElement(By.ClassName("custom-button"));
        IWebElement OkButton => WebDriver.FindElement(By.ClassName("custom-button"));

        #endregion

        public DownloadReportModal(IWebDriver driver)
            : base(driver)
        {
        }

        public ReportQueueModal GenerateReport(string bookset)
        {
            if (!string.IsNullOrEmpty(bookset))
            {
                var selectElement = new SelectElement(BookSetSelect);
                LogHelpers.Write(string.Format("Select \"{0}\" as Book Set.", bookset));
                selectElement.SelectByText(bookset);
            }
            LogHelpers.Write(string.Format("Click \"Download report\" button."));
            GenerateAndDownloadReportButton.ClickAndWaitForAjax();

            LogHelpers.Write(string.Format("Click \"Ok\" button."));
            OkButton.ClickAndWaitForAjax();
            return NewPage<ReportQueueModal>();
        }

        public ReportQueueModal GenerateReport(string fullYear, int period, string bookset = null, string layerType = null)
        {
            if (!string.IsNullOrEmpty(bookset))
            {
                var selectElement = new SelectElement(BookSetSelect);
                LogHelpers.Write(string.Format("Select \"{0}\" as Book Set.", bookset));
                selectElement.SelectByText(bookset);
            }
            if (!string.IsNullOrEmpty(layerType))
            {
                var selectElement = new SelectElement(LayerTypeSelect);
                LogHelpers.Write(string.Format("Select \"{0}\" as Layer Type.", layerType));
                selectElement.SelectByText(layerType);
            }
            if(!(ScopeSelector == null))
            {
                SelectScopeYearOrPeriod(fullYear);
            }
            if (fullYear.ToLower() == "false")
            {
                LogHelpers.Write(string.Format("Select \"{0}\" as period.", period));
                SelectEventsorPeriodForReportButton(period).ClickAndWaitForAjax();
            }
            if (!period.Equals(-1) || !string.IsNullOrEmpty(layerType) || !string.IsNullOrEmpty(bookset) || !string.IsNullOrEmpty(fullYear))
            {
                LogHelpers.Write(string.Format("Click \"Download report\" button."));
                GenerateAndDownloadReportButton.ClickAndWaitForAjax();
            }
            LogHelpers.Write(string.Format("Click \"Ok\" button."));
            OkButton.ClickAndWaitForAjax();
            return NewPage<ReportQueueModal>();
        }

        private void SelectScopeYearOrPeriod(string fullYear)
        {
            if (fullYear.ToLower() == "true")
                ClickFullTaxYearButton();
            else
                ClickEventOrPeriodButton();
        }

        public void ClickFullTaxYearButton()
        {
            LogHelpers.Write(string.Format("Click \"Full Tax Year\" button."));
            FullTaxYearButton.ClickAndWaitForAjax();
        }
        public void ClickEventOrPeriodButton()
        {
            LogHelpers.Write(string.Format("Click \"Event or Period\" as button."));
            EventOrPeriodButton.ClickAndWaitForAjax();
        }
    }
}
