using AllokateCJV.Pages.Elements;
using AllokateCJV.Pages.Modals;
using EYWebDriverFramework.Selenium;
using EYWebDriverFramework.Utils;
using OpenQA.Selenium;

namespace AllokateCJV.Pages
{
    public class ReportsPage: WebDriverBaseAbstractPageObject
    {
        #region Locators
        // Left Menu
        private ReportElements LeftMenuButton => new ReportElements(WebDriver, By.XPath("//div[@class='left-content']"));
        private ReportElements GenerateReportButton => new ReportElements(WebDriver, By.XPath("//div[@class='right-contents']"));

        #endregion

        public ReportsPage(IWebDriver driver)
            : base(driver)
        {
        }

        #region Actions
        public DownloadReportModal SelectReport(string menu, string report)
        {
            LogHelpers.Write(string.Format("Click on button category \"{0}\" on left menu to display reports.", menu.Replace(":","")));
            LeftMenuButton.GetLeftMenuElement(menu).ScrollAndClick();
            LogHelpers.Write(string.Format("Click \"{0}\" report's download button.", report));
            GenerateReportButton.GetDownloadReportElement(report).ScrollAndClick();
            return NewPage<DownloadReportModal>();
        }
        #endregion
    }
}
