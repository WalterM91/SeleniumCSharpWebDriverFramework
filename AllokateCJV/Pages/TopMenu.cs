using EYWebDriverFramework.Selenium;
using EYWebDriverFramework.Utils;
using OpenQA.Selenium;
using System.Diagnostics;

namespace AllokateCJV.Pages
{
    public class TopMenu : WebDriverBaseAbstractPageObject
    {
        #region Locators
        IWebElement ReviewMenuButton => WebDriver.FindElement(By.CssSelector("#projectnavigation-container .sub-menu-section > div > div:nth-of-type(3)"));
        IWebElement CalculationMenuButton => WebDriver.FindElement(By.XPath("//div[@class='sub-menu-section']/div/div[contains(string(), 'Calculation')]"));
        public IWebElement NavigationButton => WebDriver.FindElement(By.ClassName("navigate-icon"));
        IWebElement ReportsLink => WebDriver.FindElement(By.PartialLinkText("Reports"));

        public IWebElement ProjectLink => WebDriver.FindElement(By.XPath("//a[@class='nav-link-part']/div[contains(string(), 'Project')]"));
        public IWebElement ReviewLink => WebDriver.FindElement(By.PartialLinkText("Review"));
        public IWebElement AssetsLink => WebDriver.FindElement(By.PartialLinkText("Assets"));
        public IWebElement PurchasesButton =>
            WebDriver.FindElement(By.XPath("//div[@class='sub-menu-section']/div/div[contains(string(), 'Purchases')]"));
        #endregion

        public TopMenu(IWebDriver driver)
            : base(driver)
        {
        }

        #region Actions

        public TopMenu GoToProjectMenu()
        {
            LogHelpers.Write(string.Format("Access top Menu."));
            NavigationButton.ClickAndWaitForAjax();
            Waits.GetWait().Until(drv => ProjectLink);
            LogHelpers.Write(string.Format("Click \"Project\" link in top Menu."));
            ProjectLink.ClickAndWaitForAjax();
            Waits.GetWait().Until(drv => ReviewMenuButton);
            return this;
        }


        public ReviewPage GoToReviewPage()
        {
            LogHelpers.Write(string.Format("Click \"Review\" button in sub menu."));
            ReviewMenuButton.ClickAndWaitForAjax();
            return NewPage<ReviewPage>();
        }
        
        public CalculationPage GoToCalculationPage()
        {
            LogHelpers.Write(string.Format("Click \"Calculation\" button in sub menu."));
            CalculationMenuButton.ClickAndWaitForAjax();
            return NewPage<CalculationPage>();
        }

        public AssetsPurchasesPage GoToAssetsPurchasePage()
        {
            LogHelpers.Write(string.Format("Access top menu."));
            NavigationButton.ClickAndWaitForAjax();
            Waits.GetWait().Until(drv => AssetsLink);

            LogHelpers.Write(string.Format("Click \"Assets\" link in top menu."));
            AssetsLink.ClickAndWaitForAjax();
            Waits.GetWait().Until(drv => PurchasesButton);

            LogHelpers.Write(string.Format("Click \"Purchases\" button in sub menu."));
            PurchasesButton.ClickAndWaitForAjax();
            return NewPage<AssetsPurchasesPage>();
        }
        public ReportsPage GoToReportsPage()
        {
            LogHelpers.Write(string.Format("Access top menu."));
            NavigationButton.ClickAndWaitForAjax();
            Waits.GetWait().Until(drv => ReportsLink);

            LogHelpers.Write(string.Format("Click \"Reports\" link in top menu."));
            ReportsLink.ClickAndWaitForAjax();
            return NewPage<ReportsPage>();
        }

        #endregion

    }
}
