using EYWebDriverFramework.Selenium;
using OpenQA.Selenium;

namespace AllokateCJV.Pages.Elements
{
    public class ReportElements: WebDriverBaseAbstractPageObject
    {
        private readonly By _locator;

        private IWebElement _element;
        private readonly IWebDriver _webDriver;

        public ReportElements(IWebDriver webDriver, By locator)
        {
            _locator = locator;
            _webDriver = webDriver;
        }

        public IWebElement Context
        {
            get
            {
                return _element ?? (_element = _webDriver.FindElement(_locator));
            }
        }

        public IWebElement GetLeftMenuElement(string variable)
        {
            IWebElement element = Context.FindElement(By.XPath(string.Format("./div/div[contains(text(),'{0}')]", variable)));
            element.ScrollTo();
            return element;
        }

        public IWebElement GetDownloadReportElement(string variable)
        {
            IWebElement element = Context.FindElement(By.XPath(string.Format("./div/div/div/div[contains(text(),'{0}')]/parent::*/following-sibling::div/div/div[@class='download-package-button']", variable)));
            return element;
        }
    }
}
