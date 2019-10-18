using OpenQA.Selenium;
using System;

namespace EYWebDriverFramework.Selenium
{
    public class WebDriverBaseAbstractPageObject
    {
        protected readonly IWebDriver WebDriver;

        protected WebDriverBaseAbstractPageObject()
        {
            // Wait 2 seconds for page to become stale.
            Waits.ForExtJsAjax();
            Waits.WaitForAngular();
        }
        protected WebDriverBaseAbstractPageObject(IWebDriver driver)
        {
            WebDriver = driver;
        }

        public void Navigate(string url)
        {
            Driver.Instance.Navigate().GoToUrl(url);
            
        }

        protected internal T NewPage<T>() where T : WebDriverBaseAbstractPageObject
        {
            return (T)Activator.CreateInstance(typeof(T), Driver.Instance);

        }

        public void Quit()
        {
            Driver.Instance.Quit();
        }

        protected void RunScript(string script, params object[] args)
        {
            var jsExecutor = (IJavaScriptExecutor)Driver.Instance;
            jsExecutor.ExecuteScript(script, args);
        }


    }
}
