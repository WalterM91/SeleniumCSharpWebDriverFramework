using EYWebDriverFramework.Config;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;
using System;
using System.IO;
using System.Linq;

namespace EYWebDriverFramework.Selenium
{
    public static class Waits
    {
        static IClock clock = new SystemClock();

        private static WebDriverWait _explicitWait => new WebDriverWait(
            clock,
            Driver.Instance,
            TimeSpan.FromSeconds(Settings.Timeouts.Explicit),
            TimeSpan.FromMilliseconds(Settings.Timeouts.SleepIntervalInMillis));

        private static WebDriverWait _implicitWait => new WebDriverWait(
            clock,
            Driver.Instance,
            TimeSpan.FromSeconds(Settings.Timeouts.Implicit),
            TimeSpan.FromMilliseconds(Settings.Timeouts.SleepIntervalInMillis));

        private static WebDriverWait _processWait => new WebDriverWait(
            clock,
            Driver.Instance,
            TimeSpan.FromSeconds(Settings.Timeouts.Process),
            TimeSpan.FromMilliseconds(Settings.Timeouts.SleepIntervalInMillis));

        private static WebDriverWait _customWait(int value) => new WebDriverWait(
            clock,
            Driver.Instance,
            TimeSpan.FromSeconds(value),
            TimeSpan.FromMilliseconds(Settings.Timeouts.SleepIntervalInMillis));

        public static WebDriverWait GetWait()
        {
            return _explicitWait;
        }

        public static void ForExtJsAjax()
        {
            const string isJqueryAjaxComplete = "try { return document.readyState == 'complete' && !Ext.Ajax.isLoading(); } catch(e) { return false; }";

            IJavaScriptExecutor js = (IJavaScriptExecutor)Driver.Instance;
            bool isPageDisplayed = (bool)js.ExecuteScript(isJqueryAjaxComplete);

        }


        public static void WaitForAngular()
        {
            _implicitWait.Until(drv =>
                (drv as IJavaScriptExecutor).ExecuteScript(@"
                    try { 
                        return (
                            (window.angular !== undefined) 
                            && (angular.element(document).injector() !== undefined) 
                            && (angular.element(document).injector().get('$http').pendingRequests.length === 0)); 
                    } catch(e) { 
                        return true; 
                    }"
            ));
        }
        
        public static void ClickJS(IWebElement elementFinder)
        {
            (Driver.Instance as IJavaScriptExecutor).ExecuteScript(@"arguments[0].click()", elementFinder);
        }
        public static void TypeJS(IWebElement elementFinder, string value)
        {
            (Driver.Instance as IJavaScriptExecutor).ExecuteScript(@"arguments[0].setAttribute('value', '" + value + "')", elementFinder);
        }

        public static IWebElement WaitUntilElementPresent(Func<IWebDriver, IWebElement> elementFinder)
        {
            return _implicitWait.Until(elementFinder);
        }

        public static void WaitForProcessToFinish(Func<IWebDriver, IWebElement> elementFinder)
        {
            var element = WaitUntilElementPresent(elementFinder);
             _processWait.Until(drv =>
            {
                try
                {
                    return ! element.Displayed;
                }
                catch (NotFoundException)
                {
                    return true;
                }
                catch (StaleElementReferenceException)
                {
                    return true;
                }
            });
        }

        public static void WaitForElementIfChanges(Func<IWebDriver, IWebElement> elementFinder, int time)
        {
            string oldValue = elementFinder.Invoke(Driver.Instance).Text;
            _customWait(time).Until(drv =>
            {
                try
                {
                    return elementFinder.Invoke(Driver.Instance).Text != oldValue;
                }
                catch (StaleElementReferenceException)
                {
                    return true;
                }
            });
        }

        /**
         * TODO: This need to be double checked to make sure it waits until the file is download
         * 
         */
        public static void WaitDownloadDocument()
        {
            int numberOfFiles = ConfigReader.NumberOfFiles;

            _implicitWait.Until(drv =>
            {
                int downloadedCount = (
                    from
                        file in Directory.EnumerateFiles(
                            ConfigReader.DownloadsPath,
                            "*.xlsm",
                            SearchOption.AllDirectories)
                    select
                        file).Count();
                return downloadedCount > numberOfFiles;
            });
        }
    }
}
