using EYWebDriverFramework.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using System;
using System.Threading;

namespace EYWebDriverFramework.Selenium
{
    public static class WebElementExtensions
    {
        private static void DynamicSendKeys(this IWebElement element, string text)
        {
            foreach (char ch in text.ToCharArray())
            {
                element.SendKeys(Keys.End);
                Thread.Sleep(200);
                element.SendKeys(ch.ToString());
            }
        }
        public static void Type(this IWebElement element, string text)
        {
            element.Clear();
            element.SendKeys(text);
            LogHelpers.Write("Write \"" + text + "\".");
        }
        public static void DynamicType(this IWebElement element, string text)
        {
            element.Clear();
            element.DynamicSendKeys(text);
            LogHelpers.Write("Write \"" + text + "\".");
        }


        public static string Read(this IWebElement element)
        {
            string textRead = element.Text;
            LogHelpers.Write("Read \"" + textRead + "\".");
            return textRead;
        }

        // Clicks
        public static void ClickAndWaitForAjax(this IWebElement element)
        {
            Waits.WaitUntilElementPresent(drv => element);
            element.ClickJS();
            Thread.Sleep(2000);
            Waits.ForExtJsAjax();
        }
        public static void ClickJS(this IWebElement element)
        {
            Waits.ClickJS(element);
        }

        public static void DoubleClick(this IWebElement element)
        {
            Actions action = new Actions(Driver.Instance);
            action.DoubleClick(element).Perform();
        }

        public static void ScrollAndClick(this IWebElement element)
        {
            element.ScrollTo();
            element.ClickAndWaitForAjax();
        }
        public static void ScrollTo(this IWebElement element)
        {

            var aux = 0;
            ((IJavaScriptExecutor)Driver.Instance).ExecuteScript("arguments[0].scrollIntoView(false);", element);
            ((IJavaScriptExecutor)Driver.Instance).ExecuteScript("window.scrollBy(0, 60);", element);

            while (element.Displayed.Equals(false) && (aux < 10))
            {
                ((IJavaScriptExecutor)Driver.Instance).ExecuteScript("window.scrollBy(0, 30);", element);

                aux += 1;
            }
        }

        public static void MouseHover(this IWebElement element)
        {
            Actions action = new Actions(Driver.Instance);
            action.MoveToElement(element, element.Size.Width * 85 / 100, element.Size.Height * 85 / 100).Build().Perform();
        }

        // Select
        public static void SelectByText(this IWebElement element, string text)
        {
            Waits.WaitUntilElementPresent(drv => element);
            var options = element.FindElements(By.TagName("option"));
            foreach (var option in options)
            {
                bool isSelected = option.GetAttribute("selected") != null && option.GetAttribute("selected").ToLower() == "true";
                bool isMatch = option.Text.Contains(text);

                //Select the correct match and deselect all others
                if (isMatch != isSelected)
                    option.Click();
            }
        }

        public static bool IsDisplayed(this IWebElement element)
        {
            try
            {
                bool displayed = false;
                if (element != null)
                    displayed = element.Displayed;
                return displayed;
            }
            catch (StaleElementReferenceException)
            {
                return false;
            }
            catch (InvalidOperationException e)
            {
                if (e.Message.ToUpper().Contains("DETERMINING IF ELEMENT IS DISPLAYED"))
                {
                    return false;
                }
                throw e;
            }
        }
    }
}