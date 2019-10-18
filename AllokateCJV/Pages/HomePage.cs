using EYWebDriverFramework.Selenium;
using OpenQA.Selenium;
using System;

namespace AllokateCJV.Pages
{
    public class HomePage: TopMenu
    {
        #region Locators
        IWebElement SearchProjectTextBox => WebDriver.FindElement(By.CssSelector("[placeholder='Search Projects']"));
        #endregion

        public HomePage(IWebDriver driver)
            : base(driver)
        {
    

        }
        #region Actions

        public ProjectSettingsPage GoToProjectLink(string url)
        {
            Waits.WaitUntilElementPresent(drv => SearchProjectTextBox);
            Uri ProjectURL = new Uri(url);
            Driver.Instance.Navigate().GoToUrl(ProjectURL);
            return NewPage<ProjectSettingsPage>().LoadProject();
        }

        #endregion

    }
}
