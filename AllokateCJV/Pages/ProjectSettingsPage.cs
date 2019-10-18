using EYWebDriverFramework.Selenium;
using OpenQA.Selenium;

namespace AllokateCJV.Pages
{
    public class ProjectSettingsPage : TopMenu
    {
        IWebElement ProjectSettingsLabel => 
            WebDriver.FindElement(By.ClassName("project-settings-box-header"));
        
        public ProjectSettingsPage(IWebDriver driver)
            : base(driver)
        {
        }

        public ProjectSettingsPage LoadProject()
        {
            Waits.GetWait().Until(drv => ProjectSettingsLabel);
            return this;
        }

    }
}
