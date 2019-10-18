using EYWebDriverFramework.Config;
using EYWebDriverFramework.Selenium;
using EYWebDriverFramework.Utils;
using OpenQA.Selenium;
using System;

namespace AllokateCJV.Pages
{
    public class AssetsPurchasesPage : TopMenu
    {
        #region Locators

        IWebElement CommitChangesButton => WebDriver.FindElement(By.ClassName("commit-label"));
        IWebElement CommitChangesTextBox => WebDriver.FindElement(By.ClassName("input-Box-style"));
        IWebElement CommitChangesSaveCommentButton => 
        WebDriver.FindElement(By.XPath("//div[@class='cancel-button-modal' and contains(text(),'Save Comment')]"));
        IWebElement ImportDataButton => WebDriver.FindElement(By.ClassName("import-label"));

        #region Regular Tax Book Set
        IWebElement RegularTaxBookSetTextBoxGrossOrCostBasisRecoverableRow => 
            WebDriver.FindElement(By.CssSelector(".cost-basis-amount.float-left[data-identifier=C_1_R_1]"));
        #endregion

        #region AMT Book Set
        IWebElement AMTBookSetTextBoxGrossOrCostBasisRecoverableRow => 
            WebDriver.FindElement(By.CssSelector(".cost-basis-amount.float-left.edit-background[data-identifier=C_4_R_1]"));
        #endregion

        #region E&P Book Set
        IWebElement EandPBookSetTextBoxGrossOrCostBasisRecoverableRow => 
            WebDriver.FindElement(By.CssSelector(".cost-basis-amount.float-left.edit-background[data-identifier=C_7_R_1]"));
        #endregion

        #region State Book Set
        IWebElement StateBookSetTextboxGrossOrCostBasisRecoverableRow => 
            WebDriver.FindElement(By.CssSelector(".cost-basis-amount.float-left.edit-background[data-identifier=C_10_R_1]"));
        #endregion
        
        #endregion

        public AssetsPurchasesPage(IWebDriver driver) : base(driver)
        {
        }

        public bool UpdateValues()
        {
            ActivateUpdateScreenMode();
            bool upDayValue = SetUpDay(RegularTaxBookSetTextBoxGrossOrCostBasisRecoverableRow);
            LogHelpers.Write(string.Format("Set \"Regular Tax Book Set\" value on Recoverable row to {0}.", upDayValue ? Settings.Eyds.DayAValue : Settings.Eyds.DayBValue ));
            IncreaseOrDecreaseValue(RegularTaxBookSetTextBoxGrossOrCostBasisRecoverableRow, upDayValue);
            LogHelpers.Write(string.Format("Set \"AMT Book Set\" value on Recoverable row to {0}.", upDayValue ? Settings.Eyds.DayAValue : Settings.Eyds.DayBValue));
            IncreaseOrDecreaseValue(AMTBookSetTextBoxGrossOrCostBasisRecoverableRow, upDayValue);
            LogHelpers.Write(string.Format("Set \"E&P Book Set\" value on Recoverable row to {0}.", upDayValue ? Settings.Eyds.DayAValue : Settings.Eyds.DayBValue));
            IncreaseOrDecreaseValue(EandPBookSetTextBoxGrossOrCostBasisRecoverableRow, upDayValue);
            LogHelpers.Write(string.Format("Set \"State Book Set\" value on Recoverable row to {0}.", upDayValue ? Settings.Eyds.DayAValue : Settings.Eyds.DayBValue));
            IncreaseOrDecreaseValue(StateBookSetTextboxGrossOrCostBasisRecoverableRow, upDayValue);
            ConfirmChanges();           
            return upDayValue;
        }

        private bool SetUpDay(IWebElement webElement)
        {
            LogHelpers.Write(string.Format("Read \"Regular Tax Book Set\" text box on \"Recoverable row\" to check which golden files will be used."));
            webElement.ScrollAndClick();
            int value = Convert.ToInt32(Convert.ToDecimal(webElement.Read()));
            return !(value == Settings.Eyds.DayAValue);
        }

        private void IncreaseOrDecreaseValue(IWebElement webElement, bool dayA)
        {
            webElement.ScrollAndClick();
            int value;
            if (dayA) 
            {
                value = Settings.Eyds.DayAValue;
            }
            else
            {
                value = Settings.Eyds.DayBValue;
            }
            webElement.DynamicType(value.ToString());
        }

        private void ActivateUpdateScreenMode()
        {
            LogHelpers.Write(string.Format("Perform double click on \"Regular Tax Book Set\" text box on \"Recoverable row\" to activate edit mode."));
            RegularTaxBookSetTextBoxGrossOrCostBasisRecoverableRow.DoubleClick();
            Waits.WaitUntilElementPresent(drv => CommitChangesButton); 
        }

        private void ConfirmChanges()
        {
            LogHelpers.Write(string.Format("Click \"Commit changes\" button."));
            CommitChangesButton.ClickAndWaitForAjax();
            Waits.WaitUntilElementPresent(drv => CommitChangesTextBox);

            LogHelpers.Write(string.Format("Write text on \"Commit changes\" text box."));
            CommitChangesTextBox.Type("Text to have this text box with data.");

            LogHelpers.Write(string.Format("Click \"Save comment\" button on \"Commit changes\" modal."));
            CommitChangesSaveCommentButton.ClickAndWaitForAjax();
            Waits.WaitUntilElementPresent(drv => ImportDataButton);
        }
    }
}
