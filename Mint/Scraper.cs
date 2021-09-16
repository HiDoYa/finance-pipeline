using System;
using System.IO;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;

namespace Mint
{
    public class Scraper : IDisposable
    {
        private IWebDriver _driver;
        private WebDriverWait _wait;
        private bool _loggedIn;

        public Scraper(string downloadPath, bool debug = false)
        {
            _loggedIn = false;
            this.DirSetup(downloadPath);

            ChromeOptions options = new ChromeOptions();
            if (!debug)
            {
                options.AddArgument("headless");
                options.AddArgument("no-sandbox");
            }

            options.AddUserProfilePreference("download.default_directory", downloadPath);
            options.AddUserProfilePreference("download.prompt_for_download", false);

            _driver = new ChromeDriver(options);
            _wait = new WebDriverWait(_driver, TimeSpan.FromSeconds(5));
        }

        public void Dispose()
        {
            Console.WriteLine("Cleaning up Selenium");
            _driver.Quit();
        }

        private void DirSetup(string downloadPath)
        {
            if (!Directory.Exists(downloadPath))
            {
                Directory.CreateDirectory(downloadPath);
            }

            string downloadFilePath = Path.Join(downloadPath, "transactions.csv");
            if (File.Exists(downloadFilePath))
            {
                Directory.CreateDirectory(downloadPath);
                File.Delete(downloadFilePath);
            }
        }

        private void TimeWait(double value)
        {
            var timestamp = DateTime.Now;
            var delay = TimeSpan.FromSeconds(value);
            _wait.Until(webdriver => (DateTime.Now - timestamp) > delay);
        }

        public void Login(string username, string password)
        {
            _driver.Manage().Cookies.DeleteAllCookies();

            // Go to mint and click on sign in
            _driver.Url = "https://mint.intuit.com";
            IWebElement signInButton = _driver.FindElement(By.XPath("//a[@data-identifier='sign-in']"));
            signInButton.Click();

            // Input email and click next
            By emailFieldBy = By.XPath("//input[@id='ius-identifier']");
            _wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.VisibilityOfAllElementsLocatedBy(emailFieldBy));
            IWebElement emailField = _driver.FindElement(emailFieldBy);
            emailField.Click();
            emailField.SendKeys(username);

            IWebElement emailNext = _driver.FindElement(By.XPath("//button[@id='ius-sign-in-submit-btn']"));
            emailNext.Click();

            // Input password and click next
            By passwordFieldBy = By.XPath("//input[@id='ius-sign-in-mfa-password-collection-current-password']");
            _wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.VisibilityOfAllElementsLocatedBy(passwordFieldBy));
            IWebElement passwordField = _driver.FindElement(passwordFieldBy);
            passwordField.Click();
            passwordField.SendKeys(password);

            IWebElement loginButton = _driver.FindElement(By.XPath("//input[@id='ius-sign-in-mfa-password-collection-continue-btn']"));
            loginButton.Click();

            // Wait until mint main page is showing
            By mintLogoBy = By.XPath("//a[@id='logo-link']");
            _wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.VisibilityOfAllElementsLocatedBy(mintLogoBy));

            _loggedIn = true;
        }

        public void DownloadTransactions()
        {
            if (!_loggedIn)
            {
                throw new Exception("not logged in");
            }

            // Download file
            _driver.Url = "https://mint.intuit.com/transactionDownload.event?queryNew=&offset=0&filterType=cash&comparableType=8";

            TimeWait(5);
        }
    }
}
