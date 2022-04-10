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

        // Setup selenium driver
        public Scraper(string downloadPath, string driverPath)
        {
            _loggedIn = false;
            this.DirSetup(downloadPath);

            ChromeOptions options = new ChromeOptions();

            //// WIP: Headless is not working correctly
            //options.AddArgument("headless");
            //options.AddArgument("no-sandbox");
            //options.AddArgument("disable-dev-shm-usage");
            //options.AddArgument("disable-gpu");

            options.AddUserProfilePreference("download.default_directory", downloadPath);
            options.AddUserProfilePreference("download.prompt_for_download", false);

            _driver = new ChromeDriver(driverPath, options);
            _wait = new WebDriverWait(_driver, TimeSpan.FromSeconds(30));
        }

        // Cleanup selenium driver
        public void Dispose()
        {
            Console.WriteLine("Cleaning up Selenium");
            _driver.Quit();
        }

        // Create new directory for download file
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

        // Wait for a certain amount of time
        private void TimeWait(double value)
        {
            var timestamp = DateTime.Now;
            var delay = TimeSpan.FromSeconds(value);
            _wait.Until(webdriver => (DateTime.Now - timestamp) > delay);
        }

        // Login to mint
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

            IWebElement emailNext = _driver.FindElement(By.XPath("//button[@id='ius-identifier-first-submit-btn']"));
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
            By mintMainBy = By.XPath("//div[@id='mintNavigation']");
            _wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.VisibilityOfAllElementsLocatedBy(mintMainBy));

            _loggedIn = true;
        }

        // Download transactions from mint while logged in
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
