using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OpenQA.Selenium.Remote;
using OpenQA.Selenium.Firefox;
using OpenQA.Selenium.Support.UI;
using OpenQA.Selenium;
using System.IO;
using System.Text.RegularExpressions;
using System.Net.Mail;

namespace cabinet.alitrack.ru
{
    class Program
    {
        static void Main(string[] args)
        {
            RemoteWebDriver browser = new FirefoxDriver();

            try
            {
                browser.Navigate().GoToUrl("http://alitrack.ru/");

                #region Login
                if (browser.FindElementByName("login") != null)
                {
                    browser.FindElementByName("login").SendKeys(cabinet.alitrack.ru.Properties.Settings.Default.UserName);
                    browser.FindElementByName("pass").SendKeys(cabinet.alitrack.ru.Properties.Settings.Default.Password);
                    browser.FindElementByXPath("//input[@type='submit']").Click();
                }
                #endregion

                WaitForCompleteXPath(browser, "//a[contains(@href,'tracking')]");
                browser.Navigate().GoToUrl("http://cabinet.alitrack.ru/tracking/");
                WaitForCompleteXPath(browser, "//a[contains(@href,'tracking')]");


                if (browser.PageSource.Contains("trackingCheckAllItems"))
                {
                    #region update
                    browser.FindElementById("trackingCheckAllItems").Click();
                    var alert = browser.SwitchTo().Alert();
                    alert.Accept();
                    WaitForCompleteXPath(browser, "//div[@id='processingWindow' and contains(@style,'display: none;')]");
                    var trackingRows = browser.FindElementsByClassName("tracking-row");

                    foreach (var trackingRow in trackingRows)
                    {
                        Console.WriteLine(trackingRow.FindElement(By.ClassName("name")).Text);
                        string number = trackingRow.GetAttribute("id").Replace("trow", string.Empty);
                        if (browser.PageSource.Contains(string.Format("id=\"refresh-button{0}\"", number)))
                        {
                            browser.FindElementById("refresh-button" + number).Click();
                            WaitForCompleteXPath(browser, "//div[@id='processingWindow' and contains(@style,'display: none;')]");
                        }
                    }
                    #endregion

                    #region find new
                    browser.Navigate().GoToUrl("http://cabinet.alitrack.ru/tracking/");
                    WaitForCompleteXPath(browser, "//a[contains(@href,'tracking')]");
                    var trackingNewRows = browser.FindElementsByClassName("tracking-row");
                    Regex regexButtons = new Regex("<div[^>]*class=\"buttons\"[^>]*>(.*?)</div>", RegexOptions.Compiled | RegexOptions.IgnoreCase);
                    Regex regexFull = new Regex("<div[^>]*class=\"fullinfo\"[^>]*>(.*?)</div>", RegexOptions.Compiled | RegexOptions.IgnoreCase);
                    int newCount = 0;
                    StringBuilder newItems = new StringBuilder();
                    foreach (var trackingRow in trackingNewRows)
                    {
                        Console.WriteLine(trackingRow.FindElement(By.ClassName("name")).Text);
                        if (!string.IsNullOrEmpty(trackingRow.FindElement(By.ClassName("newSteps")).Text))
                        {
                            newCount++;
                            newItems.AppendLine(regexFull.Replace(regexButtons.Replace(trackingRow.GetAttribute("outerHTML"), string.Empty), string.Empty));
                        }
                    }
                    #endregion

                    if (newCount > 0)
                    {
                        #region Processing
                        newItems.Replace("class=\"clear\"", "style=\"clear: both;\"");
                        newItems.Replace("class=\"newSteps\"", "style=\"font-size: 11px;font-weight: normal;color: #a00;display: inline-block;padding: 0 0 0 4px;margin: -5px 0 0 0;vertical-align: top;\"");
                        newItems.Replace("class=\"info\"", "style=\"padding: 0;	overflow: hidden;	white-space: nowrap;color: #50749b;\"");
                        newItems.Replace("class=\"meta\"", "style=\"padding: 3px 0 0 0;	display:inline-block;overflow: hidden;color: #97adc3;white-space: nowrap;\"");
                        newItems.Replace("class=\"name\"", "style=\"font-size: 18px; padding: 9px 0 5px 0; overflow: hidden; color: #50749b; cursor: pointer; white-space: nowrap;\"");

                        newItems.Replace("class=\"tracking-row hasmoves\"", "style=\"font-family: Arial, Verdana, Tahoma; font-size: 12px; line-height: 1em; margin: 0 0 2px 0; background-color: #d0e7ff;	border-radius: 5px;\"");
                        newItems.Replace("class=\"tracking-row hasexport\"", "style=\"font-family: Arial, Verdana, Tahoma; font-size: 12px; line-height: 1em; margin: 0 0 2px 0; background-color: #bad2ff;	border-radius: 5px;\"");
                        newItems.Replace("class=\"tracking-row haslong\"", "style=\"font-family: Arial, Verdana, Tahoma; font-size: 12px; line-height: 1em; margin: 0 0 2px 0; background-color: #c0b4ff;	border-radius: 5px;\"");
                        newItems.Replace("class=\"tracking-row hascustom\"", "style=\"font-family: Arial, Verdana, Tahoma; font-size: 12px; line-height: 1em; margin: 0 0 2px 0; background-color: #e6ff8c;	border-radius: 5px;\"");
                        newItems.Replace("class=\"tracking-row hasnotrace\"", "style=\"font-family: Arial, Verdana, Tahoma; font-size: 12px; line-height: 1em; margin: 0 0 2px 0; background-color: #ffc994; border-radius: 5px;\"");
                        newItems.Replace("class=\"tracking-row hasnative\"", "style=\"font-family: Arial, Verdana, Tahoma; font-size: 12px; line-height: 1em; margin: 0 0 2px 0; background-color: #b3fab3;	border-radius: 5px;\"");
                        newItems.Replace("class=\"tracking-row hasarrive\"", "style=\"font-family: Arial, Verdana, Tahoma; font-size: 12px; line-height: 1em; margin: 0 0 2px 0; background-color: #161; color: #b3fab3; border-radius: 5px;\"");
                        newItems.Replace("class=\"tracking-row haswarning\"", "style=\"font-family: Arial, Verdana, Tahoma; font-size: 12px; line-height: 1em; margin: 0 0 2px 0; background-color: #ffd0d0; color: #50749B; border-radius: 5px;\"");
                        newItems.Replace("class=\"tracking-row hasrecieved\"", "style=\"font-family: Arial, Verdana, Tahoma; font-size: 12px; line-height: 1em; margin: 0 0 2px 0; background-color: #DDD;	border-radius: 5px;\"");


                        newItems.Replace("class=\"tracking-row\"", "style=\"font-family: Arial, Verdana, Tahoma; font-size: 12px; line-height: 1em; margin: 0 0 2px 0; background-color: #ebeff3; border-radius: 5px;\"");

                        newItems.Replace("class=\"main\"", "style=\"padding: 0 0 8px 8px; float: left; background-color: #ebeff3; border-radius: 0 5px 5px 0; cursor: pointer; position: relative;\"");
                        newItems.Replace("class=\"time\"", "style=\"font-size: 10px; line-height: 10px; color: #50749b; text-align: center; padding: 35px 0 0;\"");
                        newItems.Replace("class=\"icon\" style=\"", "style=\"float: left;width: 60px;height: 60px;border-radius: 5px 0 0 5px; overflow: hidden; cursor: default;background: url(trackicons/0.png) 14px 5px no-repeat;");
                        newItems.Replace(" class=\"infopos\"", string.Empty);
                        #endregion

                        #region Send report
                        string report = cabinet.alitrack.ru.Properties.Resources.Template.Replace("{NEW}", newItems.ToString());
                        SmtpClient client = new SmtpClient("localhost");
                        MailAddress from = new MailAddress(cabinet.alitrack.ru.Properties.Settings.Default.EmailFrom);
                        MailAddress to = new MailAddress(cabinet.alitrack.ru.Properties.Settings.Default.EmailTo);
                        MailMessage message = new MailMessage(from, to);
                        message.Subject = string.Format("{0} new update from alitrack", newCount);
                        message.IsBodyHtml = true;
                        message.Body = report;
                        client.Send(message);
                        #endregion
                    }
                }                
            }
            catch(Exception e)
            {
                SmtpClient client = new SmtpClient("localhost");
                MailAddress from = new MailAddress(cabinet.alitrack.ru.Properties.Settings.Default.EmailFrom);
                MailAddress to = new MailAddress(cabinet.alitrack.ru.Properties.Settings.Default.EmailTo);
                MailMessage message = new MailMessage(from, to);
                message.Subject = "new update from alitrack Exception";                
                message.Body = e.ToString();
                client.Send(message);
            }
            finally
            {
                browser.Quit();
            }
        }

        static void WaitForCompleteXPath(RemoteWebDriver browser, string xPath)
        {
            var wait = new WebDriverWait(browser, TimeSpan.FromSeconds(3600));
            wait.Until(driver => driver.FindElement(By.XPath(xPath)));
        }

        static void WaitForVisibleXPath(RemoteWebDriver browser, string xPath)
        {
            var wait = new WebDriverWait(browser, TimeSpan.FromSeconds(3600));
            wait.Until(driver => driver.FindElement(By.XPath(xPath)).Displayed);
        }
    }
}
