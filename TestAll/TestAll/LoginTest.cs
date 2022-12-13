using manager.Models;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenQA.Selenium;
using OpenQA.Selenium.Edge;
using System.Collections.Generic;
using System.Threading;
using Excel = Microsoft.Office.Interop.Excel;
using OpenQA.Selenium.Support.UI;

namespace manager
{
    [TestClass]
    public class LoginTest
    {
        
        [TestMethod]
        public void Login()
        {
            
            //var listLoginAccount = new List<LoginModel>();
            //for (int i = 0; i < 3; i++)
            //{
            //    var so = 121;
            //    LoginModel obj = new LoginModel();
            //    obj.tk = "test1@gmail.com";
            //    obj.pw = $"{so + i}";
            //    listLoginAccount.Add(obj);
            //}

            int waitingTime1 = 1000;
            int waitingTime2 = 700;
            int waitingTime = 9000;

            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(@"C:\Users\Windows\OneDrive\Desktop\DATN\TestAll\TestAll\Excel\test.xlsx");

            Excel.Range xlTestRange;
            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            xlTestRange = xlWorksheet.UsedRange;

            int row_used = xlTestRange.Cells[1, 1].End(Excel.XlDirection.xlDown).Row;

            var listLoginAccount = new List<LoginModel>();
            for (int i = 2; i <= row_used; i++)
            {
                LoginModel obj = new LoginModel();
                obj.email = xlWorksheet.Range[$"A{i}"].Value.ToString();
                obj.pass = xlWorksheet.Range[$"B{i}"].Value.ToString();
                listLoginAccount.Add(obj);
            }

            IWebDriver webDriver = new EdgeDriver();
            webDriver.Manage().Window.Maximize();
            //link url
            webDriver.Navigate().GoToUrl("http://localhost:53451/login");

            foreach ( var item in listLoginAccount )
            {
                By clicklogine = By.XPath("/html/body/app-root/app-auth-layout/div/app-login/div[2]/div/div/div/div/form/div[4]/button");
                webDriver.FindElement(clicklogine).Click();

                Thread.Sleep( waitingTime2 );
                //email
                By username = By.XPath("/html/body/app-root/app-auth-layout/div/app-login/div[2]/div/div/div/div/form/div[1]/div/input");
                webDriver.FindElement(username).SendKeys(item.email);
                //password
                By password = By.XPath("/html/body/app-root/app-auth-layout/div/app-login/div[2]/div/div/div/div/form/div[2]/div/input");
                webDriver.FindElement(password).SendKeys(item.pass);
                //nút đăng nhập
                Thread.Sleep(waitingTime1);
                By clicklogin = By.XPath("/html/body/app-root/app-auth-layout/div/app-login/div[2]/div/div/div/div/form/div[4]/button");
                webDriver.FindElement(clicklogin).Click();

                Thread.Sleep(waitingTime1);
                webDriver.FindElement(username).Clear();
                webDriver.FindElement(password).Clear();


                Thread.Sleep(waitingTime);
                webDriver.Quit();
                webDriver.Close();
            }
        }    

            //foreach (var item in listLoginAccount)
            //{
            //    Thread.Sleep(waitingTime1);

                //    By username = By.XPath("/html/body/app-root/app-auth-layout/div/app-login/div[2]/div/div/div/div/form/div[1]/div/input");
                //    webDriver.FindElement(username).SendKeys(item.tk);

                //    Thread.Sleep(waitingTime2);

                //    //webDriver.FindElement(username).Clear();
                //    webDriver.FindElement(username).Click();

                //    //
                //    Thread.Sleep(waitingTime1);

                //    By password = By.XPath("/html/body/app-root/app-auth-layout/div/app-login/div[2]/div/div/div/div/form/div[2]/div/input");
                //    webDriver.FindElement(password).SendKeys(item.pw);

                //    Thread.Sleep(waitingTime2);

                //    //webDriver.FindElement(password).Clear();
                //    webDriver.FindElement(password).Click();

                //    //Thread.Sleep(waitingTime1);

                //    By clicklogin = By.XPath("/html/body/app-root/app-auth-layout/div/app-login/div[2]/div/div/div/div/form/div[4]/button");
                //    webDriver.FindElement(clicklogin).Click();

                //    //Thread.Sleep(waitingTime1);

                //    webDriver.FindElement(username).Clear();
                //    webDriver.FindElement(password).Clear();

                //    //Thread.Sleep(waitingTime);
                //    //webDriver.Quit();
                //    //webDriver.Close();
                //}
                //Thread.Sleep(waitingTime);
                //webDriver.Quit();
                //webDriver.Close();
    }
}


