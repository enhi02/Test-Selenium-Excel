//using System;
//using System.Collections.Generic;
//using System.Linq;
//using System.Text;
//using System.Threading.Tasks;
//using OpenQA.Selenium;
//using OpenQA.Selenium.Edge;
//using System.Threading;

//namespace TestAll
//{
//    class ExcelData
//    {
//        public static void ClearLoginInformation(IWebDriver Driver)
//        {
//            Driver.FindElement(By.XPath("/html/body/app-root/app-auth-layout/div/app-login/div[2]/div/div/div/div/form/div[1]/div/input")).Clear();
//            Driver.FindElement(By.XPath("/html/body/app-root/app-auth-layout/div/app-login/div[2]/div/div/div/div/form/div[2]/div/input")).Clear();
//        }

//        public static void LoginInformation(string UserEmail, string password, IWebDriver Driver)
//        {
//            Driver.FindElement(By.XPath("/html/body/app-root/app-auth-layout/div/app-login/div[2]/div/div/div/div/form/div[1]/div/input")).SendKeys(UserEmail);
//            Driver.FindElement(By.XPath("/html/body/app-root/app-auth-layout/div/app-login/div[2]/div/div/div/div/form/div[2]/div/input")).SendKeys(password);
//        }
//    }
//}

