using Microsoft.VisualStudio.TestTools.UnitTesting;
using NUnit.Framework;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System;
using System.Threading;
using Excel = Microsoft.Office.Interop.Excel;
using Assert = Microsoft.VisualStudio.TestTools.UnitTesting.Assert;
using System.IO;
using Microsoft.Office.Interop.Excel;
using System.Collections.Generic;
using ExcelDataReader;
using OpenQA.Selenium.Support.UI;
using static Plan_Test.Module01;

namespace Plan_Test
{
    [TestFixture]
    public class Module03
    {
        IWebDriver driver;
        Excel.Application dataApp;
        Excel.Workbook dataBook;
        Excel.Worksheet dataSheet;
        Excel.Range xlRange;
        

        //Đăng nhập User
        public void LoginUser()
        {
            //btn User
            driver.FindElement(By.XPath("//a[@class='fas fa-user']")).Click();
            Thread.Sleep(2000);


            driver.FindElement(By.Id("EmailCus")).Clear();
            driver.FindElement(By.Id("EmailCus")).SendKeys("hang@gmail.com");
            driver.FindElement(By.Id("PassCus")).Clear();
            driver.FindElement(By.Id("PassCus")).SendKeys("bhang204");
            driver.FindElement(By.XPath("//input[@value='Đăng nhập']")).Click();
            Thread.Sleep(2000);
        }

        [SetUp]
        public void Setup()
        {
            //Mở web
            driver = new ChromeDriver();
            driver.Navigate().GoToUrl("http://localhost:54077/Home");
            driver.Manage().Window.Maximize();
            Thread.Sleep(3000);
            
            //Mở excel
            dataApp = new Excel.Application();
            dataBook = dataApp.Workbooks.Open(@"D:\\Code\\DBCLPM\\DoAn\\C07_Copy.xlsx");
            dataSheet = dataBook.Sheets[4];
            xlRange = dataSheet.UsedRange;

        }

        //Thêm sản phẩm vào giỏ hàng khi chưa đăng nhập
        [Test]
        public void AddToCartNoLogin()
        {
            try
            {
                //cuộn tìm phần tử
                IJavaScriptExecutor js = (IJavaScriptExecutor)driver;
                IWebElement element = driver.FindElement(By.XPath("//body/div/section[3]/div[1]/div[1]/div[1]/div[2]"));
                js.ExecuteScript("arguments[0].scrollIntoView(true);", element);
                Thread.Sleep(2000);

                //btn chi tiết
                driver.FindElement(By.XPath("//body/div/section[3]/div[1]/div[1]/div[1]/div[2]/a[1]")).Click();
                Thread.Sleep(2000);

                //btn thêm vào giỏ hàng
                driver.FindElement(By.XPath("//a[@id='addToCartBtn']")).Click();
                Thread.Sleep(2000);

                Assert.IsTrue(driver.Url.Contains("http://localhost:54077/Home/Login"));
                Console.WriteLine("Successful");
                xlRange.Cells[8][5].Value = "[2] Successful";
            }
            catch (Exception ex)
            {
                Console.WriteLine("Failed\n" + ex);
                xlRange.Cells[8][5].Value = "[2] Failed";
            }
        }

        //Thêm sản phẩm vào giỏ hàng khi đã đăng nhập
        [Test]
        public void AddToCart()
        {
            try
            {
                LoginUser();

                //cuộn tìm phần tử
                IJavaScriptExecutor js = (IJavaScriptExecutor)driver;
                IWebElement element = driver.FindElement(By.XPath("//body/div/section[3]/div[1]/div[1]/div[1]/div[2]"));
                js.ExecuteScript("arguments[0].scrollIntoView(true);", element);
                Thread.Sleep(2000);

                //btn chi tiết
                driver.FindElement(By.XPath("//body/div/section[3]/div[1]/div[1]/div[1]/div[2]/a[1]")).Click();
                Thread.Sleep(2000);

                //btn thêm vào giỏ hàng
                driver.FindElement(By.XPath("//a[@id='addToCartBtn']")).Click();
                Thread.Sleep(2000);

                Assert.IsTrue(driver.Url.Contains("http://localhost:54077/Home/ViewCart"));
                Console.WriteLine("Successful");
                xlRange.Cells[8][6].Value = "[3] Successful";
            }
            catch (Exception ex)
            {
                Console.WriteLine("Failed\n" + ex);
                xlRange.Cells[8][6].Value = "[3] Failed";
            }
        }

        //Kiểm tra xem sản phẩm vừa thêm có trong giỏ hàng hay không
        [Test]
        public void TestIntegrated_AddToCart()
        {
            AddToCart();

            try
            {
                //quay lại trang chủ
                driver.FindElement(By.XPath("//a[@href='Index']")).Click();
                Thread.Sleep(2000);

                //vào giỏ hàng
                driver.FindElement(By.XPath("//a[@id='cart']")).Click();
                Thread.Sleep(2000);

                IWebElement Element = driver.FindElement(By.XPath("//div[@class='product-name']"));
                Assert.IsTrue(Element.Displayed, "Successfull");
                Console.WriteLine("Successful");
                xlRange.Cells[8][7].Value = "[3] Successful";
            }
            catch (Exception ex)
            {
                Console.WriteLine("Failed\n" + ex);
                xlRange.Cells[8][7].Value = "[3] Failed";
            }
        }

        //Thêm sản phẩm khi đã có sản phẩm trong giỏ hàng
        [Test]
        public void AddToCart2()
        {
            try
            {
                AddToCart();

                //quay lại trang chủ
                driver.FindElement(By.XPath("//a[@href='Index']")).Click();
                Thread.Sleep(2000);

                //cuộn tìm phần tử
                IJavaScriptExecutor js1 = (IJavaScriptExecutor)driver;
                IWebElement element1 = driver.FindElement(By.XPath("//body/div/section[3]/div[1]/div[1]/div[2]/div[1]"));
                js1.ExecuteScript("arguments[0].scrollIntoView(true);", element1);
                Thread.Sleep(2000);

                //btn chi tiết
                driver.FindElement(By.XPath("//body/div/section[3]/div[1]/div[1]/div[2]/div[2]/a[1]")).Click();
                Thread.Sleep(2000);

                //btn thêm vào giỏ hàng
                driver.FindElement(By.XPath("//a[@id='addToCartBtn']")).Click();
                Thread.Sleep(2000);

                IWebElement Element = driver.FindElement(By.XPath("//div[@class='product-name']"));
                Assert.IsTrue(Element.Displayed, "Successfull");
                Console.WriteLine("Successful");
                xlRange.Cells[8][8].Value = "[3] Successful";
            }
            catch (Exception ex)
            {
                Console.WriteLine("Failed\n" + ex);
                xlRange.Cells[8][8].Value = "[3] Failed";
            }
        }

        //Xóa sản phẩm khỏi giỏ hàng
        [Test]
        public void DeleteProduct()
        {
            try
            {
                AddToCart();

                //btn delete
                driver.FindElement(By.XPath("//i[@class='fas fa-trash-alt']")).Click();
                Thread.Sleep(2000);
                Console.WriteLine("Successful");
                xlRange.Cells[8][9].Value = "[3] Successful";
            }
            catch (Exception ex)
            {
                Console.WriteLine("Failed\n" + ex);
                xlRange.Cells[8][9].Value = "[3] Failed";
            }
        }

        //Kiểm tra hiển thị số lượng sản phẩm
        [Test]
        public void TestIntegrated_DislayQuantity()
        {
            try
            {
                AddToCart();

                //quay lại trang chủ
                driver.FindElement(By.XPath("//a[@href='Index']")).Click();
                Thread.Sleep(2000);

                IWebElement element = driver.FindElement(By.Id("cart-quantity"));
                string quantity = element.Text.Trim();
                if (quantity == "1")
                {
                    Console.WriteLine("Successful");
                    
                }
                else
                {
                    Console.WriteLine("Failed");
                }
                xlRange.Cells[8][10].Value = "[3] Successful";
            }
            catch (Exception ex)
            {
                Console.WriteLine("Failed\n" + ex);
                xlRange.Cells[8][10].Value = "[3] Failed";
            }
        }

        //Kiểm tra giỏ hàng sau khi đăng xuất
        [Test]
        public void CheckCart()
        {
            try
            {
                AddToCart();

                //btn logout
                driver.FindElement(By.XPath("//i[@class='fa-solid fa-right-from-bracket']")).Click();
                Thread.Sleep(2000);

                LoginUser();

                //vào giỏ hàng
                driver.FindElement(By.XPath("//a[@id='cart']")).Click();
                Thread.Sleep(2000);

                IWebElement Element = driver.FindElement(By.XPath("//div[@class='product-name']"));
                Assert.IsTrue(Element.Displayed, "Successfull");
                Console.WriteLine("Successful");
                xlRange.Cells[8][11].Value = "[3] Successful";
            }
            catch (Exception ex)
            {
                Console.WriteLine("Failed\n" + ex);
                xlRange.Cells[8][11].Value = "[3] Failed";
            }
        }

        //Đặt hàng thành công
        [Test]
        public void Order()
        {
            try
            {
                AddToCart();

                //đặt hàng
                driver.FindElement(By.XPath("//a[@class='btn btn-danger']")).Click();
                Thread.Sleep(2000);

                Assert.IsTrue(driver.Url.Contains("http://localhost:54077/Home/ThanhCong"));
                Console.WriteLine("Successful");
                xlRange.Cells[8][12].Value = "[3] Successful";
            }
            catch (Exception ex)
            {
                Console.WriteLine("Failed\n" + ex);
                xlRange.Cells[8][12].Value = "[3] Failed";
            }
        }

        //Kiểm tra đặt hàng thành công
        [Test]
        public void Integrated_CheckOrder()
        {
            driver = new ChromeDriver();
            driver.Navigate().GoToUrl("http://localhost:54077/AdminUser/Login");
            driver.Manage().Window.Maximize();
            Thread.Sleep(3000);

            driver.FindElement(By.Id("NameUser")).Clear();
            driver.FindElement(By.Id("NameUser")).SendKeys("admin");
            driver.FindElement(By.Id("PasswordUser")).Clear();
            driver.FindElement(By.Id("PasswordUser")).SendKeys("admin");
            driver.FindElement(By.XPath("//button[@type='submit']")).Click();
            Thread.Sleep(2000);

            try
            {
                //đơn hàng
                driver.FindElement(By.XPath("//a[contains(text(),'Đơn hàng')]")).Click();
                Thread.Sleep(2000);

                IWebElement Element = driver.FindElement(By.XPath("//td[normalize-space()='01/04/2025 12:00:00 SA']"));
                Assert.IsTrue(Element.Displayed, "Successfull");
                Console.WriteLine("Successful");
                xlRange.Cells[8][13].Value = "[3] Successful";
            }
            catch (Exception ex)
            {
                Console.WriteLine("Failed\n" + ex);
                xlRange.Cells[8][13].Value = "[3] Failed";
            }
        }

        //Đặt hàng khi chưa đăng nhập
        [Test]
        public void OrderNoLogin()
        {
            try
            {
                AddToCartNoLogin();

                //đặt hàng
                driver.FindElement(By.XPath("//a[@class='btn btn-danger']")).Click();
                Thread.Sleep(2000);

                Assert.IsTrue(driver.Url.Contains("http://localhost:54077/Home/Login?returnUrl=%2FHome%2FThanhToan"));
                Console.WriteLine("Successful");
                xlRange.Cells[8][14].Value = "[3] Successful";
            }
            catch (Exception ex)
            {
                Console.WriteLine("Failed\n" + ex);
                xlRange.Cells[8][14].Value = "[3] Failed";
            }
        }

        [TearDown]
        public void TearDown()
        {
            dataBook.Save();
            dataBook.Close();
            dataApp.Quit();
            driver.Close();
        }
    }
}
