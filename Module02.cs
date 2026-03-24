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

namespace Plan_Test
{
    [TestFixture]
    public class Module02
    {
        IWebDriver driver;
        IWebElement element, inputField;
        Excel.Application dataApp;
        Excel.Workbook dataBook;
        Excel.Worksheet dataSheet;
        Excel.Range xlRange;

        //Đăng nhập Admin
        public void LoginAdmin()
        {
            driver.FindElement(By.Id("NameUser")).Clear();
            driver.FindElement(By.Id("NameUser")).SendKeys("admin");
            driver.FindElement(By.Id("PasswordUser")).Clear();
            driver.FindElement(By.Id("PasswordUser")).SendKeys("admin");
            driver.FindElement(By.XPath("//button[@type='submit']")).Click();
            Thread.Sleep(2000);
        }

        [SetUp]
        public void Setup()
        {
            //Mở web
            driver = new ChromeDriver();
            driver.Navigate().GoToUrl("http://localhost:54077/AdminUser/Login");
            driver.Manage().Window.Maximize();
            Thread.Sleep(3000);

            //Mở excel
            dataApp = new Excel.Application();
            dataBook = dataApp.Workbooks.Open(@"D:\\Code\\DBCLPM\\DoAn\\C07_Copy.xlsx");
            dataSheet = dataBook.Sheets[3];
            xlRange = dataSheet.UsedRange;

        }

        [Test]
        public void AddCategory()
        {
            try
            {
                LoginAdmin();

                //btn thể loại
                driver.FindElement(By.XPath("//a[contains(text(),'Thể loại')]")).Click();
                Thread.Sleep(2000);

                //btn create
                driver.FindElement(By.XPath("//a[contains(text(),'TẠO MỚI')]")).Click();
                Thread.Sleep(2000);

                //create
                driver.FindElement(By.Id("IDCate")).Clear();
                driver.FindElement(By.Id("NameCate")).Clear();

                driver.FindElement(By.Id("IDCate")).SendKeys(xlRange.Cells[6][6]?.Value2?.ToString());
                driver.FindElement(By.Id("NameCate")).SendKeys(xlRange.Cells[7][6]?.Value2?.ToString());

                Assert.IsTrue(driver.Url.Contains("http://localhost:54077/Categories"));
                Console.WriteLine("Successful");
                xlRange.Cells[9][6].Value = "[3] Successful";
            }
            catch (Exception ex)
            {
                Console.WriteLine("Failed\n" + ex);
                xlRange.Cells[9][6].Value = "[3] Failed";
            }
        }

        [Test]
        public void TestIntegrated_DisplayCategoryList()
        {
            try
            {
                //Mở web
                driver = new ChromeDriver();
                driver.Navigate().GoToUrl("http://localhost:54077/Home/Index");
                driver.Manage().Window.Maximize();
                Thread.Sleep(3000);

                //btn thể loại
                driver.FindElement(By.XPath("//a[@class='dropbtn']")).Click();
                Thread.Sleep(2000);

                IWebElement Element = driver.FindElement(By.XPath("//a[normalize-space()='Dân Gian']"));
                Assert.IsTrue(Element.Displayed, "");
                xlRange.Cells[9][5].Value = "[3] Successful";
            }
            catch (Exception ex)
            {
                Console.WriteLine("Failed\n" + ex);
                xlRange.Cells[9][5].Value = "[3] Failed";
            }
        }

        //Tạo mới khi bỏ trống filed (thiếu data)
        [Test]
        public void CreateMissingData()
        {
            try
            {
                LoginAdmin();

                //btn thể loại
                driver.FindElement(By.XPath("//a[contains(text(),'Thể loại')]")).Click();
                Thread.Sleep(2000);

                //btn create
                driver.FindElement(By.XPath("//a[contains(text(),'TẠO MỚI')]")).Click();
                Thread.Sleep(2000);

                //create
                driver.FindElement(By.Id("IDCate")).Clear();
                driver.FindElement(By.Id("NameCate")).Clear();
                driver.FindElement(By.XPath("//input[@value='TẠO MỚI']")).Click();
                Thread.Sleep(2000);

                Assert.IsTrue(driver.Url.Contains("http://localhost:54077/Categories/Create"));
                Console.WriteLine("Successful");
                xlRange.Cells[9][11].Value = "[3] Successful";
            }
            catch (Exception ex)
            {
                Console.WriteLine("Failed\n" + ex);
                xlRange.Cells[9][11].Value = "[3] Failed";
            }
        }

        //Thể loại tạo bị trùng
        [Test]
        public void CreateExistedCategory()
        {
            try
            {
                LoginAdmin();

                //btn thể loại
                driver.FindElement(By.XPath("//a[contains(text(),'Thể loại')]")).Click();
                Thread.Sleep(2000);

                //btn create
                driver.FindElement(By.XPath("//a[contains(text(),'TẠO MỚI')]")).Click();
                Thread.Sleep(2000);

                //create
                driver.FindElement(By.Id("IDCate")).Clear();
                driver.FindElement(By.Id("NameCate")).Clear();

                driver.FindElement(By.Id("IDCate")).SendKeys(xlRange.Cells[6][10]?.Value2?.ToString());
                driver.FindElement(By.Id("NameCate")).SendKeys(xlRange.Cells[7][10]?.Value2?.ToString());
                driver.FindElement(By.XPath("//input[@value='TẠO MỚI']")).Click();
                Thread.Sleep(2000);

                Assert.IsTrue(driver.Url.Contains("http://localhost:54077/Categories/Create"));
                Console.WriteLine("Successful");
                xlRange.Cells[9][10].Value = "[3] Successful";

            }
            catch (Exception ex)
            {
                Console.WriteLine("Failed\n" + ex);
                xlRange.Cells[9][10].Value = "[3] Failed";
            }
        }

        //Tạo thể loại có chứa kí tự đặc biệt
        [Test]
        public void CreateSpeChar()
        {
            try
            {
                LoginAdmin();

                //btn thể loại
                driver.FindElement(By.XPath("//a[contains(text(),'Thể loại')]")).Click();
                Thread.Sleep(2000);

                //btn create
                driver.FindElement(By.XPath("//a[contains(text(),'TẠO MỚI')]")).Click();
                Thread.Sleep(2000);

                //create
                driver.FindElement(By.Id("IDCate")).Clear();
                driver.FindElement(By.Id("NameCate")).Clear();

                driver.FindElement(By.Id("IDCate")).SendKeys(xlRange.Cells[6][12]?.Value2?.ToString());
                driver.FindElement(By.Id("NameCate")).SendKeys(xlRange.Cells[7][12]?.Value2?.ToString());
                driver.FindElement(By.XPath("//input[@value='TẠO MỚI']")).Click();
                Thread.Sleep(2000);

                Assert.IsTrue(driver.Url.Contains("http://localhost:54077/Categories/Create"));
                Console.WriteLine("Successful");
                xlRange.Cells[9][12].Value = "[3] Successful";

            }
            catch (Exception ex)
            {
                Console.WriteLine("Failed\n" + ex);
                xlRange.Cells[9][12].Value = "[3] Failed";
            }
        }

        //Tạo thể loại với ID vượt quá giới hạn
        [Test]
        public void CreateBeyond()
        {
            try
            {
                LoginAdmin();

                //btn thể loại
                driver.FindElement(By.XPath("//a[contains(text(),'Thể loại')]")).Click();
                Thread.Sleep(2000);

                //btn create
                driver.FindElement(By.XPath("//a[contains(text(),'TẠO MỚI')]")).Click();
                Thread.Sleep(2000);

                //create
                driver.FindElement(By.Id("IDCate")).Clear();
                driver.FindElement(By.Id("NameCate")).Clear();

                driver.FindElement(By.Id("IDCate")).SendKeys(xlRange.Cells[6][13]?.Value2?.ToString());
                driver.FindElement(By.Id("NameCate")).SendKeys(xlRange.Cells[7][13]?.Value2?.ToString());
                driver.FindElement(By.XPath("//input[@value='TẠO MỚI']")).Click();
                Thread.Sleep(2000);

                Assert.IsTrue(driver.Url.Contains("http://localhost:54077/Categories/Create"));
                Console.WriteLine("Successful");
                xlRange.Cells[9][13].Value = "[3] Successful";

            }
            catch (Exception ex)
            {
                Console.WriteLine("Failed\n" + ex);
                xlRange.Cells[9][13].Value = "[3] Failed";
            }
        }

        //Thao tác hủy tạo thể loại
        [Test]
        public void CancelCreate()
        {
            try
            {
                LoginAdmin();

                //btn thể loại
                driver.FindElement(By.XPath("//a[contains(text(),'Thể loại')]")).Click();
                Thread.Sleep(2000);

                //btn create
                driver.FindElement(By.XPath("//a[contains(text(),'TẠO MỚI')]")).Click();
                Thread.Sleep(2000);

                //create
                driver.FindElement(By.Id("IDCate")).Clear();
                driver.FindElement(By.Id("NameCate")).Clear();

                driver.FindElement(By.Id("IDCate")).SendKeys(xlRange.Cells[6][6]?.Value2?.ToString());
                driver.FindElement(By.Id("NameCate")).SendKeys(xlRange.Cells[7][6]?.Value2?.ToString());
                driver.FindElement(By.XPath("//a[contains(text(),'Trở lại')]")).Click();
                Thread.Sleep(2000);

                Assert.IsTrue(driver.Url.Contains("http://localhost:54077/Categories"));
                Console.WriteLine("Successful");
                xlRange.Cells[9][14].Value = "[3] Successful";

            }
            catch (Exception ex)
            {
                Console.WriteLine("Failed\n" + ex);
                xlRange.Cells[9][14].Value = "[3] Failed";
            }
        }

        //Xem chi tiết 
        [Test]
        public void ViewDetail()
        {
            try
            {
                LoginAdmin();

                //btn thể loại
                driver.FindElement(By.XPath("//a[contains(text(),'Thể loại')]")).Click();
                Thread.Sleep(2000);

                //btn detail
                driver.FindElement(By.XPath("//tbody/tr[2]/td[3]/a[2]")).Click();
                Thread.Sleep(2000);

                Assert.IsTrue(driver.Url.Contains("http://localhost:54077/Categories/Details/11"));
                Console.WriteLine("Successful");
                xlRange.Cells[9][7].Value = "[3] Successful";
            }
            catch (Exception ex)
            {
                Console.WriteLine("Failed\n" + ex);
                xlRange.Cells[9][7].Value = "[3] Failed";
            }
        }

        //Sửa thể loại 
        [Test]
        public void EditDetail()
        {
            try
            {
                LoginAdmin();

                //btn thể loại
                driver.FindElement(By.XPath("//a[contains(text(),'Thể loại')]")).Click();
                Thread.Sleep(2000);

                //btn edit
                driver.FindElement(By.XPath("//tbody/tr[3]/td[3]/a[1]")).Click();
                Thread.Sleep(2000);

                //edit
                driver.FindElement(By.Id("IDCate")).Clear();
                driver.FindElement(By.Id("NameCate")).Clear();

                driver.FindElement(By.Id("IDCate")).SendKeys(xlRange.Cells[6][7]?.Value2?.ToString());
                driver.FindElement(By.Id("NameCate")).SendKeys(xlRange.Cells[7][7]?.Value2?.ToString());
                driver.FindElement(By.XPath("//input[@value='LƯU']")).Click();
                Thread.Sleep(2000);

                Assert.IsTrue(driver.Url.Contains("http://localhost:54077/Categories"));
                Console.WriteLine("Successful");
                xlRange.Cells[9][8].Value = "[3] Successful";
            }
            catch(Exception ex)
            {
                Console.WriteLine("Failed\n" + ex);
                xlRange.Cells[9][8].Value = "[3] Failed";
            }
        }

        //Thao tác xóa thể loại
        [Test]
        public void Delete()
        {
            try
            {
                LoginAdmin();

                //btn thể loại
                driver.FindElement(By.XPath("//a[contains(text(),'Thể loại')]")).Click();
                Thread.Sleep(2000);

                //btn delete
                driver.FindElement(By.XPath("//tbody/tr[5]/td[3]/a[3]")).Click();
                Thread.Sleep(2000);

                IWebElement bhangElement = driver.FindElement(By.XPath("//td[normalize-space()='Dân Gian']"));
                Console.WriteLine("Successful");
                xlRange.Cells[9][9].Value = "[3] Successful";
            }
            catch (Exception ex)
            {
                Console.WriteLine("Failed\n" + ex);
                xlRange.Cells[9][9].Value = "[3] Failed";
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
