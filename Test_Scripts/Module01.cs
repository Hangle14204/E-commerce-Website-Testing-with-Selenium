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
    public class Module01
    {
        IWebDriver driver;
        IWebElement element, inputField;
        Excel.Application dataApp;
        Excel.Workbook dataBook;
        Excel.Worksheet dataSheet;
        Excel.Range xlRange;

        [SetUp]
        public void Setup()
        {
            //Mở web
            driver = new ChromeDriver();
            driver.Navigate().GoToUrl("http://localhost:54077/");
            driver.Manage().Window.Maximize();
            Thread.Sleep(3000);

            //Mở excel
            dataApp = new Excel.Application();
            dataBook = dataApp.Workbooks.Open(@"D:\\Code\\DBCLPM\\AutoTestforPlan\\Plan_Test\\Plan_Test\\Data_Test\\Data_Report.xlsx");
            dataSheet = dataBook.Sheets[2];
            xlRange = dataSheet.UsedRange;

        }

        //Đăng kí tài khoản User mới
        [Test]
        public void CreateAccount()
        {
            try
            {
                //btn User
                driver.FindElement(By.XPath("//a[@class='fas fa-user']")).Click();
                Thread.Sleep(2000);

                //chuyển trang đăng kí 
                driver.FindElement(By.XPath("//a[contains(text(),'Tạo tài khoản')]")).Click();
                Thread.Sleep(2000);

                //Đăng kí
                driver.FindElement(By.Id("NameCus")).Clear();
                driver.FindElement(By.Id("EmailCus")).Clear();
                driver.FindElement(By.Id("PhoneCus")).Clear();
                driver.FindElement(By.Id("PassCus")).Clear();

                driver.FindElement(By.Id("NameCus")).SendKeys(xlRange.Cells[6][5]?.Value2?.ToString());
                driver.FindElement(By.Id("EmailCus")).SendKeys(xlRange.Cells[7][5]?.Value2?.ToString());
                driver.FindElement(By.Id("PhoneCus")).SendKeys(xlRange.Cells[8][5]?.Value2?.ToString());
                driver.FindElement(By.Id("PassCus")).SendKeys(xlRange.Cells[9][5]?.Value2?.ToString());

                driver.FindElement(By.XPath("//input[@value='Đăng ký']")).Click();
                Thread.Sleep(2000);
                Console.WriteLine("Successful");
                xlRange.Cells[11][5].Value = "[3] Successful";
            }
            catch(Exception ex)
            {
                Console.WriteLine("Failed\n" + ex);
                xlRange.Cells[11][5].Value = "[3] Failed";
            }
        }

        //Kiểm tra account vừa tạo có nằm trong list account bên user không
        [Test]
        public void TestIntegrated_CreateAccount()
        {
            LoginAdmin();

            //btn khách hàng
            driver.FindElement(By.XPath("//a[normalize-space()='Khách hàng']")).Click();
            Thread.Sleep(2000);

            try
            {
                IWebElement bhangElement = driver.FindElement(By.XPath("//td[normalize-space()='Bhang']"));
                Assert.IsTrue(bhangElement.Displayed, "Tài khoản đã được tạo thành công.");
                xlRange.Cells[11][10].Value = "[3] Successful";
            }
            catch (NoSuchElementException)
            {
                Assert.Fail("Không tìm thấy tài khoản Bhang. Kiểm tra thất bại.");
                xlRange.Cells[11][10].Value = "[3] Failed";
            }
        }

        //Đăng kí khi thiếu data
        [Test]
        public void CreateUserMissingData()
        {
            try
            {
                //btn User
                driver.FindElement(By.XPath("//a[@class='fas fa-user']")).Click();
                Thread.Sleep(2000);

                //chuyển trang đăng kí 
                driver.FindElement(By.XPath("//a[contains(text(),'Tạo tài khoản')]")).Click();
                Thread.Sleep(2000);

                //Đăng kí
                driver.FindElement(By.Id("NameCus")).Clear();
                driver.FindElement(By.Id("EmailCus")).Clear();
                driver.FindElement(By.Id("PhoneCus")).Clear();
                driver.FindElement(By.Id("PassCus")).Clear();
                
                driver.FindElement(By.Id("EmailCus")).SendKeys(xlRange.Cells[7][6]?.Value2?.ToString());
                driver.FindElement(By.Id("PassCus")).SendKeys(xlRange.Cells[9][6]?.Value2?.ToString());


                driver.FindElement(By.XPath("//input[@value='Đăng ký']")).Click();
                Thread.Sleep(2000);

                Assert.IsTrue(driver.Url.Contains("http://localhost:54077/Home/Register"));
                Console.WriteLine("Successful");
                xlRange.Cells[11][6].Value = "[3] Successful";
            }
            catch (Exception ex)
            {
                Console.WriteLine("Failed");
                xlRange.Cells[11][6].Value = "[3] Failed";
            }
        }

        //Đăng kí với UserName đã tồn tại
        [Test]
        public void CreateExistedUsername()
        {
            try
            {
                //btn User
                driver.FindElement(By.XPath("//a[@class='fas fa-user']")).Click();
                Thread.Sleep(2000);

                //chuyển trang đăng kí 
                driver.FindElement(By.XPath("//a[contains(text(),'Tạo tài khoản')]")).Click();
                Thread.Sleep(2000);

                //Đăng kí
                driver.FindElement(By.Id("NameCus")).Clear();
                driver.FindElement(By.Id("EmailCus")).Clear();
                driver.FindElement(By.Id("PhoneCus")).Clear();
                driver.FindElement(By.Id("PassCus")).Clear();

                driver.FindElement(By.Id("NameCus")).SendKeys(xlRange.Cells[6][7]?.Value2?.ToString());
                driver.FindElement(By.Id("EmailCus")).SendKeys(xlRange.Cells[7][7]?.Value2?.ToString());
                driver.FindElement(By.Id("PhoneCus")).SendKeys(xlRange.Cells[8][7]?.Value2?.ToString());
                driver.FindElement(By.Id("PassCus")).SendKeys(xlRange.Cells[9][7]?.Value2?.ToString());

                driver.FindElement(By.XPath("//input[@value='Đăng ký']")).Click();
                Thread.Sleep(2000);

                Assert.IsTrue(driver.Url.Contains("http://localhost:54077/Home/Register"));
                Console.WriteLine("Successful");
                xlRange.Cells[11][7].Value = "[3] Successful";
            }
            catch (Exception ex)
            {
                Console.WriteLine("Failed\n" + ex);
                xlRange.Cells[11, 7].Value = "[3] Failed";
            }
        }

        //Đăng kí với Email đã tồn tại
        [Test]
        public void CreateExistedEmail()
        {
            try
            {
                //btn User
                driver.FindElement(By.XPath("//a[@class='fas fa-user']")).Click();
                Thread.Sleep(2000);

                //chuyển trang đăng kí 
                driver.FindElement(By.XPath("//a[contains(text(),'Tạo tài khoản')]")).Click();
                Thread.Sleep(2000);

                //Đăng kí
                driver.FindElement(By.Id("NameCus")).Clear();
                driver.FindElement(By.Id("EmailCus")).Clear();
                driver.FindElement(By.Id("PhoneCus")).Clear();
                driver.FindElement(By.Id("PassCus")).Clear();

                driver.FindElement(By.Id("NameCus")).SendKeys(xlRange.Cells[6][8]?.Value2?.ToString());
                driver.FindElement(By.Id("EmailCus")).SendKeys(xlRange.Cells[7][8]?.Value2?.ToString());
                driver.FindElement(By.Id("PhoneCus")).SendKeys(xlRange.Cells[8][8]?.Value2?.ToString());
                driver.FindElement(By.Id("PassCus")).SendKeys(xlRange.Cells[9][8]?.Value2?.ToString());

                driver.FindElement(By.XPath("//input[@value='Đăng ký']")).Click();
                Thread.Sleep(2000);

                Assert.IsTrue(driver.Url.Contains("http://localhost:54077/Home/Register"));
                Console.WriteLine("Successful");
                xlRange.Cells[11][8].Value = "[3] Successful";
            }
            catch (Exception ex)
            {
                Console.WriteLine("Failed\n" + ex);
                xlRange.Cells[11][8].Value = "[3] Failed";
            }
        }

        //Đăng kí với số điện thoại đã tồn tại
        [Test]
        public void CreateExistedPhone()
        {
            try
            {
                //btn User
                driver.FindElement(By.XPath("//a[@class='fas fa-user']")).Click();
                Thread.Sleep(2000);

                //chuyển trang đăng kí 
                driver.FindElement(By.XPath("//a[contains(text(),'Tạo tài khoản')]")).Click();
                Thread.Sleep(2000);

                //Đăng kí
                driver.FindElement(By.Id("NameCus")).Clear();
                driver.FindElement(By.Id("EmailCus")).Clear();
                driver.FindElement(By.Id("PhoneCus")).Clear();
                driver.FindElement(By.Id("PassCus")).Clear();

                driver.FindElement(By.Id("NameCus")).SendKeys(xlRange.Cells[6][9]?.Value2?.ToString());
                driver.FindElement(By.Id("EmailCus")).SendKeys(xlRange.Cells[7][9]?.Value2?.ToString());
                driver.FindElement(By.Id("PhoneCus")).SendKeys(xlRange.Cells[8][9]?.Value2?.ToString());
                driver.FindElement(By.Id("PassCus")).SendKeys(xlRange.Cells[9][9]?.Value2?.ToString());

                driver.FindElement(By.XPath("//input[@value='Đăng ký']")).Click();
                Thread.Sleep(2000);

                Assert.IsTrue(driver.Url.Contains("http://localhost:54077/Home/Register"));
                Console.WriteLine("Successful");
                xlRange.Cells[11][9].Value = "[3] Successful";
            }
            catch (Exception ex)
            {
                Console.WriteLine("Failed\n" + ex);
                xlRange.Cells[11][9].Value = "[3] Failed";
            }
        }

        // Đăng nhập với thông tin đúng
        [Test]
        public void LoginUser()
        {
            try
            {
                //btn User
                driver.FindElement(By.XPath("//a[@class='fas fa-user']")).Click();
                Thread.Sleep(2000);

                //đăng nhập
                driver.FindElement(By.Id("EmailCus")).Clear();
                driver.FindElement(By.Id("PassCus")).Clear();
                driver.FindElement(By.Id("EmailCus")).SendKeys(xlRange.Cells[7][17]?.Value2?.ToString());
                driver.FindElement(By.Id("PassCus")).SendKeys(xlRange.Cells[9][17]?.Value2?.ToString());
                driver.FindElement(By.XPath("//input[@value='Đăng nhập']")).Click();
                Thread.Sleep(2000);

                Assert.IsTrue(driver.Url.Contains("http://localhost:54077/"));
                Console.WriteLine("Successful");
                xlRange.Cells[11][17].Value = "[3] Successful";
            }
            catch (Exception ex)
            {
                Console.WriteLine("Failed" + ex);
                xlRange.Cells[11][17].Value = "[3] Failed";
            }
        }

        // Đăng nhập với thông tin sai
        [Test]
        public void TestLogin()
        {
            try
            {
                //btn User
                driver.FindElement(By.XPath("//a[@class='fas fa-user']")).Click();
                Thread.Sleep(2000);

                //đăng nhập
                driver.FindElement(By.Id("EmailCus")).Clear();
                driver.FindElement(By.Id("PassCus")).Clear();
                driver.FindElement(By.Id("EmailCus")).SendKeys(xlRange.Cells[7][18]?.Value2?.ToString());
                driver.FindElement(By.Id("PassCus")).SendKeys(xlRange.Cells[9][18]?.Value2?.ToString());
                driver.FindElement(By.XPath("//input[@value='Đăng nhập']")).Click();
                Thread.Sleep(2000);

                Assert.IsTrue(driver.Url.Contains("http://localhost:54077/Home/Login")); 
                Console.WriteLine("Successful");
                xlRange.Cells[11][18].Value = "[3] Successful";
                xlRange.Cells[11][19].Value = "[3] Successful";
            }
            catch (Exception ex)
            {
                Console.WriteLine("Failed\n" + ex);
                xlRange.Cells[11][18].Value = "[3] Failed";
                xlRange.Cells[11][19].Value = "[3] Failed";
            }
        }

        //Xóa tài khoản User
        [Test]
        public void DeleteUser()
        {
            try
            {
                LoginAdmin();

                //btn khách hàng
                driver.FindElement(By.XPath("//a[normalize-space()='Khách hàng']")).Click();
                Thread.Sleep(2000);

                //btn delete
                driver.FindElement(By.XPath("//tbody/tr[2]/td[4]/a[1]")).Click();
                Thread.Sleep(2000);

                //delete
                driver.FindElement(By.XPath("//input[@value='Delete']")).Click();
                Thread.Sleep(2000);

                Assert.IsTrue(driver.Url.Contains("http://localhost:54077/Customer/Delete/1"));
                Console.WriteLine("Successful");
                xlRange.Cells[11][28].Value = "[3] Successful";
            }
            catch (Exception ex)
            {
                Console.WriteLine("Failed\n" + ex);
                xlRange.Cells[11][28].Value = "[3] Failed";
            }
        }

        //Kiểm tra xem tài khoản User vừa xóa còn hiển thị trên danh sách không
        [Test]
        public void TestIntegrated_DeleteUser()
        {
            LoginAdmin();

            //btn khách hàng
            driver.FindElement(By.XPath("//a[normalize-space()='Khách hàng']")).Click();
            Thread.Sleep(2000);

            try
            {
                IWebElement bhangElement = driver.FindElement(By.XPath("//td[normalize-space()='Bhang']"));
                Assert.IsTrue(bhangElement.Displayed, "Failed");
                Console.WriteLine("Failed");
                xlRange.Cells[11][32].Value = "[3] Successful";
            }
            catch (NoSuchElementException)
            {
                Assert.Fail("Successful");
                xlRange.Cells[11][32].Value = "[3] Failed";
            }
        }

        //Tạo tài khoản Admin mới
        [Test]
        public void CreateAdmin()
        {
            try
            {
                LoginAdmin();

                //btn tài khoản
                driver.FindElement(By.XPath("//a[contains(text(),'Tài khoản')]")).Click();
                Thread.Sleep(2000);

                //btn create
                driver.FindElement(By.XPath("//a[@class='btn btn-primary']")).Click();
                Thread.Sleep(2000);

                //create
                driver.FindElement(By.Id("NameUser")).Clear();
                driver.FindElement(By.Id("RoleUser")).Clear();
                driver.FindElement(By.Id("PasswordUser")).Clear();

                driver.FindElement(By.Id("NameUser")).SendKeys(xlRange.Cells[6][11]?.Value2?.ToString());
                driver.FindElement(By.Id("RoleUser")).SendKeys(xlRange.Cells[7][11]?.Value2?.ToString());
                driver.FindElement(By.Id("PasswordUser")).SendKeys(xlRange.Cells[9][11]?.Value2?.ToString());

                driver.FindElement(By.XPath("//input[@value='Create']")).Click();
                Thread.Sleep(2000);

                Console.WriteLine("Failed");
                xlRange.Cells[11][11].Value = "[3] Successful";
            }
            catch(Exception ex)
            {
                Console.WriteLine("Failed\n"+ ex.Message);
                xlRange.Cells[11][11].Value = "[3] Failed";
            }
        }

        //Kiểm tra xem tài khoản vừa tạo có trong danh sách không
        [Test]
        public void TestIntegrated_CreateAdmin()
        {
            LoginAdmin();

            //btn tài khoản
            driver.FindElement(By.XPath("//a[contains(text(),'Tài khoản')]")).Click();
            Thread.Sleep(2000);

            try
            {
                IWebElement bhangElement = driver.FindElement(By.XPath("//td[contains(text(),'Minh Hiếu')]"));
                Assert.IsTrue(bhangElement.Displayed, "Tài khoản đã được tạo thành công.");
                xlRange.Cells[11][15].Value = "[3] Successful";
            }
            catch (NoSuchElementException)
            {
                Assert.Fail("Không tìm thấy tài khoản Bhang. Kiểm tra thất bại.");
                xlRange.Cells[11][15].Value = "[3] Failed";
            }
        }

        //Đăng kí khi thiếu data
        [Test]
        public void CreateAdminMissingData()
        {
            try
            {
                LoginAdmin();

                //btn tài khoản
                driver.FindElement(By.XPath("//a[contains(text(),'Tài khoản')]")).Click();
                Thread.Sleep(2000);

                //btn create
                driver.FindElement(By.XPath("//a[@class='btn btn-primary']")).Click();
                Thread.Sleep(2000);

                //create
                driver.FindElement(By.Id("NameUser")).Clear();
                driver.FindElement(By.Id("RoleUser")).Clear();
                driver.FindElement(By.Id("PasswordUser")).Clear();

                driver.FindElement(By.XPath("//input[@value='Create']")).Click();
                Thread.Sleep(2000);

                Assert.IsTrue(driver.Url.Contains("http://localhost:54077/AdminUser/Create"));
                Console.WriteLine("Successful");
                xlRange.Cells[11][12].Value = "[3] Successful";
            }
            catch (Exception ex)
            {
                Console.WriteLine("Failed");
                xlRange.Cells[11][12].Value = "[3] Failed";
            }
        }

        //Tạo tài khoản Admin nhưng trùng UserName
        [Test]
        public void CreateAdminExistedUsername()
        {
            try
            {
                LoginAdmin();

                //btn tài khoản
                driver.FindElement(By.XPath("//a[contains(text(),'Tài khoản')]")).Click();
                Thread.Sleep(2000);

                //btn create
                driver.FindElement(By.XPath("//a[@class='btn btn-primary']")).Click();
                Thread.Sleep(2000);

                //create
                driver.FindElement(By.Id("NameUser")).Clear();
                driver.FindElement(By.Id("RoleUser")).Clear();
                driver.FindElement(By.Id("PasswordUser")).Clear();

                driver.FindElement(By.Id("NameUser")).SendKeys(xlRange.Cells[6][13]?.Value2?.ToString());
                driver.FindElement(By.Id("RoleUser")).SendKeys(xlRange.Cells[7][13]?.Value2?.ToString());
                driver.FindElement(By.Id("PasswordUser")).SendKeys(xlRange.Cells[9][13]?.Value2?.ToString());

                driver.FindElement(By.XPath("//input[@value='Create']")).Click();
                Thread.Sleep(2000);

                Assert.IsTrue(driver.Url.Contains("http://localhost:54077/AdminUser/Create"));
                Console.WriteLine("Failed");
                xlRange.Cells[11][13].Value = "[3] Successful";
            }
            catch (Exception ex)
            {
                Console.WriteLine("Failed\n" + ex.Message);
                xlRange.Cells[11][13].Value = "[3] Failed";
            }
        }

        //Tạo tài khoản Admin nhưng khác role
        [Test]
        public void CreateAdminWrongRole()
        {
            try
            {
                LoginAdmin();

                //btn tài khoản
                driver.FindElement(By.XPath("//a[contains(text(),'Tài khoản')]")).Click();
                Thread.Sleep(2000);

                //btn create
                driver.FindElement(By.XPath("//a[@class='btn btn-primary']")).Click();
                Thread.Sleep(2000);

                //create
                driver.FindElement(By.Id("NameUser")).Clear();
                driver.FindElement(By.Id("RoleUser")).Clear();
                driver.FindElement(By.Id("PasswordUser")).Clear();

                driver.FindElement(By.Id("NameUser")).SendKeys(xlRange.Cells[6][14]?.Value2?.ToString());
                driver.FindElement(By.Id("RoleUser")).SendKeys(xlRange.Cells[7][14]?.Value2?.ToString());
                driver.FindElement(By.Id("PasswordUser")).SendKeys(xlRange.Cells[9][14]?.Value2?.ToString());

                driver.FindElement(By.XPath("//input[@value='Create']")).Click();
                Thread.Sleep(2000);

                Assert.IsTrue(driver.Url.Contains("http://localhost:54077/AdminUser/Create"));
                Console.WriteLine("Failed");
                xlRange.Cells[11][14].Value = "[3] Successful";
            }
            catch (Exception ex)
            {
                Console.WriteLine("Failed\n" + ex.Message);
                xlRange.Cells[11][14].Value = "[3] Failed";
            }
        }

        //Xóa tài khoản admin
        [Test]
        public void DeleteAdmin()
        {
            try
            {
                LoginAdmin();

                //btn tài khoản
                driver.FindElement(By.XPath("//a[contains(text(),'Tài khoản')]")).Click();
                Thread.Sleep(2000);

                //btn delete
                driver.FindElement(By.XPath("//tbody/tr[5]/td[4]/a[2]]")).Click();
                Thread.Sleep(2000);

                //delete
                driver.FindElement(By.XPath("//input[@value='Delete']")).Click();
                Thread.Sleep(2000);

                Assert.IsTrue(driver.Url.Contains("http://localhost:54077/AdminUser"));
                Console.WriteLine("Successfull");
                xlRange.Cells[11][30].Value = "[3] Successful";
            }
            catch (Exception ex)
            {
                Console.WriteLine("Failed\n" + ex);
                xlRange.Cells[11][30].Value = "[3] Failed";
            }
        }


        //Đăng nhập Admin
        [Test]
        public void LoginAdmin()
        {
            try
            {
                driver = new ChromeDriver();
                driver.Navigate().GoToUrl("http://localhost:54077/AdminUser/Login");
                driver.Manage().Window.Maximize();
                Thread.Sleep(3000);

                driver.FindElement(By.Id("NameUser")).Clear();
                driver.FindElement(By.Id("NameUser")).SendKeys(xlRange.Cells[6][20]?.Value2?.ToString());
                driver.FindElement(By.Id("PasswordUser")).Clear();
                driver.FindElement(By.Id("PasswordUser")).SendKeys(xlRange.Cells[9][20]?.Value2?.ToString());
                driver.FindElement(By.XPath("//button[@type='submit']")).Click();
                Thread.Sleep(2000);
                xlRange.Cells[11][20].Value = "[3] Successful";
            }
            catch(Exception ex)
            {
                Console.WriteLine("Failed\n" + ex);
                xlRange.Cells[11][20].Value = "[3] Failed";
            }
        }

        [Test]
        public void TestLoginAdmin()
        {
            try
            {
                driver = new ChromeDriver();
                driver.Navigate().GoToUrl("http://localhost:54077/AdminUser/Login");
                driver.Manage().Window.Maximize();
                Thread.Sleep(3000);

                //đăng nhập
                driver.FindElement(By.Id("NameUser")).Clear();
                driver.FindElement(By.Id("PasswordUser")).Clear();
                driver.FindElement(By.Id("NameUser")).SendKeys(xlRange.Cells[6][22]?.Value2?.ToString());
                driver.FindElement(By.Id("PasswordUser")).SendKeys(xlRange.Cells[9][21]?.Value2?.ToString());
                driver.FindElement(By.XPath("//button[@type='submit']")).Click();
                Thread.Sleep(2000);

                Assert.IsTrue(driver.Url.Contains("http://localhost:54077/AdminUser/Login"));
                Console.WriteLine("Successful");
                xlRange.Cells[11][21].Value = "[3] Successful";
                xlRange.Cells[11][22].Value = "[3] Successful";
            }
            catch (Exception ex)
            {
                Console.WriteLine("Failed\n" + ex);
                xlRange.Cells[11][21].Value = "[3] Failed";
                xlRange.Cells[11][22].Value = "[3] Failed";
            }
        }

        //Kiểm tra xem tài khoản còn trong danh sách không
        /*[Test]
        public void TestIntegrated_DeleteAdmin()
        {
            LoginAdmin();

            //btn tài khoản
            driver.FindElement(By.XPath("//a[contains(text(),'Tài khoản')]")).Click();
            Thread.Sleep(2000);

            try
            {
                IWebElement bhangElement = driver.FindElement(By.XPath("//td[normalize-space()='BichHang']"));
                Assert.IsTrue(bhangElement.Displayed, "Failed");
                xlRange.Cells[11][34].Value = "[3] Successful";
            }
            catch (NoSuchElementException)
            {
                Assert.Fail("Successful");
                xlRange.Cells[11][34].Value = "[3] Failed";
            }
        }*/

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
