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
    public class Module04
    {
        IWebDriver driver;
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
            dataSheet = dataBook.Sheets[5];
            xlRange = dataSheet.UsedRange;

        }

        //Tạo sản phẩm mới
        [Test]
        public void CreateProduct()
        {
            try
            {
                LoginAdmin();

                //btn sản phẩm
                driver.FindElement(By.XPath("//a[contains(text(),'Sản phẩm')]")).Click();
                Thread.Sleep(2000);

                //btn create
                driver.FindElement(By.XPath("//a[contains(text(),'TẠO MỚI')]")).Click();
                Thread.Sleep(2000);

                driver.FindElement(By.Id("NamePro")).Clear();
                driver.FindElement(By.Id("DecriptionPro")).Clear();
                driver.FindElement(By.Id("Price")).Clear();

                driver.FindElement(By.Id("NamePro")).SendKeys(xlRange.Cells[6][5]?.Value2?.ToString());
                driver.FindElement(By.Id("DecriptionPro")).SendKeys(xlRange.Cells[7][5]?.Value2?.ToString());
                Thread.Sleep(2000);
                IWebElement dropdownElement = driver.FindElement(By.Id("Category"));
                SelectElement dropdown = new SelectElement(dropdownElement);
                dropdown.SelectByText("Thiếu nhi");
                driver.FindElement(By.Id("Price")).SendKeys(xlRange.Cells[9][5]?.Value2?.ToString());
                driver.FindElement(By.XPath("//input[@name='ImagePro']")).Click();
                Thread.Sleep(2000);

                // Tìm file ảnh trong thư mục Pictures
                string picturesPath = Environment.GetFolderPath(Environment.SpecialFolder.MyPictures);
                string[] imageFiles = Directory.GetFiles(picturesPath, "*.jpg"); // Lấy ảnh JPG (có thể thêm *.png)
                if (imageFiles.Length == 0)
                {
                    Console.WriteLine("Không tìm thấy ảnh nào trong thư mục Pictures!");
                    driver.Quit();
                    return;
                }
                string imagePath = imageFiles[0]; // Chọn ảnh đầu tiên để upload
                Console.WriteLine("Uploading: " + imagePath);
                // Tìm input file upload trên trang (thay đổi nếu cần)
                IWebElement fileInput = driver.FindElement(By.CssSelector("input[name='driver.FindElement(By.XPath(']"));
                fileInput.SendKeys(imagePath); // Gửi đường dẫn file ảnh vào ô upload
                Thread.Sleep(2000);

                driver.FindElement(By.ClassName("//input[@value='TẠO MỚI']")).Click();
                Thread.Sleep(2000);
                //Assert.IsTrue(driver.Url.Contains("http://localhost:54077/Home/Login"));
                Console.WriteLine("Successful");
                xlRange.Cells[11][5].Value = "[3] Successful";
            }
            catch (Exception ex)
            {
                Console.WriteLine("Failed\n" + ex);
                xlRange.Cells[11][5].Value = "[3] Failed";
            }
        }

        //Chỉnh sửa sản phẩm
        [Test]
        public void EditProduct()
        {
            try
            {
                LoginAdmin();

                //btn sản phẩm
                driver.FindElement(By.XPath("//a[contains(text(),'Sản phẩm')]")).Click();
                Thread.Sleep(2000);

                //btn edit
                driver.FindElement(By.XPath("//tbody/tr[2]/td[6]/a[1]")).Click();
                Thread.Sleep(2000);

                driver.FindElement(By.Id("NamePro")).Clear();
                driver.FindElement(By.Id("NamePro")).SendKeys(xlRange.Cells[6][9]?.Value2?.ToString());
                
                //btn edit
                driver.FindElement(By.XPath("//input[@value='LƯU']")).Click();
                Thread.Sleep(2000);
                Assert.IsTrue(driver.Url.Contains("http://localhost:54077/Products"));
                Console.WriteLine("Successful");
                xlRange.Cells[11][9].Value = "[3] Successful";
            }
            catch (Exception ex)
            {
                Console.WriteLine("Failed\n" + ex);
                xlRange.Cells[11][9].Value = "[3] Failed";
            }
        }

        //Kiểm tra xem sau khi chỉnh sửa, sản phẩm có được hiển thị chính xác không
        [Test]
        public void TestIntegrated_EditProduct()
        {
            driver = new ChromeDriver();
            driver.Navigate().GoToUrl("http://localhost:54077/Home");
            driver.Manage().Window.Maximize();
            Thread.Sleep(3000);

            try
            {
                //cuộn tìm phần tử
                IJavaScriptExecutor js = (IJavaScriptExecutor)driver;
                IWebElement element = driver.FindElement(By.XPath("//body/div/section[3]/div[1]/div[1]/div[1]/div[2]"));
                js.ExecuteScript("arguments[0].scrollIntoView(true);", element);
                Thread.Sleep(2000);

                IWebElement Element = driver.FindElement(By.XPath("//h3[contains(text(),'Bà lão')]"));
                Assert.IsTrue(Element.Displayed, "Successfull");
                Console.WriteLine("Successful");
                xlRange.Cells[11][10].Value = "[3] Successful";
            }
            catch (Exception ex)
            {
                Console.WriteLine("Failed\n" + ex);
                xlRange.Cells[11][10].Value = "[3] Failed";
            }
        }

        //Hủy thao tác sửa thông tin sản phẩm
        [Test]
        public void CancelEditProduct()
        {
            try
            {
                LoginAdmin();

                //btn sản phẩm
                driver.FindElement(By.XPath("//a[contains(text(),'Sản phẩm')]")).Click();
                Thread.Sleep(2000);

                //btn edit
                driver.FindElement(By.XPath("//tbody/tr[2]/td[6]/a[1]")).Click();
                Thread.Sleep(2000);

                driver.FindElement(By.Id("NamePro")).Clear();
                driver.FindElement(By.Id("NamePro")).SendKeys("Bà lão");

                //btn cancel
                driver.FindElement(By.XPath("//a[contains(text(),'Trở lại')]")).Click();
                Thread.Sleep(2000);
                Assert.IsTrue(driver.Url.Contains("http://localhost:54077/Products"));
                Console.WriteLine("Successful");
                xlRange.Cells[11][11].Value = "[3] Successful";
            }
            catch (Exception ex)
            {
                Console.WriteLine("Failed\n" + ex);
                xlRange.Cells[11][11].Value = "[3] Failed";
            }
        }

        //Tạo sản phẩm mới nhưng thiếu data
        [Test]
        public void CreateMissingData()
        {
            try
            {
                LoginAdmin();

                //btn sản phẩm
                driver.FindElement(By.XPath("//a[contains(text(),'Sản phẩm')]")).Click();
                Thread.Sleep(2000);

                //btn create
                driver.FindElement(By.XPath("//a[contains(text(),'TẠO MỚI')]")).Click();
                Thread.Sleep(2000);

                driver.FindElement(By.Id("NamePro")).Clear();
                driver.FindElement(By.Id("DecriptionPro")).Clear();
                driver.FindElement(By.Id("Price")).Clear();

                driver.FindElement(By.Id("NamePro")).SendKeys(xlRange.Cells[6][5]?.Value2?.ToString());
                driver.FindElement(By.Id("DecriptionPro")).SendKeys(xlRange.Cells[7][5]?.Value2?.ToString());
                Thread.Sleep(2000);
                IWebElement dropdownElement = driver.FindElement(By.Id("Category"));
                SelectElement dropdown = new SelectElement(dropdownElement);
                dropdown.SelectByText("Thiếu nhi");

                driver.FindElement(By.ClassName("//input[@value='TẠO MỚI']")).Click();
                Thread.Sleep(2000);
                Assert.IsTrue(driver.Url.Contains("http://localhost:54077/Products/Create"));
                Console.WriteLine("Successful");
                xlRange.Cells[11][12].Value = "[3] Successful";
            }
            catch (Exception ex)
            {
                Console.WriteLine("Failed\n" + ex);
                xlRange.Cells[11][12].Value = "[3] Failed";
            }
        }

        //Hủy tạo sản phẩm
        [Test]
        public void CancelCreate()
        {
            try
            {
                LoginAdmin();

                //btn sản phẩm
                driver.FindElement(By.XPath("//a[contains(text(),'Sản phẩm')]")).Click();
                Thread.Sleep(2000);

                //btn create
                driver.FindElement(By.XPath("//a[contains(text(),'TẠO MỚI')]")).Click();
                Thread.Sleep(2000);

                driver.FindElement(By.Id("NamePro")).Clear();
                driver.FindElement(By.Id("DecriptionPro")).Clear();
                driver.FindElement(By.Id("Price")).Clear();

                driver.FindElement(By.Id("NamePro")).SendKeys(xlRange.Cells[6][5]?.Value2?.ToString());
                driver.FindElement(By.Id("DecriptionPro")).SendKeys(xlRange.Cells[7][5]?.Value2?.ToString());
                Thread.Sleep(2000);

                //btn cancel
                driver.FindElement(By.XPath("//a[contains(text(),'Trở lại')]")).Click();
                Thread.Sleep(2000);
                Assert.IsTrue(driver.Url.Contains("http://localhost:54077/Products"));
                Console.WriteLine("Successful");
                xlRange.Cells[11][6].Value = "[3] Successful";
            }
            catch (Exception ex)
            {
                Console.WriteLine("Failed\n" + ex);
                xlRange.Cells[11][6].Value = "[3] Failed";
            }
        }

        //Xem thông tin chi tiết sản phẩm
        [Test]
        public void DisplayDetail()
        {
            try
            {
                LoginAdmin();

                //btn sản phẩm
                driver.FindElement(By.XPath("//a[contains(text(),'Sản phẩm')]")).Click();
                Thread.Sleep(2000);

                //btn view
                driver.FindElement(By.XPath("//tbody/tr[2]/td[6]/a[2]")).Click();
                Thread.Sleep(2000);
                Assert.IsTrue(driver.Url.Contains("http://localhost:54077/Products/Details/7"));
                Console.WriteLine("Successful");
                xlRange.Cells[11][13].Value = "[3] Successful";
            }
            catch(Exception ex)
            {
                Console.WriteLine("Failed\n" + ex);
                xlRange.Cells[11][13].Value = "[3] Failed";
            }
        }

        //Xóa sản phẩm
        [Test]
        public void DeleteProduct()
        {
            try
            {
                LoginAdmin();

                //btn sản phẩm
                driver.FindElement(By.XPath("//a[contains(text(),'Sản phẩm')]")).Click();
                Thread.Sleep(2000);

                //btn delete
                driver.FindElement(By.XPath("//tbody/tr[3]/td[6]/a[3]")).Click();
                Thread.Sleep(2000);
                Console.WriteLine("Successful");
                xlRange.Cells[11][14].Value = "[3] Successful";
            }
            catch (Exception ex)
            {
                Console.WriteLine("Failed\n" + ex);
                xlRange.Cells[11][14].Value = "[3] Failed";
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
