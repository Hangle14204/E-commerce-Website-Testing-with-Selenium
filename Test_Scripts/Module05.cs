using Microsoft.Office.Interop.Excel;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using NUnit.Framework;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using System;
using System.Threading;
using static NUnit.Framework.Constraints.Tolerance;
using Excel = Microsoft.Office.Interop.Excel;

namespace Plan_Test
{
    [TestFixture]
    public class Module05
    {
        IWebDriver driver;
        Excel.Application dataApp;
        Excel.Workbook dataBook;
        Excel.Worksheet dataSheet;
        Excel.Range xlRange;

        [SetUp]
        public void SetUp()
        {
            //Mở trang web test 
            driver = new ChromeDriver();
            driver.Navigate().GoToUrl("http://localhost:54077/Home");
            driver.Manage().Window.Maximize();
            Thread.Sleep(3000);

            //Mở file excel 
            dataApp = new Excel.Application();
            dataBook = dataApp.Workbooks.Open(@"D:\\Code\\DBCLPM\\AutoTestforPlan\\Plan_Test\\Plan_Test\\Data_Test\\Data_Report.xlsx");
            dataSheet = dataBook.Sheets[6]; // chọn sheet số 6 trong file excel 
            xlRange = dataSheet.Cells;

        }
        [Test]
        public void TimkiemId46()
        {
            string keyword = xlRange.Cells[8, 7].Value?.ToString(); // Lấy dữ liệu từ dòng 8 cột 7
            if (string.IsNullOrEmpty(keyword))
            {
                Console.WriteLine("Không có từ khóa để tìm kiếm.");
                return;
            }

            Console.WriteLine($"Đang tìm kiếm: {keyword}");

            try
            {
                // Nhập từ khóa vào ô tìm kiếm
                IWebElement searchField = driver.FindElement(By.XPath("//input[@id='search-box']"));
                searchField.Clear();
                searchField.SendKeys(keyword);
                Thread.Sleep(500);

                // Click vào icon tìm kiếm
                driver.FindElement(By.XPath("//i[@class='fas fa-search']")).Click();
                Thread.Sleep(2000);

                // Kiểm tra danh sách sản phẩm hiển thị
                var products = driver.FindElements(By.ClassName("box")); // Thay bằng class thực tế
                if (products.Count > 0)
                {
                    xlRange.Cells[8, 9].Value = "[2]Passed"; // Ghi vào dòng 8 cột 9
                }
                else
                {
                    xlRange.Cells[8, 9].Value = "[2]Failed";
                }
                Console.WriteLine("Kết quả: " + xlRange.Cells[8, 9].Value);
            }
            catch (Exception ex)
            {
                // Ghi "Failed" nếu có lỗi
                xlRange.Cells[8, 9].Value = "[2]Failed";
                Console.WriteLine($"Lỗi xảy ra: {ex.Message}");
            }

        }
        [Test]
        public void TimkiemId47()
        {
            string keyword = xlRange.Cells[9, 7].Value?.ToString(); // Lấy dữ liệu từ dòng 9 cột 7 
            if (string.IsNullOrEmpty(keyword))
            {
                Console.WriteLine("Không có từ khóa để tìm kiếm.");
                return;
            }

            Console.WriteLine($"Đang tìm kiếm: {keyword}");

            try
            {
                // Nhập từ khóa vào ô tìm kiếm
                IWebElement searchField = driver.FindElement(By.XPath("//input[@id='search-box']"));
                searchField.Clear();
                searchField.SendKeys(keyword);
                Thread.Sleep(500);

                // Click vào icon tìm kiếm
                driver.FindElement(By.XPath("//i[@class='fas fa-search']")).Click();
                Thread.Sleep(2000);

                // Kiểm tra danh sách sản phẩm hiển thị
                var products = driver.FindElements(By.ClassName("box")); // Thay bằng class thực tế

                if (products.Count > 0)
                {
                    xlRange.Cells[9, 9].Value = "[2]Passed"; // Ghi vào dòng 9 cột 9
                }
                else
                {
                    xlRange.Cells[9, 9].Value = "[2]Failed";
                }
                Console.WriteLine("Kết quả: " + xlRange.Cells[9, 9].Value);
            }
            catch (Exception ex)
            {
                // Ghi "Failed" nếu có lỗi
                xlRange.Cells[9, 9].Value = "[2]Failed";
                Console.WriteLine($"Lỗi xảy ra: {ex.Message}");
            }

        }
        [Test]
        public void TimkiemId48()
        {
            string keyword = xlRange.Cells[10, 7].Value?.ToString(); // Lấy dữ liệu từ dòng 10 cột 7
            if (string.IsNullOrEmpty(keyword))
            {
                Console.WriteLine("Không có từ khóa để tìm kiếm.");
                return;
            }

            Console.WriteLine($"Đang tìm kiếm: {keyword}");

            try
            {
                // Nhập từ khóa vào ô tìm kiếm
                IWebElement searchField = driver.FindElement(By.XPath("//input[@id='search-box']"));
                searchField.Clear();
                searchField.SendKeys(keyword);
                Thread.Sleep(500);

                // Click vào icon tìm kiếm
                driver.FindElement(By.XPath("//i[@class='fas fa-search']")).Click();
                Thread.Sleep(2000);

                // Kiểm tra danh sách sản phẩm hiển thị
                var products = driver.FindElements(By.ClassName("box")); // Thay bằng class thực tế
                if (products.Count > 0)
                {
                    xlRange.Cells[10, 9].Value = "[2]Passed"; // Ghi vào dòng 10 cột 9
                }
                else
                {
                    xlRange.Cells[10, 9].Value = "[2]Failed";
                }

                Console.WriteLine("Kết quả: " + xlRange.Cells[10, 9].Value);
            }
            catch (Exception ex)
            {
                // Ghi "Failed" nếu có lỗi
                xlRange.Cells[10, 9].Value = "[2]Failed";
                Console.WriteLine($"Lỗi xảy ra: {ex.Message}");
            }
        }
        [Test]
        public void TimkiemId49()
        {
            string keyword = xlRange.Cells[11, 7].Value?.ToString(); // Lấy dữ liệu từ dòng 11 cột 7 
            if (string.IsNullOrEmpty(keyword))
            {
                Console.WriteLine("Không có từ khóa để tìm kiếm.");
                return;
            }

            Console.WriteLine($"Đang tìm kiếm: {keyword}");

            try
            {
                // Nhập từ khóa vào ô tìm kiếm
                IWebElement searchField = driver.FindElement(By.XPath("//input[@id='search-box']"));
                searchField.Clear();
                searchField.SendKeys(keyword);
                Thread.Sleep(500);

                // Click vào icon tìm kiếm
                driver.FindElement(By.XPath("//i[@class='fas fa-search']")).Click();
                Thread.Sleep(2000);

                // Kiểm tra danh sách sản phẩm hiển thị
                var products = driver.FindElements(By.ClassName("box")); // Thay bằng class thực tế
                if (products.Count > 0)
                {
                    xlRange.Cells[11, 9].Value = "[2]Passed"; // Ghi vào dòng 11 cột 9
                }
                else
                {
                    xlRange.Cells[11, 9].Value = "[2]Failed";
                }

                Console.WriteLine("Kết quả: " + xlRange.Cells[11, 9].Value);
            }
            catch (Exception ex)
            {
                // Ghi "Failed" nếu có lỗi
                xlRange.Cells[11, 9].Value = "[2]Failed";
                Console.WriteLine($"Lỗi xảy ra: {ex.Message}");
            }
        }
        [Test]
        public void TimkiemId50()
        {
            string keyword = xlRange.Cells[12, 7].Value?.ToString(); // Lấy dữ liệu từ dòng 12 cột 7 
            if (string.IsNullOrEmpty(keyword))
            {
                Console.WriteLine("Không có từ khóa để tìm kiếm.");
                return;
            }

            Console.WriteLine($"Đang tìm kiếm: {keyword}");

            try
            {
                // Nhập từ khóa vào ô tìm kiếm
                IWebElement searchField = driver.FindElement(By.XPath("//input[@id='search-box']"));
                searchField.Clear();
                searchField.SendKeys(keyword);
                Thread.Sleep(500);

                // Click vào icon tìm kiếm
                driver.FindElement(By.XPath("//i[@class='fas fa-search']")).Click();
                Thread.Sleep(2000);

                // Kiểm tra danh sách sản phẩm hiển thị
                var products = driver.FindElements(By.ClassName("box"));

                if (products.Count > 0)
                {
                    xlRange.Cells[12, 9].Value = "[2]Passed"; // Ghi vào dòng 12 cột 9
                }
                else
                {
                    xlRange.Cells[12, 9].Value = "[2]Failed";
                }
                Console.WriteLine("Kết quả: " + xlRange.Cells[12, 9].Value);
            }
            catch (Exception ex)
            {
                // Ghi "Failed" nếu có lỗi
                xlRange.Cells[12, 9].Value = "[2]Failed";
                Console.WriteLine($"Lỗi xảy ra: {ex.Message}");
            }
        }
        [Test]
        public void TimkiemId51()
        {
            string keyword = xlRange.Cells[13, 7].Value?.ToString(); // Lấy dữ liệu từ dòng 13 cột 7 
            if (string.IsNullOrEmpty(keyword))
            {
                Console.WriteLine("Không có từ khóa để tìm kiếm.");
                return;
            }

            Console.WriteLine($"Đang tìm kiếm: {keyword}");

            try
            {
                // Nhập từ khóa vào ô tìm kiếm
                IWebElement searchField = driver.FindElement(By.XPath("//input[@id='search-box']"));
                searchField.Clear();
                searchField.SendKeys(keyword);
                Thread.Sleep(500);

                // Click vào icon tìm kiếm
                driver.FindElement(By.XPath("//i[@class='fas fa-search']")).Click();
                Thread.Sleep(2000);

                // Kiểm tra danh sách sản phẩm hiển thị
                var products = driver.FindElements(By.ClassName("box"));
                if (products.Count > 0)
                {
                    xlRange.Cells[13, 9].Value = "[2]Passed"; // Ghi vào dòng 13 cột 10
                }
                else
                {
                    xlRange.Cells[13, 9].Value = "[2]Failed";
                }

                Console.WriteLine("Kết quả: " + xlRange.Cells[13, 9].Value);
            }
            catch (Exception ex)
            {
                // Ghi "Failed" nếu có lỗi
                xlRange.Cells[13, 9].Value = "[2]Failed";
                Console.WriteLine($"Lỗi xảy ra: {ex.Message}");
            }
        }
        [Test]
        public void TimkiemId52()
        {
            string keyword = xlRange.Cells[14, 7].Value?.ToString(); // Lấy dữ liệu từ dòng 14 cột 7 

            try
            {
                // Nhập từ khóa vào ô tìm kiếm
                IWebElement searchField = driver.FindElement(By.XPath("//input[@id='search-box']"));
                searchField.Clear();
                searchField.SendKeys(keyword);
                Thread.Sleep(500);

                // Click vào icon tìm kiếm
                driver.FindElement(By.XPath("//i[@class='fas fa-search']")).Click();
                Thread.Sleep(2000);

                // Kiểm tra danh sách sản phẩm hiển thị
                var products = driver.FindElements(By.ClassName("box")); // Thay bằng class thực tế

                if (products.Count > 0)
                {
                    xlRange.Cells[14, 9].Value = "[2]Passed"; // Ghi vào dòng 14 cột 10
                }
                else
                {
                    xlRange.Cells[14, 9].Value = "[2]Failed";
                }

                Console.WriteLine("Kết quả: " + xlRange.Cells[14, 9].Value);
            }
            catch (Exception ex)
            {
                // Ghi "Failed" nếu có lỗi
                xlRange.Cells[14, 9].Value = "[2]Failed";
                Console.WriteLine($"Lỗi xảy ra: {ex.Message}");
            }
        }
        [Test]
        public void XemspId53()
        {
            try
            {
                var products = driver.FindElements(By.ClassName("box")); // Thay bằng class thực tế
                if (products.Count > 0)
                {
                    xlRange.Cells[16, 9].Value = "Passed"; // Ghi vào dòng 16 cột 10
                }
                else
                {
                    xlRange.Cells[16, 9].Value = "Failed";
                }

                Console.WriteLine("Kết quả: " + xlRange.Cells[16, 9].Value);
            }
            catch (Exception ex)
            {
                // Ghi "Failed" nếu có lỗi
                xlRange.Cells[16, 9].Value = "Failed";
                Console.WriteLine($"Lỗi xảy ra: {ex.Message}");
            }

        }
        [Test]
        public void XemsptheodanhmucId54()
        {
            string categoryName = xlRange.Cells[17, 7].Value?.ToString(); // Lấy dữ liệu từ dòng 14 cột 7 
            if (string.IsNullOrEmpty(categoryName))
            {
                Console.WriteLine("Không có từ khóa để tìm kiếm danh mục .");
                return;
            }

            Console.WriteLine($"Đang tìm kiếm: {categoryName}");

            try
            {
                // Tìm danh mục trên trang web
                driver.FindElement(By.XPath("//a[@class='dropbtn']")).Click();
                IWebElement category = driver.FindElement(By.XPath($"//a[contains(text(), '{categoryName}')]"));

                // Click vào danh mục nếu tìm thấy
                category.Click();
                Console.WriteLine($"✅ Click vào danh mục: {categoryName}");
                xlRange.Cells[17, 9].Value = "[2]Passed";
            }
            catch (NoSuchElementException)
            {
                Console.WriteLine($"⚠️ Không tìm thấy danh mục: {categoryName}");
                xlRange.Cells[17, 9].Value = "[2]Failed";
            }
        }
        [Test]
        public void XemctspId55()
        {
            string keyword = xlRange.Cells[19, 7].Value?.ToString(); // Lấy dữ liệu từ dòng 8 cột 7
            if (string.IsNullOrEmpty(keyword))
            {
                Console.WriteLine("Không có từ khóa để tìm kiếm.");
                return;
            }

            Console.WriteLine($"Đang tìm kiếm: {keyword}");

            try
            {
                // Nhập từ khóa vào ô tìm kiếm
                IWebElement searchField = driver.FindElement(By.XPath("//input[@id='search-box']"));
                searchField.Clear();
                searchField.SendKeys(keyword);
                Thread.Sleep(500);

                // Click vào icon tìm kiếm
                driver.FindElement(By.XPath("//i[@class='fas fa-search']")).Click();
                Thread.Sleep(2000);

                // Kiểm tra danh sách sản phẩm hiển thị
                var products = driver.FindElements(By.ClassName("box")); // Thay bằng class thực tế
                if (products.Count > 0)
                {
                    xlRange.Cells[19, 9].Value = "[2]Passed"; // Ghi vào dòng 8 cột 9
                }
                else
                {
                    xlRange.Cells[19, 9].Value = "[2]Failed";
                }
                Console.WriteLine("Kết quả: " + xlRange.Cells[19, 9].Value);
                driver.FindElement(By.XPath("//a[@class='btn btn-primary']")).Click();
            }
            catch (Exception ex)
            {
                // Ghi "Failed" nếu có lỗi
                xlRange.Cells[19, 9].Value = "[2]Failed";
                Console.WriteLine($"Lỗi xảy ra: {ex.Message}");
            }


        }

        [TearDown]
        public void TearDown()
        {
            try
            {
                if (dataBook != null)
                {
                    dataBook.Save();
                    dataBook.Close(false); // Đặt false để tránh hiển thị hộp thoại lưu
                }

                if (dataApp != null)
                {
                    dataApp.Quit();
                }

                if (driver != null)
                {
                    driver.Quit(); // `Quit()` tốt hơn `Close()` vì nó đóng cả trình duyệt và driver
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Lỗi khi đóng tài nguyên: {ex.Message}");
            }
            finally
            {
                // Giải phóng bộ nhớ của Excel để tránh lỗi tiến trình treo
                if (dataSheet != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(dataSheet);
                if (dataBook != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(dataBook);
                if (dataApp != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(dataApp);

                dataSheet = null;
                dataBook = null;
                dataApp = null;
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }
    }
}