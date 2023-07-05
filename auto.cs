using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OfficeOpenXml;
using System.IO;

class Program
{
    static void Main(string[] args)
    {

        var chromeDriverPath = @"D:\Downloads\Pgs\chromedriver.exe";

        var webDriver = new ChromeDriver(chromeDriverPath);

        webDriver.Navigate().GoToUrl("https://www.datahub.com");

        var headingElement = webDriver.FindElement(By.XPath("//h1"));
        var headingText = headingElement.Text;
        Console.WriteLine("Heading: " + headingText);

        var filePath = @"D:\Projects\c--automation\data.txt";
        using (var excelPackage = new ExcelPackage())
        {
            var worksheet = excelPackage.Workbook.Worksheets.Add("Data");
            worksheet.Cells["A1"].Value = "Heading";
            worksheet.Cells["B1"].Value = headingText;

            excelPackage.SaveAs(new FileInfo(filePath));
        }
        webDriver.Quit();
    }
}
