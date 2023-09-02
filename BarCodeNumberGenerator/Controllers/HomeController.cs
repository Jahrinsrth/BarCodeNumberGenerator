using BarCodeNumberGenerator.Models;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using OfficeOpenXml;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Security.Policy;

namespace BarCodeNumberGenerator.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;

        private List<long>  globalList = new List<long>();
       
        public HomeController(ILogger<HomeController> logger)
        {
            _logger = logger;
        }

        public IActionResult Index(BarCodeNum barCodeNum)
        {

            BarCodeNum barCodeNumResult = new BarCodeNum();
            List<long> numbers = new List<long>();
            barCodeNumResult.IsValid = false;

            if (barCodeNum.NumberGap > 0 && barCodeNum.StartingNumber > 0 && barCodeNum.StartingNumber.ToString().Length == 13)
            {
                numbers = GenerateNumbers(barCodeNum.StartingNumber, barCodeNum.NumberGap);
                barCodeNumResult.NumberList = numbers;
                globalList = numbers;

                //return View(barCodeNumResult);

                ResultData resultData = new ResultData();
                resultData.PassData = numbers;
                return RedirectToAction("ExportToExcel",resultData);
            }
            else
            {
                if (barCodeNum.StartingNumber != 0 && (barCodeNum.StartingNumber.ToString().Length < 13 || barCodeNum.StartingNumber.ToString().Length > 13))
                {
                    barCodeNumResult.IsValid = true;
                }

                barCodeNumResult.ErrorMsg = "Starting number length should be 13 characters";
                return View(barCodeNumResult);
            }
        }

        public IActionResult ExportToExcel(ResultData resultData)
        {

            //BarCodeNum barCodeNumResult = new BarCodeNum();
            //List<long> numbers = new List<long>();
            //barCodeNumResult.IsValid = false;

            // Create a new Excel package
            using (var package = new ExcelPackage())

            {

                //List<long> message = ViewBag.Result;

                //var package = new ExcelPackage();
                // Add a new worksheet
                var worksheet = package.Workbook.Worksheets.Add("Sample_BarCodes");

                // Define the column headers (if needed)
                worksheet.Cells["A1"].Value = "Sample BarCodes";
                //worksheet.Cells["A1"].Columns.w

                worksheet.Column(1).Width = 30;


                //worksheet.Cells["B1"].Value = "Column2";

                // Fill in the data
                int row = 2; // Start from the second row to leave room for headers
                foreach (long number in resultData.PassData)
                {
                    worksheet.Cells["A" + row].Style.Numberformat.Format = "0";
                    worksheet.Cells["A" + row].Value = number; // Replace with actual property names
                                                               //worksheet.Cells["B" + row].Value = number.Property2; // Replace with actual property names
                                                               // Add more columns as needed
                    row++;
                }

                // Save the Excel package to a memory stream
                var stream = new MemoryStream();
                package.SaveAs(stream);

                // Set the position to the beginning of the stream
                stream.Position = 0;

                // Define the content type and file name for the response
                string contentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                string fileName = "SampleBarCodes.xlsx";

                // Return the Excel file as a file download
                return File(stream, contentType, fileName);

            }
        }

        public List<long> GenerateNumbers(long startingNumber, long gap)
        {
            List<long> numbers = new List<long>();
            for (int i = 0; i < 350; i++)
            {
                numbers.Add(startingNumber + gap);
                startingNumber += gap;
            }
            return numbers;
        }

        public IActionResult Privacy()
        {
            return View();
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }
    }
}
