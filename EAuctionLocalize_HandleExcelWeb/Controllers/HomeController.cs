using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using EAuctionLocalize_HandleExcelWeb.Models;
using Microsoft.AspNetCore.Http;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace EAuctionLocalize_HandleExcelWeb.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;

        public HomeController(ILogger<HomeController> logger)
        {
            _logger = logger;
        }

        public IActionResult Index()
        {
            return View();
        }

        public IActionResult UploadExcel(IFormFile excelfile)
        {
            if (excelfile != null && excelfile.Length > 0)
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                var records = new List<record>();

                using (var package = new ExcelPackage(excelfile.OpenReadStream()))
                {


                    // Get the work book in the file
                    var workBook = package.Workbook;
                    if (workBook != null)
                    {
                        if (workBook.Worksheets.Count > 0)
                        {
                            // Get the first worksheet
                            var currentWorksheet = workBook.Worksheets.FirstOrDefault();
                            if (currentWorksheet != null)
                            {
                                for (int i = 2; i < currentWorksheet.Dimension.Rows; i++)
                                {
                                    for (int j = 3; j < currentWorksheet.Dimension.Columns; j++)
                                    {
                                        var cell = currentWorksheet.Cells[i, j];
                                        var vals = cell.Value?.ToString().Replace("\n", "").Split(new string[] { "Key:", "En:", "Ar:" }, options: StringSplitOptions.RemoveEmptyEntries);
                                        if (vals != null && vals.Length == 3)
                                            records.Add(new record { key = vals[0].Trim(), en = vals[1].Trim(), ar = vals[2].Trim() });
                                        //


                                    }
                                }
                            }
                            // read some data
                            //object col1Header = currentWorksheet.[1, 1].Value;
                        }
                    }
                    ExcelPackage ExcelPkg = new ExcelPackage();
                    ExcelWorksheet wsSheetResult = ExcelPkg.Workbook.Worksheets.Add("Result");

                    for (int i = 1; i < records.Count; i++)
                    {
                        wsSheetResult.Cells[i, 1].Value = records[i - 1].key;
                        wsSheetResult.Cells[i, 2].Value = records[i - 1].en;
                        wsSheetResult.Cells[i, 3].Value = records[i - 1].ar;

                    }

                    wsSheetResult.Cells[wsSheetResult.Dimension.Address].AutoFitColumns();
                    wsSheetResult.Cells[wsSheetResult.Dimension.Address].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;

                    Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                    var fileContents = ExcelPkg.GetAsByteArray();
                    return File(
                        fileContents,
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        DateTime.Now.ToString("d") + "_" + excelfile.FileName
                    );

                    //ExcelPkg.SaveAs(new FileInfo(@"D:\Work\EAuction\EAuctionLocalize_HandleExcel\EAuctionLocalize_HandleExcel\res.xlsx"));



                }

            }

            return View();
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

        public class record
        {
            public string key { get; set; }
            public string en { get; set; }
            public string ar { get; set; }
        }
    }
}
