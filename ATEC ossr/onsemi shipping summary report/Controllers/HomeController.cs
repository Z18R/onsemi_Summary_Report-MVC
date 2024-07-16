using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;
using OfficeOpenXml;
using onsemi_shipping_summary_report.Models;
using System.Data;
using System.Diagnostics;
using System.IO;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;

namespace onsemi_shipping_summary_report.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;
        private readonly IConfiguration _configuration;

        public HomeController(ILogger<HomeController> logger, IConfiguration configuration)
        {
            _logger = logger;
            _configuration = configuration;
        }

        public IActionResult Index()
        {
            var model = new DateFilterViewModel
            {
                FromDate = new DateTime(2024, 1, 1),
                ToDate = new DateTime(2024, 6, 19)
            };
            return View(model);
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

        [HttpPost]
        public IActionResult ExportToExcel(DateFilterViewModel model)
        {
            // This is Report1
            return ExportData(model, "dbo.sp_ReportDataExport", "ShippingReport", true);
        }

        [HttpPost]
        public IActionResult ExportToExcelSummary(DateFilterViewModel model)
        {
            // This is Report2
            return ExportData(model, "dbo.usp_RPT_Onsemi_Shipped_Lot", "SummaryReport", false);
        }

        private IActionResult ExportData(DateFilterViewModel model, string storedProcedureName, string reportType, bool isReport1)
        {
            DataTable dataTable = new DataTable();
            string connectionString = _configuration.GetConnectionString("MES_ATEC_Connection");

            try
            {
                using (SqlConnection connection = new SqlConnection(connectionString))
                {
                    connection.Open();

                    using (SqlCommand command = new SqlCommand(storedProcedureName, connection))
                    {
                        command.CommandType = CommandType.StoredProcedure;

                        command.Parameters.AddWithValue("@FromDate", model.FromDate);
                        command.Parameters.AddWithValue("@ToDate", model.ToDate);

                        using (SqlDataAdapter adapter = new SqlDataAdapter(command))
                        {
                            adapter.Fill(dataTable);
                        }
                    }
                }

                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                using (ExcelPackage package = new ExcelPackage())
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("ExportedData");
                    worksheet.Cells["A1"].LoadFromDataTable(dataTable, true);

                    if (isReport1)
                    {
                        int[] dateColumns = { 6, 7, 8, 9, 10 }; // Adjust based on your actual date columns indices
                        foreach (int colIndex in dateColumns)
                        {
                            for (int rowIndex = 2; rowIndex <= dataTable.Rows.Count + 1; rowIndex++)
                            {
                                if (double.TryParse(worksheet.Cells[rowIndex, colIndex].Text, out double dateNumber))
                                {
                                    DateTime date = DateTime.FromOADate(dateNumber);
                                    worksheet.Cells[rowIndex, colIndex].Value = date;
                                    worksheet.Cells[rowIndex, colIndex].Style.Numberformat.Format = "yyyy-MM-dd"; 
                                }
                            }
                            int assyCtColumnIndex = 11; // Adjust to your AssyCT column index
                            int testCtColumnIndex = 12; // Adjust to your TestCT column index
                            for (int rowIndex = 2; rowIndex <= dataTable.Rows.Count + 1; rowIndex++)
                            {
                                worksheet.Cells[rowIndex, assyCtColumnIndex].Formula = $"G{rowIndex}-F{rowIndex}"; // G - F
                                worksheet.Cells[rowIndex, testCtColumnIndex].Formula = $"J{rowIndex}-H{rowIndex}"; // J - H
                            }
                        }
               
                    }

                    var stream = new MemoryStream();
                    package.SaveAs(stream);
                    stream.Position = 0;

                    string excelName = $"{reportType}-{model.FromDate:yyyyMMdd}_{model.ToDate:yyyyMMdd}.xlsx";
                    return File(stream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", excelName);
                }
            }
            catch (SqlException sqlEx)
            {
                _logger.LogError(sqlEx, "SQL Exception occurred while exporting data to Excel.");
                return StatusCode(500, "An error occurred while exporting data. Please try again later.");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Exception occurred while exporting data to Excel.");
                return StatusCode(500, "An error occurred while exporting data. Please try again later.");
            }
        }
    }
}
