using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;
using OfficeOpenXml;
using onsemi_shipping_summary_report.Models;
using System.Data;
using System.Diagnostics;
using System.IO;
using OfficeOpenXml;
using Microsoft.Extensions.Configuration;

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
            // Populate dropdowns with initial dates or any default values
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
            return ExportData(model, "dbo.sp_ReportDataExport", "ShippingReport");
        }

        [HttpPost]
        public IActionResult ExportToExcelSummary(DateFilterViewModel model)
        {
            return ExportData(model, "dbo.usp_RPT_Onsemi_Shipped_Lot", "SummaryReport");
        }

        private IActionResult ExportData(DateFilterViewModel model, string storedProcedureName, string reportType)
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

                // Set the license context for EPPlus
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                using (ExcelPackage package = new ExcelPackage())
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("ExportedData");
                    worksheet.Cells["A1"].LoadFromDataTable(dataTable, true);

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
