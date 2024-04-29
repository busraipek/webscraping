using HtmlAgilityPack;
using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;
using System.Diagnostics;
using webscraping.Models;

namespace webscraping.Controllers
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
            var flightDataList = new List<FlightData>();

            var web = new HtmlWeb();
            var doc = web.Load("https://www.avionio.com/en/airport/saw/departures?ts=1714212000000");

            var rows = doc.DocumentNode.SelectNodes("//tr[@class='tt-row ']");

            foreach (var row in rows)
            {
                var flight = new FlightData
                {
                    Time = row.SelectSingleNode(".//td[contains(@class, 'tt-t')]")?.InnerText.Trim(),
                    Date = row.SelectSingleNode(".//td[contains(@class, 'tt-d')]")?.InnerText.Trim(),
                    IATA = row.SelectSingleNode(".//td[contains(@class, 'tt-i')]//a")?.InnerText.Trim(),
                    Destination = row.SelectSingleNode(".//td[contains(@class, 'tt-ap')]")?.InnerText.Trim(),
                    Flight = row.SelectSingleNode(".//td[contains(@class, 'tt-f')]//a")?.InnerText.Trim(),
                    Airline = row.SelectSingleNode(".//td[contains(@class, 'tt-al')]")?.InnerText.Trim(),
                    Status = row.SelectSingleNode(".//td[contains(@class, 'tt-s')]")?.InnerText.Trim()
                };

                flightDataList.Add(flight);
            }

            // Export data to Excel
            ExportToExcel(flightDataList);

            return View(flightDataList);
        }

        private void ExportToExcel(List<FlightData> flightDataList)
        {
            // Create Excel package
            var package = new ExcelPackage();
            var worksheet = package.Workbook.Worksheets.Add("Flight Data");

            // Add headers
            worksheet.Cells["A1"].Value = "Time";
            worksheet.Cells["B1"].Value = "Date";
            worksheet.Cells["C1"].Value = "IATA";
            worksheet.Cells["D1"].Value = "Destination";
            worksheet.Cells["E1"].Value = "Flight";
            worksheet.Cells["F1"].Value = "Airline";
            worksheet.Cells["G1"].Value = "Status";

            // Add data
            for (int i = 0; i < flightDataList.Count; i++)
            {
                var flight = flightDataList[i];
                worksheet.Cells[i + 2, 1].Value = flight.Time;
                worksheet.Cells[i + 2, 2].Value = flight.Date;
                worksheet.Cells[i + 2, 3].Value = flight.IATA;
                worksheet.Cells[i + 2, 4].Value = flight.Destination;
                worksheet.Cells[i + 2, 5].Value = flight.Flight;
                worksheet.Cells[i + 2, 6].Value = flight.Airline;
                worksheet.Cells[i + 2, 7].Value = flight.Status;
            }

            // Save the Excel package to a file
            var filePath = Path.Combine(Directory.GetCurrentDirectory(), "FlightData.xlsx");
            using (var stream = new FileStream(filePath, FileMode.Create))
            {
                package.SaveAs(stream);
            }
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }
    }
}
