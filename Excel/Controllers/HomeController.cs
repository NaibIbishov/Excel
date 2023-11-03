using Excel.Models;
using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;
using System.Diagnostics;

namespace Excel.Controllers
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

        public IActionResult Privacy()
        {
            return View();
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }
        
        //Excele melumat yazma
        public IActionResult CreateExcelFile()
        {
            var Students=new List<Student>{
                new Student{Name="Naib",Surname="Ibishov",TelNumber="0509818558"},
                new Student{Name="Nicat",Surname="Memmedli",TelNumber="0501234567"},
                new Student{Name="Agami",Surname="Rehimov",TelNumber="0509876543"},
            };
            
            var stream=new MemoryStream();
            using(var xlPackage= new ExcelPackage(stream))
            {
                var worksheet = xlPackage.Workbook.Worksheets.Add("Students");
                worksheet.Cells["A1"].Value = "Ad";
                worksheet.Cells["B1"].Value = "Soyad";
                worksheet.Cells["C1"].Value = "Telefon nomresi";
                worksheet.Cells["A1:C1"].Style.Font.Bold = true;

                int row = 2;
                foreach (var item in Students)
                {
                    worksheet.Cells[row, 1].Value = item.Name;
                    worksheet.Cells[row,2].Value=item.Surname;
                    worksheet.Cells[row,3].Value=item.TelNumber;
                    row++;
                }
                xlPackage.Save();
                stream.Position = 0;
                return File(stream, "application/vnd." +
                    "openxmlformats-officedocument.spreadsheetml." +
                    "sheet", "Toplu.xlsx");
            }
        }


    }
}