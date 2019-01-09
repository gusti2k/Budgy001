using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Budgy001.Models;
using System.IO;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using Microsoft.AspNetCore.Hosting;

namespace Budgy001.Controllers
{
    public class HomeController : Controller
    {
        private IHostingEnvironment _hostingEnvironment;
        public HomeController(IHostingEnvironment hostingEnvironment)
        {
            _hostingEnvironment = hostingEnvironment;
        }
        public static List<Stuff> stuffs = new List<Stuff>();
        public IActionResult Index()
        {
            return View();
        }

        public IActionResult About()
        {
            ViewData["Message"] = "Your application description page.";

            return View();
        }

        public IActionResult Contact()
        {
            ViewData["Message"] = "Your contact page.";

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
        
        public async Task<IActionResult> OnPostExport()
        {
            string sWebRootFolder = _hostingEnvironment.WebRootPath;
            string sFileName = @"demo1.xlsx";
            string URL = string.Format("{0}://{1}/{2}", Request.Scheme, Request.Host, sFileName);
            FileInfo file = new FileInfo(Path.Combine(sWebRootFolder, sFileName));
            var memory = new MemoryStream();
            using (var fs = new FileStream(Path.Combine(sWebRootFolder, sFileName), FileMode.Create, FileAccess.Write))
            {
                IWorkbook workbook;
                workbook = new XSSFWorkbook();
                ISheet excelSheet = workbook.CreateSheet("Demo");

                //col = excelSheet.CreateRow(0);

                IRow row = excelSheet.CreateRow(0);


                stuffs.Add(
                    new Stuff() { Indkomst = "Indkomst", _Indkomst = 330000, Bilforsikring = "Bilfors", _Bilforsikring = 3000 }
                );

                Stuff s = new Stuff();              

                row = excelSheet.CreateRow(1);
                row.CreateCell(0).SetCellValue(stuffs.FirstOrDefault().Indkomst);
                row.CreateCell(1).SetCellValue(stuffs.FirstOrDefault()._Indkomst);
                row.CreateCell(2).SetCellValue(stuffs.FirstOrDefault().Bilforsikring);
                row.CreateCell(3).SetCellValue(stuffs.FirstOrDefault()._Indkomst);


                //row = excelSheet.CreateRow(2);
                //row.CreateCell(0).SetCellValue(2);
                //row.CreateCell(1).SetCellValue("Martin Guptil");
                //row.CreateCell(2).SetCellValue(33);

                //row = excelSheet.CreateRow(3);
                //row.CreateCell(0).SetCellValue(3);
                //row.CreateCell(1).SetCellValue("Colin Munro");
                //row.CreateCell(2).SetCellValue(23);

                workbook.Write(fs);
            }
            using (var stream = new FileStream(Path.Combine(sWebRootFolder, sFileName), FileMode.Open))
            {
                await stream.CopyToAsync(memory);
            }
            memory.Position = 0;
            return File(memory, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", sFileName);
        }
    }
}
