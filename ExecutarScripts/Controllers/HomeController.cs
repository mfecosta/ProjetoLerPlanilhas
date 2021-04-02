using ExecutarScripts.Models;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;
using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using Microsoft.Extensions.Configuration;

namespace ExecutarScripts.Controllers
{
    public class HomeController : Controller
    {
        private readonly IWebHostEnvironment _hostingEnvironment;
        public HomeController(IWebHostEnvironment hostingEnvironment)
        {

            _hostingEnvironment = hostingEnvironment;

        }

        public IActionResult Index()
        {
            FileUploadViewModel model = new FileUploadViewModel();

            return View(model);
        }
        [HttpPost]
        public ActionResult Index(FileUploadViewModel model)
        {
            string rootfolder = _hostingEnvironment.WebRootPath;
            
            string filename = Guid.NewGuid().ToString() + model.Planilha.FileName;
            if (String.IsNullOrEmpty(rootfolder))
            {
                ModelState.AddModelError("", "Carregue a planilha");
            }

            
            var AppName = new ConfigurationBuilder().AddJsonFile("appsettings.json").Build().GetSection("ConnectionStrings")["DefaultConnection"];

            FileInfo file = new FileInfo(Path.Combine(rootfolder, filename));
            using (var stream = new MemoryStream())
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                model.Planilha.CopyToAsync(stream);
                using (var package = new ExcelPackage(stream))
                {
                    package.SaveAs(file);
                }

            }

            using (ExcelPackage package = new ExcelPackage(file))
            {
                //ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                ExcelWorksheet worksheet = package.Workbook.Worksheets.FirstOrDefault();
                if (worksheet == null)
                {
                    ModelState.AddModelError("", "Planilha está vazia");
                }
                else
                {
                    var countrow = worksheet.Dimension.Rows;
                    for (int row = 2; row < countrow; row++)
                    {
                        model.StaffInfoViewModel.StaffList.Add(new StaffInfoViewModel
                        {
                            Scripts = (worksheet.Cells[row, 1].Value ?? string.Empty).ToString().Trim(),

                        });
                    }
                }
            }

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
    }
}

