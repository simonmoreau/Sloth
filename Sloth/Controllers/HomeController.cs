using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using System.IO;
using Microsoft.AspNetCore.Http;
using Xbim.BCF;

namespace Sloth.Controllers
{
    public class HomeController : Controller
    {

        public IActionResult Index()
        {
            return View();
        }

        public IActionResult Topics()
        {
            if (HttpContext.Session.Keys.Count() != 0)
            {
                Models.DisplayBCF displayBCF = Services.SessionExtensionMethods.GetObject<Models.DisplayBCF>(HttpContext.Session, "BCF");

                List<Models.DisplayTopic> DisplayTopics = displayBCF.BCF.Topics.Select(o => new Models.DisplayTopic(o)).ToList();

                return View("Topics", DisplayTopics);
            }
            else
            {
                return View("Index");
            }
        }

        [HttpPost("Home/Upload")]
        public async Task<IActionResult> Post(IFormFile file)
        {
            if (file == null || file.Length == 0) return Content("Item not found");

            string fileName = file.FileName;
            
            BCF bcf = BCF.Deserialize(file.OpenReadStream());

            Models.DisplayBCF displayBCF = new Models.DisplayBCF(bcf, fileName);

            Services.SessionExtensionMethods.SetObject(HttpContext.Session, "BCF", displayBCF);

            List<Models.DisplayTopic>  DisplayTopics = bcf.Topics.Select(o => new Models.DisplayTopic(o)).ToList();

            return View("Topics", DisplayTopics);
        }


        [HttpPost]
        [ValidateAntiForgeryToken]
        public IActionResult WordExport()
        {

            Models.DisplayBCF displayBCF = Services.SessionExtensionMethods.GetObject<Models.DisplayBCF>(HttpContext.Session, "BCF");

            string wordFileName = Path.GetFileNameWithoutExtension(displayBCF.FileName) + ".docx";

            Response.Headers.Add("content-disposition", "attachment; filename=" + wordFileName);

            MemoryStream ms = displayBCF.ExportAsWord();

            byte[] bytesInStream = ms.ToArray(); // simpler way of converting to array
            ms.Close();

            return File(bytesInStream, "application/octet-stream"); // or "application/x-rar-compressed"
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public IActionResult ExcelExport()
        {
            Models.DisplayBCF displayBCF = Services.SessionExtensionMethods.GetObject<Models.DisplayBCF>(HttpContext.Session, "BCF");

            string excelFileName = Path.GetFileNameWithoutExtension(displayBCF.FileName) + ".xlsx";

            Response.Headers.Add("content-disposition", "attachment; filename=" + excelFileName);

            MemoryStream ms = displayBCF.ExportAsExcel();

            byte[] bytesInStream = ms.ToArray(); // simpler way of converting to array
            ms.Close();

            return File(bytesInStream, "application/octet-stream"); // or "application/x-rar-compressed"
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public IActionResult PDFExport()
        {
            Response.Headers.Add("content-disposition", "attachment; filename=test.bcfzip");

            BCF bcf = Services.SessionExtensionMethods.GetObject<BCF>(HttpContext.Session, "BCF");

            return File(bcf.Serialize(),
                        "application/octet-stream"); // or "application/x-rar-compressed"
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

        public IActionResult Error()
        {
            return View();
        }
    }
}
