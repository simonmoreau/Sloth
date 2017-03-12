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

        [HttpPost("Home/Upload")]
        public async Task<IActionResult> Post(IFormFile file)
        {
            if (file == null || file.Length == 0) return Content("Item not found");

            long size = file.Length;
            string fileName = file.FileName;

            // full path to file in temp location
            var tempFilePath = Path.GetTempFileName();

            BCF bcf = BCF.Deserialize(file.OpenReadStream());
            //if (file.Length > 0)
            //{

            //    using (var stream = new FileStream(tempFilePath, FileMode.Create))
            //    {
            //        await file.CopyToAsync(stream);
            //    }
            //}

            List<Models.DisplayTopic> topics = bcf.Topics.Select(o => new Models.DisplayTopic(o)).ToList();

            return View("Topics", topics);
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
