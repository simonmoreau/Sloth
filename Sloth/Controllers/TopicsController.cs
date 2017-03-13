using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;

namespace Sloth.Controllers
{
    public class TopicsController : Controller
    {
        public IActionResult Index()
        {
            return View();
        }

        public IActionResult Export()
        {
            Console.WriteLine("it works !");

            return View();
        }
    }
}