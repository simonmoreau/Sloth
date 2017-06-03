using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.IO;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Identity;
using Sloth.Data;
using Sloth.Models;
using Xbim.BCF;

namespace Sloth.Controllers
{
    public class HomeController : Controller
    {
        private readonly IHostingEnvironment _env;
        private readonly ApplicationDbContext _context;
        private readonly UserManager<ApplicationUser> _userManager;
        private readonly SignInManager<ApplicationUser> _signInManager;

        public HomeController(
            UserManager<ApplicationUser> userManager,
            SignInManager<ApplicationUser> signInManager, 
            IHostingEnvironment env, 
            ApplicationDbContext context)
        {
            _userManager = userManager;
            _signInManager = signInManager;
            _env = env;
            _context = context;
        }

        public IActionResult Index()
        {
            return View();
        }

        private List<Models.DisplayTopic> LoadTopics(string bcfFilePath, string BCFFileName)
        {
            using (FileStream stream = System.IO.File.Open(bcfFilePath, FileMode.Open))
            {
                BCF bcf = BCF.Deserialize(stream);

                Models.DisplayBCF displayBCF = new Models.DisplayBCF(bcf, BCFFileName);

                List<Models.DisplayTopic> DisplayTopics = bcf.Topics.Select(o => new Models.DisplayTopic(o)).ToList();

                return DisplayTopics;
            }
        }

        public async Task<IActionResult> Topics()
        {
            ////Find the current user
            //ApplicationUser user = await _userManager.GetUserAsync(User);

            //if (user == null)
            //{

            //}
            //else
            //{
            //    FilePath file = _context.FilePath.Where(i => i.ApplicationUserId == user.Id).FirstOrDefault();
            //    //user.FilePaths.FirstOrDefault();
            //    return View("Topics", LoadTopics(file.FullFilePath,file.FileName));
            //}

            if (HttpContext.Session.Keys.Count() != 0)
            {
                if (HttpContext.Session.Keys.Contains("BCFFiles"))
                {
                    //Get the current list of files
                    List<FilePath> BCFFiles = Services.SessionExtensionMethods.GetObject<List<FilePath>>(HttpContext.Session, "BCFFiles");

                    //Create a list of topics
                    List<DisplayTopic> topics = new List<DisplayTopic>();

                    foreach (FilePath BCFfile in BCFFiles)
                    {
                        topics.AddRange(LoadTopics(BCFfile.FullFilePath, BCFfile.FileName));
                    }

                    return View("Topics", topics);
                }
                else
                {
                    return View("Topics");
                }
            }
            else
            {
                return View("Topics");
            }
        }

        [HttpPost("Home/Upload")]
        public async Task<IActionResult> UploadBCF(IFormFile file)
        {
            if (file == null || file.Length == 0) return Content("Item not found");

            //witch user using th app
            //write its file on him

            //Generate an GUID for the file and use it to create the local file path
            Guid BcfGuid = Guid.NewGuid();
            string fullFilePath = Path.Combine(_env.WebRootPath, "files", "bcf_" + BcfGuid.ToString());

            //Save the file on the server
            using (FileStream stream = System.IO.File.OpenWrite(fullFilePath))
            {
                await file.CopyToAsync(stream);
            }

            //Create a new filePath for the uploaded BCF
            FilePath BcfFile = new FilePath
            {
                FileName = System.IO.Path.GetFileName(file.FileName),
                FileType = FileType.BCF,
                FullFilePath = fullFilePath
            };

            ////Find the current user
            //ApplicationUser user = await _userManager.GetUserAsync(User);

            //if (user == null)
            //{

            //}
            //else
            //{
            //    if (user.FilePaths == null) user.FilePaths = new List<FilePath>();
            //    user.FilePaths.Add(BcfFile);

            //    if (ModelState.IsValid)
            //    {
            //        _context.ApplicationUser.Update(user);
            //        _context.SaveChanges();
            //    }
            //}


            if (HttpContext.Session.Keys.Contains("BCFFiles"))
            {
                //Get the current list of files
                List<FilePath> BCFFiles = Services.SessionExtensionMethods.GetObject<List<FilePath>>(HttpContext.Session, "BCFFiles");
                BCFFiles.Add(BcfFile);
                //Store it back in Session
                Services.SessionExtensionMethods.SetObject(HttpContext.Session, "BCFFiles", BCFFiles);
            }
            else
            {
                //Create a new list of BCF files
                List<FilePath> BCFFiles = new List<FilePath>();
                BCFFiles.Add(BcfFile);
                //Store it in Session
                Services.SessionExtensionMethods.SetObject(HttpContext.Session, "BCFFiles", BCFFiles);
            }

            return await Topics();
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

        [HttpPost("Home/WordTemplate")]
        public async Task<IActionResult> WordTemplate(IFormFile file)
        {

            if (file == null || file.Length == 0) return Content("Item not found");

            string fileName = file.FileName;

            // process uploaded files
            // Don't rely on or trust the FileName property without validation.

            return Ok(new { count = file.FileName, file.Length });
        }

        [HttpPost("Home/ExcelTemplate")]
        public async Task<IActionResult> ExcelTemplate(IFormFile file)
        {

            if (file == null || file.Length == 0) return Content("Item not found");

            string fileName = file.FileName;

            // process uploaded files
            // Don't rely on or trust the FileName property without validation.

            return Ok(new { count = file.FileName, file.Length });
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
