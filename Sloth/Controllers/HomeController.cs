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

                    return View("Index",topics);
                }
                else
                {
                    return View("Index");
                }
            }
            else
            {
                return View("Index");
            }
        }

        private List<Models.DisplayTopic> LoadTopics(string bcfFilePath, string BCFFileName)
        {
            using (FileStream stream = System.IO.File.Open(bcfFilePath, FileMode.Open))
            {
                try
                {
                    BCF bcf = BCF.Deserialize(stream);

                    List<Models.DisplayTopic> DisplayTopics = bcf.Topics.Select(o => new Models.DisplayTopic(o)).ToList();

                    return DisplayTopics;
                }
                catch (Exception ex)
                {
                    //Fail silently :-(
                    return new List<Models.DisplayTopic>();
                }
            }
        }

        //public async Task<IActionResult> Topics()
        //{
        //    ////Find the current user
        //    //ApplicationUser user = await _userManager.GetUserAsync(User);

        //    //if (user == null)
        //    //{

        //    //}
        //    //else
        //    //{
        //    //    FilePath file = _context.FilePath.Where(i => i.ApplicationUserId == user.Id).FirstOrDefault();
        //    //    //user.FilePaths.FirstOrDefault();
        //    //    return View("Topics", LoadTopics(file.FullFilePath,file.FileName));
        //    //}


        //}

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

            return Index();
        }


        [HttpPost("Home/WordExport")]
        [ValidateAntiForgeryToken]
        public ActionResult WordExport()
        {
            if (HttpContext.Session.Keys.Contains("BCFFiles"))
            {
                //Get the current list of files
                List<FilePath> BCFFiles = Services.SessionExtensionMethods.GetObject<List<FilePath>>(HttpContext.Session, "BCFFiles");

                //Get the word template
                string wordTemplatePath = "";
                if (HttpContext.Session.Keys.Contains("wordTemplate"))
                {
                    wordTemplatePath = Services.SessionExtensionMethods.GetObject<string>(HttpContext.Session, "wordTemplate");
                }

                string wordFileName = Path.GetFileNameWithoutExtension(BCFFiles.FirstOrDefault().FileName) + ".docx";

                Response.Headers.Add("content-disposition", "attachment; filename=" + wordFileName);

                MemoryStream ms = Services.ExportServices.ExportAsWord(BCFFiles, wordTemplatePath);

                byte[] bytesInStream = ms.ToArray(); // simpler way of converting to array
                ms.Close();

                return File(bytesInStream, "application/octet-stream"); // or "application/x-rar-compressed"
            }
            else
            {
                return View("Index");
            }

        }

        [HttpPost("Home/ExcelExport")]
        [ValidateAntiForgeryToken]
        public ActionResult ExcelExport()
        {
            if (HttpContext.Session.Keys.Contains("BCFFiles"))
            {
                //Get the current list of files
                List<FilePath> BCFFiles = Services.SessionExtensionMethods.GetObject<List<FilePath>>(HttpContext.Session, "BCFFiles");

                string excelFileName = Path.GetFileNameWithoutExtension(BCFFiles.FirstOrDefault().FileName) + ".xlsx";

                Response.Headers.Add("content-disposition", "attachment; filename=" + excelFileName);

                MemoryStream ms = Services.ExportServices.ExportAsExcel(BCFFiles);

                byte[] bytesInStream = ms.ToArray(); // simpler way of converting to array
                ms.Close();

                return File(bytesInStream, "application/octet-stream"); // or "application/x-rar-compressed"
            }
            else
            {
                return View("Index");
            }
        }

        [HttpPost("Home/WordTemplate")]
        public async Task<IActionResult> WordTemplate(IFormFile file)
        {
            //Generate an GUID for the file and use it to create the local file path
            Guid BcfGuid = Guid.NewGuid();
            string fullFilePath = Path.Combine(_env.WebRootPath, "files", "word_" + BcfGuid.ToString());

            //Save the file on the server
            using (FileStream stream = System.IO.File.OpenWrite(fullFilePath))
            {
                await file.CopyToAsync(stream);
            }

            //Create a new filePath for the uploaded BCF
            FilePath BcfFile = new FilePath
            {
                FileName = System.IO.Path.GetFileName(file.FileName),
                FileType = FileType.WordTemplate,
                FullFilePath = fullFilePath
            };

            //Store it in Session
            Services.SessionExtensionMethods.SetObject(HttpContext.Session, "wordTemplate", fullFilePath);

            ////Find the current user
            //ApplicationUser user = await _userManager.GetUserAsync(User);

            //if (user.FilePaths == null) user.FilePaths = new List<FilePath>();
            //user.FilePaths.Add(BcfFile);

            //if (ModelState.IsValid)
            //{
            //    _context.ApplicationUser.Update(user);
            //    _context.SaveChanges();
            //}

            return Index();
        }

        [HttpPost("Home/ExcelTemplate")]
        public async Task<IActionResult> ExcelTemplate(IFormFile file)
        {
            //Generate an GUID for the file and use it to create the local file path
            Guid BcfGuid = Guid.NewGuid();
            string fullFilePath = Path.Combine(_env.WebRootPath, "files", "excel_" + BcfGuid.ToString());

            //Save the file on the server
            using (FileStream stream = System.IO.File.OpenWrite(fullFilePath))
            {
                await file.CopyToAsync(stream);
            }

            //Create a new filePath for the uploaded BCF
            FilePath BcfFile = new FilePath
            {
                FileName = System.IO.Path.GetFileName(file.FileName),
                FileType = FileType.ExcelTemplate,
                FullFilePath = fullFilePath
            };

            //Find the current user
            ApplicationUser user = await _userManager.GetUserAsync(User);

            if (user.FilePaths == null) user.FilePaths = new List<FilePath>();
            user.FilePaths.Add(BcfFile);

            if (ModelState.IsValid)
            {
                _context.ApplicationUser.Update(user);
                _context.SaveChanges();
            }

            return View("Index");
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
