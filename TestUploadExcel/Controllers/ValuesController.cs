using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Hosting;
using System.IO;
using OfficeOpenXml;
using TestUploadExcel.Models;
using TestUploadExcel.Extensions;

namespace TestUploadExcel.Controllers
{
    [Route("api/[controller]")]
    public class ValuesController : Controller
    {

        private IHostingEnvironment _hostingEnvironment;
        public ValuesController(IHostingEnvironment hostingEnvironment)
        {
            _hostingEnvironment = hostingEnvironment;
        }
        ////[HttpGet("upload-check-in-sheet")]
        //public IActionResult UploadFile(IFormFile file)
        //{
        //    IEnumerable<Drivers> drivers;
        //    List<Error> errorss =new List<Error>();
        //    string fileExt = System.IO.Path.GetExtension(file.FileName).Substring(1);
        //    //get root
        //    string webRootPath = _hostingEnvironment.WebRootPath;
        //    //Folder Name
        //    string folderName = "UploadFiles";
        //    string path = Path.Combine(webRootPath, folderName);
        //    if (fileExt != "xlsx")
        //    {
        //        //AppErrorDto = new ErrorDetails(Enums.ErrorType.NotValidExtention);
        //        //return AppResult(null);
        //    }
        //    if (!Directory.Exists(path))
        //    {
        //        Directory.CreateDirectory(path);
        //    }
        //    Guid id = Guid.NewGuid();
        //    string fullPath = path + "\\" + id.ToString() + "." + fileExt;
        //    using (var fileStream = new FileStream(fullPath, FileMode.Create))
        //    {
        //        file.CopyTo(fileStream);
        //    }

        //    using (FileStream fileStream = new FileStream(fullPath, FileMode.Open))
        //    {
        //        using (ExcelPackage package = new ExcelPackage(fileStream))
        //        {
        //            ExcelWorksheet workSheet = package.Workbook.Worksheets[1];
        //            drivers = workSheet.ToList<Drivers>(out errorss);
        //        }
        //    }
        //        return Json(drivers);
        //}

        [HttpGet("upload-check-in-sheet")]
        public IActionResult UploadFile(IFormFile file)
        {
            List<Drivers> drivers;
            List<Error> errorss = new List<Error>();
            string fileExt = System.IO.Path.GetExtension(file.FileName).Substring(1);
            //get root
            string webRootPath = _hostingEnvironment.WebRootPath;
            //Folder Name
            string folderName = "UploadFiles";
            string path = Path.Combine(webRootPath, folderName);
            if (fileExt != "xlsx")
            {
                //AppErrorDto = new ErrorDetails(Enums.ErrorType.NotValidExtention);
                //return AppResult(null);
            }
            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }
            Guid id = Guid.NewGuid();
            string fullPath = path + "\\" + id.ToString() + "." + fileExt;
            using (var fileStream = new FileStream(fullPath, FileMode.Create))
            {
                file.CopyTo(fileStream);
            }

            using (FileStream fileStream = new FileStream(fullPath, FileMode.Open))
            {
                using (ExcelPackage package = new ExcelPackage(fileStream))
                {
                    ExcelWorksheet workSheet = package.Workbook.Worksheets[1];
                    drivers = workSheet.ToList<Drivers>(out errorss);
                }
            }
            return Json(drivers);
        }

        // GET api/values
        [HttpGet]
        public IActionResult Get()
        {
            //List<Error> res = new List<Error>()
            //{
            //    new Error(){ErrorName = "Test" ,ErrorType ="type"},
            //    new Error(){ErrorName = "Test2" ,ErrorType ="type2"},

            //};
            return Json("Done");
        }

        // GET api/values/5
        [HttpGet("{id}")]
        public string Get(int id)
        {
            return "value";
        }

        // POST api/values
        [HttpPost]
        public void Post([FromBody]string value)
        {
        }

        // PUT api/values/5
        [HttpPut("{id}")]
        public void Put(int id, [FromBody]string value)
        {
        }

        // DELETE api/values/5
        [HttpDelete("{id}")]
        public void Delete(int id)
        {
        }
    }
}
