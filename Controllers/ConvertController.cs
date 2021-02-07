using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Mvc;
using NetCoreWeb.Common;
using NetCoreWeb.Interfaces;
using NetCoreWeb.Models;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace NetCoreWeb.Controllers
{


    public class ConvertController : Controller
    {

        private IPathProvider pathProvider;

        private const string pathDirectory = @"\DataFile\";
        private string excelfileName = "Buyer.xlsx";
        private string docfileName = "car协议书.doc";
        List<Model_Car> list = new List<Model_Car>();

        public ConvertController(IPathProvider pathProvider)
        {
            this.pathProvider = pathProvider;
        }



        public IActionResult Index()
        {
            try
            {
                var path = Directory.GetCurrentDirectory()+ pathDirectory;

                if (!Directory.Exists(path))
                {
                    //若文件目录不存在 则创建
                    Directory.CreateDirectory(path);
                }
                string excelfullpath = path + excelfileName;
                DataTable inputdt = ExcelHelper.ExeclToDataTable(excelfullpath);

                string docfullpath = path+docfileName;

                WordHelper.DtToWord(inputdt, docfullpath);

            }
            catch (Exception ex)
            {

            }

            //string webRootPath = _hostingEnvironment.WebRootPath;
            //string contentRootPath = _hostingEnvironment.ContentRootPath;
            //return Content(webRootPath + "\n" + contentRootPath);

            //var path = Path.Combine(_hostingEnvironment.WebRootPath, "Sample.PNG");


            return View();
        }
    }
}
