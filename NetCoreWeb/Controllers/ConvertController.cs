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
                // 获取当前目录并与路径合并
                var path = Directory.GetCurrentDirectory() + pathDirectory;

                // 检查目录是否存在，如不存在则创建
                if (!Directory.Exists(path))
                {
                    Directory.CreateDirectory(path);
                }

                string excelfullpath = Path.Combine(path, excelfileName);

                // 将Excel转换为DataTable
                DataTable inputdt = ExcelHelper.ExeclToDataTable(excelfullpath);

                string docfullpath = Path.Combine(path, docfileName);

                // 将DataTable转换为Word
                WordHelper.DtToWord(inputdt, docfullpath);
            }
            catch (Exception ex)
            {
                // 处理异常
            }
            //string webRootPath = _hostingEnvironment.WebRootPath;
            //string contentRootPath = _hostingEnvironment.ContentRootPath;
            //return Content(webRootPath + "\n" + contentRootPath);
            //var path = Path.Combine(_hostingEnvironment.WebRootPath, "Sample.PNG");


            return View();
        }

        public IActionResult DynamicGenerateWord()
        {
            try
            {
                WordHelper.GenerateWordDynamically();
            }
            catch (Exception ex)
            {
                return Content(ex.Message);
            }
          
            return View();
        }
    }
}
