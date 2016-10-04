using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using ZuelligPharma.App_Start;

namespace ZuelligPharma.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            return View();
        }

        public ActionResult About()
        {
            ViewBag.Message = "Your application description page.";

            return View();
        }

        public ActionResult Contact()
        {
            ViewBag.Message = "Your contact page.";

            return View();
        }

        [HttpPost]
        public ActionResult abc(HttpPostedFileBase file)
        {
            string path = String.Empty;
            if (file.ContentLength > 0)
            {
                var fileName = Path.GetFileName(file.FileName);
                path = Path.Combine(Server.MapPath("~/App_Data/uploads"), fileName);
                if (System.IO.Directory.Exists(path)== true)
                {
                    System.IO.Directory.Delete(path, true);
                }
                file.SaveAs(path);

                ExcelLibrary objExcel = new ExcelLibrary(path);
                objExcel.ReadData();
                objExcel.Quit();
                //System.IO.Directory.Delete(path, true);
            }
            return RedirectToAction("Index");
        }

    }
}