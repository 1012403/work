using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using ZuelligPharma.App_Start;
using ZuelligPharma.Models;
using System.Web.Services;

namespace ZuelligPharma.Controllers
{
    public class HomeController : Controller
    {
        private String pathName = String.Empty;
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

        //[HttpPost]
        //public ActionResult abc(HttpPostedFileBase file)
        //{
        //    string path = String.Empty;
        //    if (file.ContentLength > 0)
        //    {
        //        var fileName = Path.GetFileName(file.FileName);
        //        path = Path.Combine(Server.MapPath("~/App_Data/uploads"), fileName);
        //        if (System.IO.Directory.Exists(path)== true)
        //        {
        //            System.IO.Directory.Delete(path, true);
        //        }
        //        file.SaveAs(path);

        //        ExcelLibrary objExcel = new ExcelLibrary(path);
        //        objExcel.ReadData();
        //        objExcel.Quit();
        //        //System.IO.Directory.Delete(path, true);
        //    }
        //    return RedirectToAction("Index");
        //}

        [HttpPost]
        public ActionResult abc(HttpPostedFileBase file)
        {
            string path = String.Empty;
            ZuelligPharmaModel result = new ZuelligPharmaModel();
            if (file.ContentLength > 0)
            {
                var fileName = Path.GetFileName(file.FileName);
                try
                {
                    path = Path.Combine(Server.MapPath("~/uploads"), fileName);
                    if (System.IO.Directory.Exists(path) == true)
                    {
                        System.IO.Directory.Delete(path, true);
                    }
                    if (System.IO.Directory.Exists(Path.GetDirectoryName(path)) == false)
                    {
                        System.IO.Directory.CreateDirectory(Path.GetDirectoryName(path));
                    }
                    file.SaveAs(path);
                    pathName = path;
                    ExcelLibrary objExcel = new ExcelLibrary(path);

                    try
                    {
                        result = objExcel.ReadData();
                    }
                    catch (InvalidCastException e)
                    {
                        if (e.Data == null)
                        {
                            objExcel.Quit();
                            throw;
                        }
                    }
                    finally
                    {

                        objExcel.Quit();
                    }
                }
                catch (IOException e)
                {
                    // Extract some information from this exception, and then 
                    // throw it to the parent method.
                    Console.WriteLine("IOException source: {0}", e.Source);
                    throw;
                }
                
                
                //System.IO.Directory.Delete(path, true);
            }
            //return PartialView("Index", path);
            Session["FilePath"] = path;
            //ViewData["Data"] = result;
            return RedirectToAction("Index");
        }

        [HttpPost]
        public ActionResult GetData()
        {
            ZuelligPharmaModel result = new ZuelligPharmaModel();
            string filename = Session["FilePath"].ToString();
            if (filename.Length > 0)
            {
                pathName = Path.Combine(Server.MapPath("~/uploads"), filename);
                ExcelLibrary objExcel = new ExcelLibrary(pathName);
                try
                {
                    result = objExcel.ReadData();
                }
                catch (InvalidCastException e)
                {
                    if (e.Data == null)
                    {
                        objExcel.Quit();
                        throw;
                    }
                }
                finally
                {
                    objExcel.Quit();
                }
                //System.IO.Directory.Delete(path, true);
            }
            return Json(result);
        }

    }
}