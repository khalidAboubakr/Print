using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml;
using Microsoft.Office.Interop.Excel;
using Syncfusion.XlsIO;
using System;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Web.Mvc;
using WebApplication2.Models;
using Workbook = Microsoft.Office.Interop.Excel.Workbook;
using Worksheet = Microsoft.Office.Interop.Excel.Worksheet;

namespace WebApplication2.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            var model = new formsubmito();
            return View(model);
        }
        [HttpPost]
        [ValidateAntiForgeryToken]
        public FileResult Index(formsubmito model)
        {
           return MapAndUpdateExcelInterop(model);
        }

        //private object MapAndUpdateExcel(formsubmito model)
        //{
        //    //Create an instance of ExcelEngine
        //    using (ExcelEngine excelEngine = new ExcelEngine())
        //    {
        //        //Instantiate the Excel application object
        //        IApplication application = excelEngine.Excel;

        //        //Set the default application version
        //        application.DefaultVersion = ExcelVersion.Xlsx;

        //        //Load the existing Excel workbook into IWorkbook
        //        IWorkbook workbook = application.Workbooks.Open(Server.MapPath("template.xlsx"));

        //        //Get the first worksheet in the workbook into IWorksheet
        //        IWorksheet worksheet = workbook.Worksheets[0];

        //        worksheet.Replace("FindValue", "NewValue");

        //        worksheet.Replace(
        //       "xx1",
        //       model.Name);
        //        worksheet.Replace(
        //            "xx2",
        //            model.Name);

        //        worksheet.Replace(
        //            "xx3",
        //            model.FileNumber);

        //        worksheet.Replace(
        //            "xx4",
        //            model.Age);

        //        worksheet.Replace(
        //            "xx5",
        //            model.Nationality);

        //        worksheet.Replace(
        //            "xx6",
        //            model.Date);

        //        worksheet.Replace(
        //            "xx7",
        //            model.Clinic);

        //        worksheet.Replace(
        //            "xx8",
        //            GetEnglishCountryName(model));

        //        //Save the Excel document
        //        workbook.SaveAs("Output.xlsx", HttpContext.ApplicationInstance.Response, ExcelDownloadType.Open);
        //        return null;
        //    }
        //}

        private FileResult MapAndUpdateExcelInterop(formsubmito model)
        {
            object m = Type.Missing;

            // open excel.
            Application app = new ApplicationClass();

            // open the workbook. 
            Workbook wb = app.Workbooks.Open(
                Server.MapPath(@"~\template.xlsx"),
                m, false, m, m, m, m, m, m, m, m, m, m, m, m);

            // get the active worksheet. (Replace this if you need to.) 
            Microsoft.Office.Interop.Excel.Worksheet ws = (Worksheet)wb.ActiveSheet;

            // get the used range. 
            Range r = (Range)ws.UsedRange;
      
            // call the replace method to replace instances. 
            r.Replace(
                 "xx1",
                 model.Name,
                 XlLookAt.xlWhole);
            r.Replace(
                "xx2",
                model.Name,
                XlLookAt.xlWhole,
                XlSearchOrder.xlByRows,
                true, m, m, m);

            r.Replace(
                "xx3",
                model.FileNumber,
                XlLookAt.xlWhole,
                XlSearchOrder.xlByRows,
                true, m, m, m);

            r.Replace(
                "xx4",
                model.Age,
                XlLookAt.xlWhole,
                XlSearchOrder.xlByRows,
                true, m, m, m);

            r.Replace(
                "xx5",
                model.Nationality,
                XlLookAt.xlWhole,
                XlSearchOrder.xlByRows,
                true, m, m, m);

            r.Replace(
                "xx6",
                model.Date,
                XlLookAt.xlWhole,
                XlSearchOrder.xlByRows,
                true, m, m, m);

            r.Replace(
                "xx7",
                model.Clinic,
                XlLookAt.xlWhole,
                XlSearchOrder.xlByRows,
                true, m, m, m);

            r.Replace(
                "xx8",
                GetEnglishCountryName(model),
                XlLookAt.xlWhole,
                XlSearchOrder.xlByRows,
                true, m, m, m);
            // save and close. 
            // wb.SaveAs("FileName.xlsx", XlFileFormat.xlOpenXMLWorkbook);
            // Return FileResult
            object misValue = System.Reflection.Missing.Value;
      

            var Parentpath = $@"C:\out";
            System.IO.Directory.CreateDirectory(Parentpath);
            var path = $@"{Parentpath}\csharp-Excel{DateTime.Now.Ticks}.xlsx";
            wb.SaveAs(path, XlFileFormat.xlOpenXMLWorkbook);
            wb.Close(true, misValue, misValue);
            app.Quit();

            Marshal.ReleaseComObject(ws);
            Marshal.ReleaseComObject(wb);
            Marshal.ReleaseComObject(app);
            return File(path, "application/vnd.ms-excel", "WidgetData.xlsx");
        }

        private string GetEnglishCountryName(formsubmito model)
        {
            var names = RestCountriesService.GetAllCountries();
            var needed = names.Where(c => c.Translations["ara"].Common == model.Nationality).FirstOrDefault();
            return needed.Name.Common;
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
    }
}