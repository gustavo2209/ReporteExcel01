using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

using Microsoft.Office.Interop.Excel;

namespace ReporteExcel01.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Excel()
        {
            ExcelCrea("C:\\temp\\demo.xlsx");
            return File("C:\\temp\\demo.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "MiExcel.xlsx");
        }

        public virtual void ExcelCrea(string doc_excel)
        {
            Application excelApp = new Application();
            Workbook wb = excelApp.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);
            Worksheet sheet = wb.Sheets.Add();
            //--------------------------------------------------------------------------

            sheet.Columns.AutoFit();

            sheet.Cells[1, 1] = "Hola Mundo";

            //--------------------------------------------------------------------------
            wb.SaveAs(doc_excel,
                 XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing,
                 Type.Missing, Type.Missing, XlSaveAsAccessMode.xlExclusive,
                 Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            wb.Close();
            excelApp.Quit();
        }

        public ActionResult Vertical()
        {
            VerticalCrea("C:\\temp\\vertical.xlsx");
            return File("C:\\temp\\vertical.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "Vertical.xlsx");
        }

        public virtual void VerticalCrea(string doc_excel)
        {
            Application excelApp = new Application();
            Workbook wb = excelApp.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);
            Worksheet sheet = wb.Sheets.Add();
            //--------------------------------------------------------------------------

            for(int i=1; i<=100; i++)
            {
                sheet.Cells[i, 1] = i;
            }

            //--------------------------------------------------------------------------
            wb.SaveAs(doc_excel,
                 XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing,
                 Type.Missing, Type.Missing, XlSaveAsAccessMode.xlExclusive,
                 Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            wb.Close();
            excelApp.Quit();
        }

        public ActionResult Horizontal()
        {
            HorizontalCrea("C:\\temp\\horizontal.xlsx");
            return File("C:\\temp\\horizontal.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "Horizontal.xlsx");
        }

        public virtual void HorizontalCrea(string doc_excel)
        {
            Application excelApp = new Application();
            Workbook wb = excelApp.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);
            Worksheet sheet = wb.Sheets.Add();
            //--------------------------------------------------------------------------

            for (int i = 1; i <= 100; i++)
            {
                sheet.Cells[1, i] = i;
            }

            //--------------------------------------------------------------------------
            wb.SaveAs(doc_excel,
                 XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing,
                 Type.Missing, Type.Missing, XlSaveAsAccessMode.xlExclusive,
                 Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            wb.Close();
            excelApp.Quit();
        }

        public ActionResult Matricial()
        {
            MatricialCrea("C:\\temp\\matricial.xlsx");
            return File("C:\\temp\\matricial.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "Matricial.xlsx");
        }

        public virtual void MatricialCrea(string doc_excel)
        {
            Application excelApp = new Application();
            Workbook wb = excelApp.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);
            Worksheet sheet = wb.Sheets.Add();
            //--------------------------------------------------------------------------

            for (int i = 1; i <= 100; i++)
            {
                for (int j = 1; j <= 100; j++)
                {
                    sheet.Cells[i, j] = i;
                }
                
            }

            //--------------------------------------------------------------------------
            wb.SaveAs(doc_excel,
                 XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing,
                 Type.Missing, Type.Missing, XlSaveAsAccessMode.xlExclusive,
                 Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            wb.Close();
            excelApp.Quit();
        }

        public ActionResult Formulas()
        {
            FormulasCrea("C:\\temp\\formulas.xlsx");
            return File("C:\\temp\\formulas.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "Formulas.xlsx");
        }

        public virtual void FormulasCrea(string doc_excel)
        {
            Application excelApp = new Application();
            Workbook wb = excelApp.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);
            Worksheet sheet = wb.Sheets.Add();
            //--------------------------------------------------------------------------

            for(int i = 1; i <= 20; i++)
            {
                sheet.Cells[i, 1] = i;
                sheet.Cells[i, 2] = "=SUM(A1:A" + i + ")";
                sheet.Cells[i, 4] = "=AVERAGE(A1:A" + i + ")";
            }

            sheet.Cells[21, 1] = "=SUM(A1:A20)";

            //--------------------------------------------------------------------------
            wb.SaveAs(doc_excel,
                 XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing,
                 Type.Missing, Type.Missing, XlSaveAsAccessMode.xlExclusive,
                 Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            wb.Close();
            excelApp.Quit();
        }

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
    }
}