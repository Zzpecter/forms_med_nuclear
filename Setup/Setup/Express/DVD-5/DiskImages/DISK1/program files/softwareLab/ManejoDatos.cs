using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Net.SourceForge.Koogra;
using System.IO;
using System.Data;
using ClosedXML;
using ClosedXML.Excel;
using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

namespace DAL
{
    public class ManejoDatos
    {
        public string WriteExcelHema(System.Data.DataTable dtDatos, string nombre, string doctor, DateTime fecha)
        {
            XLWorkbook wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Hematologia");

            ws.Cell("B2").Value = "Universidad San Francisco Xavier";
            ws.Cells("B2").Style.Font.FontSize = 16;
            ws.Cells("B2").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            ws.Cell("B3").Value = "Instituto Medicina Nuclear";
            ws.Cells("B3").Style.Font.FontSize = 16;
            ws.Cells("B3").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            ws.Cell("B5").Value = "Hematología";
            ws.Cells("B5").Style.Font.FontSize = 14;
            ws.Cells("B5").Style.Font.SetBold(true);
            ws.Cell("B6").Value = "Nombre: " + nombre;
            ws.Cell("E6").Value = "Fecha: " + fecha.ToShortDateString();
            ws.Cell("B7").Value = "Doctor: " + doctor;
            ws.Range("B2:E2").Merge();
            ws.Range("B3:E3").Merge();
            ws.Cell("B2").Style.Alignment.WrapText = true;
            ws.Cell("B3").Style.Alignment.WrapText = true;

            var tableWithData = ws.Cell("B9").InsertTable(dtDatos.AsEnumerable());

            ws.Columns().AdjustToContents();

            string ruta = @"E:\Examenes\hematologia" + fecha.Year + fecha.Month + fecha.Day + fecha.Hour + fecha.Minute + ".xlsx";
            wb.SaveAs(ruta);
            return ruta;
        }

        public string WriteExcelQuimina(System.Data.DataTable dtDatos, string nombre, string doctor, DateTime fecha)
        {
            XLWorkbook wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Quimica");

            ws.Cell("B2").Value = "Universidad San Francisco Xavier";
            ws.Cells("B2").Style.Font.FontSize = 16;
            ws.Cells("B2").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            ws.Cell("B3").Value = "Instituto Medicina Nuclear";
            ws.Cells("B3").Style.Font.FontSize = 16;
            ws.Cells("B3").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            ws.Cell("B5").Value = "Química Sanguínea";
            ws.Cells("B5").Style.Font.FontSize = 11;
            ws.Cells("B5").Style.Font.SetBold(true);
            ws.Cell("B6").Value = "Nombre: " + nombre;
            ws.Cell("E6").Value = "Fecha: " + fecha.ToShortDateString();
            ws.Cell("B7").Value = "Doctor: " + doctor;
            ws.Range("B2:E2").Merge();
            ws.Range("B3:E3").Merge();
            ws.Cell("B2").Style.Alignment.WrapText = true;
            ws.Cell("B3").Style.Alignment.WrapText = true;

            var tableWithData = ws.Cell("B9").InsertTable(dtDatos.AsEnumerable());

            ws.Columns().AdjustToContents();

            string ruta = @"E:\Examenes\quimica" + fecha.Year + fecha.Month + fecha.Day + fecha.Hour + fecha.Minute + ".xlsx";
            wb.SaveAs(ruta);
            return ruta;
        }

        public string WriteExcelRadio(System.Data.DataTable dtDatos, string nombre, string doctor, DateTime fecha)
        {
            XLWorkbook wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Radio");

            ws.Cell("B2").Value = "Universidad San Francisco Xavier";
            ws.Cells("B2").Style.Font.FontSize = 16;
            ws.Cells("B2").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            ws.Cell("B3").Value = "Instituto Medicina Nuclear";
            ws.Cells("B3").Style.Font.FontSize = 16;
            ws.Cells("B3").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            ws.Cell("B5").Value = "RadioinmunoAnálisis";
            ws.Cells("B5").Style.Font.FontSize = 14;
            ws.Cells("B5").Style.Font.SetBold(true);
            ws.Cell("B6").Value = "Nombre: " + nombre;
            ws.Cell("E6").Value = "Fecha: " + fecha.ToShortDateString();
            ws.Cell("B7").Value = "Doctor: " + doctor;
            ws.Range("B2:E2").Merge();
            ws.Range("B3:E3").Merge();
            ws.Cell("B2").Style.Alignment.WrapText = true;
            ws.Cell("B3").Style.Alignment.WrapText = true;

            var tableWithData = ws.Cell("B9").InsertTable(dtDatos.AsEnumerable());

            ws.Style.Font.FontSize = 9;
            ws.Columns().AdjustToContents();

            string ruta = @"E:\Examenes\radio" + fecha.Year + fecha.Month + fecha.Day + fecha.Hour + fecha.Minute + ".xlsx";
            wb.SaveAs(ruta);
            return ruta;
        }

        public System.Data.DataTable ReadExcelContent(string filePath)
        {
            System.Data.DataTable dtDatos = new System.Data.DataTable();
            dtDatos.Columns.Add("unidades");
            dtDatos.Columns.Add("VR");
            string a = null;
            string b = null;

            var data = new StringBuilder();
            try
            {
                IWorkbook wb = null;
                    wb = Net.SourceForge.Koogra.WorkbookFactory.GetExcel2007Reader(filePath);

                IWorksheet ws = wb.Worksheets.GetWorksheetByIndex(0);

                for (uint r = ws.FirstRow; r <= ws.LastRow; ++r)
                {
                    IRow row = ws.Rows.GetRow(r);
                    if (row != null)
                        for (uint colCount = ws.FirstCol; colCount <= ws.LastCol; ++colCount)
                            if (row.GetCell(colCount) != null && row.GetCell(colCount).Value != null)
                            {
                               if( colCount == ws.FirstCol)
                                   a = row.GetCell(colCount).Value.ToString();
                               else
                                   b = row.GetCell(colCount).Value.ToString();
                            }
                    dtDatos.Rows.Add(a, b);
                }

            }
            catch (Exception ex)
            {
                Exception exception = ex;
                exception.Source = string.Format("Error occured");
                throw ex;
            }

            return dtDatos;
        }

        public void PrintExcel(string filePath, bool horizontal)
        {
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            // Open the Workbook:
            Microsoft.Office.Interop.Excel.Workbook wb = excelApp.Workbooks.Open(
                filePath,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            // Get the first worksheet.
            // (Excel uses base 1 indexing, not base 0.)
            Microsoft.Office.Interop.Excel.Worksheet ws = (Microsoft.Office.Interop.Excel.Worksheet)wb.Worksheets[1];

            if (horizontal)
                ws.PageSetup.Orientation = XlPageOrientation.xlLandscape;
            // Print out 1 copy to the default printer:
            ws.PrintOut(Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            // Cleanup:
            GC.Collect();
            GC.WaitForPendingFinalizers();

            Marshal.FinalReleaseComObject(ws);

            wb.Close(false, Type.Missing, Type.Missing);
            Marshal.FinalReleaseComObject(wb);

            excelApp.Quit();
            Marshal.FinalReleaseComObject(excelApp);
        }
    }
}
