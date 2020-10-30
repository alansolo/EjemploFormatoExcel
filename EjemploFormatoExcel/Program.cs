using LinqToExcel;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Drawing;
using Excel = Microsoft.Office.Interop.Excel;
using OfficeOpenXml;
using System.IO;
using OfficeOpenXml.Style;
using System.Data;

namespace EjemploFormatoExcel
{
    class Program
    {
        static void Main(string[] args)
        {

            //ExcelQueryFactory excel = new ExcelQueryFactory("C:\\Users\\k697344\\Documents\\Comex PPG\\Documentacion\\CEO-2019-NR-11.xls");

            //Worksheet workSheet = (Worksheet)excel.Worksheet("Table1");

            //workSheet.Range["A1", "A2"].Interior.Color = Color.Green;

            //var excelApp = new Excel.Application();
            //excelApp.Visible = true;

            //excelApp.Worksheets()

            System.Data.DataTable dtDatos = new System.Data.DataTable();
            dtDatos.Columns.Add("ID");
            dtDatos.Columns.Add("FirstName");
            dtDatos.Columns.Add("LastName");
            dtDatos.Columns.Add("DOB");

            dtDatos.Rows.Add("1", "Alan", "Solorzano", "alan_solo");

            using (ExcelPackage excel = new ExcelPackage())
            {
                excel.Workbook.Worksheets.Add("Ejemplo");

                var worksheet = excel.Workbook.Worksheets["Ejemplo"];

                var headerRow = new List<string[]>()
                {
                    new string[] { "ID", "First Name", "Last Name", "DOB" }
                };

                worksheet.Cells["A1:" + Char.ConvertFromUtf32(headerRow[0].Length + 64) + "1"].LoadFromArrays(headerRow);
                worksheet.Cells["A1:" + Char.ConvertFromUtf32(headerRow[0].Length + 64) + "1"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                worksheet.Cells["A1:" + Char.ConvertFromUtf32(headerRow[0].Length + 64) + "1"].Style.Fill.BackgroundColor.SetColor(1, 38, 130, 221);
                worksheet.Cells["A1:" + Char.ConvertFromUtf32(headerRow[0].Length + 64) + "1"].Style.Font.Color.SetColor(Color.White);
            
                //LoadFromDataTableworksheet.Cells["A1:A4"].LoadFromDataTable
                worksheet.Cells["A3:" + Char.ConvertFromUtf32(headerRow[0].Length + 64) + "3"].LoadFromDataTable(dtDatos, true, OfficeOpenXml.Table.TableStyles.Light2);


                FileInfo excelFile = new FileInfo("C:\\Users\\k697344\\Documents\\Comex PPG\\Documentacion\\Ejemplo.xlsx");
                excel.SaveAs(excelFile);
            }
        }
    }
}
