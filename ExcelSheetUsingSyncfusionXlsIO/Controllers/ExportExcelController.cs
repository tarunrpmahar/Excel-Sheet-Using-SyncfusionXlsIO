using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore.Metadata.Internal;
using OfficeOpenXml;
using OfficeOpenXml.DataValidation;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;
using Syncfusion.XlsIO;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Web;
using static System.Net.Mime.MediaTypeNames;
using Syncfusion.Drawing;
using ExcelSheetUsingSyncfusionXlsIO.Models;

namespace ExportExcelDemo.Controllers
{
    public class ExcelController : Controller
    {
        public readonly ExcelSheetDBContext _excelDBContext;
        public ExcelController(ExcelSheetDBContext excelDBContext)
        {
            _excelDBContext = excelDBContext;
        }
        public ActionResult Index()
        {
            return View();
        }
        public IActionResult PlanSheet()
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;

                application.DefaultVersion = ExcelVersion.Excel2016;

                IWorkbook workbook = application.Workbooks.Create(1);
                IWorksheet worksheet = workbook.Worksheets[0];

                /*worksheet.IsGridLinesVisible = false;*/

                worksheet.Range["A7:A27"].CellStyle.Color = Color.FromArgb(144, 238, 144);

                worksheet.Range["A7"].Text = "P0T9 + P0TD";
                worksheet.Range["B7"].Text = "jan-18";
                worksheet.Range["C7"].Text = "feb-18";
                worksheet.Range["D7"].Text = "mar-18";
                worksheet.Range["E7"].Text = "apr-18";
                worksheet.Range["F7"].Text = "May-18";
                worksheet.Range["G7"].Text = "Jun-18";
                worksheet.Range["H7"].Text = "Jul-18";
                worksheet.Range["I7"].Text = "Aug-18";
                worksheet.Range["J7"].Text = "Sep-18";
                worksheet.Range["K7"].Text = "Oct-18";
                worksheet.Range["L7"].Text = "Nov-18";
                worksheet.Range["M7"].Text = "Dec-18";

                worksheet.Range["B7:M7"].CellStyle.Color = Color.FromArgb(144, 238, 144);

                worksheet.Range["A7"].Text = "P0T9 + P0TD";
                worksheet.Range["A8"].Text = "AOP";
                worksheet.Range["A9"].Text = "Track 10";
                /*worksheet.Range["A10"].Text = "Statistical Forecast";*/
                worksheet.Range["A11"].Text = "Jan-18";
                worksheet.Range["A12"].Text = "feb-18";
                worksheet.Range["A13"].Text = "mar-18";
                worksheet.Range["A14"].Text = "apr-18";
                worksheet.Range["A15"].Text = "May-18";
                worksheet.Range["A16"].Text = "Jun-18";
                worksheet.Range["A17"].Text = "Jul-18";
                worksheet.Range["A18"].Text = "Aug-18";
                worksheet.Range["A19"].Text = "Sep-18";
                worksheet.Range["A20"].Text = "Oct-18";
                worksheet.Range["A21"].Text = "Nov-18";
                worksheet.Range["A22"].Text = "Dec-18";
                worksheet.Range["A23"].Text = "Actuals";
                worksheet.Range["A24"].Text = "KSQM";
                worksheet.Range["A25"].Text = "Sales accuracy";
                worksheet.Range["A26"].Text = "Impact on inventory";
                worksheet.Range["A27"].Text = "Statistical SFA";
                

                worksheet.Range["A8:A9"].CellStyle.Font.RGBColor = Color.FromArgb(255, 0, 0);
                worksheet.Range["A8:A9"].CellStyle.Font.Bold = true;

                worksheet.Range["B7:M27"].CellStyle.HorizontalAlignment = ExcelHAlign.HAlignRight;
                worksheet.Range["B7:M27"].CellStyle.VerticalAlignment = ExcelVAlign.VAlignBottom;

                worksheet.Range["A8:A23"].CellStyle.HorizontalAlignment = ExcelHAlign.HAlignRight;
                worksheet.Range["A8:A23"].CellStyle.VerticalAlignment = ExcelVAlign.VAlignCenter;

                worksheet.Range["A24:A27"].CellStyle.HorizontalAlignment = ExcelHAlign.HAlignLeft;

                worksheet.Range["A7:A27"].ColumnWidth = 25;
                worksheet.Range["A7:M7"].RowHeight = 25;

                MemoryStream stream = new MemoryStream();

                workbook.SaveAs(stream);

                //Set the position as '0'.
                stream.Position = 0;

                //Download the Excel file in the browser
                FileStreamResult fileStreamResult = new FileStreamResult(stream, "application/excel");

                fileStreamResult.FileDownloadName = "planSheet.xlsx";

                return fileStreamResult;
            }
        }

        public IActionResult CreateDocument()
        {
            //Create an instance of ExcelEngine
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;

                application.DefaultVersion = ExcelVersion.Excel2016;

                //Create a workbook
                IWorkbook workbook = application.Workbooks.Create(1);
                IWorksheet worksheet = workbook.Worksheets[0];

                //Adding a picture
                /*FileStream imageStream = new FileStream("AdventureCycles-Logo.png", FileMode.Open, FileAccess.Read);
                IPictureShape shape = worksheet.Pictures.AddPicture(1, 1, imageStream);*/

                //Disable gridlines in the worksheet
                worksheet.IsGridLinesVisible = false;

                //Enter values to the cells from A3 to A5
                worksheet.Range["A3"].Text = "46036 Michigan Ave";
                worksheet.Range["A4"].Text = "Canton, USA";
                worksheet.Range["A5"].Text = "Phone: +1 231-231-2310";

                //Make the text bold
                worksheet.Range["A3:A5"].CellStyle.Font.Bold = true;

                //Merge cells
                worksheet.Range["D1:E1"].Merge();

                //Enter text to the cell D1 and apply formatting.
                worksheet.Range["D1"].Text = "INVOICE";
                worksheet.Range["D1"].CellStyle.Font.Bold = true;
                worksheet.Range["D1"].CellStyle.Font.RGBColor = Color.FromArgb(42, 118, 189);
                worksheet.Range["D1"].CellStyle.Font.Size = 35;

                //Apply alignment in the cell D1
                worksheet.Range["D1"].CellStyle.HorizontalAlignment = ExcelHAlign.HAlignRight;
                worksheet.Range["D1"].CellStyle.VerticalAlignment = ExcelVAlign.VAlignTop;

                //Enter values to the cells from D5 to E8
                worksheet.Range["D5"].Text = "INVOICE#";
                worksheet.Range["E5"].Text = "DATE";
                worksheet.Range["D6"].Number = 1028;
                worksheet.Range["E6"].Value = "12/31/2018";
                worksheet.Range["D7"].Text = "CUSTOMER ID";
                worksheet.Range["E7"].Text = "TERMS";
                worksheet.Range["D8"].Number = 564;
                worksheet.Range["E8"].Text = "Due Upon Receipt";

                //Apply RGB backcolor to the cells from D5 to E8
                worksheet.Range["D5:E5"].CellStyle.Color = Color.FromArgb(42, 118, 189);
                worksheet.Range["D7:E7"].CellStyle.Color = Color.FromArgb(42, 118, 189);

                //Apply known colors to the text in cells D5 to E8
                worksheet.Range["D5:E5"].CellStyle.Font.Color = ExcelKnownColors.White;
                worksheet.Range["D7:E7"].CellStyle.Font.Color = ExcelKnownColors.White;

                //Make the text as bold from D5 to E8
                worksheet.Range["D5:E8"].CellStyle.Font.Bold = true;

                //Apply alignment to the cells from D5 to E8
                worksheet.Range["D5:E8"].CellStyle.HorizontalAlignment = ExcelHAlign.HAlignCenter;
                worksheet.Range["D5:E5"].CellStyle.VerticalAlignment = ExcelVAlign.VAlignCenter;
                worksheet.Range["D7:E7"].CellStyle.VerticalAlignment = ExcelVAlign.VAlignCenter;
                worksheet.Range["D6:E6"].CellStyle.VerticalAlignment = ExcelVAlign.VAlignTop;

                //Enter value and applying formatting in the cell A7
                worksheet.Range["A7"].Text = "  BILL TO";
                worksheet.Range["A7"].CellStyle.Color = Color.FromArgb(42, 118, 189);
                worksheet.Range["A7"].CellStyle.Font.Bold = true;
                worksheet.Range["A7"].CellStyle.Font.Color = ExcelKnownColors.White;

                //Apply alignment
                worksheet.Range["A7"].CellStyle.HorizontalAlignment = ExcelHAlign.HAlignLeft;
                worksheet.Range["A7"].CellStyle.VerticalAlignment = ExcelVAlign.VAlignCenter;

                //Enter values in the cells A8 to A12
                worksheet.Range["A8"].Text = "Steyn";
                worksheet.Range["A9"].Text = "Great Lakes Food Market";
                worksheet.Range["A10"].Text = "20 Whitehall Rd";
                worksheet.Range["A11"].Text = "North Muskegon,USA";
                worksheet.Range["A12"].Text = "+1 231-654-0000";

                //Create a Hyperlink for e-mail in the cell A13
                IHyperLink hyperlink = worksheet.HyperLinks.Add(worksheet.Range["A13"]);
                hyperlink.Type = ExcelHyperLinkType.Url;
                hyperlink.Address = "Steyn@greatlakes.com";
                hyperlink.ScreenTip = "Send Mail";

                //Merge column A and B from row 15 to 22
                worksheet.Range["A15:B15"].Merge();
                worksheet.Range["A16:B16"].Merge();
                worksheet.Range["A17:B17"].Merge();
                worksheet.Range["A18:B18"].Merge();
                worksheet.Range["A19:B19"].Merge();
                worksheet.Range["A20:B20"].Merge();
                worksheet.Range["A21:B21"].Merge();
                worksheet.Range["A22:B22"].Merge();

                //Enter details of products and prices
                worksheet.Range["A15"].Text = "  DESCRIPTION";
                worksheet.Range["C15"].Text = "QTY";
                worksheet.Range["D15"].Text = "UNIT PRICE";
                worksheet.Range["E15"].Text = "AMOUNT";
                worksheet.Range["A16"].Text = "Cabrales Cheese";
                worksheet.Range["A17"].Text = "Chocos";
                worksheet.Range["A18"].Text = "Pasta";
                worksheet.Range["A19"].Text = "Cereals";
                worksheet.Range["A20"].Text = "Ice Cream";
                worksheet.Range["C16"].Number = 3;
                worksheet.Range["C17"].Number = 2;
                worksheet.Range["C18"].Number = 1;
                worksheet.Range["C19"].Number = 4;
                worksheet.Range["C20"].Number = 3;
                worksheet.Range["D16"].Number = 21;
                worksheet.Range["D17"].Number = 54;
                worksheet.Range["D18"].Number = 10;
                worksheet.Range["D19"].Number = 20;
                worksheet.Range["D20"].Number = 30;
                worksheet.Range["D23"].Text = "Total";

                //Apply number format
                worksheet.Range["D16:E22"].NumberFormat = "$.00";
                worksheet.Range["E23"].NumberFormat = "$.00";

                //Apply incremental formula for column Amount by multiplying Qty and UnitPrice
                application.EnableIncrementalFormula = true;
                worksheet.Range["E16:E20"].Formula = "=C16*D16";

                //Formula for Sum the total
                worksheet.Range["E23"].Formula = "=SUM(E16:E22)";

                //Apply borders
                worksheet.Range["A16:E22"].CellStyle.Borders[ExcelBordersIndex.EdgeTop].LineStyle = ExcelLineStyle.Thin;
                worksheet.Range["A16:E22"].CellStyle.Borders[ExcelBordersIndex.EdgeBottom].LineStyle = ExcelLineStyle.Thin;
                worksheet.Range["A16:E22"].CellStyle.Borders[ExcelBordersIndex.EdgeTop].Color = ExcelKnownColors.Grey_25_percent;
                worksheet.Range["A16:E22"].CellStyle.Borders[ExcelBordersIndex.EdgeBottom].Color = ExcelKnownColors.Grey_25_percent;
                worksheet.Range["A23:E23"].CellStyle.Borders[ExcelBordersIndex.EdgeTop].LineStyle = ExcelLineStyle.Thin;
                worksheet.Range["A23:E23"].CellStyle.Borders[ExcelBordersIndex.EdgeBottom].LineStyle = ExcelLineStyle.Thin;
                worksheet.Range["A23:E23"].CellStyle.Borders[ExcelBordersIndex.EdgeTop].Color = ExcelKnownColors.Black;
                worksheet.Range["A23:E23"].CellStyle.Borders[ExcelBordersIndex.EdgeBottom].Color = ExcelKnownColors.Black;

                //Apply font setting for cells with product details
                worksheet.Range["A3:E23"].CellStyle.Font.FontName = "Arial";
                worksheet.Range["A3:E23"].CellStyle.Font.Size = 10;
                worksheet.Range["A15:E15"].CellStyle.Font.Color = ExcelKnownColors.White;
                worksheet.Range["A15:E15"].CellStyle.Font.Bold = true;
                worksheet.Range["D23:E23"].CellStyle.Font.Bold = true;

                //Apply cell color
                worksheet.Range["A15:E15"].CellStyle.Color = Color.FromArgb(42, 118, 189);

                //Apply alignment to cells with product details
                worksheet.Range["A15"].CellStyle.HorizontalAlignment = ExcelHAlign.HAlignLeft;
                worksheet.Range["C15:C22"].CellStyle.HorizontalAlignment = ExcelHAlign.HAlignCenter;
                worksheet.Range["D15:E15"].CellStyle.HorizontalAlignment = ExcelHAlign.HAlignCenter;

                //Apply row height and column width to look good
                worksheet.Range["A1"].ColumnWidth = 36;
                worksheet.Range["B1"].ColumnWidth = 11;
                worksheet.Range["C1"].ColumnWidth = 8;
                worksheet.Range["D1:E1"].ColumnWidth = 18;
                worksheet.Range["A1"].RowHeight = 47;
                worksheet.Range["A2"].RowHeight = 15;
                worksheet.Range["A3:A4"].RowHeight = 15;
                worksheet.Range["A5"].RowHeight = 18;
                worksheet.Range["A6"].RowHeight = 29;
                worksheet.Range["A7"].RowHeight = 18;
                worksheet.Range["A8"].RowHeight = 15;
                worksheet.Range["A9:A14"].RowHeight = 15;
                worksheet.Range["A15:A23"].RowHeight = 18;


                //Saving the Excel to the MemoryStream 
                MemoryStream stream = new MemoryStream();

                workbook.SaveAs(stream);

                //Set the position as '0'.
                stream.Position = 0;

                //Download the Excel file in the browser
                FileStreamResult fileStreamResult = new FileStreamResult(stream, "application/excel");

                fileStreamResult.FileDownloadName = "Output.xlsx";

                return fileStreamResult;
            }
        }
    }
}
