using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using ClosedXML.Excel;

namespace ClosedXmlStudy
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnExcelExport_Click(object sender, EventArgs e)
        {
            var wb = new XLWorkbook(); // Creating a new workbook
            var ws = wb.Worksheets.Add("Contacts"); // Adding a worksheet

            // Adding text

            // Title
            ws.Cell("B2").Value = "Contacts";

            // First Names
            ws.Cell("B3").Value = "FName";
            ws.Cell("B4").Value = "John";
            ws.Cell("B5").Value = "Hank";
            ws.Cell("B6").Value = "Dagny";

            // Last Names
            ws.Cell("C3").Value = "LName";
            ws.Cell("C4").Value = "Galt";
            ws.Cell("C5").Value = "Rearden";
            ws.Cell("C6").Value = "Taggart";

            // Adding more data types

            // Boolean
            ws.Cell("D3").Value = "OutCast";
            ws.Cell("D4").Value = true;
            ws.Cell("D5").Value = false;
            ws.Cell("D6").Value = false;

            // DateTime
            ws.Cell("E3").Value = "DOB";
            ws.Cell("E4").Value = new DateTime(1919, 1, 21);
            ws.Cell("E5").Value = new DateTime(1907, 3, 4);
            ws.Cell("E6").Value = new DateTime(1921, 12, 25);

            // Numeric 
            ws.Cell("F3").Value = "Income";
            ws.Cell("F4").Value = 2000;
            ws.Cell("F5").Value = 40000;
            ws.Cell("F6").Value = 10000;

            // Defining ranges

            // From worksheet
            var rngTable = ws.Range("B2:F6");

            // From another range
            var rngDates = rngTable.Range("D3:D5"); // The address is relative to rngTable (NOT the worksheet)
            var rngNumbers = rngTable.Range("E3:E5"); // The address is relative to rngTable (NOT the worksheet)

            // Formatting dates and numbers

            // Using OpenXML's predefined formats
            rngDates.Style.NumberFormat.NumberFormatId = 15;

            // Using a custom format
            rngNumbers.Style.NumberFormat.Format = "$ #,##0";

            // Formating headers 
            var rngHeaders = rngTable.Range("A2:E2"); // The address is relative to rngTable (NOT the worksheet)
            rngHeaders.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            rngHeaders.Style.Font.Bold = true;
            rngHeaders.Style.Fill.BackgroundColor = XLColor.Aqua;

            // Adding grid lines
            rngTable.Style.Border.BottomBorder = XLBorderStyleValues.Thin;

            // Format title cell
            rngTable.Cell(1, 1).Style.Font.Bold = true;
            rngTable.Cell(1, 1).Style.Fill.BackgroundColor = XLColor.CornflowerBlue;
            rngTable.Cell(1, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

            // Merge title cells
            rngTable.Row(1).Merge(); // We could've also used: rngTable.Range("A1:E1").Merge()

            // Add thick borders
            // Add a thick outside border
            rngTable.Style.Border.OutsideBorder = XLBorderStyleValues.Thick;

            // You can also specify the border for each side with:
            // rngTable.FirstColumn().Style.Border.LeftBorder = XLBorderStyleValues.Thick;
            // rngTable.LastColumn().Style.Border.RightBorder = XLBorderStyleValues.Thick;
            // rngTable.FirstRow().Style.Border.TopBorder = XLBorderStyleValues.Thick;
            // rngTable.LastRow().Style.Border.BottomBorder = XLBorderStyleValues.Thick;

            // Adjust column widths to their content
            ws.Columns(2, 6).AdjustToContents();


            ws = wb.Worksheets.Add("Inserting Data");

            // From a list of arrays
            var listOfArr = new List<Int32[]>();
            listOfArr.Add(new Int32[] { 1, 2, 3 });
            listOfArr.Add(new Int32[] { 1 });
            listOfArr.Add(new Int32[] { 1, 2, 3, 4, 5, 6 });

            ws.Cell(1, 3).Value = "From Arrays";
            ws.Range(1, 3, 1, 8).Merge().AddToNamed("Titles");
            var rangeWithArrays = ws.Cell(2, 3).InsertData(listOfArr);

            // Saving the workbook
            wb.SaveAs("D:\\Projects\\ClosedXmlStudy\\ClosedXmlStudy\\BasicTable.xlsx");
        }
    }
}
