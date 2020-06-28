using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

using Excel = Microsoft.Office.Interop.Excel;

namespace AuditSheet
{
    public class ExcelHelper
    {
        public static void setBorderCell(Excel._Worksheet xlWorksheet, int row, int col)
        {
            Excel.Range cell1 = xlWorksheet.Cells[row, col];
            cell1.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            cell1.Borders.Weight = Excel.XlBorderWeight.xlThin;
        }

        public static void setCellWidth(Excel._Worksheet xlWorksheet, int col, int width = 16)
        {
            xlWorksheet.Columns[col].ColumnWidth = width;
        }

        public static void setCurrencyCell(Excel._Worksheet xlWorksheet, int row, int cell)
        {
            //Excel.Range cell1 = xlWorksheet.Cells[row, cell];
            //cell1.NumberFormat = "#,###.00";
        }
    }
}