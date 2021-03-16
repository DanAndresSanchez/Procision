using System;
using System.Data;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Microsoft.Office.Interop.Excel;
using _Excel = Microsoft.Office.Interop.Excel;

namespace Procision
{
    class DataClense
    {
        int lastUsedRow, lastUsedCol;
        _Excel.Worksheet final_report;

        public void centerAlign()
        {
            var rng = final_report.get_Range("A1:R300", Type.Missing);
            rng.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
            rng.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
            rng.Cells.Columns.AutoFit();
            rng.Cells.Rows.AutoFit();
        }

        public void clearSheet()
        {
            getSheetInfo();
            var rng = final_report.get_Range("A2:Z300", Type.Missing);
            rng.Cells.Clear();
            Console.WriteLine("Sheet Cleared");
        }

        //Working on this in the reportFromDB() method, feel free to use this instead if its easier
        public void createBorder()
        {
            string previous = "";
            string current = "";

            for (int row = 2; row < lastUsedRow; row++)
            {
                current = "";
                current += final_report.Cells[row, 2].Value() + "," + final_report.Cells[row, 3].Value() + "," + final_report.Cells[row, 4].Value();
                if (row != 2)
                {
                    if (!(previous.Equals(current)))
                    {
                        var rng = "A" + row + ":R" + lastUsedRow;
                        final_report.Cells.Range[rng].Borders[_Excel.XlBordersIndex.xlEdgeTop].Weight = 3d;
                    }
                }

                previous = current;
            }
        }

        //Initializes final report sheet
        public void getSheetInfo()
        {
            final_report = Globals.ThisAddIn.Application.ActiveSheet;
            _Excel.Range last = final_report.Cells.SpecialCells(_Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
            lastUsedRow = last.Row;
            lastUsedCol = last.Column;
        }
    }
}
