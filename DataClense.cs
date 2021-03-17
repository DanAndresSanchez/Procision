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
        _Excel.Worksheet sheet;


        public void centerAlign()
        {
            getSheetInfo();
            var rng = sheet.get_Range(sheet.UsedRange, Type.Missing);
            //rng.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
            rng.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
            rng.Cells.Columns.AutoFit();
            rng.Cells.Rows.AutoFit();
        }

        public void clearSheet()
        {
            getSheetInfo();
            var rng = sheet.get_Range("A2:Z300", Type.Missing);
            rng.Cells.Clear();
            Console.WriteLine("Sheet Cleared");
        }


        public void formatNames()
        {
            getSheetInfo();
            //capture user selection
            _Excel.Range oRange = Globals.ThisAddIn.Application.Selection;
            //Adding new Column
            oRange.EntireColumn.Insert();
        }

        //Initializes final report sheet
        public void getSheetInfo()
        {
            sheet = Globals.ThisAddIn.Application.ActiveSheet;
            _Excel.Range last = sheet.Cells.SpecialCells(_Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
            lastUsedRow = last.Row;
            lastUsedCol = last.Column;
        }
    }
}
