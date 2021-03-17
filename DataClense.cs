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
        _Excel.Worksheet sheet;
        int lastUsedRow, lastUsedCol;


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

            //string[][] allNames = new string[lastUsedRow][];

            //Adding new Column
            oRange.EntireColumn.Insert();
            oRange.EntireColumn.Insert();
            oRange.EntireColumn.Insert();

            for (int i = 0; i < 3; i++)
            {
                var cell = (Range)sheet.Cells[1, oRange.Column - (i + 1)];
                switch (i)
                {
                    case 0:
                        cell.Value = "Club Title";
                        break;
                    case 1:
                        cell.Value = "Last Name";
                        break;
                    case 2:
                        cell.Value = "First Name";
                        break;
                }
            }

            for (int i = 1; i < lastUsedRow; i++)
            {
                string fullName = ReadCell(oRange.Row + i, oRange.Column, sheet);
                string[] names = fullName.Split(',');

                if (names[0] == "")
                    return;

                if (names[1].Contains('*'))
                {
                    string[] club = names[1].Split('*');
                    var cell = (Range)sheet.Cells[oRange.Row + i, oRange.Column - 2];
                    cell.Value = club[0];

                    cell = (Range)sheet.Cells[oRange.Row + i, oRange.Column -1];
                    cell.Value = club[1];

                    cell = (Range)sheet.Cells[oRange.Row + i, oRange.Column - 3];
                    cell.Value = club[1];
                }
                else
                {
                    var cell = (Range)sheet.Cells[oRange.Row + i, oRange.Column - 2];
                    cell.Value = names[0];

                    cell = (Range)sheet.Cells[oRange.Row + i, oRange.Column - 3];
                    cell.Value = names[1];
                }

            }
        }

        //Initializes final report sheet
        public void getSheetInfo()
        {
            sheet = Globals.ThisAddIn.Application.ActiveSheet;
            _Excel.Range last = sheet.Cells.SpecialCells(_Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
            lastUsedRow = last.Row;
            lastUsedCol = last.Column;
        }

        private string ReadCell(int row, int col, Worksheet sheet)
        {
            if (sheet.Cells[row, col].Value2 != null)
                return sheet.Cells[row, col].Value2 + "";
            return "";
        }
    }
}
