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

            //capture user selection
            _Excel.Range oRange = Globals.ThisAddIn.Application.Selection;

            var rng = sheet.get_Range("A2:Z" + lastUsedRow, Type.Missing);
            rng.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;
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

        public void formatAddress()
        {
            getSheetInfo();

            //capture user selection
            _Excel.Range oRange = Globals.ThisAddIn.Application.Selection;

            //Adding new Column
            oRange.EntireColumn.Insert();
            oRange.EntireColumn.Insert();
            oRange.EntireColumn.Insert();
            oRange.EntireColumn.Insert();
            oRange.EntireColumn.Insert();


            for (int i = 0; i < 5; i++)
            {
                var cell = (Range)sheet.Cells[1, oRange.Column - (i + 1)];
                switch (i)
                {
                    case 0:
                        cell.Value = "Zip";
                        break;
                    case 1:
                        cell.Value = "State";
                        break;
                    case 2:
                        cell.Value = "City";
                        break;
                    case 3:
                        cell.Value = "Address 2";
                        break;
                    case 4:
                        cell.Value = "Address 1";
                        break;
                }
            }

            for (int i = 1; i < lastUsedRow; i++)
            {
                string fullAddress = ReadCell(oRange.Row + i, oRange.Column, sheet);
                string[] parts = fullAddress.Split(',');

                if (parts.Length == 5)
                {
                    // Address 1
                    var cell = (Range)sheet.Cells[oRange.Row + i, oRange.Column - 5];
                    cell.Value = parts[0].Trim();

                    // Address 2
                    cell = (Range)sheet.Cells[oRange.Row + i, oRange.Column - 4];
                    cell.Value = parts[1].Trim();

                    // City
                    cell = (Range)sheet.Cells[oRange.Row + i, oRange.Column - 3];
                    cell.Value = parts[2].Trim();

                    // State
                    cell = (Range)sheet.Cells[oRange.Row + i, oRange.Column - 2];
                    cell.Value = parts[3].Trim();

                    // Zip
                    cell = (Range)sheet.Cells[oRange.Row + i, oRange.Column - 1];
                    cell.Value = parts[4].Trim();
                }
            }

            oRange.EntireColumn.Delete();
        }

        public void formatNames()
        {
            getSheetInfo();
            
            //capture user selection
            _Excel.Range oRange = Globals.ThisAddIn.Application.Selection;

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

                    // First name
                    var cell = (Range)sheet.Cells[oRange.Row + i, oRange.Column - 3];
                    cell.Value = club[1].Trim();

                    // Last name
                    cell = (Range)sheet.Cells[oRange.Row + i, oRange.Column - 2];
                    cell.Value = club[0].Trim();

                    // Club Title
                    cell = (Range)sheet.Cells[oRange.Row + i, oRange.Column -1];
                    cell.Value = club[1].Trim();

                }
                else
                {
                    // First name
                    var cell = (Range)sheet.Cells[oRange.Row + i, oRange.Column - 2];
                    cell.Value = names[0].Trim();


                    // Last name
                    cell = (Range)sheet.Cells[oRange.Row + i, oRange.Column - 3];
                    cell.Value = names[1].Trim();
                }

            }

            oRange.EntireColumn.Delete();
        }

        public void formatPhoneNumber()
        {
            getSheetInfo();

            //capture user selection
            _Excel.Range oRange = Globals.ThisAddIn.Application.Selection;

            //Adding new Column
            oRange.EntireColumn.Insert();
            oRange.EntireColumn.Insert();


            for (int i = 0; i < 2; i++)
            {
                var cell = (Range)sheet.Cells[1, oRange.Column - (i + 1)];
                switch (i)
                {
                    case 0:
                        cell.Value = "Work Number";
                        break;
                    case 1:
                        cell.Value = "Cell Number";
                        break;
                }
            }

            for (int i = 1; i < lastUsedRow; i++)
            {
                string fullNumber = ReadCell(oRange.Row + i, oRange.Column, sheet);
                var cell = (Range)sheet.UsedRange;

                int value;
                if(int.TryParse(fullNumber, out value))
                {
                    cell = (Range)sheet.Cells[oRange.Row + i, oRange.Column - 1];
                    cell.Value = fullNumber;
                    continue;
                }

                string[] number = fullNumber.Split(new char[] { ' ' }, 2);

                //if (number[0] == "")
                //    break;
                if (number.Length == 1)
                {
                    if(number[0] != "")
                    {
                        cell = (Range)sheet.Cells[oRange.Row + i, oRange.Column - 1];
                        cell.Value = number[0].Trim();
                        continue;
                    }
                    else continue;
                }
                    
                else if (number[0].Contains('W'))
                    cell = (Range)sheet.Cells[oRange.Row + i, oRange.Column - 1];
                else
                    cell = (Range)sheet.Cells[oRange.Row + i, oRange.Column - 2];
                
                cell.Value = number[1].Trim();
            }

            oRange.EntireColumn.Delete();
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
