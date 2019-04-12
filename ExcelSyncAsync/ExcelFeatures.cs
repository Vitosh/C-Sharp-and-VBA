namespace TriedExcel
{
    using System;
    using System.Runtime.InteropServices;
    using Excel = Microsoft.Office.Interop.Excel;

    public static class ExcelFeatures
    {
        public static int LastRowPerColumn(int column, Excel.Worksheet wks)
        {
            int lastRow = LastRowTotal(wks);
            while (((wks.Cells[lastRow, column]).Text == "") && (lastRow != 1))
            {
                lastRow--;
            }
            return lastRow;
        }

        public static void CloseExcelExe(Excel.Application excel)
        {
            Marshal.ReleaseComObject(excel);
        }

        public static int LastRowTotal(Excel.Worksheet wks)
        {
            Excel.Range lastCell = wks.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
            return lastCell.Row;
        }

        public static Excel.Workbook Open(Excel.Application excelInstance,
                        string fileName, bool readOnly = false,
                        bool editable = true, bool updateLinks = true)
        {
            return excelInstance.Workbooks.Open(fileName, updateLinks, readOnly, Editable: editable);
        }
    }
}
