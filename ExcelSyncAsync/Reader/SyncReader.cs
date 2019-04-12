namespace TriedExcel.Reader
{
    using System;
    using System.Collections.Generic;
    using Excel = Microsoft.Office.Interop.Excel;

    class SyncReader
    {
        public string FilePath { get; set; }
        public SyncReader(string filePath)
        {
            FilePath = filePath;
        }

        public void MainSync()
        {
            var excel = new Excel.Application { Visible = true, EnableAnimations = false };
            var wkb = ExcelFeatures.Open(excel, this.FilePath);

            List<Information> calculation = CalculateAllWorksheetsSync(wkb);
            Information.PrintInformation(calculation);

            excel.EnableAnimations = true;
            wkb.Close(true);
            excel.Quit();
            ExcelFeatures.CloseExcelExe(excel);
        }

        public Information CalculateSingleWorksheetSync(Excel.Worksheet wks)
        {
            int resultFromCalculation = 0;
            int resultTryParse;
            int lastRow = ExcelFeatures.LastRowPerColumn(1, wks);

            for (int i = 1; i < lastRow; i++)
            {
                if (Int32.TryParse(wks.Cells[i, 1].Text, out resultTryParse))
                {
                    resultFromCalculation += resultTryParse;
                }
            }

            Information infoToReturn = new Information(wks.Name, resultFromCalculation, lastRow);
            Console.WriteLine(infoToReturn.ToString());
            return infoToReturn;
        }

        public List<Information> CalculateAllWorksheetsSync(Excel.Workbook wkb)
        {
            List<Information> tasks = new List<Information>();
            foreach (Excel.Worksheet wks in wkb.Worksheets)
            {
                tasks.Add(CalculateSingleWorksheetSync(wks));
            }
            return tasks;
        }
    }
}
