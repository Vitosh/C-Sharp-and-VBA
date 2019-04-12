namespace TriedExcel.Reader
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Threading.Tasks;
    using Excel = Microsoft.Office.Interop.Excel;
    class AsyncReader
    {
        public string FilePath { get; set; }

        public AsyncReader(string filePath)
        {
            FilePath = filePath;
        }

        public async Task MainAsync()
        {
            var excel = new Excel.Application { Visible = true, EnableAnimations = false };
            var wkb = ExcelFeatures.Open(excel, this.FilePath);

            var calculation = await CalculateAllWorksheetsAsync(wkb);
            Information.PrintInformation(calculation);

            excel.EnableAnimations = true;
            wkb.Close(true);
            excel.Quit();
            ExcelFeatures.CloseExcelExe(excel);
        }

        public async Task<List<Information>> CalculateAllWorksheetsAsync(Excel.Workbook wkb)
        {
            var tasks = wkb.Worksheets.Cast<Excel.Worksheet>().Select(CalculateSingleWorksheetAsync);
            var results = await Task.WhenAll(tasks);
            return results.ToList();
        }

        public async Task<Information> CalculateSingleWorksheetAsync(Excel.Worksheet wks)
        {
            int lastRow = ExcelFeatures.LastRowPerColumn(1, wks);
            int result = await Task.Run(() =>
            {
                int resultFromCalculation = 0;
                int resultTryParse;
                for (int i = 1; i < lastRow; i++)
                {
                    if (Int32.TryParse(wks.Cells[i, 1].Text, out resultTryParse))
                    {
                        resultFromCalculation += resultTryParse;
                    }
                }
                return resultFromCalculation;
            });

            Information infoToReturn = new Information(wks.Name, result, lastRow);
            Console.WriteLine(infoToReturn.ToString());
            return infoToReturn;
        }
    }
}
