namespace Downtime_and_service
{
    class FileAction
    {
        

        public static void FuncCopy(Excel.Application ExcelObj, Dictionary<string, string> date, string d_briefly, Excel.Workbook ExcelWorkBook_report, Dictionary<string, Excel.Workbook> ExcelWorkBook_sources, Excel.Workbook ExcelWorkBook_rating)
        {
            /*
            

            Excel.Worksheet sheet_rating = (Excel.Worksheet)ExcelWorkBook_report.Worksheets["Рейтинг"];
            sheet_rating.Activate();
            sheet_rating.Range["R2"].Value = date2;
            ExcelObj.Run("Сортировка_рейтинга");
            sheet_rating.Range["A22", "B34"].Copy();

            Excel.Worksheet sort_rating = (Excel.Worksheet)ExcelWorkBook_rating.Worksheets[d_briefly];
            var paste = Excel.XlPasteType.xlPasteValues;
            sort_rating.Range["A3"].PasteSpecial(paste);
            var sort = Excel.XlSortOrder.xlAscending;
            dynamic range2 = sort_rating.Range["A3", "B15"];
            range2.Sort(range2.Columns[1], sort);
            //range2.Sort(range2.Columns.Item[1], sort);

            */
        }

        
    }
}