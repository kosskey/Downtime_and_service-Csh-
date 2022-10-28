namespace Downtime_and_service
{
    class FileAction
    {
        

        public static void FuncCopy(Excel.Application ExcelObj, Dictionary<string, string> date, string d_briefly, Excel.Workbook ExcelWorkBook_report, Dictionary<string, Excel.Workbook> ExcelWorkBook_sources, Excel.Workbook ExcelWorkBook_rating)
        {
            FilePath Path = new FilePath(date);
            DateTime date2 = new DateTime(Int32.Parse(date["year"]), Int32.Parse(date["month"]), Int32.Parse(date["day"]));

            for (int index = 0; index < Path.source_name_rus.Length; index++)
            {
                var value1 = (Excel.Worksheet)ExcelWorkBook_sources[Path.source_name_eng[index]].Sheets["Установочные"];
                string range = (string)value1.Range["B2"].Text;

                value1.Range[range].Copy();
                
                Excel.Worksheet sheet_installation = (Excel.Worksheet)ExcelWorkBook_report.Worksheets["Установочные"];
                //int amount_line_old = Convert.ToInt32(sheet_installation.Cells.Item[index + 3, 2].Text);
                var q = (Excel.Range)sheet_installation.Cells.Item[index + 3, 2];
                int amount_line_old = Convert.ToInt32(q.Text);
                //workbook.Activate();
                Excel.Worksheet worksheet = (Excel.Worksheet)ExcelWorkBook_report.Worksheets[Path.source_name_rus[index]];
                worksheet.Range["C" + (amount_line_old + 1)].PasteSpecial(Microsoft.Office.Interop.Excel.XlPasteType.xlPasteValues);

                //int amount_line_new = Convert.ToInt32(sheet_installation.Cells.Item[index + 3, 2].Text);
                var q2 = (Excel.Range)sheet_installation.Cells.Item[index + 3, 2];
                int amount_line_new = Convert.ToInt32(q2.Text);
                
                worksheet.Range["B" + (amount_line_old + 1), "B" + amount_line_new].Value = date2;
                worksheet.Range["A" + amount_line_old].Copy();
                worksheet.Range["A" + (amount_line_old + 1), "A" + amount_line_new].PasteSpecial();
            }

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
        }

        
    }
}