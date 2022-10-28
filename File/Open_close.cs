using Excel = Microsoft.Office.Interop.Excel;
namespace File;
public class Open_close
{
    private static Excel.Application Start_Excel()
    {
        var ExcelObj = new Excel.Application();
        ExcelObj.Visible = true;
        ExcelObj.WindowState = Excel.XlWindowState.xlMaximized;

        return ExcelObj;
    }

    public static Excel.Workbook Open_file(Dictionary<string, string> config)
        {
            //FilePath Path = new FilePath(date);

            var date = new Date.Get_date(config["date_current_report"]);
            Start_Excel().Workbooks.Open(config["path_directory"] + date.yaer + "\\" + date.month + ". " + date.month_name[date.month] + "\\" + config["ExcelWorkBook_report"] + date.d_full + ".xlsm");

           /*

            var ExcelWorkSheet_rating = (Excel.Worksheet)ExcelWorkBook_report!.Sheets.Item["Источник Рейтинг"];
            ExcelWorkSheet_rating.Visible = Excel.XlSheetVisibility.xlSheetVisible;
            ExcelObj.Run("Выкладки");

            var ExcelWorkSheet_install = (Excel.Worksheet)ExcelWorkBook_report.Sheets["Установочные"];
            string? amount_line_old_start = (string)ExcelWorkSheet_install.Range["B2"].Text;
            string? amount_line_old_end = Convert.ToString(Int32.Parse(amount_line_old_start) - 13);
            string? amount_line_next_start = Convert.ToString(Convert.ToInt32(amount_line_old_start) + 1);
            string? amount_line_next_end = Convert.ToString(Convert.ToInt32(amount_line_old_start) + 14);
            
            ExcelWorkSheet_rating.Range["A" + amount_line_old_start + ":AC" + amount_line_old_end].Copy();
            ExcelWorkSheet_rating.Paste(ExcelWorkSheet_rating.Range["A" + amount_line_next_start]);
            DateTime date1 = new DateTime(Int32.Parse(date["year"]), Int32.Parse(date["month"]), Int32.Parse(date["day"]));
            ExcelWorkSheet_rating.Range["B" + amount_line_next_start + ":B" + amount_line_next_end].Value = date1;
    */
            return null!;
            
        }

    /*
        public static object FuncOpen2(Excel.Application ExcelObj, Dictionary<string, string> date, string d_briefly, Dictionary<string, Excel.Workbook> ExcelWorkBook_sources)
        {
            FilePath Path = new FilePath(date);
            //string o = Path.path_directory
            //string o2 = Path.file()

            //var source = new Dictionary<string, Excel.Workbook>();
            int index = 0;
            foreach (string node in Path.source_name_eng)
            {
                ExcelWorkBook_sources.Add(node, ExcelObj.Workbooks.Open(Path.path_directory + Path.path_source + Path.source_name_rus[index] + ".xlsx"));
                int last_sheet = ExcelWorkBook_sources[node].Worksheets.Count;

                int checker = 0;
                foreach (Excel.Worksheet sheet_existing in ExcelWorkBook_sources[node].Sheets)
                {
                    if (sheet_existing.Name == d_briefly)
                    {
                        checker = 1;
                    }
                }
                if (checker == 1)
                {
                    var ExcelWorkSheet_sheet1 = (Excel.Worksheet)ExcelWorkBook_sources[node].Worksheets.Add(System.Reflection.Missing.Value, ExcelWorkBook_sources[node].Worksheets[last_sheet]);
                    ExcelWorkSheet_sheet1.Name = "Лист1";
                }
                else
                {
                    var ExcelWorkSheet_brefly = (Excel.Worksheet)ExcelWorkBook_sources[node].Worksheets.Add(System.Reflection.Missing.Value, ExcelWorkBook_sources[node].Worksheets[last_sheet]);
                    ExcelWorkSheet_brefly.Name = d_briefly;
                }

                var ExcelWorkSheet_install = (Excel.Worksheet)ExcelWorkBook_sources[node].Worksheets["Установочные"];
                ExcelWorkSheet_install.Range["B1"].Value = d_briefly;

                index += 1;
            }
            return ExcelWorkBook_sources;
        }

        public static Excel.Workbook FuncOpen3(Excel.Application ExcelObj, Dictionary<string, string> date, string d_briefly, string d_full)
        {
            FilePath Path = new FilePath(date);

            var ExcelWorkBook_rating = ExcelObj.Workbooks.Open(Path.path_directory + "!!! Рейтинги_" + Path.path_month_name + " " + Path.path_yaer + ".xlsx");
            int last_sheet = ExcelWorkBook_rating.Worksheets.Count;
            var ExcelWorkSheet_new_list = (Excel.Worksheet)ExcelWorkBook_rating.Worksheets.Add(System.Reflection.Missing.Value, ExcelWorkBook_rating.Worksheets[last_sheet]);
            ExcelWorkSheet_new_list.Name = d_briefly;
            ExcelWorkSheet_new_list.Range["A1"].Value = "Рейтинг подразделений по 4 показателям за " + d_full;

            return ExcelWorkBook_rating;
        }



    public static void FuncClose(Excel.Application ExcelObj, Dictionary<string, string> date, string d_full, Excel.Workbook ExcelWorkBook_report, Dictionary<string, Excel.Workbook> ExcelWorkBook_sources, Excel.Workbook ExcelWorkBook_rating)
        {
            FilePath Path = new FilePath(date);

            var ExcelWorkSheet = (Excel.Worksheet)ExcelWorkBook_report.Sheets.Item["Источник Рейтинг"];
            ExcelWorkSheet.Activate();
            ExcelObj.Run("Выкладки");
            ExcelWorkSheet.Visible = Excel.XlSheetVisibility.xlSheetHidden;
            //Format-List -Property Name, Index -InputObject $ExcelWorkBook_report.Sheets.Item("Рейтинг")
            var ExcelWorkSheet_active = (Excel.Worksheet)ExcelWorkBook_report.Sheets[1];
            ExcelWorkSheet_active.Activate();
            ExcelObj.DisplayAlerts = false;
            ExcelWorkBook_report.SaveAs(Path.path_directory + "Отчет по простоям и сервису_" + d_full + ".xlsm", Microsoft.Office.Interop.Excel.XlFileFormat.xlOpenXMLWorkbookMacroEnabled);
            ExcelWorkBook_report.Close();

            //foreach (key in ExcelWorkBook_sources.Keys) {ExcelWorkBook_sources[key].Save() ExcelWorkBook_sources[key].Close()}

            foreach (string node in Path.source_name_eng)
            {
                ExcelWorkBook_sources[node].Save();
                ExcelWorkBook_sources[node].Close();
            };

            ExcelWorkBook_rating.Save();
            ExcelWorkBook_rating.Close();
        }

    */

}
