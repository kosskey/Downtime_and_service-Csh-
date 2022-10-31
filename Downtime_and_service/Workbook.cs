namespace Downtime_and_service;

class ClassWorkbook
{
    public static void Report_create(Dictionary<string, string> config, Excel.Application excel, File.ClassFile report_config, Excel.Workbook excelWorkBook_report)
    {
        var excelWorkSheet_rating = report_config.Activate_sheet(excelWorkBook_report, "Источник Рейтинг");
        excelWorkSheet_rating.Visible = Excel.XlSheetVisibility.xlSheetVisible;
        excel.Run("Выкладки");

        var excelWorkSheet_install = report_config.Activate_sheet(excelWorkBook_report, "Установочные");
        string? amount_line_old_start = (string)excelWorkSheet_install.Range["B2"].Text;
        string? amount_line_old_end = Convert.ToString(Int32.Parse(amount_line_old_start) - 13);
        string? amount_line_next_start = Convert.ToString(Convert.ToInt32(amount_line_old_start) + 1);
        string? amount_line_next_end = Convert.ToString(Convert.ToInt32(amount_line_old_start) + 14);
            
        excelWorkSheet_rating.Range["A" + amount_line_old_start + ":AC" + amount_line_old_end].Copy();
        excelWorkSheet_rating.Paste(excelWorkSheet_rating.Range["A" + amount_line_next_start]);
        DateTime date1 = new DateTime(Int32.Parse(config["year"]), Int32.Parse(config["month"]), Int32.Parse(config["day"]));
        excelWorkSheet_rating.Range["B" + amount_line_next_start + ":B" + amount_line_next_end].Value = date1;
    }

    public static void Report_copy()
    {

    }

    public void Operators_CA()
    {
        

        //excelWorkBook_operators_CA.Close();
    }

    public void Operators_CC()
    {


        //excelWorkBook_operators_CC.Close();
    }

    public void Technician()
    {


        //excelWorkBook_technician.Close();
    }

    public void Revenue()
    {


        //excelWorkBook_revenue.Close();
    }

    public void Amount()
    {


        //excelWorkBook_amount.Close();
    }

    public void Not_connection()
    {


        //excelWorkBook_not_connection.Close();
    }

    public void Not_work()
    {


        //excelWorkBook_not_work.Close();
    }

    public void Rating()
    {


        //excelWorkBook_rating.Close();
    }
}