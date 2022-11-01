namespace Downtime_and_service;

class ClassProgram
{
    static void Main(string[] args)
    {
        var config = Config.ClassConfig.Get_Config();

        Excel.Workbook? excelWorkBook_report = null;
        Excel.Workbook? excelWorkBook_operators_CA = null;
        Excel.Workbook? excelWorkBook_operators_CC = null;
        Excel.Workbook? excelWorkBook_technician = null;
        Excel.Workbook? excelWorkBook_revenue = null;
        Excel.Workbook? excelWorkBook_amount = null;
        Excel.Workbook? excelWorkBook_not_connection = null;
        Excel.Workbook? excelWorkBook_not_work = null;
        Excel.Workbook? excelWorkBook_rating = null;

        File.ClassFile? report_config = null;
        File.ClassFile? operators_CA_config = null;

        Console.WriteLine("Дата текущего отчета: " + config["date_current_report"]);
        Console.WriteLine("Дата предыдущего отчета: " + config["date_previous_report"]);

        while (true) {
            Console.WriteLine("=====================================================================");
            Console.WriteLine("Выберите функцию:");
            Console.WriteLine("1. Запустить Excel, открыть все файлы отчетов, создать новые выкладки");
            Console.WriteLine("2. Копировать сведения");
            Console.WriteLine("3. Сохранить и закрыть все файлы отчетов");
            Console.WriteLine("4. Зактрыть Excel, завершить скрипт");

            string? v = Console.ReadLine();

            if (v == "1")
            {
                var excel = File.ClassFile.Start_Excel();
                
                report_config = new File.ClassFile("report", config);
                excelWorkBook_report = report_config.Open_file(excel);
                ClassWorkbook.Report_create(config, excel, report_config, excelWorkBook_report);

                operators_CA_config = new File.ClassFile("operators_CA", config);
                excelWorkBook_operators_CA = operators_CA_config.Open_file(excel);
                ClassWorkbook.Sources_create(config, operators_CA_config, excelWorkBook_operators_CA);

                var operators_CC_config = new File.ClassFile("operators_CC", config);
                excelWorkBook_operators_CC = operators_CC_config.Open_file(excel);
                ClassWorkbook.Sources_create(config, operators_CC_config, excelWorkBook_operators_CC);

                var technician_config = new File.ClassFile("technician", config);
                excelWorkBook_technician = technician_config.Open_file(excel);
                ClassWorkbook.Sources_create(config, technician_config, excelWorkBook_technician);

                var revenue_config = new File.ClassFile("revenue", config);
                excelWorkBook_revenue = revenue_config.Open_file(excel);
                ClassWorkbook.Sources_create(config, revenue_config, excelWorkBook_revenue);

                var amount_config = new File.ClassFile("amount", config);
                excelWorkBook_amount = amount_config.Open_file(excel);
                ClassWorkbook.Sources_create(config, amount_config, excelWorkBook_amount);

                var not_connection_config = new File.ClassFile("not_connection", config);
                excelWorkBook_not_connection = not_connection_config.Open_file(excel);
                ClassWorkbook.Sources_create(config, not_connection_config, excelWorkBook_not_connection);

                var not_work_config = new File.ClassFile("not_work", config);
                excelWorkBook_not_work = not_work_config.Open_file(excel);
                ClassWorkbook.Sources_create(config, not_work_config, excelWorkBook_not_work);

                var rating_config = new File.ClassFile("rating", config);
                excelWorkBook_rating = rating_config.Open_file(excel);
                ClassWorkbook.Rating_create(config, excel, rating_config, excelWorkBook_rating);
            }
            else if(v == "2")
            {
                //FileAction.FuncCopy(ExcelObj, date, d_briefly, ExcelWorkBook_report!, ExcelWorkBook_sources, //ExcelWorkBook_rating!);

                ClassWorkbook.Sources_copy(config, operators_CA_config!, excelWorkBook_operators_CA!, excelWorkBook_report!, 0);
            }
            else if(v == "3")
            {
                File.ClassFile.Close_file(excelWorkBook_report!);
                File.ClassFile.Close_file(excelWorkBook_operators_CA!);
                File.ClassFile.Close_file(excelWorkBook_operators_CC!);
                File.ClassFile.Close_file(excelWorkBook_technician!);
                File.ClassFile.Close_file(excelWorkBook_revenue!);
                File.ClassFile.Close_file(excelWorkBook_amount!);
                File.ClassFile.Close_file(excelWorkBook_not_connection!);
                File.ClassFile.Close_file(excelWorkBook_not_work!);
                File.ClassFile.Close_file(excelWorkBook_rating!);

                Console.WriteLine("-- Файлы закрыты");
            }
            else if(v == "4")
            {
                //System.Diagnostics.Process ExcelProcess = new System.Diagnostics.Process();

                var Ex = System.Diagnostics.Process.GetProcessesByName("EXCEL");
                Ex[0].Kill();
                break;
            }
        }
    }
}