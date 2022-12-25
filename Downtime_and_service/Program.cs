namespace Downtime_and_service;

class ClassProgram
{
    static void Main(string[] args)
    {
        var config = Config.ClassConfig.Get_Config();
        var date = new Date.ClassDate(config["date_current_report"]);

        Excel.Application? excel = null;

        Excel.Workbook? excelWorkBook_report = null;
        Excel.Workbook? excelWorkBook_operators_CA = null;
        Excel.Workbook? excelWorkBook_operators_CC = null;
        Excel.Workbook? excelWorkBook_technician = null;
        Excel.Workbook? excelWorkBook_revenue = null;
        Excel.Workbook? excelWorkBook_amount = null;
        Excel.Workbook? excelWorkBook_not_connection = null;
        Excel.Workbook? excelWorkBook_not_work = null;
        Excel.Workbook? excelWorkBook_rating = null;

        TypeFile.FileXLSX? report_config = null;
        TypeFile.FileXLSX? operators_CA_config = null;
        TypeFile.FileXLSX? operators_CC_config = null;
        TypeFile.FileXLSX? technician_config = null;
        TypeFile.FileXLSX? revenue_config = null;
        TypeFile.FileXLSX? amount_config = null;
        TypeFile.FileXLSX? not_connection_config = null;
        TypeFile.FileXLSX? not_work_config = null;

        Console.WriteLine("Дата текущего отчета: " + config["date_current_report"]);
        Console.WriteLine("Дата предыдущего отчета: " + config["date_previous_report"]);

        while (true) {
            Console.WriteLine("=====================================================================");
            Console.WriteLine("Выберите функцию:");
            Console.WriteLine("1. Запустить Excel, конвертировать из исходников XLSX в CSV");
            Console.WriteLine("2. Открыть файл отчета, создать новые выкладки");
            Console.WriteLine("3. Копировать сведения из CSV в базу Базу данных");
            Console.WriteLine("4. Копировать сведения из Базы данных в файл отчета");
            Console.WriteLine("5. Копировать рейтинг из файла отчета в Базу данных и в файл рейтинга в TXT");
            Console.WriteLine("6. Сохранить и закрыть файл отчета");
            Console.WriteLine("7. Зактрыть Excel, удалить файлы CSV, завершить скрипт");

            string? v = Console.ReadLine();

            if (v == "1")
            {
                //Database.ClassDatabase.Open();
                //Database.ClassDatabase.Insert();

                TypeFile.FileCSV.Open_CSV("operators_CA", config);
                TypeFile.FileCSV.Open_CSV("operators_CC", config);
                TypeFile.FileCSV.Open_CSV("technician", config);

                excel = TypeFile.FileXLSX.Start_Excel();

                operators_CA_config = new TypeFile.FileXLSX("operators_CA", config, date);
                excelWorkBook_operators_CA = operators_CA_config.Open_file(excel);
                var save_format = Excel.XlFileFormat.xlCSV;
                excelWorkBook_operators_CA.SaveAs("q1", save_format);
                

            }
            else if(v == "2")
            {
                ClassWorkbook.Sources_copy(date, operators_CA_config!, excelWorkBook_operators_CA!, excelWorkBook_report!, 3);
                ClassWorkbook.Sources_copy(date, operators_CC_config!, excelWorkBook_operators_CC!, excelWorkBook_report!, 4);
                ClassWorkbook.Sources_copy(date, technician_config!, excelWorkBook_technician!, excelWorkBook_report!, 5);
                ClassWorkbook.Sources_copy(date, revenue_config!, excelWorkBook_revenue!, excelWorkBook_report!, 6);
                ClassWorkbook.Sources_copy(date, amount_config!, excelWorkBook_amount!, excelWorkBook_report!, 7);
                ClassWorkbook.Sources_copy(date, not_connection_config!, excelWorkBook_not_connection!, excelWorkBook_report!, 8);
                ClassWorkbook.Sources_copy(date, not_work_config!, excelWorkBook_not_work!, excelWorkBook_report!, 9);

                ClassWorkbook.Report_copy(date, excel!, report_config!, excelWorkBook_report!, excelWorkBook_rating!);


                if (excel != null & report_config != null & excelWorkBook_report != null & excelWorkBook_rating != null)
                {
                    Console.WriteLine("-- Информация скопирована");
                }
                else
                {
                    Console.WriteLine("-- Ошибка копирования, файлы отчета не открыты");
                }
            }
            else if(v == "3")
            {
                ClassWorkbook.Report_save(excel!, report_config!, excelWorkBook_report!);
                TypeFile.FileXLSX.Close_file(excelWorkBook_report!);

                ClassWorkbook.Source_and_rating_save(excelWorkBook_operators_CA!);
                TypeFile.FileXLSX.Close_file(excelWorkBook_operators_CA!);

                ClassWorkbook.Source_and_rating_save(excelWorkBook_operators_CC!);
                TypeFile.FileXLSX.Close_file(excelWorkBook_operators_CC!);

                ClassWorkbook.Source_and_rating_save(excelWorkBook_technician!);
                TypeFile.FileXLSX.Close_file(excelWorkBook_technician!);

                ClassWorkbook.Source_and_rating_save(excelWorkBook_revenue!);
                TypeFile.FileXLSX.Close_file(excelWorkBook_revenue!);

                ClassWorkbook.Source_and_rating_save(excelWorkBook_amount!);
                TypeFile.FileXLSX.Close_file(excelWorkBook_amount!);

                ClassWorkbook.Source_and_rating_save(excelWorkBook_not_connection!);
                TypeFile.FileXLSX.Close_file(excelWorkBook_not_connection!);

                ClassWorkbook.Source_and_rating_save(excelWorkBook_not_work!);
                TypeFile.FileXLSX.Close_file(excelWorkBook_not_work!);

                ClassWorkbook.Source_and_rating_save(excelWorkBook_rating!);
                TypeFile.FileXLSX.Close_file(excelWorkBook_rating!);


                if (excel != null & report_config != null & excelWorkBook_report != null)
                {
                    Console.WriteLine("-- Файлы закрыты");
                }
                else
                {
                    Console.WriteLine("-- Ошибка сохранения, файлы отчета не открыты");
                }
            }
            else if(v == "4")
            {
                //System.Diagnostics.Process ExcelProcess = new System.Diagnostics.Process();

                var Ex = System.Diagnostics.Process.GetProcessesByName("EXCEL");
                if (Ex.Count() != 0)
                {
                    Ex[0].Kill();
                }
                
                break;
            }
            else if (v == "5")
            {

            }
            else if (v == "6")
            {

            }
            else if (v == "7")
            {


                report_config = new TypeFile.FileXLSX("report", config, date);
                excelWorkBook_report = report_config.Open_file(excel!);
                ClassWorkbook.Report_create(date, excel!, report_config, excelWorkBook_report);

                operators_CA_config = new TypeFile.FileXLSX("operators_CA", config, date);
                excelWorkBook_operators_CA = operators_CA_config.Open_file(excel!);
                ClassWorkbook.Sources_create(date, operators_CA_config, excelWorkBook_operators_CA);

                operators_CC_config = new TypeFile.FileXLSX("operators_CC", config, date);
                excelWorkBook_operators_CC = operators_CC_config.Open_file(excel!);
                ClassWorkbook.Sources_create(date, operators_CC_config, excelWorkBook_operators_CC);

                technician_config = new TypeFile.FileXLSX("technician", config, date);
                excelWorkBook_technician = technician_config.Open_file(excel!);
                ClassWorkbook.Sources_create(date, technician_config, excelWorkBook_technician);

                revenue_config = new TypeFile.FileXLSX("revenue", config, date);
                excelWorkBook_revenue = revenue_config.Open_file(excel!);
                ClassWorkbook.Sources_create(date, revenue_config, excelWorkBook_revenue);

                amount_config = new TypeFile.FileXLSX("amount", config, date);
                excelWorkBook_amount = amount_config.Open_file(excel!);
                ClassWorkbook.Sources_create(date, amount_config, excelWorkBook_amount);

                not_connection_config = new TypeFile.FileXLSX("not_connection", config, date);
                excelWorkBook_not_connection = not_connection_config.Open_file(excel!);
                ClassWorkbook.Sources_create(date, not_connection_config, excelWorkBook_not_connection);

                not_work_config = new TypeFile.FileXLSX("not_work", config, date);
                excelWorkBook_not_work = not_work_config.Open_file(excel!);
                ClassWorkbook.Sources_create(date, not_work_config, excelWorkBook_not_work);

                var rating_config = new TypeFile.FileXLSX("rating", config, date);
                excelWorkBook_rating = rating_config.Open_file(excel!);
                ClassWorkbook.Rating_create(date, excel!, rating_config, excelWorkBook_rating);

                Console.WriteLine("-- Файлы открыты");



            }
        }
    }
}