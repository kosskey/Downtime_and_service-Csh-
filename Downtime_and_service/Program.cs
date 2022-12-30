namespace Downtime_and_service;

class ClassProgram
{
    static void Main(string[] args)
    {
        var config = Config.ClassConfig.Get_Config();
        var date = new Date.ClassDate(config["date_current_report"]);

        Excel.Application? excel = null;

        Excel.Workbook? excelWorkBook_report = null;

        Files.Path? report_path = null;
        Files.Path? operators_CA_path_csv = null;
        Files.Path? operators_CC_path_csv = null;
        Files.Path? technician_path_csv = null;
        Files.Path? revenue_path_csv = null;
        Files.Path? amount_path_csv = null;
        Files.Path? not_connection_path = null;
        Files.Path? not_work_path = null;

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
                var save_format = Excel.XlFileFormat.xlCSV;

                excel = Files.FileXLSX.Start_Excel();

                var operators_CA_path_xlsx = new Files.Path("operators_CA", config, date, "xlsx");
                operators_CA_path_csv = new Files.Path("operators_CA", config, date, "csv");
                var excelWorkBook_operators_CA = Files.FileXLSX.Open_file(excel, operators_CA_path_xlsx.path_file);
                excelWorkBook_operators_CA.SaveAs(operators_CA_path_csv.path_file, FileFormat:Excel.XlFileFormat.xlCSV, Local:true);
                excelWorkBook_operators_CA.Close(SaveChanges:false);
                //TypeFile.FileCSV.Convert_ANSI_UTF8("operators_CA", config, save_link);

                var operators_CC_path_xlsx = new Files.Path("operators_CC", config, date, "xlsx");
                operators_CC_path_csv = new Files.Path("operators_CC", config, date, "csv");
                var excelWorkBook_operators_CC = Files.FileXLSX.Open_file(excel, operators_CC_path_xlsx.path_file);
                excelWorkBook_operators_CC.SaveAs(operators_CC_path_csv.path_file, save_format, Local:true);
                excelWorkBook_operators_CC.Close(SaveChanges:false);

                var technician_path_xlsx = new Files.Path("technician", config, date, "xlsx");
                technician_path_csv = new Files.Path("technician", config, date, "csv");
                var excelWorkBook_technician = Files.FileXLSX.Open_file(excel, technician_path_xlsx.path_file);
                excelWorkBook_technician.SaveAs(technician_path_csv.path_file, save_format, Local:true);
                excelWorkBook_technician.Close(SaveChanges:false);

                var revenue_path_xlsx = new Files.Path("revenue", config, date, "xlsx");
                revenue_path_csv = new Files.Path("revenue", config, date, "csv");
                var excelWorkBook_revenue = Files.FileXLSX.Open_file(excel, revenue_path_xlsx.path_file);
                excelWorkBook_revenue.SaveAs(revenue_path_csv.path_file, save_format, Local:true);
                excelWorkBook_revenue.Close(SaveChanges:false);

                var amount_path_xlsx = new Files.Path("amount", config, date, "xlsx");
                amount_path_csv = new Files.Path("amount", config, date, "csv");
                var excelWorkBook_amount = Files.FileXLSX.Open_file(excel, amount_path_xlsx.path_file);
                excelWorkBook_amount.SaveAs(amount_path_csv.path_file, save_format, Local:true);
                excelWorkBook_amount.Close(SaveChanges:false);
            }
            else if(v == "2")
            {
                report_path = new Files.Path("report", config, date);
                excelWorkBook_report = Files.FileXLSX.Open_file(excel!, report_path.path_file);
                ClassWorkbook.Report_create(date, excel!, excelWorkBook_report);

                Console.WriteLine("-- Файл Отчета открыт");
            }
            else if(v == "3")
            {
                Files.FileCSV.Open_CSV("operators_CA", config, operators_CA_path_csv!.path_file);
                Files.FileCSV.Open_CSV("operators_CC", config, operators_CC_path_csv!.path_file);
                Files.FileCSV.Open_CSV("technician", config, technician_path_csv!.path_file);
                Files.FileCSV.Open_CSV("revenue", config, revenue_path_csv!.path_file);
                Files.FileCSV.Open_CSV("amount", config, amount_path_csv!.path_file);

                not_connection_path = new Files.Path("not_connection", config, date);
                Files.FileTXT.Open_TXT("not_connection", config, not_connection_path!.path_file);
                not_work_path = new Files.Path("not_work", config, date);
                Files.FileTXT.Open_TXT("not_work", config, not_work_path!.path_file);
            }
            else if(v == "4")
            {
                Database.ClassDatabase.Extract("operators_CA");

            }
            else if (v == "5")
            {

            }
            else if (v == "6")
            {

            }
            else if (v == "7")
            {

            }
            else
            {
                /*
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




                //System.Diagnostics.Process ExcelProcess = new System.Diagnostics.Process();

                var Ex = System.Diagnostics.Process.GetProcessesByName("EXCEL");
                if (Ex.Count() != 0)
                {
                    Ex[0].Kill();
                }
                
                break;
                */
            }
        }
    }
}