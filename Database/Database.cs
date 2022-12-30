namespace Database;

public class ClassDatabase
{
    public static void Insert(string name_eng, string[] slist, string slist2, string date_current_report)
    {
        using (var connection = new SqliteConnection("Data Source=" + Directory.GetCurrentDirectory() + "\\database.db"))
        {
            connection.Open();
            
            using (var transaction = connection.BeginTransaction())
            {
                var command = connection.CreateCommand();
                
                if (name_eng == "operators_CA")
                {
                    command.CommandText =
                    @"
                        INSERT INTO Operators_CA
                        VALUES ($ID, $Date, $City, $Plan, $Fact, $Report_i, $Plan_execute)
                    ";
                    
                }
                else if (name_eng == "operators_CC")
                {
                    command.CommandText =
                    @"
                        INSERT INTO Operators_CC
                        VALUES ($ID, $Date, $City, $Plan, $Fact, $Report_i, $Plan_execute)
                    ";
                }
                else if (name_eng == "technician")
                {
                    command.CommandText =
                    @"
                        INSERT INTO Technician
                        VALUES ($ID, $Date, $City, $Plan, $Fact, $Report_i, $Plan_execute)
                    ";
                }
                else if (name_eng == "revenue")
                {
                    command.CommandText =
                    @"
                        INSERT INTO Revenue
                        VALUES ($ID, $Date, $City, $Percent_loss)
                    ";
                }
                else if (name_eng == "amount")
                {
                    command.CommandText =
                    @"
                        INSERT INTO Amount
                        VALUES ($ID, $Date, $City, $Amount_TA)
                    ";
                }
                else if (name_eng == "not_connection")
                {
                    command.CommandText =
                    @"
                        INSERT INTO Not_connection
                        VALUES ($ID, $Date, $City, $Amount_TA)
                    ";
                }
                else if (name_eng == "not_work")
                {
                    command.CommandText =
                    @"
                        INSERT INTO Not_work
                        VALUES ($ID, $Date, $City, $Amount_TA)
                    ";
                }

                if (name_eng == "operators_CA" || name_eng == "operators_CC" || name_eng == "technician")
                {
                    command.Parameters.AddWithValue("$ID", Convert.ToDateTime(slist[1]).Ticks / 1000000000 + slist[0]);
                    command.Parameters.AddWithValue("$Date", slist[1]);
                    command.Parameters.AddWithValue("$City", slist[0]);
                    command.Parameters.AddWithValue("$Plan", Convert.ToInt32(slist[2]));
                    command.Parameters.AddWithValue("$Fact", Convert.ToInt32(slist[3]));
                    command.Parameters.AddWithValue("$Report_i", Convert.ToInt32(slist[4]));
                    command.Parameters.AddWithValue("$Plan_execute", Convert.ToInt32(slist[7]));
                }
                else if (name_eng == "revenue")
                {
                    if (slist[0] != "Итого")
                    {
                        command.Parameters.AddWithValue("$ID", Convert.ToDateTime(slist[1]).Ticks / 1000000000 + slist[0]);
                        command.Parameters.AddWithValue("$Date", slist[1].Substring(0, 10));
                        command.Parameters.AddWithValue("$City", slist[0]);
                    }
                    else
                    {
                        command.Parameters.AddWithValue("$ID", Convert.ToDateTime(date_current_report).Ticks / 1000000000 + "ФС");
                        command.Parameters.AddWithValue("$Date", date_current_report);
                        command.Parameters.AddWithValue("$City", "ФС");
                    }
                    command.Parameters.AddWithValue("$Percent_loss", Convert.ToDouble(slist[2]));
                }
                else if (name_eng == "amount")
                {
                    command.Parameters.AddWithValue("$ID", Convert.ToDateTime(slist[1]).Ticks / 1000000000 + slist[0]);
                    command.Parameters.AddWithValue("$Date", slist[1]);
                    command.Parameters.AddWithValue("$City", slist[0]);
                    command.Parameters.AddWithValue("$Amount_TA", Convert.ToInt32(slist[2]));
                }
                else if (name_eng == "not_connection" || name_eng == "not_work")
                {
                    command.Parameters.AddWithValue("$ID", Convert.ToDateTime(date_current_report).Ticks / 1000000000 + slist[1]);
                    command.Parameters.AddWithValue("$Date", date_current_report);
                    if (slist.Length == 3)
                    {
                        command.Parameters.AddWithValue("$City", slist[1] + " " + slist[2]);
                    }
                    else
                    {
                        command.Parameters.AddWithValue("$City", slist[1]);
                    }
                    command.Parameters.AddWithValue("$Amount_TA", Convert.ToInt32(slist2));
                }

                //var parameter = command.CreateParameter();
                //parameter.ParameterName = "$value";
                //command.Parameters.Add(parameter);
                //parameter.Value = "4, 4";

                command.ExecuteNonQuery();

                transaction.Commit();
            }
        }
    }

    public static void Extract(string name_eng)
    {
        using (var connection = new SqliteConnection("Data Source=" + Directory.GetCurrentDirectory() + "\\database.db"))
        {
            connection.Open();

            var command = connection.CreateCommand();
            
            if (name_eng == "operators_CA")
            {
                command.CommandText =
                @"
                    SELECT *
                    FROM Operators_CA
                    WHERE Date = $Date
                    ORDER BY = City
                ";
                command.Parameters.AddWithValue("$Date", "15.12.2022");
                //command.Parameters.AddWithValue("$id2", 2);
                //command.Parameters.Add("$id2", SqliteType.Integer);
                //command.Parameters["$id2"].Value = 2;

                using (var reader = command.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        var id = reader.GetString(0);
                        var date = reader.GetString(1);
                        var city = reader.GetString(2);
                        var plan = reader.GetInt32(3);
                        var fact = reader.GetInt32(4);
                        var report_i = reader.GetInt32(5);
                        var plan_execute = reader.GetInt32(6);

                        //Console.WriteLine($"Hello, {id} {plan}!");


                    }
                }
            }
        }
    }
}
