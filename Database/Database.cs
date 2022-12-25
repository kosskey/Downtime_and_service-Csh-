namespace Database;

public class ClassDatabase
{
    public static void Open()
    {
        using (var connection = new SqliteConnection("Data Source=D:\\IT\\Development\\GitHub\\kosskey\\Downtime_and_service-Csh-\\Downtime_and_service\\database2.db"))
        {
            connection.Open();

            var command = connection.CreateCommand();
            command.CommandText =
            @"
                SELECT *
                FROM NewTable
                WHERE Column2 = $id OR Column2 = $id2
            ";
            command.Parameters.AddWithValue("$id", 1);
            //command.Parameters.AddWithValue("$id2", 2);
            command.Parameters.Add("$id2", SqliteType.Integer);
            command.Parameters["$id2"].Value = 2;

            using (var reader = command.ExecuteReader())
            {
                while (reader.Read())
                {
                    var name = reader.GetInt32(0);
                    var name2 = reader.GetInt32(1);

                    Console.WriteLine($"Hello, {name} {name2}!");
                }
            }
        }
    }
    
    public static void Insert(string name_eng, string[] slist)
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

                command.Parameters.AddWithValue("$ID", Convert.ToDateTime(slist[1]).Ticks / 1000000000 + slist[0]);
                command.Parameters.AddWithValue("$Date", slist[1]);
                command.Parameters.AddWithValue("$City", slist[0]);
                command.Parameters.AddWithValue("$Plan", Convert.ToInt32(slist[2]));
                command.Parameters.AddWithValue("$Fact", Convert.ToInt32(slist[3]));
                command.Parameters.AddWithValue("$Report_i", Convert.ToInt32(slist[4]));
                command.Parameters.AddWithValue("$Plan_execute", Convert.ToInt32(slist[7]));

                //var parameter = command.CreateParameter();
                //parameter.ParameterName = "$value";
                //command.Parameters.Add(parameter);
                //parameter.Value = "4, 4";

                command.ExecuteNonQuery();

                transaction.Commit();
            }
        }
    }
}
