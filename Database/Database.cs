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
    
    public static void Insert()
    {
        using (var connection = new SqliteConnection("Data Source=D:\\IT\\Development\\GitHub\\kosskey\\Downtime_and_service-Csh-\\Downtime_and_service\\database2.db"))
        {
            connection.Open();
            
            using (var transaction = connection.BeginTransaction())
            {
                var command = connection.CreateCommand();
                command.CommandText =
                @"
                    INSERT INTO NewTable
                    VALUES ($value, $value2)
                ";

                var parameter = command.CreateParameter();
                command.Parameters.AddWithValue("$value", 4);
                command.Parameters.AddWithValue("$value2", 4);
                
                //parameter.ParameterName = "$value";
                //command.Parameters.Add(parameter);
                
                //parameter.Value = "4, 4";
                command.ExecuteNonQuery();

                transaction.Commit();
            }
        }
    }
}
