namespace Database;

public class ClassDatabase
{
    public static void Open()
    {
        using (var connection = new SqliteConnection("Data Source=usersdata.db"))
        {
            connection.Open();

            var command = connection.CreateCommand();
            command.CommandText =
            @"
                SELECT Column1
                FROM NewTable
                WHERE Column2 = '2'
            ";
            //command.Parameters.AddWithValue("$id", id);

            using (var reader = command.ExecuteReader())
            {
                while (reader.Read())
                {
                    var name = reader.GetString(0);

                    Console.WriteLine($"Hello, {name}!");
                }
            }
        }
    }
}
