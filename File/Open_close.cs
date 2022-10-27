//Read a Text File
using System;
using System.IO;
namespace File
{
    public class Open_close
    {
        public static void Op()
        {
            StreamReader sr = new StreamReader(Directory.GetCurrentDirectory() + "\\Config.txt");
            string? line = sr.ReadLine();

            while (line != null)
            {
                Console.WriteLine(line);
                line = sr.ReadLine();
            }
            sr.Close();
            Console.ReadLine();
        }
    }
}