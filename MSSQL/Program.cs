using System;

namespace MSSQL
{
    class Program
    {
        static void Main(string[] args)
        {
            AutoCreateModels.Start();
            Console.WriteLine("MSSQL");
            Console.ReadLine();
        }
    }
}
