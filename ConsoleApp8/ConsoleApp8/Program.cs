using ConsoleApp8;
using ConsoleApp8.Utils;
using Microsoft.Extensions.Configuration;
using OfficeOpenXml;
using System.ComponentModel;

namespace ConsoleApp8
{
    public class Program
    {
        
        static void Main(string[] args)
        {
            
            new Excel();
            Console.ReadLine();
            
        }
        
        //private static void Config()
        //{
        //    var configuration = new ConfigurationBuilder().AddJsonFile($"appsettings.json");
        //    var config = configuration.Build();
        //    AppSettings.DatabaseConfig = config.GetConnectionString(Name);
        //}

    }
}
