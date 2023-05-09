using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApp8.Utils
{
    public class AppSettings
    {
        public static DatabaseConfigOption DatabaseConfig { get; set; } = new();
    }

    public class DatabaseConfigOption
    {
        public string ConnectionString { get; set; } = string.Empty;
    }
}
