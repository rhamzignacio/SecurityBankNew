using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace sbtc
{
    public static class DatabaseConnection
    {
        public static string ConnectionString = "datasource=localhost;port=3306;username=root;password=secret;";

        public static string DumpString = "mysqldump.exe --user=root --password=secret --host=localhost captive_database ";

        public static string ArchiveOutPut = "D:\\";
    }
}
