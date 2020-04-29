using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace sbtc
{
    public static class DatabaseConnection
    {
        public static string ConnectionString = "datasource=192.168.0.254;port=3306;username=root;password=CorpCaptive;";

        public static string DumpString = "mysqldump.exe --user=root --password=CorpCaptive --host=192.168.0.254 captive_database ";

        public static string ArchiveOutPut = "\\\\192.168.0.254\\captive\\Zips\\sbtc\\" + DateTime.Now.Year.ToString();

        public static string LiveStartPath = "\\\\192.168.0.254\\captive\\Auto\\SBTC\\SBTC_2.0\\Output";
    }
}
