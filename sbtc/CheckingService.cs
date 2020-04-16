using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using MySql.Data.MySqlClient;
namespace sbtc
{
    public static class CheckingService
    {
        public static bool CheckBatchIfDuplicate(string _batch)
        {
            MySqlConnection conn = new MySqlConnection(DatabaseConnection.ConnectionString);

            conn.Open();

            string query = "SELECT BRSTN FROM captive_database.sbtc_history WHERE FinalBatch='" + _batch + "' LIMIT 1";

            MySqlCommand cmd = new MySqlCommand(query, conn);

            MySqlDataReader reader = cmd.ExecuteReader();

            string brstn = "";

            while (reader.Read())
            {
                brstn = !reader.IsDBNull(0) ? reader.GetString(0) : "";

                if (brstn != "")
                    return true;
                else
                    return false;
            }

            return false;
        }
    }
}
