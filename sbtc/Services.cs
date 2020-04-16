using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using MySql.Data.MySqlClient;

namespace sbtc
{
    public class Services
    {
        public MySqlConnection myConnect;

        public void DBConnect()
        {
            string DBConnection = DatabaseConnection.ConnectionString;

            myConnect = new MySqlConnection(DBConnection);

            myConnect.Open();
        }

        public void DBClosed()
        {
            myConnect.Close();
        }

        public void UpdateSerial(List<BranchesModel> _branches)
        {
            try
            {
                DBConnect();

                _branches.ForEach(b =>
                {                  
                    MySqlCommand cmd = new MySqlCommand("UPDATE captive_database.sbtc_branches SET LastNo_PA =" + b.LastNo_PA +
                        " LastNo_CA =" + b.LastNo_CA + " LastNo_MC =" + b.LastNo_MC + " LastNo_Power_PA =" + b.LastNo_Power_PA + " LastNo_Power_CA =" + b.LastNo_Power_CA
                        + " LastNo_GC =" + b.LastNo_GC + " LastNo_CheckOne_PA =" + b.LastNo_CheckOne_PA + " LastNo_CheckOne_CA =" + b.LastNo_CheckOne_CA + " " +
                        "ModifiedDate ='" + DateTime.Now.ToString("yyyyMMddHHmmss") + "' WHERE BRSTN ='" + b.BRSTN + "'", myConnect);

                    cmd.ExecuteNonQuery();
                });//END OF FOREACH

                DBClosed();
            }//END OF TRY
            catch
            {

            }
        }

        public List<BranchesModel> GetAllBranch()
        {
            try
            {
                DBConnect();

                List<BranchesModel> Branches = new List<BranchesModel>();

                MySqlCommand cmd = new MySqlCommand("SELECT * FROM captive_database.sbtc_branches", myConnect);

                MySqlDataReader reader = cmd.ExecuteReader();

                while(reader.Read())
                {
                    BranchesModel branch = new BranchesModel();

                    branch.BRSTN = reader.GetString(0);

                    branch.Address1 = !reader.IsDBNull(1) ? reader.GetString(1) : "";

                    branch.Address2 = !reader.IsDBNull(2) ? reader.GetString(2) : "";

                    branch.Address3 = !reader.IsDBNull(3) ? reader.GetString(3) : "";

                    branch.Address4 = !reader.IsDBNull(4) ? reader.GetString(4) : "";
                    
                    branch.Address5 = !reader.IsDBNull(5) ? reader.GetString(5) : "";
                    
                    branch.Address6 = !reader.IsDBNull(6) ? reader.GetString(6) : "";

                    branch.LastNo_PA = !reader.IsDBNull(7) ? reader.GetInt64(7) : 0;

                    branch.LastNo_CA = !reader.IsDBNull(8) ? reader.GetInt64(8) : 0;

                    branch.LastNo_MC = !reader.IsDBNull(9) ? reader.GetInt64(9) : 0;

                    branch.LastNo_Power_PA = !reader.IsDBNull(10) ? reader.GetInt64(10) : 0;

                    branch.LastNo_Power_CA = !reader.IsDBNull(11) ? reader.GetInt64(11) : 0;

                    branch.LastNo_GC = !reader.IsDBNull(12) ? reader.GetInt64(12) : 0;

                    branch.LastNo_CheckOne_CA = !reader.IsDBNull(13) ? reader.GetInt64(13) : 0;

                    branch.LastNo_CheckOne_PA = !reader.IsDBNull(14) ? reader.GetInt64(14) : 0;

                    Branches.Add(branch);
                }//END OF WHILE

                DBClosed();

                return Branches;
            }//END OF TRY
            catch(Exception error)
            {
                return null;
            }
        }
    }
}
