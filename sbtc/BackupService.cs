using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using MySql.Data.MySqlClient;

namespace sbtc
{
    public static class BackupService
    {
        public static void SaveHistory(OrderSorted _orders, string _batch, string _ext, DateTime _deliveryDate)
        {
            MySqlConnection myConn = new MySqlConnection(DatabaseConnection.ConnectionString);

            myConn.Open();

            MySqlCommand cmd;

            string query = "INSERT INTO captive_database.sbtc_history (Date, Time, DeliveryDate, ChkType, ChequeName, BRSTN, " +
                "AccountNo, Name1, Name2, StartingSerial, ENdingSerial, Batch, FinalBatch, Address1, " +
                "Address2, Address3, Address4, Address5, Address6) ";

            #region Regular Personal
            if (_orders.RegularPersonal.Count > 0)
            {
                _orders.RegularPersonal.ForEach(order =>
                {
                    string start = order.StartingSerial.ToString(), end = order.EndingSerial.ToString();

                    while (start.Length < 7)
                        start = "0" + start;

                    while (end.Length < 7)
                        end = "0" + end;

                    string values = "VALUES('" + DateTime.Now.ToString("yyyy-MM-dd") + "','" + DateTime.Now.ToString("H:mm") + "','" +
                        _deliveryDate.ToString("yyyy-MM-dd") + "','" + order.CheckType + "','Regular Personal','" +
                        order.BRSTN + "','" + order.AccountNo + "','" + order.Name.Replace("'", "''") + "','" + order.Name2.Replace("'", "''") + "','" +
                        start + "','" + end + "','" + order.FileName + "','" + _batch + "','" + order.Address1.Replace("'", "''") + "','" + order.Address2.Replace("'", "''") + "','" +
                        order.Address3.Replace("'", "''") + "','" + order.Address4.Replace("'", "''") + "','" + order.Address5.Replace("'", "''") + "','" + order.Address6.Replace("'", "''") + "');";

                    cmd = new MySqlCommand(query + values, myConn);

                    cmd.ExecuteNonQuery();
                });
            }//END IF
            #endregion

            #region Regular Commercial
            if (_orders.RegularCommercial.Count > 0)
            {
                _orders.RegularCommercial.ForEach(order =>
                {
                    string start = order.StartingSerial.ToString(), end = order.EndingSerial.ToString();

                    while (start.Length < 10)
                        start = "0" + start;

                    while (end.Length < 10)
                        end = "0" + end;

                    string values = "VALUES('" + DateTime.Now.ToString("yyyy-MM-dd") + "','" + DateTime.Now.ToString("H:mm") + "','" +
                        _deliveryDate.ToString("yyyy-MM-dd") + "','" + order.CheckType + "','Regular Commercial','" +
                        order.BRSTN + "','" + order.AccountNo + "','" + order.Name.Replace("'", "''") + "','" + order.Name2.Replace("'", "''") + "','" +
                        start + "','" + end + "','" + order.FileName + "','" + _batch + "','" + order.Address1.Replace("'", "''") + "','" + order.Address2.Replace("'", "''") + "','" +
                        order.Address3.Replace("'", "''") + "','" + order.Address4.Replace("'", "''") + "','" + order.Address5.Replace("'", "''") + "','" + order.Address6.Replace("'", "''") + "');";

                    cmd = new MySqlCommand(query + values, myConn);

                    cmd.ExecuteNonQuery();
                });
            }//END IF
            #endregion

            #region Regular Personal Pre-Encoded
            if (_orders.PersonalPreEncoded.Count > 0)
            {
                _orders.PersonalPreEncoded.ForEach(order =>
                {
                    string start = order.StartingSerial.ToString(), end = order.EndingSerial.ToString();

                    while (start.Length < 7)
                        start = "0" + start;

                    while (end.Length < 7)
                        end = "0" + end;

                    string values = "VALUES('" + DateTime.Now.ToString("yyyy-MM-dd") + "','" + DateTime.Now.ToString("H:mm") + "','" +
                        _deliveryDate.ToString("yyyy-MM-dd") + "','" + order.CheckType + "','Personal Pre-Encoded','" +
                        order.BRSTN + "','" + order.AccountNo + "','" + order.Name.Replace("'", "''") + "','" + order.Name2.Replace("'", "''") + "','" +
                        start + "','" + end + "','" + order.FileName + "','" + _batch + "','" + order.Address1.Replace("'", "''") + "','" + order.Address2.Replace("'", "''") + "','" +
                        order.Address3.Replace("'", "''") + "','" + order.Address4.Replace("'", "''") + "','" + order.Address5.Replace("'", "''") + "','" + order.Address6.Replace("'", "''") + "');";

                    cmd = new MySqlCommand(query + values, myConn);

                    cmd.ExecuteNonQuery();
                });
            }//END IF
            #endregion

            #region Regular Commercial Pre-Encoded
            if (_orders.CommercialPreEncoded.Count > 0)
            {
                _orders.CommercialPreEncoded.ForEach(order =>
                {
                    string start = order.StartingSerial.ToString(), end = order.EndingSerial.ToString();

                    while (start.Length < 10)
                        start = "0" + start;

                    while (end.Length < 10)
                        end = "0" + end;

                    string values = "VALUES('" + DateTime.Now.ToString("yyyy-MM-dd") + "','" + DateTime.Now.ToString("H:mm") + "','" +
                        _deliveryDate.ToString("yyyy-MM-dd") + "','" + order.CheckType + "','Commercial Pre-Encoded','" +
                        order.BRSTN + "','" + order.AccountNo + "','" + order.Name.Replace("'", "''") + "','" + order.Name2.Replace("'", "''") + "','" +
                        start + "','" + end + "','" + order.FileName + "','" + _batch + "','" + order.Address1.Replace("'", "''") + "','" + order.Address2.Replace("'", "''") + "','" +
                        order.Address3.Replace("'", "''") + "','" + order.Address4.Replace("'", "''") + "','" + order.Address5.Replace("'", "''") + "','" + order.Address6.Replace("'", "''") + "');";

                    cmd = new MySqlCommand(query + values, myConn);

                    cmd.ExecuteNonQuery();
                });
            }//END IF
            #endregion

            #region CheckOne Personal
            if (_orders.CheckOnePersonal.Count > 0)
            {
                _orders.CheckOnePersonal.ForEach(order =>
                {
                    string start = order.StartingSerial.ToString(), end = order.EndingSerial.ToString();

                    while (start.Length < 7)
                        start = "0" + start;

                    while (end.Length < 7)
                        end = "0" + end;

                    string values = "VALUES('" + DateTime.Now.ToString("yyyy-MM-dd") + "','" + DateTime.Now.ToString("H:mm") + "','" +
                        _deliveryDate.ToString("yyyy-MM-dd") + "','" + order.CheckType + "','CheckOne Personal','" +
                        order.BRSTN + "','" + order.AccountNo + "','" + order.Name.Replace("'", "''") + "','" + order.Name2.Replace("'", "''") + "','" +
                        start + "','" + end + "','" + order.FileName + "','" + _batch + "','" + order.Address1.Replace("'", "''") + "','" + order.Address2.Replace("'", "''") + "','" +
                        order.Address3.Replace("'", "''") + "','" + order.Address4.Replace("'", "''") + "','" + order.Address5.Replace("'", "''") + "','" + order.Address6.Replace("'", "''") + "');";

                    cmd = new MySqlCommand(query + values, myConn);

                    cmd.ExecuteNonQuery();
                });
            }//END IF
            #endregion

            #region CheckOne Commercial
            if (_orders.CheckOneCommerical.Count > 0)
            {
                _orders.CheckOneCommerical.ForEach(order =>
                {
                    string start = order.StartingSerial.ToString(), end = order.EndingSerial.ToString();

                    while (start.Length < 10)
                        start = "0" + start;

                    while (end.Length < 10)
                        end = "0" + end;

                    string values = "VALUES('" + DateTime.Now.ToString("yyyy-MM-dd") + "','" + DateTime.Now.ToString("H:mm") + "','" +
                        _deliveryDate.ToString("yyyy-MM-dd") + "','" + order.CheckType + "','CheckOne Commercial','" +
                        order.BRSTN + "','" + order.AccountNo + "','" + order.Name.Replace("'", "''") + "','" + order.Name2.Replace("'", "''") + "','" +
                        start + "','" + end + "','" + order.FileName + "','" + _batch + "','" + order.Address1.Replace("'", "''") + "','" + order.Address2.Replace("'", "''") + "','" +
                        order.Address3.Replace("'", "''") + "','" + order.Address4.Replace("'", "''") + "','" + order.Address5.Replace("'", "''") + "','" + order.Address6.Replace("'", "''") + "');";

                    cmd = new MySqlCommand(query + values, myConn);

                    cmd.ExecuteNonQuery();
                });
            }//END IF
            #endregion

            #region CheckPower Personal
            if (_orders.CheckPowerPersonal.Count > 0)
            {
                _orders.CheckPowerPersonal.ForEach(order =>
                {
                    string start = order.StartingSerial.ToString(), end = order.EndingSerial.ToString();

                    while (start.Length < 7)
                        start = "0" + start;

                    while (end.Length < 7)
                        end = "0" + end;

                    string values = "VALUES('" + DateTime.Now.ToString("yyyy-MM-dd") + "','" + DateTime.Now.ToString("H:mm") + "','" +
                        _deliveryDate.ToString("yyyy-MM-dd") + "','" + order.CheckType + "','CheckPower Personal','" +
                        order.BRSTN + "','" + order.AccountNo + "','" + order.Name.Replace("'", "''") + "','" + order.Name2.Replace("'", "''") + "','" +
                        start + "','" + end + "','" + order.FileName + "','" + _batch + "','" + order.Address1.Replace("'", "''") + "','" + order.Address2.Replace("'", "''") + "','" +
                        order.Address3.Replace("'", "''") + "','" + order.Address4.Replace("'", "''") + "','" + order.Address5.Replace("'", "''") + "','" + order.Address6.Replace("'", "''") + "');";

                    cmd = new MySqlCommand(query + values, myConn);

                    cmd.ExecuteNonQuery();
                });
            }//END IF
            #endregion

            #region CheckPower Commercial
            if (_orders.CheckPowerCommercial.Count > 0)
            {
                _orders.CheckPowerCommercial.ForEach(order =>
                {
                    string start = order.StartingSerial.ToString(), end = order.EndingSerial.ToString();

                    while (start.Length < 10)
                        start = "0" + start;

                    while (end.Length < 10)
                        end = "0" + end;

                    string values = "VALUES('" + DateTime.Now.ToString("yyyy-MM-dd") + "','" + DateTime.Now.ToString("H:mm") + "','" +
                        _deliveryDate.ToString("yyyy-MM-dd") + "','" + order.CheckType + "','CheckPower Commercial','" +
                        order.BRSTN + "','" + order.AccountNo + "','" + order.Name.Replace("'", "''") + "','" + order.Name2.Replace("'", "''") + "','" +
                        start + "','" + end + "','" + order.FileName + "','" + _batch + "','" + order.Address1.Replace("'", "''") + "','" + order.Address2.Replace("'", "''") + "','" +
                        order.Address3.Replace("'", "''") + "','" + order.Address4.Replace("'", "''") + "','" + order.Address5.Replace("'", "''") + "','" + order.Address6.Replace("'", "''") + "');";

                    cmd = new MySqlCommand(query + values, myConn);

                    cmd.ExecuteNonQuery();
                });
            }//END IF
            #endregion

            #region Regular Commercial Pre-Encoded
            if (_orders.CommercialPreEncoded.Count > 0)
            {
                _orders.CommercialPreEncoded.ForEach(order =>
                {
                    string start = order.StartingSerial.ToString(), end = order.EndingSerial.ToString();

                    while (start.Length < 7)
                        start = "0" + start;

                    while (end.Length < 7)
                        end = "0" + end;

                    string values = "VALUES('" + DateTime.Now.ToString("yyyy-MM-dd") + "','" + DateTime.Now.ToString("H:mm") + "','" +
                        _deliveryDate.ToString("yyyy-MM-dd") + "','" + order.CheckType + "','Regular Personal','" +
                        order.BRSTN + "','" + order.AccountNo + "','" + order.Name.Replace("'", "''") + "','" + order.Name2.Replace("'", "''") + "','" +
                        start + "','" + end + "','" + order.FileName + "','" + _batch + "','" + order.Address1.Replace("'", "''") + "','" + order.Address2.Replace("'", "''") + "','" +
                        order.Address3.Replace("'", "''") + "','" + order.Address4.Replace("'", "''") + "','" + order.Address5.Replace("'", "''") + "','" + order.Address6.Replace("'", "''") + "');";

                    cmd = new MySqlCommand(query + values, myConn);

                    cmd.ExecuteNonQuery();
                });
            }//END IF
            #endregion

            #region Manager's Check
            if (_orders.ManagersCheck.Count > 0)
            {
                _orders.ManagersCheck.ForEach(order =>
                {
                    string start = order.StartingSerial.ToString(), end = order.EndingSerial.ToString();

                    while (start.Length < 10)
                        start = "0" + start;

                    while (end.Length < 10)
                        end = "0" + end;

                    string values = "VALUES('" + DateTime.Now.ToString("yyyy-MM-dd") + "','" + DateTime.Now.ToString("H:mm") + "','" +
                        _deliveryDate.ToString("yyyy-MM-dd") + "','" + order.CheckType + "','Manager''s Check','" +
                        order.BRSTN + "','" + order.AccountNo + "','" + order.Name.Replace("'", "''") + "','" + order.Name2.Replace("'", "''") + "','" +
                        start + "','" + end + "','" + order.FileName + "','" + _batch + "','" + order.Address1.Replace("'", "''") + "','" + order.Address2.Replace("'", "''") + "','" +
                        order.Address3.Replace("'", "''") + "','" + order.Address4.Replace("'", "''") + "','" + order.Address5.Replace("'", "''") + "','" + order.Address6.Replace("'", "''") + "');";

                    cmd = new MySqlCommand(query + values, myConn);

                    cmd.ExecuteNonQuery();
                });
            }//END IF
            #endregion

            #region Manager's Check Cont
            if (_orders.ManagersCheckCont.Count > 0)
            {
                _orders.ManagersCheckCont.ForEach(order =>
                {
                    string start = order.StartingSerial.ToString(), end = order.EndingSerial.ToString();

                    while (start.Length < 10)
                        start = "0" + start;

                    while (end.Length < 10)
                        end = "0" + end;

                    string values = "VALUES('" + DateTime.Now.ToString("yyyy-MM-dd") + "','" + DateTime.Now.ToString("H:mm") + "','" +
                        _deliveryDate.ToString("yyyy-MM-dd") + "','" + order.CheckType + "','Manager''s Check Cont','" +
                        order.BRSTN + "','" + order.AccountNo + "','" + order.Name.Replace("'", "''") + "','" + order.Name2.Replace("'", "''") + "','" +
                        start + "','" + end + "','" + order.FileName + "','" + _batch + "','" + order.Address1.Replace("'", "''") + "','" + order.Address2.Replace("'", "''") + "','" +
                        order.Address3.Replace("'", "''") + "','" + order.Address4.Replace("'", "''") + "','" + order.Address5.Replace("'", "''") + "','" + order.Address6.Replace("'", "''") + "');";

                    cmd = new MySqlCommand(query + values, myConn);

                    cmd.ExecuteNonQuery();
                });
            }//END IF
            #endregion

            myConn.Close();
        }//END FUNCTION

        public static void SaveNewSeries(List<BranchesModel> _branches)
        {
            var branches = _branches.Where(r => r.IfChanges == 1).ToList();

            if (branches.Count > 0)
            {
                MySqlConnection conn = new MySqlConnection(DatabaseConnection.ConnectionString);

                conn.Open();

                branches.ForEach(branch =>
                {
                    string query = "UPDATE captive_database.sbtc_branches SET LastNo_PA = " + branch.LastNo_PA +
                    ", LastNo_CA = " + branch.LastNo_CA + ", LastNo_MC = " + branch.LastNo_MC + ", LastNo_Power_PA =" + branch.LastNo_Power_PA +
                    ", LastNo_Power_CA = " + branch.LastNo_Power_CA + ", LastNo_GC = " + branch.LastNo_GC + ", LastNo_CheckOne_CA = " +
                    branch.LastNo_CheckOne_CA + ", LastNo_CheckOne_PA = " + branch.LastNo_CheckOne_PA + ", ModifiedDate = '" +
                    DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss") + "' WHERE BRSTN ='" + branch.BRSTN + "';";

                    MySqlCommand cmd = new MySqlCommand(query, conn);

                    cmd.ExecuteNonQuery();
                });//END FOREACH

                conn.Close();
            }//END IF
        }//END FUNCTION

        private static string GetMySQLLocator()
        {
            MySqlConnection conn = new MySqlConnection(DatabaseConnection.ConnectionString);

            conn.Open();

            MySqlCommand cmd = new MySqlCommand("SELECT * FROM captive_database.mysqldump_location", conn);

            MySqlDataReader reader = cmd.ExecuteReader();

            List<Locator> sqlLocations = new List<Locator>();

            while (reader.Read())
            {
                Locator loc = new Locator
                {
                    PrimaryKey = reader.GetInt32(0),
                    Location = reader.GetString(1)
                };

                sqlLocations.Add(loc);
            }

            conn.Close();

            foreach (var loc in sqlLocations)
            {
                if (File.Exists(loc.Location))
                    return loc.Location;
            }

            return "";
        }

        public static void ProcessSQLDump()
        {
            string dbName = "sbtc_branches";

            Process proc = new Process();

            proc.StartInfo.FileName = "cmd.exe";

            proc.StartInfo.UseShellExecute = false;

            proc.StartInfo.WorkingDirectory = GetMySQLLocator().ToUpper().Replace("MYSQLDUMP.EXE", "");

            proc.StartInfo.RedirectStandardInput = true;

            proc.StartInfo.RedirectStandardOutput = true;

            proc.Start();

            StreamWriter myStreamWriter = proc.StandardInput;

            string temp = DatabaseConnection.DumpString + dbName + " > " + DatabaseConnection.LiveStartPath + "\\" +
                DateTime.Today.ToShortDateString().Replace("/", ".") + "-" + dbName + ".SQL";

            myStreamWriter.WriteLine(temp);

            dbName = "sbtc_history";

            temp = DatabaseConnection.DumpString + dbName + " > " + DatabaseConnection.LiveStartPath + "\\" +
                DateTime.Today.ToShortDateString().Replace("/", ".") + "-" + dbName + ".SQL";

            myStreamWriter.WriteLine(temp);

            myStreamWriter.Close();

            proc.WaitForExit();

            proc.Close();

        }//END FUNCTION

        private static string GetWinZipLoc()
        {
            MySqlConnection conn = new MySqlConnection(DatabaseConnection.ConnectionString);

            conn.Open();

            MySqlCommand cmd = new MySqlCommand("SELECT * FROM captive_database.winzip_location", conn);

            MySqlDataReader read = cmd.ExecuteReader();

            List<Locator> locators = new List<Locator>();

            while (read.Read())
            {
                Locator loc = new Locator
                {
                    PrimaryKey = read.GetInt32(0),
                    Location = read.GetString(1)
                };

                locators.Add(loc);
            }

            conn.Close();

            foreach(var loc in locators)
            {
                if (File.Exists(loc.Location))
                    return loc.Location;
            }
            return "";
        }

        public static void ProcessArchiving(string _batchNumber, string _processBy, OrderSorted _orders = null)
        {
            //ProcessSQLDump();
            
            Process proc = new Process();

            proc.EnableRaisingEvents = false;

            proc.StartInfo.FileName = "\"" + GetWinZipLoc().Replace("\\", "\\\\") + "\"";
               
            proc.StartInfo.Arguments = "-u -r -p " + "\"" + DatabaseConnection.ArchiveOutPut + "\\AFT" + _batchNumber + "_" + _processBy + ".zip\"" + " " +
                "\\\\192.168.0.254\\captive\\Auto\\SBTC\\SBTC_2.0\\Output\\*.*";

            proc.Start();

            proc.WaitForExit();

            proc.Close();
        }//END FUNCTION
    }
}
