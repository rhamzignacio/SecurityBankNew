using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Data.OleDb;
using Microsoft.Win32;
using MySql.Data;
using MySql.Data.MySqlClient;
using System.Data.Odbc;




namespace sbtc
{
    public static class ReturnMe
    {
        public static Boolean CodesOnly = false;
        public static string PrinterFiles_Folder = "";
        public static string Resting_Folder = "";
        public static string TimeStart = "";
        public static string server = "";
        public static string uid = "root";
        public static string password = "CorpCaptive";
        public static string WinZipLocation = "";
        public static DateTime DateTimeToday_date;
        
        public static int mdb_status_bar;
        public static int mdb_status_bar_max;

        public static int status_bar;
        public static int status_bar_max;

        public static void ActivateMaxProgressBarCarbon()
        {
            string sql = "SELECT ChkType, FormType, SUM(OrderQty) FROM SBTC WHERE (ChkType = 'F' AND FormType = '25') OR (ChkType = 'F' AND FormType = '26') OR (ChkType = 'GC' AND FormType = '20') OR (ChkType = 'MC' AND FormType = '20') OR (ChkType = 'MC_1' AND FormType = '00') OR (ChkType = 'CUSTOM' AND FormType = '00') OR (ChkType = 'CUSTOM_PA' AND FormType = '00') OR (ChkType = 'CS') GROUP BY ChkType, FormType";

            int TotalMDB = 0;

            OleDbConnection conn1 = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + Application.StartupPath + ";Extended Properties=dBASE III;");
            OleDbDataAdapter command1 = new OleDbDataAdapter(sql, conn1);
            conn1.Open();
            DataSet dataSet = new DataSet();
            command1.Fill(dataSet);

            DataTable dt = dataSet.Tables[0];
            foreach (DataRow dr in dt.Rows)
            {
                string ChkType = dr[0].ToString();
                string FormType = dr[1].ToString();
                int OrderQty = 0;
                if (Int32.TryParse(dr[2].ToString(), out OrderQty))
                {
                }
                else { OrderQty = 0; }

                int pcsperbook = 0;
                if (ChkType == "F" && FormType == "25")	pcsperbook = 25;
                if (ChkType == "F" && FormType == "26")	pcsperbook = 50;
                if (ChkType == "GC" && FormType == "20") pcsperbook = 50;
                if (ChkType == "MC" && FormType == "20") pcsperbook = 50;
                if (ChkType == "MC_1" && FormType == "00") pcsperbook = 100;
                if (ChkType == "CUSTOM" && FormType == "00") pcsperbook = 100;
                if (ChkType == "CUSTOM_PA" && FormType == "00") pcsperbook = 50;
                if (ChkType == "CS") pcsperbook = 50;

                int total = pcsperbook * OrderQty;    //Correct
                TotalMDB =TotalMDB + total;


                
                
                
                //Check for 4 outs
                int additional = 0;
                int DataNumber = total;
                while (DataNumber % (pcsperbook * 4) != 0)
                {
                    DataNumber = DataNumber + 1;
                    additional = additional + 1;
                }
                TotalMDB = TotalMDB + additional + DataNumber;
                //End Check for 4 outs
                


                //Check for 3 outs
                additional = 0;
                DataNumber = total;
                while (DataNumber % (pcsperbook * 3) != 0)
                {
                    DataNumber = DataNumber + 1;
                    additional = additional + 1;
                }
                TotalMDB = TotalMDB + additional + DataNumber;
                //End Check for 3 outs


                //Check for 2 outs
                additional = 0;
                DataNumber = total;
                while (DataNumber % (pcsperbook * 2) != 0)
                {
                    DataNumber = DataNumber + 1;
                    additional = additional + 1;
                }
                TotalMDB = TotalMDB + additional + DataNumber;
                //End Check for 2 outs
 
            }
            conn1.Close();

            mdb_status_bar_max = TotalMDB;

            sql = "SELECT SUM(OrderQty) FROM SBTC";
            conn1 = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + Application.StartupPath + ";Extended Properties=dBASE III;");
            command1 = new OleDbDataAdapter(sql, conn1);
            conn1.Open();
            dataSet = new DataSet();
            command1.Fill(dataSet);

            dt = dataSet.Tables[0];
            foreach (DataRow dr in dt.Rows)
            {
                string status_bar_max_result = dr[0].ToString();

                if (Int32.TryParse(dr[0].ToString(), out status_bar_max))
                {
                }
                else { status_bar_max = 0; }
            }
            conn1.Close();
        }
        public static string getAddress1(string brstn, int addressline, string ChequeName)
        {
            string address = "";

            string sql = "SELECT Address" + addressline.ToString() + " FROM branches WHERE BRSTN = '" + brstn + "'";

            OleDbConnection conn1;

            if (ChequeName == "MC_1")
            {
                conn1 = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + Application.StartupPath + "\\MC\\Continues\\;Extended Properties=dBASE III;");
            }
            else
            {
                conn1 = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + Application.StartupPath + ";Extended Properties=dBASE III;");
            }
            
            OleDbDataAdapter command1 = new OleDbDataAdapter(sql, conn1);
            conn1.Open();
            DataSet dataSet = new DataSet();
            command1.Fill(dataSet);

            DataTable dt = dataSet.Tables[0];
            foreach (DataRow dr in dt.Rows)
            {
                address = dr[0].ToString();
            }
            conn1.Close();

            return address;
        }
        public static void CreateTable()
        {
            if (System.IO.File.Exists("sbtc.dbf") == true)
            {
                System.IO.File.Delete("sbtc.dbf");
            }

            if (System.IO.File.Exists("temp.dbf") == true)
            {
                System.IO.File.Delete("temp.dbf");
            }

            if (System.IO.File.Exists("errors.dbf") == true)
            {
                System.IO.File.Delete("errors.dbf");
            }

            if (System.IO.File.Exists("batch.dbf") == true)
            {
                System.IO.File.Delete("batch.dbf");
            }

            string FilePath = Application.StartupPath;
            OleDbConnection conn = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + FilePath + ";Extended Properties=dBASE III;");

            OleDbCommand command = new OleDbCommand("CREATE TABLE SBTC (ChkType Varchar(9),BRSTN Varchar(9) , AccountNo Varchar(12), Name1 Varchar(60), Name2 Varchar(60), FormType Varchar(2),OrderQty Varchar(3), Batch Varchar(30), Address1 Varchar(60), Address2 Varchar(60), Address3 Varchar(60), Address4 Varchar(60), Address5 Varchar(60), Address6 Varchar(60), PKey Numeric, BStock Varchar(50), FileName Varchar(50), StartSN Varchar(50), PcsPerBook Varchar(3) , StartSN1 numeric)", conn);
            conn.Open();
            command.ExecuteReader();

            conn.Close();


            command = new OleDbCommand("CREATE TABLE Temp (AccountNo Varchar(12), AcctName Varchar(60), Batch Varchar(30), Pkey Numeric, FileName Varchar(50))", conn);
            conn.Open();
            command.ExecuteReader();

            conn.Close();


            command = new OleDbCommand("CREATE TABLE Errors (Errors Varchar(244))", conn);
            conn.Open();
            command.ExecuteReader();

            conn.Close();

            command = new OleDbCommand("CREATE TABLE Batch (Batch Varchar(244))", conn);
            conn.Open();
            command.ExecuteReader();

            conn.Close();
        }
        public static DialogResult InputBox(string title, string promptText, ref string value)
        {
              Form form = new Form();
              Label label = new Label();
              TextBox textBox = new TextBox();
              Button buttonOk = new Button();
              Button buttonCancel = new Button();

              form.Text = title;
              label.Text = promptText;
              textBox.Text = value;

              buttonOk.Text = "OK";
              buttonCancel.Text = "Cancel";
              buttonOk.DialogResult = DialogResult.OK;
              buttonCancel.DialogResult = DialogResult.Cancel;

              label.SetBounds(9, 20, 372, 13);
              textBox.SetBounds(12, 36, 372, 20);
              buttonOk.SetBounds(228, 72, 75, 23);
              buttonCancel.SetBounds(309, 72, 75, 23);

              label.AutoSize = true;
              textBox.Anchor = textBox.Anchor | AnchorStyles.Right;
              buttonOk.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
              buttonCancel.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;

              form.ClientSize = new Size(396, 107);
              form.Controls.AddRange(new Control[] { label, textBox, buttonOk, buttonCancel });
              form.ClientSize = new Size(Math.Max(300, label.Right + 10), form.ClientSize.Height);
              form.FormBorderStyle = FormBorderStyle.FixedDialog;
              form.StartPosition = FormStartPosition.CenterScreen;
              form.MinimizeBox = false;
              form.MaximizeBox = false;
              form.AcceptButton = buttonOk;
              form.CancelButton = buttonCancel;

              DialogResult dialogResult = form.ShowDialog();
              value = textBox.Text;
              return dialogResult;
        }
        public static void CreateDirectory(string Path_Location)
        {
            if (Directory.Exists(Path_Location) == false)
            {
                Directory.CreateDirectory(Path_Location);
            }
        }
        public static void TransferAll(string Batch)
        {
            if (Directory.Exists(Application.StartupPath + "\\Archive\\" + Batch) == false)
            {
                Directory.CreateDirectory(Application.StartupPath + "\\Archive\\" + Batch);
            }

            string[] files = Directory.GetFiles(Application.StartupPath + "\\Head\\", "*.txt");

            foreach (string file in files)
            {

                string filename = "";

                int LoopCount = file.Length - 1;
                while (LoopCount > 4)
                {
                    if (file.Substring(LoopCount, 1) == "\\")
                    {
                        if (filename == "")
                        {
                            filename = file.Substring(LoopCount + 1, file.Length - LoopCount - 1);
                        }
                    }

                    LoopCount = LoopCount - 1;
                }

                
                File.Copy(file, Application.StartupPath + "\\Archive\\" + Batch + "\\" + filename);
                File.Delete(file);
            }
        }
        public static void ProcessAll2(DateTime DeliveryDate)
        {
repeatme:

            string Batch = DateTime.Now.ToString("MMddyyyy");
            if (ReturnMe.InputBox("Batch", "Enter Batch Name:", ref Batch) == DialogResult.OK)
            {

            }
            else
            {
                Batch = "";
            }

            if (Batch == "") { goto repeatme; }

            //Check if Batch exists
            if (Directory.Exists(Application.StartupPath + "\\Archive\\" + Batch) || ReturnMe.CheckBatchExists(Batch) == true)
            {
                MessageBox.Show("Batch " + Batch + " already exists", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                goto repeatme;
            }
            //End Check if Batch exists

            //For Process By
            string ProcessBy = "";
            if (ReturnMe.InputBox("", "Enter Process By:", ref ProcessBy) == DialogResult.OK)
            {

            }
            else
            {
                ProcessBy = "";
            }

           
            string CheckedBy = "";
            if (ReturnMe.InputBox("", "Enter Checked By:", ref CheckedBy) == DialogResult.OK)
            {

            }
            else
            {
                CheckedBy = "";
            }

            
            //End for Process By


            //For Zip
            string DateTimeToday = DateTime.Now.ToString("MMddyyyyHHmmss");

            Directory.CreateDirectory("C:\\Windows\\Temp\\" + DateTimeToday + "\\" + Batch);
            //End For Zip

            //Clear Folder
            int LoopCount = 0;
            while (LoopCount < 10)
            {
                string FolderLocation = "";
                if (LoopCount == 0) { FolderLocation = Application.StartupPath + "\\Charge_Slip"; }
                if (LoopCount == 1) { FolderLocation = Application.StartupPath + "\\CheckOne"; }
                if (LoopCount == 2) { FolderLocation = Application.StartupPath + "\\CheckPower"; }
                if (LoopCount == 3) { FolderLocation = Application.StartupPath + "\\Customized"; }
                if (LoopCount == 4) { FolderLocation = Application.StartupPath + "\\GiftCheck"; }
                if (LoopCount == 5) { FolderLocation = Application.StartupPath + "\\MC"; }
                if (LoopCount == 6) { FolderLocation = Application.StartupPath + "\\MC\\Continues"; }
                if (LoopCount == 7) { FolderLocation = Application.StartupPath + "\\Regular"; }
                if (LoopCount == 8) { FolderLocation = Application.StartupPath + "\\Regular\\PreEncoded"; }
                if (LoopCount == 9)
                {
                    FolderLocation = Application.StartupPath;
                }

                string[] files = Directory.GetFiles(FolderLocation, "*.*");
                foreach (string file in files)
                {
                    if (file.Substring(file.Length - 3, 3).ToUpper() == "TXT")
                    {
                        if (LoopCount >= 0 && LoopCount <= 8)
                        {
                            if (file.Substring(file.Length - 10, 10).ToUpper() != "SORTRT.TXT")
                            {
                                File.Delete(file);
                            }
                        }
                    }

                    if (file.Substring(file.Length - 3, 3).ToUpper() == "MDB")
                    {
                        if (LoopCount >= 0 && LoopCount <= 8)
                        {
                            File.Delete(file);
                        }
                    }
                    if (file.Substring(file.Length - 3, 3).ToUpper() == "ZIP")
                    {
                        if (LoopCount == 9)
                        {
                            File.Delete(file);
                        }
                    }
                }
                LoopCount = LoopCount + 1;
            }
            //End Clear Folder

            ReturnMe.ProcessAll(Batch, ProcessBy, CheckedBy, DeliveryDate, DateTimeToday);

            TransferAll(Batch);

            DialogResult result1 = MessageBox.Show("Data has been processed. Send hash total?", "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Information);
            if (result1 == DialogResult.No) 
            {
                Application.Exit();
                return;
            }

            ReturnMe.SendHashTotal(DateTime.Now.ToString("yyyy-MM-dd"));
        }
        public static void DeletePackDBF(string strTableName, string strPath)
        {

            string connectionString = @"Provider=VFPOLEDB.1;Data Source=" + strPath; 

            using (OleDbConnection connection = new OleDbConnection(connectionString)) 
            { 
                using (OleDbCommand scriptCommand = connection.CreateCommand()) 
                { 
                    connection.Open();

                    string vfpScript = @"SET EXCLUSIVE ON
                                        DELETE FROM " + strTableName.Substring(0,strTableName.Length-4) +
                                        "\r\nPACK"; 

                    scriptCommand.CommandType = CommandType.StoredProcedure; 
                    scriptCommand.CommandText = "ExecScript"; 
                    scriptCommand.Parameters.Add("myScript", OleDbType.Char).Value = vfpScript; 
                    scriptCommand.ExecuteNonQuery(); 
                } 
            } 
        }
        public static void SendHashTotal(string DeliveryDate)
        {

            //string recipient_email = "gsdpurchasing7@securitybank.com.ph,gsdpurchasing5@securitybank.com.ph,gsdpurchasing2@securitybank.com.ph,ctimusan@securitybank.com.ph,rmenguito@securitybank.com.ph,virtualsupport@securitybank.com.ph,GSDPurchasing4@securitybank.com.ph,orders@captiveprinting.com.ph,cpc_services@captiveprinting.com.ph,eguzman@securitybank.com.ph,cbgcoc@securitybank.com.ph";

            string recipient_email = "orders@captiveprinting.com.ph,cpc_services@captiveprinting.com.ph";

            //get the max
            string dbase = "";
            if (ReturnMe.CodesOnly == true)
            {
                dbase = "captive_database.sbtc_history";
                recipient_email = "orders@captiveprinting.com.ph";
            }
            if (ReturnMe.CodesOnly == false) dbase = "captive_database.sbtc_history";



            string sql = "SELECT MAX(LENGTH(ChequeName)) , MAX(LENGTH(Batch)) , MAX(LENGTH(BRSTN)) , MAX(LENGTH(Address1)) , MAX(LENGTH(AccountNo)) , MAX(LENGTH(Name1)) , MAX(LENGTH(Name2)) , MAX(LENGTH(StartingSerial)) , MAX(LENGTH(EndingSerial)) FROM " + dbase + " WHERE DeliveryDate = '" + DeliveryDate + "'";
            string MyConnection = "datasource=" + ReturnMe.server + ";port=3306;username=" + ReturnMe.uid + ";password=" + ReturnMe.password;
            MySqlConnection MyConn = new MySqlConnection(MyConnection);
            MySqlCommand MyCommand = new MySqlCommand(sql, MyConn);
            MySqlDataReader MyReader;
            MyConn.Open();
            MyReader = MyCommand.ExecuteReader();

            int max_chequename = 0;
            int max_batch = 0;
            int max_brstn = 0;
            int max_address1 = 0;
            int max_accountno = 0;
            int max_name1 = 0;
            int max_name2 = 0;
            int max_startingserial = 0;
            int max_endingserial = 0;

            while (MyReader.Read())
            {
                max_chequename = MyReader.GetInt32(0);
                max_batch = MyReader.GetInt32(1);
                max_brstn = MyReader.GetInt32(2);
                max_address1 = MyReader.GetInt32(3);
                max_accountno = MyReader.GetInt32(4);
                max_name1 = MyReader.GetInt32(5);
                max_name2 = MyReader.GetInt32(6);
                max_startingserial = MyReader.GetInt32(7);
                max_endingserial = MyReader.GetInt32(8);
            }
            //End get the max






            //All batch
            string all_batch = "";
            
            sql = "SELECT DISTINCT(Batch) FROM " + dbase + " WHERE DeliveryDate = '" + DeliveryDate + "'";
            string MyConnection9 = "datasource=" + ReturnMe.server + ";port=3306;username=" + ReturnMe.uid + ";password=" + ReturnMe.password;
            MySqlConnection MyConn9 = new MySqlConnection(MyConnection9);
            MySqlCommand MyCommand9 = new MySqlCommand(sql, MyConn9);
            MySqlDataReader MyReader9;
            MyConn9.Open();
            MyReader9 = MyCommand9.ExecuteReader();

            while (MyReader9.Read())
            {
                string batch_result = MyReader9.GetValue(0).ToString();

                if (all_batch == "")
                {
                    all_batch = batch_result;
                }
                else 
                {
                    all_batch = all_batch + "_" + batch_result;
                }
            }

            MyConn9.Close();
            //End All batch








            if (File.Exists("C:\\Windows\\Temp\\" + DateTime.Parse(DeliveryDate).ToString("MMddyyyy") + ".xls") == true)
            {
                File.Delete("C:\\Windows\\Temp\\" + DateTime.Parse(DeliveryDate).ToString("MMddyyyy") + ".xls");
            }
            File.Copy(Application.StartupPath + "\\Delivery_Report_Source.xls", "C:\\Windows\\Temp\\" + DateTime.Parse(DeliveryDate).ToString("MMddyyyy") + ".xls");


            var excelApp = new Microsoft.Office.Interop.Excel.Application();
            var workbook = excelApp.Workbooks.Open("C:\\Windows\\Temp\\" + DateTime.Parse(DeliveryDate).ToString("MMddyyyy") + ".xls");
            Microsoft.Office.Interop.Excel.Worksheet worksheet = workbook.ActiveSheet as Microsoft.Office.Interop.Excel.Worksheet;


           



            //StreamWriter sw = new StreamWriter("C:\\Windows\\Temp\\" + all_batch + ".txt");




            sql = "SELECT chequename , batch , brstn , accountno , name1 , name2 , min(startingserial) , max(endingserial) , address1, deliverydate , count(primarykey), Date FROM " + dbase + " WHERE DeliveryDate = '" + DeliveryDate + "' GROUP BY chequename , batch , brstn , accountno , name1 , name2 , address1, deliverydate, Date     ORDER BY ChequeName, Batch, BRSTN, AccountNo, Name1";
            string MyConnection2 = "datasource=" + ReturnMe.server + ";port=3306;username=" + ReturnMe.uid + ";password=" + ReturnMe.password;
            MySqlConnection MyConn2 = new MySqlConnection(MyConnection2);
            MySqlCommand MyCommand2 = new MySqlCommand(sql, MyConn2);
            MySqlDataReader MyReader2;
            MyConn2.Open();
            MyReader2 = MyCommand2.ExecuteReader();


            string subject_email = "";


            int LoopCount = 0;

            while (MyReader2.Read())
            {
                string chequename = MyReader2.GetValue(0).ToString();
                string batch = MyReader2.GetValue(1).ToString();
                string brstn = MyReader2.GetValue(2).ToString();
                string accountno = MyReader2.GetValue(3).ToString();
                string name1 = MyReader2.GetValue(4).ToString();
                string name2 = MyReader2.GetValue(5).ToString();
                string startingserial = MyReader2.GetValue(6).ToString();
                string endingserial = MyReader2.GetValue(7).ToString();
                string address1 = MyReader2.GetValue(8).ToString();
                DateTime deliverydate = MyReader2.GetDateTime(9);
                string OrderQty = MyReader2.GetString(10);
                string DateProcess = MyReader2.GetDateTime(11).ToString("yyyy-MM-dd");


                
                if (LoopCount == 0)
                {
                    //sw.WriteLine("For Delivery Date: " + deliverydate.ToString("MMMM. dd, yyyy"));
                    //sw.WriteLine("");

                    subject_email = "SBTC Hash Total for Delivery Date " + DateTime.Parse(DeliveryDate).ToString("MMMM. dd, yyyy") + " - Batch: " + all_batch.Replace("_",",");
                }
                

                Microsoft.Office.Interop.Excel.Range range = worksheet.Cells[LoopCount + 2, 1] as Microsoft.Office.Interop.Excel.Range;
                range.Value2 = DateProcess;

                range = worksheet.Cells[LoopCount + 2, 2] as Microsoft.Office.Interop.Excel.Range;
                range.Value2 = batch;

                range = worksheet.Cells[LoopCount + 2, 3] as Microsoft.Office.Interop.Excel.Range;
                range.Value2 = chequename;

                range = worksheet.Cells[LoopCount + 2, 4] as Microsoft.Office.Interop.Excel.Range;
                range.Value2 = address1;

                range = worksheet.Cells[LoopCount + 2, 5] as Microsoft.Office.Interop.Excel.Range;
                range.Value2 = brstn;

                range = worksheet.Cells[LoopCount + 2, 6] as Microsoft.Office.Interop.Excel.Range;
                range.Value2 = accountno;

                range = worksheet.Cells[LoopCount + 2, 7] as Microsoft.Office.Interop.Excel.Range;
                range.Value2 = name1;

                range = worksheet.Cells[LoopCount + 2, 8] as Microsoft.Office.Interop.Excel.Range;
                range.Value2 = name2;

                range = worksheet.Cells[LoopCount + 2, 9] as Microsoft.Office.Interop.Excel.Range;
                range.Value2 = startingserial;

                range = worksheet.Cells[LoopCount + 2, 10] as Microsoft.Office.Interop.Excel.Range;
                range.Value2 = endingserial;

                range = worksheet.Cells[LoopCount + 2, 11] as Microsoft.Office.Interop.Excel.Range;
                range.Value2 = OrderQty;
                

                LoopCount = LoopCount + 1;
            }


            //xlWorkBook.Close(true, null, null);

            workbook.Save();
            workbook.Close();




            //for summary
            string summary_batch_email = "";


            sql = "SELECT Batch, COUNT(PrimaryKey) FROM " + dbase + " WHERE DeliveryDate = '" + DeliveryDate + "' GROUP BY Batch";
            string MyConnection5 = "datasource=" + ReturnMe.server + ";port=3306;username=" + ReturnMe.uid + ";password=" + ReturnMe.password;
            MySqlConnection MyConn5 = new MySqlConnection(MyConnection5);
            MySqlCommand MyCommand5 = new MySqlCommand(sql, MyConn5);
            MySqlDataReader MyReader5;
            MyConn5.Open();
            MyReader5 = MyCommand5.ExecuteReader();

            while (MyReader5.Read())
            {
                string batch = MyReader5.GetValue(0).ToString();
                string qty = MyReader5.GetValue(1).ToString();

                while (batch.Length < 15)
                {
                    batch = batch + " ";
                }




                if (summary_batch_email == "")
                {
                    summary_batch_email = batch + qty;
                }
                else
                {
                    summary_batch_email = summary_batch_email + "\r\n" + batch + qty;
                }
            }




            string summary_chequename_email = "";


            sql = "SELECT ChequeName, COUNT(PrimaryKey) FROM " + dbase + " WHERE DeliveryDate = '" + DeliveryDate + "' GROUP BY ChequeName";
            string MyConnection6 = "datasource=" + ReturnMe.server + ";port=3306;username=" + ReturnMe.uid + ";password=" + ReturnMe.password;
            MySqlConnection MyConn6 = new MySqlConnection(MyConnection6);
            MySqlCommand MyCommand6 = new MySqlCommand(sql, MyConn6);
            MySqlDataReader MyReader6;
            MyConn6.Open();
            MyReader6 = MyCommand6.ExecuteReader();

            while (MyReader6.Read())
            {
                string chequename = MyReader6.GetValue(0).ToString();
                string qty = MyReader6.GetValue(1).ToString();

                while (chequename.Length < 25)
                {
                    chequename = chequename + " ";
                }



                if (summary_chequename_email == "")
                {
                    summary_chequename_email = chequename + qty;
                }
                else
                {
                    summary_chequename_email = summary_chequename_email + "\r\n" + chequename + qty;
                }
            }
            //End for summary


            string heading_email = "  Hello and Good Day"
                            + "\r\n"
                            + "\r\n"
                            + "\r\n"
                            + "     Kindly see the attached file for the Hash Total."
                            + "\r\n"
                            + "\r\n"
                            + "\r\n"
                            + "     Orders as of this Batch:"
                            + "\r\n"
                            + "\r\n"
                            + "\r\n"
                            + summary_batch_email
                            + "\r\n"
                            + "\r\n"
                            + "\r\n"
                            + summary_chequename_email;

            string Footer = "\r\n"
               + "\r\n"
               + "\r\n"
               + "\r\n"
               + "     This is a System Generated Message." + "\r\n"
               + "\r\n"
               + "\r\n"
               + "\r\n"
               + "\r\n"
               + "     Thanks and Best Regards,"
               + "\r\n"
               + "\r\n"
               + "\r\n"
               + "     Captive Printing Corporation    ";

            string Body_email = heading_email + "\r\n" + Footer;



            sql = "SELECT max(primarykey) FROM captive_database.emails";
            string MyConnection3 = "datasource=" + ReturnMe.server + ";port=3306;username=" + ReturnMe.uid + ";password=" + ReturnMe.password;
            MySqlConnection MyConn3 = new MySqlConnection(MyConnection3);
            MySqlCommand MyCommand3 = new MySqlCommand(sql, MyConn3);
            MySqlDataReader MyReader3;
            MyConn3.Open();
            MyReader3 = MyCommand3.ExecuteReader();

            int max_pkey = 0;

            while (MyReader3.Read())
            {
                max_pkey = MyReader3.GetInt32(0);
            }


            sql = "INSERT INTO captive_database.Emails (Bank , Recipient_Email , Subject , Body , DateRequest , TimeRequest , Status , PrimaryKey , source_email) VALUES ('SBTC','" + recipient_email + "','" + subject_email + "','" + Body_email.Replace("'", "''") + "','" + DateTime.Now.ToString("yyyy-MM-dd") + "','" + DateTime.Now.ToString("HH:mm:ss") + "','Received','" + (max_pkey + 1) + "','orders@captiveprinting.com.ph')";
            string MyConnection4 = "datasource=" + ReturnMe.server + ";port=3306;username=" + ReturnMe.uid + ";password=" + ReturnMe.password;
            MySqlConnection MyConn4 = new MySqlConnection(MyConnection4);
            MySqlCommand MyCommand4 = new MySqlCommand(sql, MyConn4);
            MySqlDataReader MyReader4;
            MyConn4.Open();
            MyReader4 = MyCommand4.ExecuteReader();






            //save to file
            byte[] rawData = File.ReadAllBytes("C:\\Windows\\Temp\\" + DateTime.Parse(DeliveryDate).ToString("MMddyyyy") + ".xls");
            FileInfo info = new FileInfo("C:\\Windows\\Temp\\" + DateTime.Parse(DeliveryDate).ToString("MMddyyyy") + ".xls");

            int fileSize = Convert.ToInt32(info.Length);


            MyConnection2 = "datasource=" + ReturnMe.server + ";port=3306;username=" + ReturnMe.uid + ";password=" + ReturnMe.password;
            MySqlConnection connection = new MySqlConnection(MyConnection2);
            MySqlCommand command = new MySqlCommand();
            command.Connection = connection;
            command.CommandText = "INSERT INTO captive_database.emails_blob (Attachment,filename,primarykey_source) VALUES (?rawData,'" + DateTime.Parse(DeliveryDate).ToString("MMddyyyy") + ".xls','" + (max_pkey + 1) + "');";


            MySqlParameter fileContentParameter = new MySqlParameter("?rawData", MySqlDbType.MediumBlob, rawData.Length);
            fileContentParameter.Direction = ParameterDirection.Input;

            fileContentParameter.Value = rawData;
            command.Parameters.Add(fileContentParameter);

            connection.Open();

            command.ExecuteNonQuery();
            //End save to file






            //Check until send
            string status = "Received";
            string message = "";

        repeatme:

            sql = "SELECT Status,ErrorMessage FROM captive_database.Emails WHERE PrimaryKey = " + (max_pkey + 1);
            string MyConnection7 = "datasource=" + ReturnMe.server + ";port=3306;username=" + ReturnMe.uid + ";password=" + ReturnMe.password;
            MySqlConnection MyConn7 = new MySqlConnection(MyConnection7);
            MySqlCommand MyCommand7 = new MySqlCommand(sql, MyConn7);
            MySqlDataReader MyReader7;
            MyConn7.Open();
            MyReader7 = MyCommand7.ExecuteReader();



            while (MyReader7.Read())
            {
                status = MyReader7.GetString(0);
                if (!MyReader7.IsDBNull(1))
                {
                    message = MyReader7.GetString(1);
                }
                else message = "";
            }

            if (status == "Received")
            {
                MyConn7.Close();
                goto repeatme;
            }

            if (status == "Sent")
            {
                sql = "UPDATE " + dbase + " SET HashSentDate = '" + DateTime.Now.ToString("yyyy-MM-dd") + "', HashSentTime = '" + DateTime.Now.ToString("HH:mm:ss") + "' WHERE DeliveryDate = '" + DeliveryDate + "'";
                string MyConnection8 = "datasource=" + ReturnMe.server + ";port=3306;username=" + ReturnMe.uid + ";password=" + ReturnMe.password;
                MySqlConnection MyConn8 = new MySqlConnection(MyConnection8);
                MySqlCommand MyCommand8 = new MySqlCommand(sql, MyConn8);
                MySqlDataReader MyReader8;
                MyConn8.Open();
                MyReader8 = MyCommand8.ExecuteReader();

                MessageBox.Show("Hash Total has been sent", " ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                Application.Exit();
            }

            if (status == "Failed")
            {
                MessageBox.Show(message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                Application.Exit();
            }
            //end Check until send
        }
        public static void ProcessAll(string Batch, string ProcessBy, string CheckedBy, DateTime  DeliveryDate, string DateTimeToday)
        {
            
            int Reg_PA = ProcessMe("A", "05", "Regular", true, Batch, DeliveryDate,ProcessBy ,DateTimeToday);
            int Reg_CA = ProcessMe("B", "16", "Regular", false, Batch, DeliveryDate,ProcessBy,DateTimeToday);

            int PreEncoded_PA = ProcessMe("AA", "05", "Regular\\PreEncoded", true , Batch, DeliveryDate,ProcessBy,DateTimeToday);
            int PreEncoded_CA = ProcessMe("BB", "16", "Regular\\PreEncoded", false, Batch, DeliveryDate,ProcessBy,DateTimeToday);

            int MC = ProcessMe("MC", "20", "MC", true, Batch, DeliveryDate,ProcessBy,DateTimeToday);
    
            int CheckOne_PA = ProcessMe("F", "25", "CheckOne", true, Batch, DeliveryDate,ProcessBy,DateTimeToday);
            int CheckOne_CA = ProcessMe("F", "26", "CheckOne", false, Batch, DeliveryDate,ProcessBy,DateTimeToday);

            int CheckPower_PA = ProcessMe("E", "23", "CheckPower", true, Batch, DeliveryDate,ProcessBy,DateTimeToday);
            int CheckPower_CA = ProcessMe("E", "22", "CheckPower", false, Batch, DeliveryDate,ProcessBy,DateTimeToday);

            int GiftCheck = ProcessMe("GC", "20", "GiftCheck", true, Batch, DeliveryDate,ProcessBy,DateTimeToday);

            int MC_Continues = ProcessMe("MC_1", "00", "MC\\Continues", true, Batch, DeliveryDate,ProcessBy,DateTimeToday);


            int Customized = ProcessMe("CUSTOM", "00", "Customized", true, Batch, DeliveryDate,ProcessBy,DateTimeToday);
            int Customized_PA = 0;
            if (Customized == 0)
            {
                Customized_PA  = ProcessMe("CUSTOM_PA", "00", "Customized", true, Batch, DeliveryDate, ProcessBy, DateTimeToday);
            }


            int Charge_Slip = ProcessMe("CS", "00", "Charge_Slip", true, Batch, DeliveryDate,ProcessBy,DateTimeToday);





            











            //For Zip
            string Temp_Zip = "";
            string ChequeName_all = "";
            string ChequeName_all_temp = "";

            if (Reg_PA + Reg_CA + PreEncoded_PA + PreEncoded_CA >= 1)
            {
                File.Copy (Application.StartupPath + "\\Regular\\Ref.dbf", "C:\\Windows\\Temp\\" + DateTimeToday + "\\Regular\\Ref.dbf");
                File.Copy (Application.StartupPath + "\\Regular\\Packing.dbf", "C:\\Windows\\Temp\\" + DateTimeToday + "\\Regular\\Packing.dbf");
                File.Copy (Application.StartupPath + "\\Regular\\SortRT.txt", "C:\\Windows\\Temp\\" + DateTimeToday + "\\Regular\\SortRT.txt");
    
                if (ReturnMe.CodesOnly == true) Temp_Zip = Resting_Folder + "\\Zips\\Codes\\SBTC\\" + DateTime.Now.ToString("yyyy");
                if (ReturnMe.CodesOnly == false) Temp_Zip = Resting_Folder + "\\Zips\\SBTC\\" + DateTime.Now.ToString("yyyy");

                ChequeName_all_temp = "Regular";
                if (ChequeName_all == "")
                {
                    ChequeName_all = ChequeName_all_temp;
                } else { ChequeName_all = ChequeName_all + "_" + ChequeName_all_temp;}
            }

            if (PreEncoded_PA + PreEncoded_CA >= 1)
            {
                File.Copy (Application.StartupPath + "\\Regular\\PreEncoded\\Packing.dbf", "C:\\Windows\\Temp\\" + DateTimeToday + "\\Regular\\PreEncoded\\Packing.dbf");
                File.Copy (Application.StartupPath + "\\Regular\\PreEncoded\\SortRT.txt", "C:\\Windows\\Temp\\" + DateTimeToday + "\\Regular\\PreEncoded\\SortRT.txt");

                if (ReturnMe.CodesOnly == true)  { Temp_Zip = Resting_Folder + "\\Zips\\Codes\\SBTC\\" + DateTime.Now.ToString("yyyy");}
                if (ReturnMe.CodesOnly == false) { Temp_Zip = Resting_Folder + "\\Zips\\SBTC\\" + DateTime.Now.ToString("yyyy");}

                ChequeName_all_temp = "PreEncoded";
                if (ChequeName_all == "")
                {
                    ChequeName_all = ChequeName_all_temp;
                }
                else { ChequeName_all = ChequeName_all + "_" + ChequeName_all_temp; }
            }

            if (MC >= 1)
            {
                File.Copy (Application.StartupPath + "\\MC\\Ref.dbf", "C:\\Windows\\Temp\\" + DateTimeToday + "\\MC\\Ref.dbf");
                File.Copy (Application.StartupPath + "\\MC\\Packing.dbf", "C:\\Windows\\Temp\\" + DateTimeToday + "\\MC\\Packing.dbf");
                File.Copy (Application.StartupPath + "\\MC\\SortRT.txt", "C:\\Windows\\Temp\\" + DateTimeToday + "\\MC\\SortRT.txt");
    
                if (ReturnMe.CodesOnly == true) {Temp_Zip = Resting_Folder + "\\Zips\\Codes\\SBTC\\" + DateTime.Now.ToString("yyyy");}
                if (ReturnMe.CodesOnly == false) {Temp_Zip = Resting_Folder + "\\Zips\\SBTC\\" + DateTime.Now.ToString("yyyy");}
                
                ChequeName_all_temp = "MC";
                if (ChequeName_all == "")
                {
                    ChequeName_all = ChequeName_all_temp;
                }
                else { ChequeName_all = ChequeName_all + "_" + ChequeName_all_temp; }
            }

            if (CheckOne_PA + CheckOne_CA >= 1)
            {
                File.Copy (Application.StartupPath + "\\CheckOne\\Ref.dbf", "C:\\Windows\\Temp\\" + DateTimeToday + "\\CheckOne\\Ref.dbf");
                File.Copy (Application.StartupPath + "\\CheckOne\\Packing.dbf", "C:\\Windows\\Temp\\" + DateTimeToday + "\\CheckOne\\Packing.dbf");
                File.Copy (Application.StartupPath + "\\CheckOne\\SortRT.txt", "C:\\Windows\\Temp\\" + DateTimeToday + "\\CheckOne\\SortRT.txt");
    
                if (ReturnMe.CodesOnly == true) { Temp_Zip = Resting_Folder + "\\Zips\\Codes\\SBTC\\" + DateTime.Now.ToString("yyyy");}
                if (ReturnMe.CodesOnly == false) { Temp_Zip = Resting_Folder + "\\Zips\\SBTC\\" + DateTime.Now.ToString("yyyy");}

                ChequeName_all_temp = "CheckOne";
                if (ChequeName_all == "")
                {
                    ChequeName_all = ChequeName_all_temp;
                }
                else { ChequeName_all = ChequeName_all + "_" + ChequeName_all_temp; }
            }

            if (CheckPower_PA + CheckPower_CA >= 1)
            {
                File.Copy (Application.StartupPath + "\\CheckPower\\Ref.dbf", "C:\\Windows\\Temp\\" + DateTimeToday + "\\CheckPower\\Ref.dbf");
                File.Copy (Application.StartupPath + "\\CheckPower\\Packing.dbf", "C:\\Windows\\Temp\\" + DateTimeToday + "\\CheckPower\\Packing.dbf");
                File.Copy (Application.StartupPath + "\\CheckPower\\SortRT.txt", "C:\\Windows\\Temp\\" + DateTimeToday + "\\CheckPower\\SortRT.txt");
    
                if (ReturnMe.CodesOnly == true) Temp_Zip = Resting_Folder + "\\Zips\\Codes\\SBTC\\" + DateTime.Now.ToString("yyyy");
                if (ReturnMe.CodesOnly == false) Temp_Zip = Resting_Folder + "\\Zips\\SBTC\\" + DateTime.Now.ToString("yyyy");

                ChequeName_all_temp = "CheckPower";
                if (ChequeName_all == "")
                {
                    ChequeName_all = ChequeName_all_temp;
                }
                else { ChequeName_all = ChequeName_all + "_" + ChequeName_all_temp; }
            }

            if (GiftCheck >= 1)
            {
                File.Copy (Application.StartupPath + "\\GiftCheck\\Ref.dbf", "C:\\Windows\\Temp\\" + DateTimeToday + "\\GiftCheck\\Ref.dbf");
                File.Copy (Application.StartupPath + "\\GiftCheck\\Packing.dbf", "C:\\Windows\\Temp\\" + DateTimeToday + "\\GiftCheck\\Packing.dbf");
                File.Copy (Application.StartupPath + "\\GiftCheck\\SortRT.txt", "C:\\Windows\\Temp\\" + DateTimeToday + "\\GiftCheck\\SortRT.txt");
    
                if (ReturnMe.CodesOnly == true) Temp_Zip = Resting_Folder + "\\Zips\\Codes\\SBTC\\" + DateTime.Now.ToString("yyyy");
                if (ReturnMe.CodesOnly == false) Temp_Zip = Resting_Folder + "\\Zips\\SBTC\\" + DateTime.Now.ToString("yyyy");

                ChequeName_all_temp = "GiftCheck";
                if (ChequeName_all == "")
                {
                    ChequeName_all = ChequeName_all_temp;
                }
                else { ChequeName_all = ChequeName_all + "_" + ChequeName_all_temp; }
            }

            if (MC_Continues >= 1)
            {
                File.Copy (Application.StartupPath + "\\MC\\Continues\\Branches.dbf", "C:\\Windows\\Temp\\" + DateTimeToday + "\\MC\\Continues\\Branches.dbf");
                File.Copy (Application.StartupPath + "\\MC\\Continues\\Packing.dbf", "C:\\Windows\\Temp\\" + DateTimeToday + "\\MC\\Continues\\Packing.dbf");
                File.Copy (Application.StartupPath + "\\MC\\Continues\\SortRT.txt", "C:\\Windows\\Temp\\" + DateTimeToday + "\\MC\\Continues\\SortRT.txt");
    
                if (ReturnMe.CodesOnly == true) Temp_Zip = Resting_Folder + "\\Zips\\Codes\\SBTC\\CONTINUOUS_MC\\" + DateTime.Now.ToString("yyyy");
                if (ReturnMe.CodesOnly == false) Temp_Zip = Resting_Folder + "\\Zips\\SBTC\\CONTINUOUS_MC\\" + DateTime.Now.ToString("yyyy");

                ChequeName_all_temp = "MC.Continues";
                if (ChequeName_all == "")
                {
                    ChequeName_all = ChequeName_all_temp;
                }
                else { ChequeName_all = ChequeName_all + "_" + ChequeName_all_temp; }
            }

            if (Customized >= 1 || Customized_PA >= 1)
            {
                File.Copy (Application.StartupPath + "\\Customized\\Packing.dbf", "C:\\Windows\\Temp\\" + DateTimeToday + "\\Customized\\Packing.dbf");
                File.Copy (Application.StartupPath + "\\Customized\\SortRT.txt", "C:\\Windows\\Temp\\" + DateTimeToday + "\\Customized\\SortRT.txt");
    
                if (ReturnMe.CodesOnly == true) Temp_Zip = Resting_Folder + "\\Zips\\Codes\\SBTC\\CUSTOMIZED\\" + DateTime.Now.ToString("yyyy");
                if (ReturnMe.CodesOnly == false) Temp_Zip = Resting_Folder + "\\Zips\\SBTC\\CUSTOMIZED\\" + DateTime.Now.ToString("yyyy");

                ChequeName_all_temp = "Customized";
                if (ChequeName_all == "")
                {
                    ChequeName_all = ChequeName_all_temp;
                }
                else { ChequeName_all = ChequeName_all + "_" + ChequeName_all_temp; }
            }


            if (Charge_Slip >= 1)
            {
                if (ReturnMe.CodesOnly == true ) Temp_Zip = Resting_Folder + "\\Zips\\Codes\\SBTC\\Charge_Slip\\" + DateTime.Now.ToString("yyyy") ;
                if (ReturnMe.CodesOnly == false) Temp_Zip = Resting_Folder + "\\Zips\\SBTC\\Charge_Slip\\" + DateTime.Now.ToString("yyyy");

                ChequeName_all_temp = "Charge.Slip";
                if (ChequeName_all == "")
                {
                    ChequeName_all = ChequeName_all_temp;
                }
                else { ChequeName_all = ChequeName_all + "_" + ChequeName_all_temp; }
            }

            File.Copy (Application.StartupPath + "\\sbtc.exe", "C:\\Windows\\Temp\\" + DateTimeToday + "\\sbtc.exe");

            //End For Zip






            string sql = "";

            //For Batch
            string[] files = Directory.GetFiles(Application.StartupPath + "\\Head\\", "*.txt");

            foreach (string file in files)
            {

                string filename = "";

                int LoopCount = file.Length - 1;
                while (LoopCount > 4)
                {
                    if (file.Substring(LoopCount, 1) == "\\")
                    {
                        if (filename == "")
                        {
                            filename = file.Substring(LoopCount + 1, file.Length - LoopCount - 1);
                        }
                    }

                    LoopCount = LoopCount - 1;
                }

                
                //lstFiles.Items.Add(filename);
                sql = "INSERT INTO Batch (Batch) VALUES ('" + filename + "')";
                OleDbConnection conn5 = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + Application.StartupPath  + ";Extended Properties=dBASE III;");
                OleDbDataAdapter command5 = new OleDbDataAdapter(sql, conn5);
                conn5.Open();
                DataSet dataSet5 = new DataSet();
                command5.Fill(dataSet5);
                conn5.Close();
            }





            sql = "SELECT DISTINCT(Batch) FROM Batch";
            OleDbConnection conn1 = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + Application.StartupPath + ";Extended Properties=dBASE III;");
            OleDbDataAdapter command1 = new OleDbDataAdapter(sql, conn1);
            conn1.Open();
            DataSet dataSet = new DataSet();
            command1.Fill(dataSet);

            DataTable dt = dataSet.Tables[0];
            string File_Batch = "";

            foreach (DataRow dr in dt.Rows)
            {
                string temp = dr[0].ToString().Substring(0,dr[0].ToString().Length -4);
                temp = temp.ToUpper();

                if (File_Batch == "")
                {
                    File_Batch = temp;
                }
                else
                {
                    File_Batch = File_Batch + "_" + temp;
                }
            }
            conn1.Close();

            if (File_Batch == "") { File_Batch = Batch; }
            //End For Batch








            //Zip File
            System.Diagnostics.Process proc = new System.Diagnostics.Process();
            proc.EnableRaisingEvents = false;
            proc.StartInfo.FileName = "\"" + WinZipLocation.Replace("\\","\\\\") + "\""  ;
            proc.StartInfo.Arguments = " -u -r -p " + "\"" + Application.StartupPath + "\\AFT" + "_" + File_Batch + "_" + ChequeName_all + "_Process.by_" + ProcessBy + "__Checked.By_" + CheckedBy + ".zip" + "\"" + " C:\\Windows\\Temp\\" + DateTimeToday + "\\*.*";
            proc.Start();
            proc.WaitForExit();
            //End Zip File







            //Copy the Zip File
            

            CreateDirectory (Temp_Zip);

            if (File.Exists(Temp_Zip + "\\AFT" + "_" + File_Batch + "_" + ChequeName_all + "_Process.by_" + ProcessBy + "__Checked.By_" + CheckedBy + ".zip") == true)
            {
                File.Delete(Temp_Zip + "\\AFT" + "_" + File_Batch + "_" + ChequeName_all + "_Process.by_" + ProcessBy + "__Checked.By_" + CheckedBy + ".zip");
            }
            File.Copy(Application.StartupPath + "\\AFT" + "_" + File_Batch + "_" + ChequeName_all + "_Process.by_" + ProcessBy + "__Checked.By_" + CheckedBy + ".zip", Temp_Zip + "\\AFT" + "_" + File_Batch + "_" + ChequeName_all + "_Process.by_" + ProcessBy + "__Checked.By_" + CheckedBy + ".zip");
            //End Copy the Zip File
        }
        public static void PackingList(string ChkType, string FolderName, string FormType, string RefChkType)
        {
            int PageNo = 0;

            StreamWriter sw = new StreamWriter(Application.StartupPath + "\\" + FolderName + "\\Packing"+RefChkType+".txt");

            Boolean Print_Front_Cover = false;

        RepeatMe:

            
            string sql = "SELECT RT_NO, Branch, Sum(No_Bks), BatchNo FROM Packing WHERE ChkType = '" + RefChkType + "' GROUP BY RT_NO, Branch, BatchNo ORDER BY BatchNo, RT_NO";
            OleDbConnection conn1 = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + Application.StartupPath + "\\"+FolderName + ";Extended Properties=dBASE III;");
            OleDbDataAdapter command1 = new OleDbDataAdapter(sql, conn1);
            conn1.Open();
            DataSet dataSet = new DataSet();
            command1.Fill(dataSet);


            int LoopCount = 0;
            DataTable dt = dataSet.Tables[0];
            foreach (DataRow dr in dt.Rows)
            {
                string BRSTN  = dr[0].ToString();
                string BranchName = dr[1].ToString();
                string OrderQty = dr[2].ToString();
                string Batch = dr[3].ToString();


                //For Heading
                PageNo = PageNo +1;

                if (LoopCount != 0) {sw.WriteLine("");}

                sw.WriteLine("");
                sw.WriteLine("  Page No. " +PageNo);
                sw.WriteLine("  " + DateTime.Now.ToString(" MMMM dd, yyyy"));
                sw.WriteLine("                                CAPTIVE PRINTING CORPORATION");
    
                if (ChkType == "A" && FormType == "05") {sw.WriteLine("                               SBTC - Personal Checks Summary");}
                if (ChkType == "B" && FormType == "16") {sw.WriteLine("                               SBTC - Commercial Checks Summary");}
    
                if (ChkType == "AA" && FormType == "05") {sw.WriteLine("                               SBTC - Personal PreEncoded Checks Summary");}
                if (ChkType == "BB" && FormType == "16") {sw.WriteLine("                               SBTC - Commercial PreEncoded Checks Summary");}
    
                if (ChkType == "MC" && FormType == "20") {sw.WriteLine("                               SBTC - Manager's Checks Summary");}
    
                if (ChkType == "F" && FormType == "25") {sw.WriteLine("                               SBTC - Personal CheckOne Summary");}
                if (ChkType == "F" && FormType == "26") {sw.WriteLine("                               SBTC - Commercial CheckOne Summary");}
    
                if (ChkType == "E" && FormType == "23") {sw.WriteLine("                               SBTC - Personal CheckPower Summary");}
                if (ChkType == "E" && FormType == "22") {sw.WriteLine("                               SBTC - Commercial CheckPower Summary");}
    
                if (ChkType == "GC" && FormType == "20") {sw.WriteLine("                               SBTC - Gift Checks Summary");}
                if (ChkType == "MC_1" && FormType == "00") {sw.WriteLine("                               SBTC - Manager's Checks Continues Summary");}
                if (ChkType == "CUSTOM" && FormType == "00") {sw.WriteLine("                               SBTC - Customized Checks Summary");}
    
                if (ChkType == "CS") {sw.WriteLine("                               SBTC - Charge Slip Checks Summary");}
    
                if (Print_Front_Cover == true)
                {
                    sw.WriteLine("                                 ( F R O N T  C O V E R )");
                }

                sw.WriteLine("");
                sw.WriteLine("  ACCT_NO         ACCOUNT NAME                     QTY CT START #    END #");
                sw.WriteLine("");
                sw.WriteLine("");
                sw.WriteLine(" ** ORDERS OF BRSTN " + BRSTN + " " + BranchName);
                sw.WriteLine("");
                sw.WriteLine(" * Batch #: " + Batch);
                //End For Heading




                sql = "SELECT Acct_No_P, Acct_Name1, Acct_Name2, Ck_No_B, CK_NO_E FROM Packing WHERE CHkTYpe = '" + RefChkType + "' AND RT_NO = '" + BRSTN + "' AND Branch = '" + BranchName.Replace("'", "''") + "' AND BatchNo = '" + Batch + "' ORDER BY Acct_No_P, CK_NO_B";
                
                OleDbConnection conn2 = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + Application.StartupPath + "\\"+FolderName + ";Extended Properties=dBASE III;");
                OleDbDataAdapter command2 = new OleDbDataAdapter(sql, conn2);
                conn2.Open();
                DataSet dataSet2 = new DataSet();
                command2.Fill(dataSet2);

                int LoopCount1 = 0;
                DataTable dt2 = dataSet2.Tables[0];
                foreach (DataRow dr2 in dt2.Rows)
                {
                    //string BRSTN  = dr2[0].ToString();
                    string AccountNo = dr2[0].ToString();
                    string Name1 = dr2[1].ToString();
                    string Name2 = dr2[2].ToString();
                        
                    if (Print_Front_Cover == true)
                    {
                        Name1 = "";
                        Name2 = "";
                    }
        
                    string StartingSerial = dr2[3].ToString();
                    string EndingSerial = dr2[4].ToString();

                    while (Name1.Length < 35)
                    {
                        Name1 = Name1 + " ";
                    }

                    while (StartingSerial.Length < 11)
                    {
                        StartingSerial = StartingSerial + " ";
                    }
                    
                    sw.WriteLine("  " + AccountNo + "  " + Name1 + "1 " + ChkType.Replace("MC_1", "B").Replace("CUSTOM", "B") + "  " + StartingSerial + EndingSerial);
                    if (Name2 != "") {sw.WriteLine("                  " + Name2);}

                    LoopCount1 = LoopCount1 +1;
                }
                conn2.Close();

                sw.WriteLine("");
                sw.WriteLine("");
                sw.WriteLine(" * * * Sub Total * * *                              " + OrderQty);


                LoopCount = LoopCount +1;
            }

            if (Print_Front_Cover == false)
            {
                sw.WriteLine("");
    
                Print_Front_Cover = true;

                goto RepeatMe;
            }
            sw.Close();
            conn1.Close();

        }
        public static void CopyBatchToFinalBatch(string FinalBatch)
        {
            string sql = "UPDATE SBTC SET Batch = '" + FinalBatch + "'";
            OleDbConnection conn1 = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + Application.StartupPath + ";Extended Properties=dBASE III;");
            OleDbDataAdapter command1 = new OleDbDataAdapter(sql, conn1);
            conn1.Open();
            DataSet dataSet = new DataSet();
            command1.Fill(dataSet);
            conn1.Close();
        }
        public static int ProcessMe(string ChkType, string FormType, string FolderName, Boolean DeleteDBF_Value, string FinalBatch, DateTime DeliveryDate, string ProcessBy, string DateTimeToday)
        {
            int PcsPerBook = 0;
            string ChkType2 = "";
            string Description = "";
            string FormatSerial = "";
            string MICRLine = "";
            string Ref_Location = "";
            string RefChkType = "";
            string FileName = "";
            string ChkType_1 = "";
            string ChkType_2 = "";
            string FormType_1 = "";
            string FormType_2 = "";
            string ChkType_31 = "";
            string ChkType_32 = "";
            string Temp_DriveR = "";
            string Temp_CTC = "";



            if (ChkType == "A" && FormType == "05")
            {
                PcsPerBook = 50;
                ChkType2 = "P";
                Description = "PERSONAL";
                FormatSerial = "0000000";
                MICRLine = "     ONNNNNNNO";
                Ref_Location = Application.StartupPath + "\\Regular";
                RefChkType = "A";
                FileName = FinalBatch.Substring(0, 4) + "_P12" + FinalBatch.Substring(8, FinalBatch.Length - 8);

                ChkType_1 = "A";
                ChkType_2 = "B";
                FormType_1 = "05";
                FormType_2 = "16";
                ChkType_31 = "PA";
                ChkType_32 = "CA";

                if (ReturnMe.CodesOnly == true)
                {
                    Temp_DriveR = PrinterFiles_Folder + "\\Codes\\SBTC\\" + DateTime.Now.ToString("yyyy");
                    Temp_CTC = Resting_Folder + "\\CTC\\Codes\\SBTC\\" + DateTime.Now.ToString("yyyy");
                }

                if (ReturnMe.CodesOnly == false)
                {
                    Temp_DriveR = PrinterFiles_Folder + "\\SBTC\\" + DateTime.Now.ToString("yyyy");
                    Temp_CTC = Resting_Folder + "\\CTC\\SBTC\\" + DateTime.Now.ToString("yyyy");
                }
            }



            if (ChkType == "B" && FormType == "16")
            {
                PcsPerBook = 100;
                ChkType2 = "C";
                Description = "COMMERCIAL";
                FormatSerial = "0000000000";
                MICRLine = "  ONNNNNNNNNNO";
                Ref_Location = Application.StartupPath + "\\Regular";
                RefChkType = "B";
                FileName = FinalBatch.Substring(0, 4) + "_C12" + FinalBatch.Substring(8, FinalBatch.Length - 8);

                ChkType_1 = "A";
                ChkType_2 = "B";
                FormType_1 = "05";
                FormType_2 = "16";
                ChkType_31 = "PA";
                ChkType_32 = "CA";

                if (ReturnMe.CodesOnly == true)
                {
                    Temp_DriveR = PrinterFiles_Folder + "\\Codes\\SBTC\\" + DateTime.Now.ToString("yyyy");
                    Temp_CTC = Resting_Folder + "\\CTC\\Codes\\SBTC\\" + DateTime.Now.ToString("yyyy");
                }

                if (ReturnMe.CodesOnly == false)
                {
                    Temp_DriveR = PrinterFiles_Folder + "\\SBTC\\" + DateTime.Now.ToString("yyyy");
                    Temp_CTC = Resting_Folder + "\\CTC\\SBTC\\" + DateTime.Now.ToString("yyyy");
                }
            }


            if (ChkType == "AA" && FormType == "05")
            {
                PcsPerBook = 50;
                ChkType2 = "P";
                Description = "PERSONAL";
                FormatSerial = "0000000";
                MICRLine = "     ONNNNNNNO";
                Ref_Location = Application.StartupPath + "\\Regular";
                RefChkType = "A";

                FileName = "SB" + FinalBatch.Substring(0, 4) + "P" + FinalBatch.Substring(8, FinalBatch.Length - 8);

                ChkType_1 = "AA";
                ChkType_2 = "BB";
                FormType_1 = "05";
                FormType_2 = "16";
                ChkType_31 = "PA";
                ChkType_32 = "CA";

                if (ReturnMe.CodesOnly == true)
                {
                    Temp_DriveR = PrinterFiles_Folder + "\\Codes\\SBTC\\" + DateTime.Now.ToString("yyyy");
                    Temp_CTC = Resting_Folder + "\\CTC\\Codes\\SBTC\\" + DateTime.Now.ToString("yyyy");
                }

                if (ReturnMe.CodesOnly == false)
                {
                    Temp_DriveR = PrinterFiles_Folder + "\\SBTC\\" + DateTime.Now.ToString("yyyy");
                    Temp_CTC = Resting_Folder + "\\CTC\\SBTC\\" + DateTime.Now.ToString("yyyy");
                }
            }

            if (ChkType == "BB" && FormType == "16")
            {
                PcsPerBook = 100;
                ChkType2 = "C";
                Description = "COMMERCIAL";
                FormatSerial = "0000000000";
                MICRLine = "  ONNNNNNNNNNO";
                Ref_Location = Application.StartupPath + "\\Regular";
                RefChkType = "B";
                FileName = "SB" + FinalBatch.Substring(0, 4) + "C" + FinalBatch.Substring(8, FinalBatch.Length - 8);

                ChkType_1 = "AA";
                ChkType_2 = "BB";
                FormType_1 = "05";
                FormType_2 = "16";
                ChkType_31 = "PA";
                ChkType_32 = "CA";

                if (ReturnMe.CodesOnly == true)
                {
                    Temp_DriveR = PrinterFiles_Folder + "\\Codes\\SBTC\\" + DateTime.Now.ToString("yyyy");
                    Temp_CTC = Resting_Folder + "\\CTC\\Codes\\SBTC\\" + DateTime.Now.ToString("yyyy");
                }

                if (ReturnMe.CodesOnly == false)
                {
                    Temp_DriveR = PrinterFiles_Folder + "\\SBTC\\" + DateTime.Now.ToString("yyyy");
                    Temp_CTC = Resting_Folder + "\\CTC\\SBTC\\" + DateTime.Now.ToString("yyyy");
                }
            }

            if (ChkType == "MC" && FormType == "20")
            {
                PcsPerBook = 50;
                ChkType2 = "P";
                Description = "MANAGER'S CHECK";
                FormatSerial = "0000000000";
                MICRLine = "  ONNNNNNNNNNO";
                Ref_Location = Application.StartupPath + "\\MC";
                RefChkType = "A";
                FileName = "MC" + FinalBatch.Substring(0, 4) + "P" + FinalBatch.Substring(8, FinalBatch.Length - 8);

                ChkType_1 = "MC";
                ChkType_2 = "MC";
                FormType_1 = "20";
                FormType_2 = "20";
                ChkType_31 = "MC";
                ChkType_32 = "MC";

                if (ReturnMe.CodesOnly == true)
                {
                    Temp_DriveR = PrinterFiles_Folder + "\\Codes\\SBTC\\MC\\" + DateTime.Now.ToString("yyyy");
                    Temp_CTC = Resting_Folder + "\\CTC\\Codes\\SBTC\\MC\\" + DateTime.Now.ToString("yyyy");
                }

                if (ReturnMe.CodesOnly == false)
                {
                    Temp_DriveR = PrinterFiles_Folder + "\\SBTC\\MC\\" + DateTime.Now.ToString("yyyy");
                    Temp_CTC = Resting_Folder + "\\CTC\\SBTC\\MC\\" + DateTime.Now.ToString("yyyy");
                }
            }


            if (ChkType == "MC_1" && FormType == "00")
            {
                PcsPerBook = 100;
                ChkType2 = "C";
                Description = "MANAGER'S CHECK CONTINUES";
                FormatSerial = "0000000000";
                MICRLine = "  ONNNNNNNNNNO";
                Ref_Location = Application.StartupPath + "\\MC\\Continues";
                RefChkType = "B";

                if (FinalBatch.Substring(0, 2) == "MC")
                {
                    FileName = "MCC" + FinalBatch.Substring(2, FinalBatch.Length - 2);
                }
                else
                {
                    FileName = "MCC" + FinalBatch.Substring(0, 4) + FinalBatch.Substring(8, FinalBatch.Length - 8);
                }

                ChkType_1 = "MC_1";
                ChkType_2 = "MC_1";
                FormType_1 = "00";
                FormType_2 = "00";
                ChkType_31 = "MC_1";
                ChkType_32 = "MC_1";

                if (ReturnMe.CodesOnly == true)
                {
                    Temp_DriveR = PrinterFiles_Folder + "\\Codes\\SBTC\\MC\\CONTINUOUS\\" + DateTime.Now.ToString("yyyy");
                    Temp_CTC = Resting_Folder + "\\CTC\\Codes\\SBTC\\CONTINUOUS_MC\\" + DateTime.Now.ToString("yyyy");
                }

                if (ReturnMe.CodesOnly == false)
                {
                    Temp_DriveR = PrinterFiles_Folder + "\\SBTC\\MC\\CONTINUOUS\\" + DateTime.Now.ToString("yyyy");
                    Temp_CTC = Resting_Folder + "\\CTC\\SBTC\\CONTINUOUS_MC\\" + DateTime.Now.ToString("yyyy");
                }
            }

            if (ChkType == "CUSTOM_PA" && FormType == "00")
            {
                PcsPerBook = 50;
                ChkType2 = "P";
                Description = "CUSTOMIZED CHECKS";
                FormatSerial = "0000000000";
                MICRLine = "  ONNNNNNNNNNO";
                Ref_Location = Application.StartupPath + "\\Customized";
                RefChkType = "A";

                FileName = "CUS" + FinalBatch.Substring(0, 4) + "P" + FinalBatch.Substring(8, FinalBatch.Length - 8);


                ChkType_1 = "CUSTOM_PA";
                ChkType_2 = "CUSTOM_PA";
                FormType_1 = "00";
                FormType_2 = "00";
                ChkType_31 = "CUSTOM_PA";
                ChkType_32 = "CUSTOM_PA";

                if (ReturnMe.CodesOnly == true)
                {
                    Temp_DriveR = PrinterFiles_Folder + "\\Codes\\SBTC\\CUSTOM\\" + DateTime.Now.ToString("yyyy");
                    Temp_CTC = Resting_Folder + "\\CTC\\Codes\\SBTC\\Customized\\" + DateTime.Now.ToString("yyyy");
                }

                if (ReturnMe.CodesOnly == false)
                {
                    Temp_DriveR = PrinterFiles_Folder + "\\SBTC\\CUSTOM\\" + DateTime.Now.ToString("yyyy");
                    Temp_CTC = Resting_Folder + "\\CTC\\SBTC\\Customized\\" + DateTime.Now.ToString("yyyy");
                }
            }

            if (ChkType == "CUSTOM" && FormType == "00")
            {
                PcsPerBook = 100;
                ChkType2 = "C";
                Description = "CUSTOMIZED CHECKS";
                FormatSerial = "0000000000";
                MICRLine = "  ONNNNNNNNNNO";
                Ref_Location = Application.StartupPath + "\\Customized";
                RefChkType = "B";

                FileName = "CUS" + FinalBatch.Substring(0, 4) + "C" + FinalBatch.Substring(8, FinalBatch.Length - 8);


                ChkType_1 = "CUSTOM";
                ChkType_2 = "CUSTOM";
                FormType_1 = "00";
                FormType_2 = "00";
                ChkType_31 = "CUSTOM";
                ChkType_32 = "CUSTOM";

                if (ReturnMe.CodesOnly == true)
                {
                    Temp_DriveR = PrinterFiles_Folder + "\\Codes\\SBTC\\CUSTOM\\" + DateTime.Now.ToString("yyyy");
                    Temp_CTC = Resting_Folder + "\\CTC\\Codes\\SBTC\\Customized\\" + DateTime.Now.ToString("yyyy");
                }

                if (ReturnMe.CodesOnly == false)
                {
                    Temp_DriveR = PrinterFiles_Folder + "\\SBTC\\CUSTOM\\" + DateTime.Now.ToString("yyyy");
                    Temp_CTC = Resting_Folder + "\\CTC\\SBTC\\Customized\\" + DateTime.Now.ToString("yyyy");
                }
            }

            if (ChkType == "F" && FormType == "25")
            {
                PcsPerBook = 25;
                ChkType2 = "P";
                Description = "PERSONAL CHECKONE";
                FormatSerial = "0000000";
                MICRLine = "     ONNNNNNNO";
                Ref_Location = Application.StartupPath + "\\CheckOne";
                RefChkType = "A";
                FileName = "13D" + FinalBatch.Substring(0, 4) + "P" + FinalBatch.Substring(8, FinalBatch.Length - 8);

                ChkType_1 = "F";
                ChkType_2 = "F";
                FormType_1 = "25";
                FormType_2 = "26";
                ChkType_31 = "PA";
                ChkType_32 = "CA";

                if (ReturnMe.CodesOnly == true)
                {
                    Temp_DriveR = PrinterFiles_Folder + "\\Codes\\SBTC\\CheckOne\\" + DateTime.Now.ToString("yyyy");
                    Temp_CTC = Resting_Folder + "\\CTC\\Codes\\SBTC\\CheckOne\\" + DateTime.Now.ToString("yyyy");
                }


                if (ReturnMe.CodesOnly == false)
                {
                    Temp_DriveR = PrinterFiles_Folder + "\\SBTC\\CheckOne\\" + DateTime.Now.ToString("yyyy");
                    Temp_CTC = Resting_Folder + "\\CTC\\SBTC\\CheckOne\\" + DateTime.Now.ToString("yyyy");
                }
            }

            if (ChkType == "F" && FormType == "26")
            {
                PcsPerBook = 50;
                ChkType2 = "C";
                Description = "COMMERCIAL CHECKONE";
                FormatSerial = "0000000000";
                MICRLine = "  ONNNNNNNNNNO";
                Ref_Location = Application.StartupPath + "\\CheckOne";
                RefChkType = "B";
                FileName = "13D" + FinalBatch.Substring(0, 4) + "C" + FinalBatch.Substring(8, FinalBatch.Length - 8);

                ChkType_1 = "F";
                ChkType_2 = "F";
                FormType_1 = "25";
                FormType_2 = "26";
                ChkType_31 = "PA";
                ChkType_32 = "CA";

                if (ReturnMe.CodesOnly == true)
                {
                    Temp_DriveR = PrinterFiles_Folder + "\\Codes\\SBTC\\CheckOne\\" + DateTime.Now.ToString("yyyy");
                    Temp_CTC = Resting_Folder + "\\CTC\\Codes\\SBTC\\CheckOne\\" + DateTime.Now.ToString("yyyy");
                }

                if (ReturnMe.CodesOnly == false)
                {
                    Temp_DriveR = PrinterFiles_Folder + "\\SBTC\\CheckOne\\" + DateTime.Now.ToString("yyyy");
                    Temp_CTC = Resting_Folder + "\\CTC\\SBTC\\CheckOne\\" + DateTime.Now.ToString("yyyy");
                }
            }

            if (ChkType == "E" && FormType == "23")
            {
                PcsPerBook = 50;
                ChkType2 = "P";
                Description = "PERSONAL CHECKPOWER";
                FormatSerial = "0000000";
                MICRLine = "     ONNNNNNNO";
                Ref_Location = Application.StartupPath + "\\CheckPower";
                RefChkType = "A";
                FileName = "CKP" + FinalBatch.Substring(0, 4) + "P" + FinalBatch.Substring(8, FinalBatch.Length - 8);

                ChkType_1 = "E";
                ChkType_2 = "E";
                FormType_1 = "23";
                FormType_2 = "22";
                ChkType_31 = "PA";
                ChkType_32 = "CA";

                if (ReturnMe.CodesOnly == true)
                {
                    Temp_DriveR = PrinterFiles_Folder + "\\Codes\\SBTC\\CheckPower\\" + DateTime.Now.ToString("yyyy");
                    Temp_CTC = Resting_Folder + "\\CTC\\Codes\\SBTC\\CKPOWER\\" + DateTime.Now.ToString("yyyy");
                }

                if (ReturnMe.CodesOnly == false)
                {
                    Temp_DriveR = PrinterFiles_Folder + "\\SBTC\\CheckPower\\" + DateTime.Now.ToString("yyyy");
                    Temp_CTC = Resting_Folder + "\\CTC\\SBTC\\CKPOWER\\" + DateTime.Now.ToString("yyyy");
                }
            }

            if (ChkType == "E" && FormType == "22")
            {
                PcsPerBook = 100;
                ChkType2 = "C";
                Description = "COMMERCIAL CHECKPOWER";
                FormatSerial = "0000000000";
                MICRLine = "  ONNNNNNNNNNO";
                Ref_Location = Application.StartupPath + "\\CheckPower";
                RefChkType = "B";
                FileName = "CKP" + FinalBatch.Substring(0, 4) + "C" + FinalBatch.Substring(8, FinalBatch.Length - 8);

                ChkType_1 = "E";
                ChkType_2 = "E";
                FormType_1 = "23";
                FormType_2 = "22";
                ChkType_31 = "PA";
                ChkType_32 = "CA";

                if (ReturnMe.CodesOnly == true)
                {
                    Temp_DriveR = PrinterFiles_Folder + "\\Codes\\SBTC\\CheckPower\\" + DateTime.Now.ToString("yyyy");
                    Temp_CTC = Resting_Folder + "\\CTC\\Codes\\SBTC\\CKPOWER\\" + DateTime.Now.ToString("yyyy");
                }

                if (ReturnMe.CodesOnly == false)
                {
                    Temp_DriveR = PrinterFiles_Folder + "\\SBTC\\CheckPower\\" + DateTime.Now.ToString("yyyy");
                    Temp_CTC = Resting_Folder + "\\CTC\\SBTC\\CKPOWER\\" + DateTime.Now.ToString("yyyy");
                }
            }

            if (ChkType == "GC" && FormType == "20")
            {
                PcsPerBook = 50;
                ChkType2 = "P";
                Description = "GIFT CHECK";
                FormatSerial = "000000";
                MICRLine = "  O0000NNNNNNO";
                Ref_Location = Application.StartupPath + "\\GiftCheck";
                RefChkType = "A";
                FileName = "GC" + FinalBatch.Substring(0, 4) + "P" + FinalBatch.Substring(8, FinalBatch.Length - 8);

                ChkType_1 = "GC";
                ChkType_2 = "GC";
                FormType_1 = "20";
                FormType_2 = "20";
                ChkType_31 = "GC";
                ChkType_32 = "GC";

                if (ReturnMe.CodesOnly == true)
                {
                    Temp_DriveR = PrinterFiles_Folder + "\\Codes\\SBTC\\GiftCheck\\" + DateTime.Now.ToString("yyyy");
                    Temp_CTC = Resting_Folder + "\\CTC\\Codes\\SBTC\\GC\\" + DateTime.Now.ToString("yyyy");
                }

                if (ReturnMe.CodesOnly == false)
                {
                    Temp_DriveR = PrinterFiles_Folder + "\\SBTC\\GiftCheck\\" + DateTime.Now.ToString("yyyy");
                    Temp_CTC = Resting_Folder + "\\CTC\\SBTC\\GC\\" + DateTime.Now.ToString("yyyy");
                }
            }

            if (ChkType == "CS")
            {
                PcsPerBook = 50;
                ChkType2 = "P";
                Description = "CHARGE SLIP";
                FormatSerial = "0000000000";
                MICRLine = "  O0000NNNNNNO";
                Ref_Location = Application.StartupPath + "\\Charge_Slip";
                RefChkType = "A";
                FileName = "CS" + FinalBatch.Substring(0, 4) + "P" + FinalBatch.Substring(8, FinalBatch.Length - 8);

                ChkType_1 = "CS";
                ChkType_2 = "CS";
                FormType_1 = "00";
                FormType_2 = "00";
                ChkType_31 = "CS";
                ChkType_32 = "CS";

                if (ReturnMe.CodesOnly == true)
                {
                    Temp_DriveR = PrinterFiles_Folder + "\\Codes\\SBTC\\Charge_Slip\\" + DateTime.Now.ToString("yyyy");
                    Temp_CTC = Resting_Folder + "\\CTC\\Codes\\SBTC\\Charge_Slip\\" + DateTime.Now.ToString("yyyy");
                }

                if (ReturnMe.CodesOnly == false)
                {
                    Temp_DriveR = PrinterFiles_Folder + "\\SBTC\\Charge_Slip\\" + DateTime.Now.ToString("yyyy");
                    Temp_CTC = Resting_Folder + "\\CTC\\SBTC\\Charge_Slip\\" + DateTime.Now.ToString("yyyy");
                }
            }

            int BlockCount = 0;
            int TotalData = 0;
            int DataNumber = 0;


            StreamWriter DoBlock = new StreamWriter(Application.StartupPath + "\\" + FolderName + "\\Block" + ChkType2 + ".txt");
            StreamWriter PrinterFile;
            PrinterFile = new StreamWriter("C:\\Windows\\Temp\\" + DateTimeToday + FolderName.Replace("\\","") + ChkType2 + ".txt");

            if (ChkType == "CUSTOM" || ChkType == "CUSTOM_PA")
            {
                CopyBatchToFinalBatch(FinalBatch);
            }


            if (DeleteDBF_Value == true)
            {
                DeletePackDBF("Packing.dbf", Application.StartupPath + "\\" + FolderName);
            }


            string sql = "SELECT BRSTN, AccountNo, OrderQty, Name1, Name2, Address1, Address2, Address3, Address4, Address5, Address6, Batch, BStock, StartSN, PcsPerBook , PKey FROM SBTC WHERE ChkType = '" + ChkType + "' AND FormType = '" + FormType + "' ORDER BY BRSTN, AccountNo, Name1";

            string DayOfWeekResult = "";
            string temp = "";
            string Summary_DoBlock = "";
            int LoopCount = 0;
            string Sum_DoBlock = "";

            OleDbConnection conn1 = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + Application.StartupPath + ";Extended Properties=dBASE III;");
            OleDbDataAdapter command1 = new OleDbDataAdapter(sql, conn1);
            conn1.Open();
            DataSet dataSet = new DataSet();
            command1.Fill(dataSet);

            DataTable dt = dataSet.Tables[0];

            foreach (DataRow dr in dt.Rows)
            {
                if (LoopCount == 0)
                {
                    PrinterFile.Close();
                    PrinterFile = new StreamWriter(Application.StartupPath + "\\" + FolderName + "\\" + FileName + ".txt");
                }

                string BRSTN = dr[0].ToString();
                string AccountNo = dr[1].ToString();

                int OrderQty = 0;
                if (Int32.TryParse(dr[2].ToString(), out OrderQty))
                {
                }
                else { OrderQty = 0; }

                string Name1 = dr[3].ToString();
                string Name2 = dr[4].ToString();
                string Address1 = dr[5].ToString();
                string Address2 = dr[6].ToString();
                string Address3 = dr[7].ToString();
                string Address4 = dr[8].ToString();
                string Address5 = dr[9].ToString();
                string Address6 = dr[10].ToString();
                string Batch = dr[11].ToString();
                string BStock = dr[12].ToString();
                Int64 StartingSerial = 0;
                if (Int64.TryParse(dr[13].ToString(), out StartingSerial))
                {

                }
                else { StartingSerial = 0; }

                string PKey = dr[15].ToString();

                if (LoopCount == 0)
                {
                    //Copy Printer File MDB
                    if ((ChkType == "F" && FormType == "25") || (ChkType == "F" && FormType == "26") || (ChkType == "GC" && FormType == "20") || (ChkType == "MC" && FormType == "20") || (ChkType == "MC_1" && FormType == "00") || (ChkType == "CUSTOM" && FormType == "00") || (ChkType == "CUSTOM_PA" && FormType == "00") || ChkType == "CS")
                    {
                        File.Copy(Application.StartupPath + "\\DataSource.mdb", Application.StartupPath + "\\" + FolderName + "\\" + FileName + ".mdb");
                    }
                    //End Copy Printer File MDB
                }

                //For Total
                if (TotalData == 0)
                {
                    DayOfWeekResult = DeliveryDate.DayOfWeek.ToString().Substring(0, 3).ToUpper();

                    temp = Batch;
                    while (temp.Length < 45)
                    {
                        temp = temp + " ";
                    }

                    Summary_DoBlock = "    " + temp + "DLVR: " + DeliveryDate.ToString("MM-dd") + "(" + DayOfWeekResult + ")" + "\n\n";



                    sql = "SELECT SUM(OrderQty) FROM SBTC WHERE ChkType = '" + ChkType_1 + "' AND FormType = '" + FormType_1 + "'";
                    OleDbConnection conn2 = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + Application.StartupPath + ";Extended Properties=dBASE III;");
                    OleDbDataAdapter command2 = new OleDbDataAdapter(sql, conn2);
                    conn2.Open();
                    DataSet dataSet2 = new DataSet();
                    command2.Fill(dataSet2);

                    int sum_orderQty = 0;
                    DataTable dt2 = dataSet2.Tables[0];
                    foreach (DataRow dr2 in dt2.Rows)
                    {
                        //sum_orderQty = dr2[0].ToString();
                        if (Int32.TryParse(dr2[0].ToString(), out sum_orderQty))
                        { }
                        else { sum_orderQty = 0; }
                    }

                    conn2.Close();

                    Sum_DoBlock = "    " + ChkType_31 + " = " + sum_orderQty + "                 " + FileName + ".txt";

                    if (ChkType_31 != ChkType_32)
                    {
                        sql = "SELECT SUM(OrderQty) FROM SBTC WHERE ChkType = '" + ChkType_2 + "' AND FormType = '" + FormType_2 + "'";
                        OleDbConnection conn3 = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + Application.StartupPath + ";Extended Properties=dBASE III;");
                        OleDbDataAdapter command3 = new OleDbDataAdapter(sql, conn3);
                        conn3.Open();
                        DataSet dataSet3 = new DataSet();
                        command3.Fill(dataSet3);

                        DataTable dt3 = dataSet3.Tables[0];
                        foreach (DataRow dr3 in dt3.Rows)
                        {
                            //sum_orderQty = dr3[0].ToString();
                            if (Int32.TryParse(dr3[0].ToString(), out sum_orderQty))
                            { }
                            else { sum_orderQty = 0; }
                        }
                        conn3.Close();

                        Sum_DoBlock = Sum_DoBlock + "\r\n" + "    " + ChkType_32 + " = " + sum_orderQty;
                    }

                    Summary_DoBlock = Summary_DoBlock + "\r\n" + Sum_DoBlock;

                    Summary_DoBlock = Summary_DoBlock + "\r\n" + "\r\n" + "    Prepared By : " + ProcessBy + "\r\n" + "    Updated By  : " + ProcessBy + "\r\n" + "    Time Start  : " + TimeStart + "\r\n" + "    Time Fnished:                                              RECHECKED BY:  " + "\r\n" + "    File Rcvd   :";
                }
                //End For Total

                if (FolderName == "MC" || FolderName == "GiftCheck") { Name1 = ""; }

                //For Ref.dbf Starting SN
                string BranchName = "";
                if (ChkType == "CS") { BranchName = "CHARGE SLIP"; }

                if (ChkType != "CS")
                {
                    if ((ChkType == "MC_1" && FormType == "00") || (ChkType == "CUSTOM" && FormType == "00") || (ChkType == "CUSTOM_PA" && FormType == "00"))
                    {


                        OleDbConnection conn = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + Application.StartupPath + "\\" + FolderName + ";Extended Properties=dBASE III;");
                        OleDbCommand command = new OleDbCommand("INSERT INTO Master ([Date], Batch, BRSTN, AccountNo, StartSN, EndSN, OrderQty, Address1, Address2, Address3, Address4, Address5, Address6, Name1, Name2) VALUES ('" + DateTime.Now.ToString("yyyy-MM-dd") + "','" + Batch + "','" + BRSTN + "','" + AccountNo + "','" + StartingSerial + "','" + (StartingSerial + (OrderQty * PcsPerBook) - 1) + "','" + OrderQty + "','" + Address1.Replace("'", "''") + "','" + Address2.Replace("'", "''") + "','" + Address3.Replace("'", "''") + "','" + Address4.Replace("'", "''") + "','" + Address5.Replace("'", "''") + "','" + Address6.Replace("'", "''") + "','" + Name1.Replace("'", "''") + "','" + Name2.Replace("'", "''") + "')", conn);
                        conn.Open();
                        command.ExecuteReader();

                        conn.Close();

                        BranchName = Address1;
                    }
                    else
                    {

                        sql = "SELECT LastNo,C_Before, Branch_Tex FROM REF WHERE RTNO = '" + BRSTN + "' AND ChkType = '" + RefChkType + "'";
                        OleDbConnection conn3 = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + Ref_Location + ";Extended Properties=dBASE III;");
                        OleDbDataAdapter command3 = new OleDbDataAdapter(sql, conn3);
                        conn3.Open();
                        DataSet dataSet3 = new DataSet();
                        command3.Fill(dataSet3);

                        int EndingSerial = 0;
                        int C_Before = 0;
                        int NewLastNo = 0;

                        DataTable dt3 = dataSet3.Tables[0];
                        foreach (DataRow dr3 in dt3.Rows)
                        {

                            if (Int32.TryParse(dr3[0].ToString(), out EndingSerial))
                            { }
                            else { EndingSerial = 0; }

                            if (Int32.TryParse(dr3[1].ToString(), out C_Before))
                            { }
                            else { C_Before = 0; }

                            BranchName = dr3[2].ToString();
                        }
                        conn3.Close();


                        StartingSerial = EndingSerial + 1;
                        NewLastNo = (PcsPerBook * OrderQty) + EndingSerial;


                        OleDbConnection conn = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + Ref_Location + ";Extended Properties=dBASE III;");
                        OleDbCommand command = new OleDbCommand("UPDATE REF SET [Date] = '" + DateTime.Now.ToString("yyyy-MM-dd") + "', LastNo = '" + NewLastNo + "', P_Before = '" + C_Before + "', C_Before = '" + EndingSerial + "' WHERE RTNO = '" + BRSTN + "' AND ChkType = '" + RefChkType + "'", conn);
                        conn.Open();
                        command.ExecuteReader();

                        conn.Close();
                    }
                }
                //End For Ref.dbf Starting SN

                //Update SBTC StartSN1
                OleDbConnection conn4 = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + Application.StartupPath + ";Extended Properties=dBASE III;");
                OleDbCommand command4 = new OleDbCommand("UPDATE SBTC SET StartSN1 = " + StartingSerial + ",PcsPerBook = '" + PcsPerBook + "' WHERE PKey = " + PKey, conn4);
                conn4.Open();
                command4.ExecuteReader();

                conn4.Close();
                //End Update SBTC StartSN1

                while (OrderQty > 0)
                {

                    //For Do-Block
                    if (TotalData % 32 == 0)
                    {
                        if (TotalData != 0)
                        {
                            if (TotalData == 32)
                            {
                                DoBlock.WriteLine("");
                                DoBlock.WriteLine(Summary_DoBlock);
                            }

                            DoBlock.WriteLine("");
                        }

                        DoBlock.WriteLine("");
                        DoBlock.WriteLine("        Page No. " + ((TotalData / 32) + 1));
                        DoBlock.WriteLine("        " + DateTime.Now.ToString("MMM.dd, yyyy"));
                        DoBlock.WriteLine("                   SBTC - SUMMARY OF BLOCK - " + Description);
                        if (ChkType == "AA" || ChkType == "BB") { DoBlock.WriteLine("                                Pre-Encoded"); }
                        DoBlock.WriteLine("");

                        //For Heading Carbon
                        if ((ChkType == "F" && FormType == "25") || (ChkType == "F" && FormType == "26") || (ChkType == "GC" && FormType == "20") || (ChkType == "MC" && FormType == "20"))
                        {
                            DoBlock.WriteLine("            *** With Duplicate Copy ---> " + FileName + ".mdb" + " ***");
                            DoBlock.WriteLine("");
                        }

                        if (ChkType == "MC_1" && FormType == "00")
                        {
                            DoBlock.WriteLine("            *** With Triplicate Copy ---> " + FileName + ".mdb" + " ***");
                            DoBlock.WriteLine("");
                        }
                        //End For Heading Carbon

                        if (ChkType == "CUSTOM" || ChkType == "CUSTOM_PA")
                        {
                            DoBlock.WriteLine("                    A L L  M A N U A L  E N C O D E D ! ! !");
                            DoBlock.WriteLine("");
                        }
                            
                        DoBlock.WriteLine("        BLOCK RT_NO     M ACCT_NO         START_NO.  END_NO.");
                        DoBlock.WriteLine("");
                    }


                    if (TotalData % 4 == 0)
                    {
                        BlockCount = BlockCount + 1;

                        DoBlock.WriteLine("");
                        DoBlock.WriteLine("       ** BLOCK " + BlockCount);
                    }

                    temp = BlockCount.ToString();
                    while (temp.Length < 13)
                    {
                        temp = " " + temp;
                    }

                    string temp1 = StartingSerial.ToString(FormatSerial);
                    while (temp1.Length < 11)
                    {
                        temp1 = temp1 + " ";
                    }

                    DoBlock.WriteLine(temp + " " + BRSTN + "   " + AccountNo + "    " + temp1 + (StartingSerial + PcsPerBook - 1).ToString(FormatSerial));
                    //End For Do-Block







                    //For Printer File
                    if (TotalData % 4 == 0)
                    {
                        if (TotalData == 0)
                        {
                            PrinterFile.WriteLine("3");//1
                        }
                        else
                        {
                            PrinterFile.WriteLine("3"); //1
                        }
                    }
                    else
                    {
                        PrinterFile.WriteLine("3"); //1
                    }


                    PrinterFile.WriteLine(BRSTN); //2
                    PrinterFile.WriteLine(AccountNo);//3
                    PrinterFile.WriteLine((StartingSerial + PcsPerBook).ToString(FormatSerial));//4
                    PrinterFile.WriteLine("A");//5
                    PrinterFile.WriteLine(MICRLine + BRSTN.Substring(0, 5) + "D" + BRSTN.Substring(5, 4) + "T" + AccountNo + "O");//6

                    if (ChkType == "MC_1" && FormType == "00")
                    {
                        PrinterFile.WriteLine(BRSTN.Substring(0, 5) + BRSTN.Substring(5, 3) + BRSTN.Substring(8, 1));//7
                        PrinterFile.WriteLine("");//8
                    }
                    else
                    {
                        PrinterFile.WriteLine(BRSTN.Substring(0, 5));//7
                        PrinterFile.WriteLine(" " + BRSTN.Substring(5, 4));//8
                    }

                    PrinterFile.WriteLine(AccountNo.Substring(0, 3) + "-" + AccountNo.Substring(3, 6) + "-" + AccountNo.Substring(9, 3)); //9
                    PrinterFile.WriteLine(Name1); //10
                    PrinterFile.WriteLine("SN"); //11
                    PrinterFile.WriteLine(""); //12
                    PrinterFile.WriteLine(Name2); //13
                    PrinterFile.WriteLine("C"); //14
                    PrinterFile.WriteLine("XXXX"); //15
                    PrinterFile.WriteLine(""); //16



                    if (ChkType == "GC" && FormType == "20")
                    {
                        PrinterFile.WriteLine(Address1.Replace("BRANCH", "")); //17
                    }
                    else
                    {
                        PrinterFile.WriteLine(Address1); //17
                    }



                    PrinterFile.WriteLine(Address2); //18
                    PrinterFile.WriteLine(Address3); //19
                    PrinterFile.WriteLine(Address4); //20
                    PrinterFile.WriteLine(Address5); //21
                    PrinterFile.WriteLine(Address6); //22
                    PrinterFile.WriteLine("SECURITY BANK"); //23
                    PrinterFile.WriteLine(""); //24
                    PrinterFile.WriteLine(""); //25
                    PrinterFile.WriteLine(""); //26
                    PrinterFile.WriteLine(""); //27
                    PrinterFile.WriteLine(""); //28
                    PrinterFile.WriteLine(""); //29
                    PrinterFile.WriteLine(""); //30
                    PrinterFile.WriteLine(StartingSerial.ToString(FormatSerial));//31
                    PrinterFile.WriteLine((StartingSerial + PcsPerBook - 1).ToString(FormatSerial));//32
                    //End For Printer File

                    //Save to Master_Database
                    string dbase = "";
                    if (CodesOnly == true) { dbase = "captive_database.Master_Database_SBTC_Temp"; }
                    if (CodesOnly == false) { dbase = "captive_database.Master_Database_SBTC"; }

                   // sql = "INSERT INTO " + dbase + " (Date , Time , DeliveryDate , ChkType , ChequeName , BRSTN , AccountNo , Name1 , Name2 , StartingSerial , EndingSerial , Batch , Address1 , Address2 , Address3 , Address4 , Address5 , Address6 , FinalBatch) VALUES ('" + ReturnMe.DateTimeToday_date.ToString("yyyy-MM-dd") + "','" + ReturnMe.DateTimeToday_date.ToString("HH:mm:ss") + "','" + DeliveryDate.ToString("yyyy-MM-dd") + "','" + ChkType + "','" + ReturnMe.getChequeName(ChkType, FormType).Replace("'", "''") + "','" + BRSTN + "','" + AccountNo + "','" + Name1.Replace("'", "''") + "','" + Name2.Replace("'", "''") + "','" + StartingSerial.ToString(FormatSerial) + "','" + (StartingSerial + PcsPerBook - 1).ToString(FormatSerial) + "','" + Batch + "','" + Address1.Replace("'", "''") + "','" + Address2.Replace("'", "''") + "','" + Address3.Replace("'", "''") + "','" + Address4.Replace("'", "''") + "','" + Address5.Replace("'", "''") + "','" + Address6.Replace("'", "''") + "','" + FinalBatch + "')";
                    //string MyConnection2 = "datasource=" + ReturnMe.server + ";port=3306;username=" + ReturnMe.uid + ";password=" + ReturnMe.password;
                    //MySqlConnection MyConn2 = new MySqlConnection(MyConnection2);
                   // MySqlCommand MyCommand2 = new MySqlCommand(sql, MyConn2);
                    MySqlDataReader MyReader2;
                    //MyConn2.Open();

                   // MyReader2 = MyCommand2.ExecuteReader();
                    //MyConn2.Close();
                    //End Save to Master_Database


                    //For MDB File
                    if ((ChkType == "F" && FormType == "25") || (ChkType == "F" && FormType == "26") || (ChkType == "GC" && FormType == "20") || (ChkType == "MC" && FormType == "20") || (ChkType == "MC_1" && FormType == "00") || (ChkType == "CUSTOM" && FormType == "00") || (ChkType == "CUSTOM_PA" && FormType == "00") || ChkType == "CS")
                    {
                        Int64 TempStartingSerial = StartingSerial;

                        int LoopCount1 = 0;
                        while (LoopCount1 < PcsPerBook)
                        {
                            OleDbConnection con = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + Application.StartupPath + "\\" + FolderName + "\\" + FileName + ".mdb");
                            OleDbCommand cmd = con.CreateCommand();
                            con.Open();
                            cmd.CommandText = "INSERT INTO InputFile_1Out (BRSTN, AccountNumber, RT1to5, RT6to9,AccountNumberWithHyphen,Serial,Name1,Name2,Name3,Address1,Address2,Address3,Address4,Address5,Address6,BankName, StartingSerial, EndingSerial,PcsPerBook,DataNumber) VALUES ('" + BRSTN + "','" + AccountNo + "','" + BRSTN.Substring(0, 5) + "','" + BRSTN.Substring(5, 4) + "','" + AccountNo.Substring(0, 3) + "-" + AccountNo.Substring(3, 6) + "-" + AccountNo.Substring(9, 3) + "','" + TempStartingSerial.ToString(FormatSerial) + "','" + Name1.Replace("'", "''") + "','" + Name2.Replace("'", "''") + "','" + "" + "','" + Address1.Replace("'", "''") + "','" + Address2.Replace("'", "''") + "','" + Address3.Replace("'", "''") + "','" + Address4.Replace("'", "''") + "','" + Address5.Replace("'", "''") + "','" + Address6.Replace("'", "''") + "','SECURITY BANK','" + StartingSerial.ToString(FormatSerial) + "','" + (StartingSerial + PcsPerBook - 1).ToString(FormatSerial) + "','" + PcsPerBook + "','" + (DataNumber + 1) + "')";
                            cmd.Connection = con;
                            cmd.ExecuteNonQuery();
                            con.Close();

                            TempStartingSerial = TempStartingSerial + 1;
                            DataNumber = DataNumber + 1;
                            LoopCount1 = LoopCount1 + 1;

                            mdb_status_bar = mdb_status_bar + 1;
                        }
                    }
                    //End For MDB File

                    //For Packing
                    string temp_batch = "";

                    if (ChkType == "GC" && FormType == "20")
                    {
                        temp_batch = Batch + " (GC)";
                    }
                    else
                    {
                        if ((ChkType == "MC_1" && FormType == "00") || (ChkType == "MC" && FormType == "20"))
                        {
                            temp_batch = Batch + " (MC)";
                        }
                        else
                        {
                            temp_batch = Batch;
                        }
                    }

                    sql = "INSERT INTO Packing (BatchNo, RT_NO, Branch, Acct_No, Acct_No_P, Acct_Name1, Acct_Name2, No_Bks, CK_NO_P, CK_NO_B, CK_NOE, CK_NO_E, Block, ChkType) VALUES ('" + temp_batch.ToUpper() + "','" + BRSTN + "','" + Address1.Replace("'", "''") + "','" + AccountNo + "','" + AccountNo.Substring(0, 3) + "-" + AccountNo.Substring(3, 6) + "-" + AccountNo.Substring(9, 3) + "','" + Name1.Replace("'", "''") + "','" + Name2.Replace("'", "''") + "','" + "1" + "','" + StartingSerial.ToString(FormatSerial) + "','" + StartingSerial.ToString(FormatSerial) + "','" + (StartingSerial + PcsPerBook - 1).ToString(FormatSerial) + "','" + (StartingSerial + PcsPerBook - 1).ToString(FormatSerial) + "','" + BlockCount + "','" + RefChkType + "')";
                    OleDbConnection conn5 = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + Application.StartupPath + "\\" + FolderName + ";Extended Properties=dBASE III;");
                    OleDbDataAdapter command5 = new OleDbDataAdapter(sql, conn5);
                    conn5.Open();
                    DataSet dataSet5 = new DataSet();
                    command5.Fill(dataSet5);
                    conn5.Close();
                    //End for Packing

                    TotalData = TotalData + 1;
                    StartingSerial = StartingSerial + PcsPerBook;
                    OrderQty = OrderQty - 1;
                    status_bar=status_bar+1;
                }


                LoopCount = LoopCount + 1;
            }

            if (LoopCount > 1)
            {
                PrinterFile.Write("\\");
            }

            if (LoopCount <= 32)
            {
                DoBlock.WriteLine("");
                DoBlock.WriteLine(Summary_DoBlock);
            }

            PrinterFile.Close();
            DoBlock.Close();
            conn1.Close();

            ReturnMe.PackingList(ChkType, FolderName, FormType, RefChkType);

            if (TotalData > 0)
            {

                if (Directory.Exists("C:\\Windows\\Temp\\" + DateTimeToday) == false) { Directory.CreateDirectory("C:\\Windows\\Temp\\" + DateTimeToday); }

                if (FolderName.ToUpper() == "REGULAR\\PREENCODED")
                {
                    if (Directory.Exists("C:\\Windows\\Temp\\" + DateTimeToday + "\\Regular") == false) { Directory.CreateDirectory("C:\\Windows\\Temp\\" + DateTimeToday + "\\Regular"); }
                }


                if (ChkType == "MC_1" && FormType == "00")
                {
                    if (Directory.Exists("C:\\Windows\\Temp\\" + DateTimeToday + "\\MC") == false)
                    {
                        Directory.CreateDirectory("C:\\Windows\\Temp\\" + DateTimeToday + "\\MC");
                    }
                }

                if (Directory.Exists("C:\\Windows\\Temp\\" + DateTimeToday + "\\" + FolderName) == false)
                {
                    Directory.CreateDirectory("C:\\Windows\\Temp\\" + DateTimeToday + "\\" + FolderName);
                }

                File.Copy(Application.StartupPath + "\\" + FolderName + "\\Block" + ChkType2 + ".txt", "C:\\Windows\\Temp\\" + DateTimeToday + "\\" + FolderName + "\\Block" + ChkType2 + ".txt");
                File.Copy(Application.StartupPath + "\\" + FolderName + "\\" + FileName + ".txt", "C:\\Windows\\Temp\\" + DateTimeToday + "\\" + FolderName + "\\" + FileName + ".txt");
                File.Copy(Application.StartupPath + "\\" + FolderName + "\\Packing" + RefChkType + ".txt", "C:\\Windows\\Temp\\" + DateTimeToday + "\\" + FolderName + "\\Packing" + RefChkType + ".txt");

                if (FolderName == "Regular\\PreEncoded")
                {
                    if (Directory.Exists("C:\\Windows\\Temp\\" + DateTimeToday + "\\" + FinalBatch + "\\Regular\\") == false) { Directory.CreateDirectory("C:\\Windows\\Temp\\" + DateTimeToday + "\\" + FinalBatch + "\\Regular"); }
                }

                if (ChkType == "MC_1" && FormType == "00")
                {
                    if (Directory.Exists("C:\\Windows\\Temp\\" + DateTimeToday + "\\" + FinalBatch + "\\MC") == false)
                    {
                        Directory.CreateDirectory("C:\\Windows\\Temp\\" + DateTimeToday + "\\" + FinalBatch + "\\MC");
                    }
                }

                if (Directory.Exists("C:\\Windows\\Temp\\" + DateTimeToday + "\\" + FinalBatch + "\\" + FolderName) == false)
                {
                    Directory.CreateDirectory("C:\\Windows\\Temp\\" + DateTimeToday + "\\" + FinalBatch + "\\" + FolderName);
                }

                File.Copy(Application.StartupPath + "\\" + FolderName + "\\Packing" + RefChkType + ".txt", "C:\\Windows\\Temp\\" + DateTimeToday + "\\" + FinalBatch + "\\" + FolderName + "\\Packing" + RefChkType + ".txt");

                if ((ChkType == "F" && FormType == "25") || (ChkType == "F" && FormType == "26") || (ChkType == "GC" && FormType == "20") || (ChkType == "MC" && FormType == "20") || (ChkType == "MC_1" && FormType == "00") || (ChkType == "CUSTOM" && FormType == "00") || (ChkType == "CUSTOM_PA" && FormType == "00") || ChkType == "CS")
                {

                    int Orig_DataNumber = DataNumber;
                    string All_Column = "BRSTN , AccountNumber , RT1to5 , RT6to9 , AccountNumberWithHyphen , Serial , Name1 , Name2 , Name3 , Address1 , Address2 , Address3 , Address4 , Address5 , Address6 , BankName , StartingSerial , EndingSerial , PcsPerBook , FileName, datanumber";
                    
                    //4 Outs
                    OleDbConnection con = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + Application.StartupPath + "\\" + FolderName + "\\" + FileName + ".mdb");
                    OleDbCommand cmd = con.CreateCommand();
                    con.Open();
                    cmd.CommandText = "DELETE FROM InputFile_Temp";
                    cmd.Connection = con;
                    cmd.ExecuteNonQuery();
                    con.Close();

                    con = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + Application.StartupPath + "\\" + FolderName + "\\" + FileName + ".mdb");
                    cmd = con.CreateCommand();
                    con.Open();
                    cmd.CommandText = "INSERT INTO InputFile_Temp (" + All_Column + ") SELECT " + All_Column + " FROM InputFile_1Out";
                    cmd.Connection = con;
                    cmd.ExecuteNonQuery();
                    con.Close();

                    while (DataNumber % (PcsPerBook * 4) != 0)
                    {
                        con = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + Application.StartupPath + "\\" + FolderName + "\\" + FileName + ".mdb");
                        cmd = con.CreateCommand();
                        con.Open();
                        cmd.CommandText = "INSERT INTO InputFile_Temp (DataNumber) VALUES ('" + (DataNumber + 1) + "')";
                        cmd.Connection = con;
                        cmd.ExecuteNonQuery();
                        con.Close();
                            
                            
                        DataNumber = DataNumber + 1;

                        mdb_status_bar = mdb_status_bar + 1;
                    }

                    int LineNumber1 = PcsPerBook * 0;
                    int LineNumber2 = PcsPerBook * 1;
                    int LineNumber3 = PcsPerBook * 2;
                    int LineNumber4 = PcsPerBook * 3;

                RepeatMe_4Outs:

                    LoopCount = 0;
                    while (LoopCount < PcsPerBook)
                    {
                        //Line Number 1
                        LineNumber1 = LineNumber1 + 1;

                        con = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + Application.StartupPath + "\\" + FolderName + "\\" + FileName + ".mdb");
                        cmd = con.CreateCommand();
                        con.Open();
                        cmd.CommandText = "INSERT INTO InputFile_4Outs (" + All_Column + ") SELECT " + All_Column + " FROM InputFile_Temp WHERE DataNumber = '" + LineNumber1 + "'";
                        cmd.Connection = con;
                        cmd.ExecuteNonQuery();
                        con.Close();
                        mdb_status_bar = mdb_status_bar + 1;
                        //End Line Number 1

                        //Line Number 2
                        LineNumber2 = LineNumber2 + 1;

                        con = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + Application.StartupPath + "\\" + FolderName + "\\" + FileName + ".mdb");
                        cmd = con.CreateCommand();
                        con.Open();
                        cmd.CommandText = "INSERT INTO InputFile_4Outs (" + All_Column + ") SELECT " + All_Column + " FROM InputFile_Temp WHERE DataNumber = '" + LineNumber2 + "'";
                        cmd.Connection = con;
                        cmd.ExecuteNonQuery();
                        con.Close();
                        mdb_status_bar = mdb_status_bar + 1;
                        //End Line Number 2

                        //Line Number 3
                        LineNumber3 = LineNumber3 + 1;

                        con = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + Application.StartupPath + "\\" + FolderName + "\\" + FileName + ".mdb");
                        cmd = con.CreateCommand();
                        con.Open();
                        cmd.CommandText = "INSERT INTO InputFile_4Outs (" + All_Column + ") SELECT " + All_Column + " FROM InputFile_Temp WHERE DataNumber = '" + LineNumber3 + "'";
                        cmd.Connection = con;
                        cmd.ExecuteNonQuery();
                        con.Close();
                        mdb_status_bar = mdb_status_bar + 1;
                        //End Line Number 3

                        //Line Number 4
                        LineNumber4 = LineNumber4 + 1;

                        con = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + Application.StartupPath + "\\" + FolderName + "\\" + FileName + ".mdb");
                        cmd = con.CreateCommand();
                        con.Open();
                        cmd.CommandText = "INSERT INTO InputFile_4Outs (" + All_Column + ") SELECT " + All_Column + " FROM InputFile_Temp WHERE DataNumber = '" + LineNumber4 + "'";
                        cmd.Connection = con;
                        cmd.ExecuteNonQuery();
                        con.Close();
                        mdb_status_bar = mdb_status_bar + 1;
                        //End Line Number 4

                        LoopCount = LoopCount + 1;
                    }

                    if (LineNumber4 != DataNumber)
                    {
                        LineNumber1 = LineNumber1 + (PcsPerBook * 3);
                        LineNumber2 = LineNumber2 + (PcsPerBook * 3);
                        LineNumber3 = LineNumber3 + (PcsPerBook * 3);
                        LineNumber4 = LineNumber4 + (PcsPerBook * 3);

                        goto RepeatMe_4Outs;
                    }
                    //End 4 Outs
                    //3 Outs
                    DataNumber = Orig_DataNumber;

                    con = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + Application.StartupPath + "\\" + FolderName + "\\" + FileName + ".mdb");
                    cmd = con.CreateCommand();
                    con.Open();
                    cmd.CommandText = "DELETE FROM InputFile_Temp";
                    cmd.Connection = con;
                    cmd.ExecuteNonQuery();
                    con.Close();


                    con = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + Application.StartupPath + "\\" + FolderName + "\\" + FileName + ".mdb");
                    cmd = con.CreateCommand();
                    con.Open();
                    cmd.CommandText = "INSERT INTO InputFile_Temp (" + All_Column + ") SELECT " + All_Column + " FROM InputFile_1Out";
                    cmd.Connection = con;
                    cmd.ExecuteNonQuery();
                    con.Close();

                    while (DataNumber % (PcsPerBook * 3) != 0)
                    {
                        con = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + Application.StartupPath + "\\" + FolderName + "\\" + FileName + ".mdb");
                        cmd = con.CreateCommand();
                        con.Open();
                        cmd.CommandText = "INSERT INTO InputFile_Temp (DataNumber) VALUES ('" + (DataNumber + 1) + "')";
                        cmd.Connection = con;
                        cmd.ExecuteNonQuery();
                        con.Close();

                        DataNumber = DataNumber + 1;

                        mdb_status_bar = mdb_status_bar + 1;
                    }

                    LineNumber1 = PcsPerBook * 0;
                    LineNumber2 = PcsPerBook * 1;
                    LineNumber3 = PcsPerBook * 2;

                RepeatMe_3Outs:

                    LoopCount = 0;
                    while (LoopCount < PcsPerBook)
                    {
                        //Line Number 1
                        LineNumber1 = LineNumber1 + 1;

                        con = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + Application.StartupPath + "\\" + FolderName + "\\" + FileName + ".mdb");
                        cmd = con.CreateCommand();
                        con.Open();
                        cmd.CommandText = "INSERT INTO InputFile_3Outs (" + All_Column + ") SELECT " + All_Column + " FROM InputFile_Temp WHERE DataNumber = '" + LineNumber1 + "'";
                        cmd.Connection = con;
                        cmd.ExecuteNonQuery();
                        con.Close();
                        mdb_status_bar = mdb_status_bar + 1;
                        //End Line Number 1

                        //Line Number 2
                        LineNumber2 = LineNumber2 + 1;

                        con = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + Application.StartupPath + "\\" + FolderName + "\\" + FileName + ".mdb");
                        cmd = con.CreateCommand();
                        con.Open();
                        cmd.CommandText = "INSERT INTO InputFile_3Outs (" + All_Column + ") SELECT " + All_Column + " FROM InputFile_Temp WHERE DataNumber = '" + LineNumber2 + "'";
                        cmd.Connection = con;
                        cmd.ExecuteNonQuery();
                        con.Close();
                        mdb_status_bar = mdb_status_bar + 1;
                        //End Line Number 2

                        //Line Number 3
                        LineNumber3 = LineNumber3 + 1;

                        con = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + Application.StartupPath + "\\" + FolderName + "\\" + FileName + ".mdb");
                        cmd = con.CreateCommand();
                        con.Open();
                        cmd.CommandText = "INSERT INTO InputFile_3Outs (" + All_Column + ") SELECT " + All_Column + " FROM InputFile_Temp WHERE DataNumber = '" + LineNumber3 + "'";
                        cmd.Connection = con;
                        cmd.ExecuteNonQuery();
                        con.Close();
                        mdb_status_bar = mdb_status_bar + 1;
                        //End Line Number 3                      

                        LoopCount = LoopCount + 1;
                    }

                    if (LineNumber3 != DataNumber)
                    {
                        LineNumber1 = LineNumber1 + (PcsPerBook * 2);
                        LineNumber2 = LineNumber2 + (PcsPerBook * 2);
                        LineNumber3 = LineNumber3 + (PcsPerBook * 2);


                        goto RepeatMe_3Outs;
                    }
                    //End 3 Outs

                    //2 Outs
                    DataNumber = Orig_DataNumber;

                    con = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + Application.StartupPath + "\\" + FolderName + "\\" + FileName + ".mdb");
                    cmd = con.CreateCommand();
                    con.Open();
                    cmd.CommandText = "DELETE FROM InputFile_Temp";
                    cmd.Connection = con;
                    cmd.ExecuteNonQuery();
                    con.Close();

                    con = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + Application.StartupPath + "\\" + FolderName + "\\" + FileName + ".mdb");
                    cmd = con.CreateCommand();
                    con.Open();
                    cmd.CommandText = "INSERT INTO InputFile_Temp (" + All_Column + ") SELECT " + All_Column + " FROM InputFile_1Out";
                    cmd.Connection = con;
                    cmd.ExecuteNonQuery();
                    con.Close();

                    while (DataNumber % (PcsPerBook * 2) != 0)
                    {
                        con = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + Application.StartupPath + "\\" + FolderName + "\\" + FileName + ".mdb");
                        cmd = con.CreateCommand();
                        con.Open();
                        cmd.CommandText = "INSERT INTO InputFile_Temp (DataNumber) VALUES ('" + (DataNumber + 1) + "')";
                        cmd.Connection = con;
                        cmd.ExecuteNonQuery();
                        con.Close();

                        DataNumber = DataNumber + 1;

                        mdb_status_bar = mdb_status_bar + 1;
                    }

                    LineNumber1 = PcsPerBook * 0;
                    LineNumber2 = PcsPerBook * 1;


                RepeatMe_2Outs:

                    LoopCount = 0;
                    while (LoopCount < PcsPerBook)
                    {
                        //Line Number 1
                        LineNumber1 = LineNumber1 + 1;

                        con = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + Application.StartupPath + "\\" + FolderName + "\\" + FileName + ".mdb");
                        cmd = con.CreateCommand();
                        con.Open();
                        cmd.CommandText = "INSERT INTO InputFile_2Outs (" + All_Column + ") SELECT " + All_Column + " FROM InputFile_Temp WHERE DataNumber = '" + LineNumber1 + "'";
                        cmd.Connection = con;
                        cmd.ExecuteNonQuery();
                        con.Close();
                        mdb_status_bar = mdb_status_bar + 1;
                        //End Line Number 1

                        //Line Number 2
                        LineNumber2 = LineNumber2 + 1;

                        con = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + Application.StartupPath + "\\" + FolderName + "\\" + FileName + ".mdb");
                        cmd = con.CreateCommand();
                        con.Open();
                        cmd.CommandText = "INSERT INTO InputFile_2Outs (" + All_Column + ") SELECT " + All_Column + " FROM InputFile_Temp WHERE DataNumber = '" + LineNumber2 + "'";
                        cmd.Connection = con;
                        cmd.ExecuteNonQuery();
                        con.Close();
                        mdb_status_bar = mdb_status_bar + 1;
                        //End Line Number 2


                        LoopCount = LoopCount + 1;
                    }

                    if (LineNumber2 != DataNumber)
                    {
                        LineNumber1 = LineNumber1 + (PcsPerBook * 1);
                        LineNumber2 = LineNumber2 + (PcsPerBook * 1);


                        goto RepeatMe_2Outs;
                    }
                    //End 2 Outs


                    File.Copy(Application.StartupPath + "\\" + FolderName + "\\" + FileName + ".mdb", "C:\\Windows\\Temp\\" + DateTimeToday + "\\" + FolderName + "\\" + FileName + ".mdb");

                    //Copy to Drive R
                    ReturnMe.CreateDirectory(Temp_DriveR);

                    if (File.Exists(Temp_DriveR + "\\" + FileName + ".mdb")) { File.Delete(Temp_DriveR + "\\" + FileName + ".mdb"); }
                    File.Copy(Application.StartupPath + "\\" + FolderName + "\\" + FileName + ".mdb", Temp_DriveR + "\\" + FileName + ".mdb");
                    //End Copy to Drive R


                    //Copy to CTC
                    ReturnMe.CreateDirectory(Temp_CTC);

                    if (File.Exists(Temp_CTC + "\\" + FileName + ".mdb")) { File.Delete(Temp_CTC + "\\" + FileName + ".mdb"); }
                    File.Copy(Application.StartupPath + "\\" + FolderName + "\\" + FileName + ".mdb", Temp_CTC + "\\" + FileName + ".mdb");
                    //End Copy to CTC
                }

                //Copy to Drive R
                CreateDirectory(Temp_DriveR);

                if (File.Exists(Temp_DriveR + "\\" + FileName + ".txt")) { File.Delete(Temp_DriveR + "\\" + FileName + ".txt"); }
                File.Copy(Application.StartupPath + "\\" + FolderName + "\\" + FileName + ".txt", Temp_DriveR + "\\" + FileName + ".txt");
                //Copy to Drive R

                //Copy to CTC
                CreateDirectory(Temp_CTC);

                if (File.Exists(Temp_CTC + "\\" + FileName + ".txt")) { File.Delete(Temp_CTC + "\\" + FileName + ".txt"); }
                File.Copy(Application.StartupPath + "\\" + FolderName + "\\" + FileName + ".txt", Temp_CTC + "\\" + FileName + ".txt");
                //End Copy to CTC
            }
            return TotalData;
        }
        public static void DeleteDBF(string FileName, string FolderName)
        {
            string sql = "DELETE FROM " + FileName;

            OleDbConnection conn1 = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + Application.StartupPath + "\\"+FolderName+";Extended Properties=dBASE III;");
            OleDbDataAdapter command1 = new OleDbDataAdapter(sql, conn1);
            conn1.Open();
            DataSet dataSet = new DataSet();
            command1.Fill(dataSet);
            conn1.Close();
        }
        public static Boolean CheckBatchExists(string Batch)
        {
            string dbase = "";
            if (CodesOnly == true) { dbase = "captive_database.Master_Database_SBTC_Temp"; }
            if (CodesOnly == false) { dbase = "captive_database.Master_Database_SBTC"; }

            string sql = "SELECT * FROM " + dbase + " WHERE FinalBatch = '" + Batch + "'";
            string MyConnection2 = "datasource=" + ReturnMe.server + ";port=3306;username=" + ReturnMe.uid + ";password=" + ReturnMe.password;
            MySqlConnection MyConn2 = new MySqlConnection(MyConnection2);
            MySqlCommand MyCommand2 = new MySqlCommand(sql, MyConn2);
            MySqlDataReader MyReader2;
            MyConn2.Open();

            MyReader2 = MyCommand2.ExecuteReader();
            if (MyReader2.HasRows)
            {
                MyConn2.Close();
                return true;
            }
            else
            {
                MyConn2.Close();
                return false;
            }                    
        }
        public static OrderSorted Sort(List<OrderModel> _orders)
        {
            OrderSorted sorted = new OrderSorted();

            sorted.RegularPersonal = new List<OrderModel>();

            sorted.RegularCommercial = new List<OrderModel>();

            sorted.ManagersCheck = new List<OrderModel>();

            sorted.GiftCheck = new List<OrderModel>();

            sorted.PersonalPreEncoded = new List<OrderModel>();

            sorted.CommercialPreEncoded = new List<OrderModel>();

            sorted.CheckOnePersonal = new List<OrderModel>();

            sorted.CheckOneCommerical = new List<OrderModel>();

            sorted.CheckPowerPersonal = new List<OrderModel>();

            sorted.CheckPowerCommercial = new List<OrderModel>();

            sorted.CustomizedCheck = new List<OrderModel>();

            sorted.ManagersCheckCont = new List<OrderModel>();

            _orders.ForEach(r =>{

                if (r.CheckType == "A" && r.FormType == "05")
                    sorted.RegularPersonal.Add(r);

                else if (r.CheckType == "B" && r.FormType == "16")
                    sorted.RegularCommercial.Add(r);

                else if (r.CheckType == "MC" && r.FormType == "20")
                    sorted.ManagersCheck.Add(r);

                else if (r.CheckType == "GC" && r.FormType == "20")
                    sorted.GiftCheck.Add(r);

                else if (r.CheckType == "AA" && r.FormType == "05")
                    sorted.PersonalPreEncoded.Add(r);

                else if (r.CheckType == "BB" && r.FormType == "16")
                    sorted.CommercialPreEncoded.Add(r);

                else if (r.CheckType == "F" && r.FormType == "25")
                    sorted.CheckOnePersonal.Add(r);

                else if (r.CheckType == "F" && r.FormType == "26")
                    sorted.CheckOneCommerical.Add(r);

                else if (r.CheckType == "E" && r.FormType == "23")
                    sorted.CheckPowerPersonal.Add(r);

                else if (r.CheckType == "E" && r.FormType == "22")
                    sorted.CheckPowerCommercial.Add(r);

                else if (r.CheckType == "CUSTOM" && r.FormType == "00")
                    sorted.CustomizedCheck.Add(r);

                else if (r.CheckType == "MC_1" && r.FormType == "00")
                    sorted.ManagersCheckCont.Add(r);
            });//END FOREACH

            return sorted;
        }//END OF FUNCTION      
    }
}
