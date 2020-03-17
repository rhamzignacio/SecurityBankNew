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


namespace sbtc
{
    public partial class frmMain : Form
    {        
        int pkey = 0;

        string errorMessage = "";

        List<BranchesModel> branchList = new List<BranchesModel>();

        OrderSorted sortedList = new OrderSorted();

        public frmMain()
        {
            InitializeComponent();

            GetAllBranch();

            getSettings();

            dteDeliveryDate.Value = DateTime.Today;

            ListFilesHead();

            ReturnMe.TimeStart = DateTime.Now.ToString("HH:mm");
            ReturnMe.DateTimeToday_date = DateTime.Now;

            CheckHashTotal();
        }

        private void GetAllBranch()
        {
            Services service = new Services();

            branchList = service.GetAllBranch();
        }

        private void getSettings()
        {
            string Temp = Application.StartupPath.ToUpper();
            string Temp2 = "";

            int LoopCount = 0;
            while (LoopCount < Temp.Length-5)
            {
                if (Temp.Substring(LoopCount,4) == "AUTO")
                {
                    Temp2 = Temp.Substring(0,LoopCount-1);
                }

                if (Temp.Substring(LoopCount,5) == "CODES")
                {
                    ReturnMe.CodesOnly = true;
                }
            
                LoopCount = LoopCount +1;
            }

            LoopCount = 0;

            var fileStream = new FileStream(Temp2 +"\\Auto\\Settings.ini", FileMode.Open, FileAccess.Read);
            using (var streamReader = new StreamReader(fileStream, Encoding.UTF8))
            {
                string line;
                while ((line = streamReader.ReadLine()) != null)
                {
                    LoopCount = LoopCount +1;
                    if (LoopCount == 1) ReturnMe.server = line;
                    if (LoopCount == 2) ReturnMe.Resting_Folder = line;
                    if (LoopCount == 3) ReturnMe.PrinterFiles_Folder = line;
                }
            }
            fileStream.Close();

            LoopCount = 0;
            fileStream = new FileStream("C:\\WinZip.txt", FileMode.Open, FileAccess.Read);
            using (var streamReader = new StreamReader(fileStream, Encoding.UTF8))
            {
                string line;
                while ((line = streamReader.ReadLine()) != null)
                {
                    
                    LoopCount = LoopCount + 1;
                    if (LoopCount == 1) ReturnMe.WinZipLocation = line;

                    
                }
            }
            fileStream.Close();            
        }//END OF FUNCTION

        private void frmMain_Load(object sender, EventArgs e)
        {

        }//END OF FUNCTION

        private void CheckHashTotal()
        {
            string dbase = "";
            if (ReturnMe.CodesOnly == true) dbase = "captive_database.master_database_sbtc_temp";
            if (ReturnMe.CodesOnly == false) dbase = "captive_database.master_database_sbtc";

            string sql = "SELECT DISTINCT(FinalBatch) FROM " + dbase + " WHERE HashSentDate is NULL AND HashSentTime IS NULL AND DeliveryDate < '" + DateTime.Now.ToString("yyyy-MM-dd") + "'";
            string MyConnection2 = "datasource=" + ReturnMe.server + ";port=3306;username=" + ReturnMe.uid + ";password=" + ReturnMe.password;
            MySqlConnection MyConn2 = new MySqlConnection(MyConnection2);
            MySqlCommand MyCommand2 = new MySqlCommand(sql, MyConn2);
            MySqlDataReader MyReader2;
            MyConn2.Open();
            MyReader2 = MyCommand2.ExecuteReader();

            int LoopCount = 0;

            while (MyReader2.Read())
            {
                LoopCount = LoopCount + 1;
            }

            if (LoopCount >= 1)
            {
                lblHashTotal.Visible = true;

                if (LoopCount == 1)
                {
                    lblHashTotal.Text = "1 Hash Total hasn't been sent yet";
                }
                else lblHashTotal.Text = LoopCount.ToString() + " Hash Totals hasn't been sent yet";
            }
            else
            {
                lblHashTotal.Visible = false;
            }
        }



        public void ListFilesHead()
        {
            lstFiles.Items.Clear();

            string[] files = Directory.GetFiles(Application.StartupPath+"\\Head\\","*.txt");

            foreach (string file in files)
            {

                string filename = "";

                int LoopCount = file.Length-1;
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

                lstFiles.Items.Add(filename);
            }


            if (lstFiles.Items.Count == 0)
            {
                btnCheckFiles.Enabled = false;
            }
            else
            {
                btnCheckFiles.Enabled = true;
            }
        }

        private void btnCheckFiles_Click(object sender, EventArgs e)
        {
            btnEncode.Enabled = false;

            if (dteDeliveryDate.Value.ToShortDateString() == DateTime.Today.ToShortDateString())
            {
                DialogResult result3 = MessageBox.Show("Please select Delivery Date", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                
                dteDeliveryDate.Focus();
                
                return;
            }

            if (btnCheckFiles.Text == "Check Files on Head")
            {
                DialogResult result1 = MessageBox.Show("Are you sure you want to check " + lstFiles.Items.Count + " items?", "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Information);
                
                if (result1 == DialogResult.No)
                {
                    return;
                }

                ReturnMe.CreateTable();

                List<OrderModel> OrderList = new List<OrderModel>();

                int LoopCount = 0;
                
                while (LoopCount < lstFiles.Items.Count)
                {
                    string filename = lstFiles.Items[LoopCount].ToString();

                    OrderList.AddRange(CheckFiles(filename));

                    LoopCount = LoopCount + 1;
                }

                CheckBranches(OrderList);
              
                DisplayData(OrderList);

                btnCheckFiles.Text = "Process ! ! !";

                //For Sort RT
                ReturnMe.SortRT("Regular");
                ReturnMe.SortRT("Regular\\PreEncoded");
                ReturnMe.SortRT("MC");
                ReturnMe.SortRT("CheckOne");
                ReturnMe.SortRT("CheckPower");
                ReturnMe.SortRT("GiftCheck");
                //End For Sort RT

                MessageBox.Show("Data has been Checked. No Errors found!", " ", MessageBoxButtons.OK, MessageBoxIcon.Information);              
            }
            else
            {
                btnCheckFiles.Enabled = false;
                btnEncode.Enabled = false;

                Boolean DebugMe = false;

                if (DebugMe == true)
                {
                    ReturnMe.ProcessAll2(dteDeliveryDate.Value);
                }
                else
                {
                    ReturnMe.ActivateMaxProgressBarCarbon();

                    backgroundWorker1.RunWorkerAsync();
                    
                    timer1.Start();
                }
            }
        }//END OF FUNCTION
    

        private void DisplayData(List<OrderModel> _orders)
        {
            sortedList = ReturnMe.Sort(_orders);

            string display = "";

            if (sortedList.RegularPersonal.Count > 0)
                display += "Regular Personal\t\t\t" + sortedList.RegularPersonal.Count.ToString() + "\n";

            if (sortedList.RegularCommercial.Count > 0)
                display += "Regular Commercial\t\t\t" + sortedList.RegularCommercial.Count.ToString() + "\n";

            if (sortedList.ManagersCheck.Count > 0)
                display += "Manager's Check\t\t\t" + sortedList.ManagersCheck.Count.ToString() + "\n";

            if (sortedList.GiftCheck.Count > 0)
                display += "Gift Check\t\t\t" + sortedList.GiftCheck.Count.ToString() + "\n";

            if (sortedList.PersonalPreEncoded.Count > 0)
                display += "Personal Pre-Encoded\t\t\t" + sortedList.PersonalPreEncoded.Count.ToString() + "\n";

            if (sortedList.CommercialPreEncoded.Count > 0)
                display += "Commercial Pre-Encoded\t\t\t" + sortedList.CommercialPreEncoded.Count.ToString() + "\n";

            if (sortedList.CheckOnePersonal.Count > 0)
                display += "CheckOne Personal\t\t\t" + sortedList.CheckOnePersonal.Count.ToString() + "\n";

            if (sortedList.CheckOneCommerical.Count > 0)
                display += "CheckOne Commercial\t\t\t" + sortedList.CheckOneCommerical.Count.ToString() + "\n";

            if (sortedList.CheckPowerPersonal.Count > 0)
                display += "CheckPower Personal\t\t\t" + sortedList.CheckPowerPersonal.Count.ToString() + "\n";

            if (sortedList.CheckPowerCommercial.Count > 0)
                display += "CheckPower Commercial\t\t\t" + sortedList.CheckOneCommerical.Count.ToString() + "\n";

            if(sortedList.CustomizedCheck.Count> 0)
                display += "Customized Check\t\t\t" + sortedList.CustomizedCheck.Count.ToString() + "\n";

            if (sortedList.ManagersCheckCont.Count > 0)
                display += "Manager's Check Continous\t\t\t" + sortedList.ManagersCheckCont.Count.ToString() + "\n";

            lblTotal.Text = display;
        }

        private void CheckBranches(List<OrderModel> _orders)
        {
            _orders.ForEach(order =>
            {
                var temp = branchList.FirstOrDefault(r => r.BRSTN == order.BRSTN);

                if(temp == null)
                {
                    errorMessage += "\nBRSTN=" + order.BRSTN + " is not found on BRANCHES DATABASE";
                }

            });
        }

        private List<OrderModel> CheckFiles(string filename)
        {
            string Batch = filename.Substring(0, filename.Length - 4);

            List<OrderModel> returnList = new List<OrderModel>();

            List<OrderModel> tempList = new List<OrderModel>();

            var fileStream = new FileStream(Application.StartupPath + "\\Head\\" + filename, FileMode.Open, FileAccess.Read);

            using (var streamReader = new StreamReader(fileStream, Encoding.UTF8))
            {
                string line;
                while ((line = streamReader.ReadLine()) != null)
                {
                    pkey = pkey + 1;

                    OrderModel order = new OrderModel();

                    order.CheckType = line.Substring(0, 1).Trim();

                    order.BRSTN = line.Substring(1, 9).Trim();

                    order.AccountNo = line.Substring(11, 12).Trim();

                    order.Name = line.Substring(23, 56).Trim();

                    order.Name = order.Name.Replace("Ñ", "N");

                    order.Name = order.Name.Replace("¥", "N");

                    order.Name = order.Name.Replace("NO NAME", "");

                    order.ContCode = line.Substring(79, 1).Trim();

                    order.FormType = line.Substring(80, 2).Trim();

                    string temp = line.Substring(82, 2).Trim();

                    if(int.TryParse(temp, out int test))
                    {
                        //TRY TO CHECK IF CAPTURED DATA IS INTEGER
                    }
                    else
                    {
                        errorMessage += "-Error Parsing Quantity of BRSTN=" + order.BRSTN + " AccountNo=" + order.AccountNo + " FileName=" + filename; 
                    }

                    order.OrderQuantity = int.Parse(line.Substring(82, 2).Trim());
      

                    if ((order.CheckType == "A" && order.FormType == "05") || (order.CheckType == "B" && order.FormType == "16") || 
                    (order.CheckType == "F" && order.FormType == "25") || (order.CheckType == "F" && order.FormType == "26") || 
                    (order.CheckType == "B" && order.FormType == "20" && order.AccountNo.Substring(4, 3) != "212") || 
                    (order.CheckType == "B" && order.FormType == "20" && order.AccountNo.Substring(4, 3) == "212") || (order.CheckType == "E" && order.FormType == "22") || 
                    (order.CheckType == "E" && order.FormType == "23"))
                    {
                        if (filename.Substring(0, 6).ToUpper() == "YSECPT")
                        {
                            if (order.CheckType == "A" && order.FormType == "05") 
                            {
                                order.CheckType = "AA"; 
                            }
                            if (order.CheckType == "B" && order.FormType == "16") 
                            {
                                order.CheckType = "BB"; 
                            }
                        }

                        if 
                        ((order.CheckType == "B" && order.FormType == "20" && order.AccountNo.Substring(4, 3) == "212") ||
                            (order.CheckType == "B" && order.FormType == "20" && order.AccountNo.Substring(0,1) == "9" && order.AccountNo.Substring(5, 7) == "2000022"))
                        {
                            order.CheckType = "GC";
                            order.Name = "";
                        }

                        if (order.CheckType == "B" && order.FormType == "20" && order.AccountNo.Substring(4, 3) != "212") { order.CheckType = "MC"; }


                        if (order.ContCode == "" || order.ContCode == "1")
                        {
                            returnList.Add(order);
                        }

                        if (order.ContCode == "2")
                        {
                            tempList.Add(order);
                        }
                    }//END IF

                    else 
                    {
                        MessageBox.Show("Unable to find the chequename of ChkType " + order.CheckType + " with FormType " + order.FormType + " on " + filename + ". Account No: " + 
                            order.AccountNo, "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);

                        Application.Exit();
                    }
                }
            }

            fileStream.Close();

            tempList.ForEach(temp =>
            {
                var order = returnList.Where(r => r.AccountNo == temp.AccountNo).ToList();

                if(order != null)
                {
                    order.ForEach(t =>
                    {
                        t.Name2 = temp.Name;
                    });
                }
            });//END OF FOREACH

            return returnList;
        }

        public void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            try
            {
                ReturnMe.ProcessAll2(dteDeliveryDate.Value);
            }

            catch (Exception exception)
            {
                MessageBox.Show(exception.ToString());
                Application.Exit();
            }
        }

        private void lblHashTotal_Click(object sender, EventArgs e)
        {

        }

        private void lblHashTotal_DoubleClick(object sender, EventArgs e)
        {
            this.Hide();

            frmHashTotal frmHashTotal = new frmHashTotal();
            frmHashTotal.ShowDialog();
            
            
            
        }

        private void btnEncode_Click(object sender, EventArgs e)
        {   
            this.Hide();
            
            frmCustomized frmCustomized = new frmCustomized();
            frmCustomized.ShowDialog();
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            progressBar2.Maximum = ReturnMe.mdb_status_bar_max+1;
            progressBar2.Minimum = 0;
            progressBar2.Value = ReturnMe.mdb_status_bar;

            if (progressBar2.Value != 0 && progressBar2.Value != progressBar2.Maximum)
            {
                progressBar2.Visible = true;
            }
            else
            {
                progressBar2.Visible = false;
            }





            progressBar1.Maximum = ReturnMe.status_bar_max;
            progressBar1.Minimum = 0;
            progressBar1.Value = ReturnMe.status_bar;

            if (progressBar1.Value != 0 && progressBar1.Value != progressBar1.Maximum)
            {
                progressBar1.Visible = true;
            }
            else
            {
                progressBar1.Visible = false;
            }
        }
    }
}
