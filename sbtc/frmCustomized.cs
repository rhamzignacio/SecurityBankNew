using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.OleDb;
using MySql.Data;
using MySql.Data.MySqlClient;
using System.Xml;
using ADODB;

namespace sbtc
{
    public partial class frmCustomized : Form
    {
        int tmrValue = 0;

        List<BranchesModel> branchList;

        string MyConnection2 = "datasource=192.168.0.254;port=3306;username=root;password=CorpCaptive";
        public frmCustomized(List<BranchesModel> _branches)
        {
            InitializeComponent();

            branchList = _branches;
        }

        private void frmCustomized_Load(object sender, EventArgs e)
        {
            
            dteDeliveryDate.Value = DateTime.Today;

            LoadMe();
            
        }

        public void LoadMe()
        {
            cboChequeName.Items.Clear();
            cboChequeName.Items.Add("Customized Sheeted Checks");
            cboChequeName.Items.Add("Customized Continues Checks");
            cboChequeName.Items.Add("Customized Personal Checks");
            cboChequeName.Items.Add("Dividend Checks");
            cboChequeName.Items.Add("Digibanker");

            txtAccountNo.Text = "";
            txtBRSTN.Text = "";
            txtName1.Text = "";
            txtName2.Text = "";
            txtBooks.Text = "";
            txtStartingSeries.Text = "";

            cboChequeName.Focus();

            lblBRSTN.Visible = true;
            txtBRSTN.Visible = true;

            lblBranchName.Visible = false;
            cboBranchName.Visible = false;
        }

        private void frmCustomized_FormClosed(object sender, FormClosedEventArgs e)
        {
            Application.Exit();
        }

        private void btnAdd_Click(object sender, EventArgs e)
        {      
            if (cboChequeName.Text == "")
            {
                MessageBox.Show("Please select Cheque Name", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                cboChequeName.Focus();
                return;
            }

            if (txtAccountNo.Text.Length != 12)
            {
                MessageBox.Show("Account Number is invalid", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                txtAccountNo.Text = "";
                txtAccountNo.Focus();
                return;
            }

            if (txtBRSTN.Text.Length != 9 && txtBRSTN.Visible == true)
            {
                MessageBox.Show("BRSTN is invalid", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                txtBRSTN.Text = "";
                txtBRSTN.Focus();
                return;
            }

            if (cboChequeName.Text == "" && cboBranchName.Visible == true)
            {
                MessageBox.Show("Branch Name is Invalid", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                cboBranchName.Focus();
                return;
            }

            //Books
            int books = 0;
            if (!Int32.TryParse(txtBooks.Text, out books))
            {
                MessageBox.Show("Books is invalid", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                txtBooks.Text = "";
                txtBooks.Focus();
                return;
            }
            else
            {
                if (books <= 0)
                {
                    MessageBox.Show("Books is invalid", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    txtBooks.Text = "";
                    txtBooks.Focus();
                    return;
                }
            }
            //End Books

            //Starting Serial
            Int64 startingserial = 0;
            if (!Int64.TryParse(txtStartingSeries.Text, out startingserial))
            {
                MessageBox.Show("Starting Series is invalid", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                txtStartingSeries.Text = "";
                txtStartingSeries.Focus();
                return;
            }
            else
            {
                if (books <= 0)
                {
                    MessageBox.Show("Starting Series is invalid", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    txtStartingSeries.Text = "";
                    txtStartingSeries.Focus();
                    return;
                }
            }
            //End Starting Serial


            if (dteDeliveryDate.Value.ToShortDateString() == DateTime.Today.ToShortDateString())
            {
                DialogResult result3 = MessageBox.Show("Please select Delivery Date", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                dteDeliveryDate.Focus();
                return;
            }

            txtName1.Text = txtName1.Text.ToUpper();
            txtName2.Text = txtName2.Text.ToUpper();

            string ChkType = "";
            if (cboChequeName.Text == "Customized Personal Checks")
            {
                ChkType = "CUSTOM_PA";
            }
            else
            {
                if (cboChequeName.Text == "Manager's Check Continues")
                {
                    ChkType = "MC_1";
                }
                else if (cboChequeName.Text == "Digibanker")
                {
                    ChkType = "DIGIBANKER";
                }
                else
                {
                    {
                        ChkType = "CUSTOM";
                    }
                }
                
            }
       
            dataGridView1.Rows.Add(ChkType, txtBRSTN.Text, txtAccountNo.Text, txtName1.Text, txtName2.Text, txtBooks.Text, txtStartingSeries.Text);
            MessageBox.Show("Data has been added", " ", MessageBoxButtons.OK, MessageBoxIcon.Information);
            btnProcess.Enabled = true;

            LoadMe();
        }

        private void txtBRSTN_TextChanged(object sender, EventArgs e)
        {
            if (txtBRSTN.Text.Length == 9)
            {
                OleDbConnection conn1 = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + Application.StartupPath + ";Extended Properties=dBASE III;");
                OleDbDataAdapter command1 = new OleDbDataAdapter("SELECT * FROM Branches WHERE BRSTN = '" + txtBRSTN.Text + "'", conn1);
                conn1.Open();
                DataSet dataSet = new DataSet();
                command1.Fill(dataSet);

                DataTable dt = dataSet.Tables[0];
                //foreach (DataRow dr in dt.Rows)
                //{
                //    string BRSTN = dr[0].ToString();
                //}

                if (dt.Rows.Count == 0)
                {
                    MessageBox.Show("BRSTN "+txtBRSTN.Text+ " does not exists on Branches.dbf", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    txtBRSTN.Text = "";
                    txtBRSTN.Focus();
                    return;
                }
            }
        }

        private void txtAccountNo_TextChanged(object sender, EventArgs e)
        {
            if (txtAccountNo.Text.Length == 12 && cboChequeName.Text != "Manager's Check Continues")
            {
                string sql = "SELECT BRSTN, Name1, Name2 FROM captive_database.sbtc_history WHERE AccountNo = '" + txtAccountNo.Text + "' ORDER BY PrimaryKey DESC LIMIT 1";
                MySqlConnection MyConn2 = new MySqlConnection(MyConnection2);
                MySqlCommand MyCommand2 = new MySqlCommand(sql, MyConn2);
                MySqlDataReader MyReader2;
                MyConn2.Open();

                MyReader2 = MyCommand2.ExecuteReader();

                while (MyReader2.Read())
                {
                    string BRSTN = MyReader2.GetString(0);
                    string Name1 = MyReader2.GetString(1);
                    string Name2 = MyReader2.GetString(2);

                    txtBRSTN.Text = BRSTN;
                    txtName1.Text = Name1;
                    txtName2.Text = Name2;
                }

                MyConn2.Close();
            }
        }

        private void btnProcess_Click(object sender, EventArgs e)
        {
            OrderSorted sbtcList = new OrderSorted();
            sbtcList.RegularPersonal = new List<OrderModel>();
            sbtcList.RegularCommercial = new List<OrderModel>();
            sbtcList.ManagersCheck = new List<OrderModel>();
            sbtcList.ManagersCheckCont = new List<OrderModel>();
            sbtcList.GiftCheck = new List<OrderModel>();
            sbtcList.PersonalPreEncoded = new List<OrderModel>();
            sbtcList.CommercialPreEncoded = new List<OrderModel>();
            sbtcList.CheckOnePersonal = new List<OrderModel>();
            sbtcList.CheckOneCommerical = new List<OrderModel>();
            sbtcList.CheckPowerPersonal = new List<OrderModel>();
            sbtcList.CheckPowerCommercial = new List<OrderModel>();
            sbtcList.DigiBanker = new List<OrderModel>();
            sbtcList.CustomizedCheck = new List<OrderModel>();

            for (int LoopCount = 0; LoopCount < dataGridView1.Rows.Count; LoopCount++)
            {
                string chequename = dataGridView1.Rows[LoopCount].Cells[0].Value.ToString();
                string brstn = dataGridView1.Rows[LoopCount].Cells[1].Value.ToString();
                string AccountNo = dataGridView1.Rows[LoopCount].Cells[2].Value.ToString();
                string Name1 = dataGridView1.Rows[LoopCount].Cells[3].Value.ToString();
                string Name2 = dataGridView1.Rows[LoopCount].Cells[4].Value.ToString();
                string Books = dataGridView1.Rows[LoopCount].Cells[5].Value.ToString();
                string StartingSerial = dataGridView1.Rows[LoopCount].Cells[6].Value.ToString();

                OrderModel sbtc = new OrderModel();
                sbtc.CheckType = chequename ;
                sbtc.BRSTN = brstn;
                sbtc.AccountNo = AccountNo;
                sbtc.Name = Name1;
                sbtc.Name2 = Name2;
                sbtc.FormType = "00";
                sbtc.OrderQuantity = int.Parse(Books);
                sbtc.ManualStart = int.Parse(StartingSerial);

                if (chequename == "CUSTOM")
                {
                    sbtcList.CustomizedCheck.Add(sbtc);
                }
                else if (chequename == "DIGIBANKER")
                {
                    sbtcList.DigiBanker.Add(sbtc);
                }

            }

            if (sbtcList.CustomizedCheck.Count > 0)
            {
                AssignOtherValue(sbtcList.CustomizedCheck);
            }

            if (sbtcList.DigiBanker.Count > 0)
            {
                AssignOtherValue(sbtcList.DigiBanker);
            }

            btnAdd.Enabled = false;
            btnProcess.Enabled = false;

            //PrinterFile
            GenerateService.GeneratePrinterFiles(sbtcList, txtBoxBatchNo.Text, txtBoxExt.Text);

            //DoBlock
            GenerateService.GenerateDoBlock(sbtcList, txtBoxBatchNo.Text, txtBoxExt.Text, dteDeliveryDate.Value,
                "");
        }

        private void AssignSeriels(string _type, List<OrderModel> _orders)
        {
            if(_type == "DIGIBANKER")
            {
                Int64 serial = _orders[0].ManualStart;

                _orders.ForEach(r =>
                {
                    r.StartingSerial = serial + 1;
                    r.EndingSerial = serial + 100;

                    serial = serial + 100;
                });
            }
        }

        private void AssignOtherValue(List<OrderModel> _orders)
        {
            _orders.ForEach(c =>
            {
                var branch = branchList.FirstOrDefault(r => r.BRSTN == c.BRSTN);

                c.Address1 = branch.Address1;
                c.Address2 = branch.Address2;
                c.Address3 = branch.Address3;
                c.Address4 = branch.Address4;
                c.Address5 = branch.Address5;
            });
        }

        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            //ReturnMe.GenerateSortRT("Customized");

            ReturnMe.ProcessAll2(dteDeliveryDate.Value);
        }

        private void backgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            
        }

        private void backgroundWorker1_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            
        }

        private void timer1_Tick(object sender, EventArgs e)
        {         
            tmrValue = tmrValue + 1;
            if (tmrValue == 1)
            {
                lblStatus.Text = ". ";
            }
            else
            {
                lblStatus.Text = lblStatus.Text + ". ";
            }
            
            if (tmrValue > 20)
            {
                tmrValue = 0;
            }       
        }

        private void cboChequeName_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cboChequeName.Text == "Customized Personal Checks")
            {
                lblBooks.Text = "Books: (50 pcs per Bkt):";
            }
            else
            {
                lblBooks.Text = "Books: (100 pcs per Bkt):";
            }

            if (cboChequeName.Text == "Manager's Check Continues")
            {
                lblBRSTN.Visible = false;
                txtBRSTN.Visible = false;
                lblBranchName.Visible = true;
                cboBranchName.Visible = true;

                //For Branch Name
                cboBranchName.Items.Clear();

                branchList = branchList.OrderBy(r => r.Address1).ToList();

                branchList.ForEach(r =>
                {;
                    cboBranchName.Items.Add(r.Address1);
                });
                //End For Branch Name
            }
            else
            {
                lblBRSTN.Visible = true;
                txtBRSTN.Visible = true;

                lblBranchName.Visible = false;
                cboBranchName.Visible = false;
            }
        }

        private void cboBranchName_SelectedIndexChanged(object sender, EventArgs e)
        {
            var branch = branchList.FirstOrDefault(r => r.Address1 == cboBranchName.Text);

            if(branch != null)
            {
                txtBRSTN.Text = branch.BRSTN;
            }
        }
    }
}