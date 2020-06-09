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

namespace sbtc
{
    public partial class frmCustomized : Form
    {
        int tmrValue = 0;

        List<BranchesModel> branchList;

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
            cboChequeName.Items.Add("Manager's Check Continues");
            cboChequeName.Items.Add("Dividend Checks");



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

                string dbase = "";
                if (ReturnMe.CodesOnly == true) { dbase = "captive_database.Master_Database_SBTC_Temp"; }
                if (ReturnMe.CodesOnly == false) { dbase = "captive_database.Master_Database_SBTC"; }

                string sql = "SELECT BRSTN, Name1, Name2 FROM " + dbase + " WHERE AccountNo = '" + txtAccountNo.Text + "' ORDER BY PrimaryKey DESC LIMIT 1";
                string MyConnection2 = "datasource=" + ReturnMe.server + ";port=3306;username=" + ReturnMe.uid + ";password=" + ReturnMe.password;
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
            int LoopCount = 0;

            ReturnMe.CreateTable();

            while (LoopCount < dataGridView1.Rows.Count)
            {


                string chequename = dataGridView1.Rows[LoopCount].Cells[0].Value.ToString();
                string brstn = dataGridView1.Rows[LoopCount].Cells[1].Value.ToString();
                string AccountNo = dataGridView1.Rows[LoopCount].Cells[2].Value.ToString();
                string Name1 = dataGridView1.Rows[LoopCount].Cells[3].Value.ToString();
                string Name2 = dataGridView1.Rows[LoopCount].Cells[4].Value.ToString();
                string Books = dataGridView1.Rows[LoopCount].Cells[5].Value.ToString();
                string StartingSerial = dataGridView1.Rows[LoopCount].Cells[6].Value.ToString();


                string sql = "INSERT INTO SBTC (BRSTN, ChkType, AccountNo , Name1 , Name2 , FormType,OrderQty,Pkey, StartSN , Address1 , Address2 , Address3 , Address4 , Address5 , Address6) VALUES ('" + brstn + "','" + chequename + "','" + AccountNo + "','" + Name1.Replace("'", "''") + "','" + Name2.Replace("'", "''") + "','00','" + Books + "'," + (LoopCount + 1).ToString() + ",'" + StartingSerial + "','" + ReturnMe.getAddress1(brstn, 1, chequename).Replace("'", "''") + "','" + ReturnMe.getAddress1(brstn, 2, chequename).Replace("'", "''") + "','" + ReturnMe.getAddress1(brstn, 3, chequename).Replace("'", "''") + "','" + ReturnMe.getAddress1(brstn, 4, chequename).Replace("'", "''") + "','" + ReturnMe.getAddress1(brstn, 5, chequename).Replace("'", "''") + "','" + ReturnMe.getAddress1(brstn, 6, chequename).Replace("'", "''") + "')";
                OleDbConnection conn = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + Application.StartupPath + ";Extended Properties=dBASE III;");
                OleDbCommand command = new OleDbCommand(sql, conn);
                conn.Open();
                command.ExecuteReader();
                conn.Close();
                
                LoopCount = LoopCount + 1;

            }

            btnAdd.Enabled = false;
            btnProcess.Enabled = false;

            ReturnMe.ActivateMaxProgressBarCarbon();

            backgroundWorker1.RunWorkerAsync();
            timer1.Start();

            
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
            progressBar1.Maximum = ReturnMe.mdb_status_bar_max;
            progressBar1.Minimum = 0;
            progressBar1.Value = ReturnMe.mdb_status_bar;

            if (progressBar1.Value != 0 && progressBar1.Value != progressBar1.Maximum)
            {
                progressBar1.Visible = true;
                lblStatus.Visible = true;
            }
            else
            {
                progressBar1.Visible = false;
                lblStatus.Visible = false;
            }


            
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
