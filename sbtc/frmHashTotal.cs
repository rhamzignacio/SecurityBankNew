using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using MySql.Data;
using MySql.Data.MySqlClient;
using System.IO;

namespace sbtc
{
    public partial class frmHashTotal : Form
    {
        public frmHashTotal()
        {
            InitializeComponent();
        }

        private void frmHashTotal_Load(object sender, EventArgs e)
        {
            LoadMe();
        }

        private void LoadMe()
        {
            string dbase = "";
            if (ReturnMe.CodesOnly == true) dbase = "captive_database.master_database_sbtc_temp";
            if (ReturnMe.CodesOnly == false) dbase = "captive_database.master_database_sbtc";

            
            string sql = "SELECT DISTINCT(DeliveryDate) FROM " + dbase + " WHERE HashSentDate is NULL AND HashSentTime IS NULL AND DeliveryDate <= '" + DateTime.Now.ToString("yyyy-MM-dd") + "'";
            string MyConnection2 = "datasource=" + ReturnMe.server + ";port=3306;username=" + ReturnMe.uid + ";password=" + ReturnMe.password;
            MySqlConnection MyConn2 = new MySqlConnection(MyConnection2);
            MySqlCommand MyCommand2 = new MySqlCommand(sql, MyConn2);
            MySqlDataReader MyReader2;
            MyConn2.Open();
            MyReader2 = MyCommand2.ExecuteReader();

            cboDeliveryDate.Items.Clear();


            int LoopCount = 0;

            while (MyReader2.Read())
            {
                string DeliveryDate = MyReader2.GetDateTime(0).ToString("yyyy-MM-dd");

                cboDeliveryDate.Items.Add(DeliveryDate);

                LoopCount = LoopCount + 1;
            }
        }




        private void frmHashTotal_FormClosed(object sender, FormClosedEventArgs e)
        {
            Application.Exit();
        }



        private void btnSendHashTotal_Click(object sender, EventArgs e)
        {   
            if (cboDeliveryDate.Text == "")
            {
                MessageBox.Show("Please select Batch", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                cboDeliveryDate.Focus();
                return;
            }

            ReturnMe.SendHashTotal(cboDeliveryDate.Text);
        }
    }
}
