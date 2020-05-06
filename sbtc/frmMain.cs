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
using System.Text.RegularExpressions;

namespace sbtc
{
    public partial class frmMain : Form
    {        
        int pkey = 0;

        string errorMessage = "";

        List<BranchesModel> branchList = new List<BranchesModel>();

        OrderSorted sortedList = new OrderSorted();

        string AutoBatch = "";

        public frmMain()
        {
            InitializeComponent();

            GetAllBranch();

            getSettings();

            ListFilesHead();

            lblStatus.Text = "Ready . . .";

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
            fileStream = new FileStream("C:\\Auto\\WinZip.txt", FileMode.Open, FileAccess.Read);
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
            string[] files = Directory.GetFiles(Application.StartupPath);

            foreach (string file in files)
            {
                if (file.Contains(".SQL"))
                    File.Delete(file);
            }
        }//END OF FUNCTION

        private void CheckHashTotal()
        {
            string dbase = "";
            if (ReturnMe.CodesOnly == true) dbase = "captive_database.sbtc_history";
            if (ReturnMe.CodesOnly == false) dbase = "captive_database.sbtc_history";

            string sql = "SELECT DISTINCT(FinalBatch) FROM " + dbase + " WHERE HashSentDate is NULL AND HashSentTime IS NULL";
            //string MyConnection2 = "datasource=" + ReturnMe.server + ";port=3306;username=" + ReturnMe.uid + ";password=" + ReturnMe.password;

            string MyConnection2 = "datasource=192.168.0.254;port=3306;username=root;password=CorpCaptive;";
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

            AutoBatch = files[0].Substring(6, 4);


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

            GenerateService.CheckPaths();

            if (txtBoxBatchNo.Text.ToUpper() != "0000")
            {
                if (dteDeliveryDate.Value.ToShortDateString() == DateTime.Today.ToShortDateString())
                {
                    DialogResult result3 = MessageBox.Show("Please select Delivery Date", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);

                    dteDeliveryDate.Focus();

                    return;
                }
                else if(txtBoxBatchNo.Text == "")
                {
                    DialogResult result3 = MessageBox.Show("Please input Batch No: ", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);

                    dteDeliveryDate.Focus();

                    return;
                }
                else if(txtBoxProcessBy.Text == "")
                {
                    DialogResult result3 = MessageBox.Show("Please input Process by: ", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);

                    dteDeliveryDate.Focus();

                    return;
                }
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

                lblStatus.Text = "Now Checking Files in Head Folder . . .";

                while (LoopCount < lstFiles.Items.Count)
                {
                    string filename = lstFiles.Items[LoopCount].ToString();

                    OrderList.AddRange(CheckFiles(filename));

                    LoopCount = LoopCount + 1;
                }

                CheckBranches(OrderList);
              
                DisplayData(OrderList);

                if(CheckingService.CheckBatchIfDuplicate(txtBoxBatchNo.Text))
                {
                    MessageBox.Show("Batch No is already been use ! ! !", "Error");
                }
                else
                {
                    btnCheckFiles.Text = "Process ! ! !";

                    MessageBox.Show("Data has been Checked. No Errors found!", " ", MessageBoxButtons.OK, MessageBoxIcon.Information);

                    lblStatus.Text = "Ready . . .";
                }

                if (checkBoxSortRT.Checked == true)
                {
                    //For Sort RT
                    List<OrderModel> sortTemp = new List<OrderModel>();

                    sortTemp.AddRange(sortedList.RegularPersonal);
                    sortTemp.AddRange(sortedList.RegularCommercial);
                    GenerateService.GenerateSortRT("Regular", sortTemp);

                    sortTemp = new List<OrderModel>();
                    sortTemp.AddRange(sortedList.PersonalPreEncoded);
                    sortTemp.AddRange(sortedList.CommercialPreEncoded);
                    GenerateService.GenerateSortRT("Regular\\PreEncoded", sortTemp);

                    GenerateService.GenerateSortRT("MC", sortedList.ManagersCheck);

                    sortTemp = new List<OrderModel>();
                    sortTemp.AddRange(sortedList.CheckOnePersonal);
                    sortTemp.AddRange(sortedList.CheckOneCommerical);
                    GenerateService.GenerateSortRT("CheckOne", sortTemp);

                    sortTemp = new List<OrderModel>();
                    sortTemp.AddRange(sortedList.CheckPowerPersonal);
                    sortTemp.AddRange(sortedList.CheckPowerCommercial);
                    GenerateService.GenerateSortRT("CheckPower", sortTemp);

                    GenerateService.GenerateSortRT("GiftCheck", sortedList.GiftCheck);

                    GenerateService.GenerateSortRT("MC\\Continues", sortedList.ManagersCheckCont);

                    MessageBox.Show("SortRT File has been successfully generated", " ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }                            
            }
            else
            {               
                btnCheckFiles.Enabled = false;
                btnEncode.Enabled = false;


                //GENERATE PRINTER FILES
                lblStatus.Text = "Generating Printer Files";
                Application.DoEvents();
                AddSerials();
                GenerateService.GeneratePrinterFiles(sortedList, txtBoxBatchNo.Text, txtBoxExt.Text);

                //GENERATE MDB FILE FOR Manager's Check
                if (sortedList.ManagersCheck.Count > 0 || sortedList.ManagersCheckCont.Count > 0)
                {                    
                    lblStatus.Text = "Generating MDB Files for Manager's Check";
                    Application.DoEvents();
                    GenerateService.GenerateMDBFile(sortedList, txtBoxBatchNo.Text, txtBoxExt.Text);
                }

                //GENERATE PACKING DBF
                lblStatus.Text = "Generating Packing DBF Files";
                Application.DoEvents();
                GenerateService.GeneratePackingDBF(sortedList, txtBoxBatchNo.Text, txtBoxExt.Text);

                //GENERATE DO BLOCK
                lblStatus.Text = "Generating DoBlock Files";
                Application.DoEvents();
                GenerateService.GenerateDoBlock(sortedList, txtBoxBatchNo.Text, txtBoxExt.Text, dteDeliveryDate.Value, txtBoxProcessBy.Text);

                //GENERATE PACKINGLIST
                lblStatus.Text = "Generating PackingList Files";
                Application.DoEvents();
                GenerateService.GeneratePackingList(sortedList, txtBoxBatchNo.Text, dteDeliveryDate.Value, branchList);

                if (txtBoxBatchNo.Text != "0000")
                {
                    //SAVE HISTORY
                    lblStatus.Text = "Saving History";
                    Application.DoEvents();
                    BackupService.SaveHistory(sortedList, txtBoxBatchNo.Text, txtBoxExt.Text, dteDeliveryDate.Value);

                    //SAVE NEW SERIES ON DATABASE
                    lblStatus.Text = "Saving New Serial in Database . . .";
                    Application.DoEvents();
                    BackupService.SaveNewSeries(branchList);

                    //WinZIP Process
                    lblStatus.Text = "Archiving Output Files . . .";
                    Application.DoEvents();
                    BackupService.ProcessArchiving(AutoBatch, txtBoxProcessBy.Text, sortedList);

                    DeleteHeadFiles();//DELETE FILES IN HEAD FOLDER
                }


                lblStatus.Text = "Processing Done. . .";

                MessageBox.Show("Processing Done . . .", " ", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }//END OF FUNCTION

        private void DeleteHeadFiles()
        {
            DirectoryInfo di = new DirectoryInfo(Application.StartupPath + "\\Head");

            foreach(FileInfo file in di.GetFiles())
            {
                file.Delete();
            }
        }
    

        private void DisplayData(List<OrderModel> _orders)
        {
            sortedList = ReturnMe.Sort(_orders);

            string display = "";

            if (sortedList.RegularPersonal.Count > 0)
                display += "Regular Personal - " + sortedList.RegularPersonal.Count.ToString() + "\n";

            if (sortedList.RegularCommercial.Count > 0)
                display += "Regular Commercial - " + sortedList.RegularCommercial.Count.ToString() + "\n";

            if (sortedList.ManagersCheck.Count > 0)
                display += "Manager's Check - " + sortedList.ManagersCheck.Count.ToString() + "\n";

            if (sortedList.GiftCheck.Count > 0)
                display += "Gift Check - " + sortedList.GiftCheck.Count.ToString() + "\n";

            if (sortedList.PersonalPreEncoded.Count > 0)
                display += "Personal Pre-Encoded - " + sortedList.PersonalPreEncoded.Count.ToString() + "\n";

            if (sortedList.CommercialPreEncoded.Count > 0)
                display += "Commercial Pre-Encoded - " + sortedList.CommercialPreEncoded.Count.ToString() + "\n";

            if (sortedList.CheckOnePersonal.Count > 0)
                display += "CheckOne Personal - " + sortedList.CheckOnePersonal.Count.ToString() + "\n";

            if (sortedList.CheckOneCommerical.Count > 0)
                display += "CheckOne Commercial - " + sortedList.CheckOneCommerical.Count.ToString() + "\n";

            if (sortedList.CheckPowerPersonal.Count > 0)
                display += "CheckPower Personal - " + sortedList.CheckPowerPersonal.Count.ToString() + "\n";

            if (sortedList.CheckPowerCommercial.Count > 0)
                display += "CheckPower Commercial - " + sortedList.CheckOneCommerical.Count.ToString() + "\n";

            if(sortedList.CustomizedCheck.Count> 0)
                display += "Customized Check - " + sortedList.CustomizedCheck.Count.ToString() + "\n";

            if (sortedList.ManagersCheckCont.Count > 0)
                display += "Manager's Check Continous  - " + sortedList.ManagersCheckCont.Count.ToString() + "\n";

            lblTotal.Text = display;
        }

        private List<OrderModel> CheckBranches(List<OrderModel> _orders)
        {
            _orders.ForEach(order =>
            {
                var branch = branchList.FirstOrDefault(r => r.BRSTN == order.BRSTN);

                if(branch == null)
                {
                    errorMessage += "\nBRSTN=" + order.BRSTN + " is not found on BRANCHES DATABASE";
                }
                else //ADD BRANCHES INFO
                {
                    order.Address1 = branch.Address1;

                    order.Address2 = branch.Address2;

                    order.Address3 = branch.Address3;

                    order.Address4 = branch.Address4;

                    order.Address5 = branch.Address5;

                    order.Address6 = branch.Address6;
                }
            });

            return _orders;
        }

        private void AddSerials()
        {
            if(sortedList.RegularPersonal.Count > 0)
            {
                sortedList.RegularPersonal = sortedList.RegularPersonal.OrderBy(r => r.BRSTN).ThenBy(r => r.AccountNo).ToList();

                sortedList.RegularPersonal.ForEach(r =>
                {
                    var branch = branchList.FirstOrDefault(c => c.BRSTN == r.BRSTN);

                    var series = branch.LastNo_PA;

                    r.StartingSerial = series + 1;

                    r.EndingSerial = series + QuantityPerBooklet.RegularPersonal;

                    branch.LastNo_PA += QuantityPerBooklet.RegularPersonal;

                    branch.IfChanges = 1;
                });
            }//END IF

            if(sortedList.RegularCommercial.Count > 0)
            {
                sortedList.RegularCommercial = sortedList.RegularCommercial.OrderBy(r => r.BRSTN).ThenBy(r => r.AccountNo).ToList();

                sortedList.RegularCommercial.ForEach(r =>
                {
                    var branch = branchList.FirstOrDefault(c => c.BRSTN == r.BRSTN);

                    var series = branch.LastNo_CA;

                    r.StartingSerial = series + 1;

                    r.EndingSerial = series + QuantityPerBooklet.RegularCommercial;

                    branch.LastNo_CA += QuantityPerBooklet.RegularCommercial;

                    branch.IfChanges = 1;
                });
            }//END IF

            if(sortedList.PersonalPreEncoded.Count > 0)
            {
                sortedList.PersonalPreEncoded = sortedList.PersonalPreEncoded.OrderBy(r => r.BRSTN).ThenBy(r => r.AccountNo).ToList();

                sortedList.PersonalPreEncoded.ForEach(r =>
                {
                    var branch = branchList.FirstOrDefault(c => c.BRSTN == r.BRSTN);

                    var series = branch.LastNo_PA;

                    r.StartingSerial = series + 1;

                    r.EndingSerial = series + QuantityPerBooklet.RegularPersonalPre;

                    branch.LastNo_PA += QuantityPerBooklet.RegularPersonalPre;

                    branch.IfChanges = 1;
                });
            }//END IF

            if(sortedList.CommercialPreEncoded.Count > 0)
            {
                sortedList.CommercialPreEncoded = sortedList.CommercialPreEncoded.OrderBy(r => r.BRSTN).ThenBy(r => r.AccountNo).ToList();

                sortedList.CommercialPreEncoded.ForEach(r =>
                {
                    var branch = branchList.FirstOrDefault(c => c.BRSTN == r.BRSTN);

                    var series = branch.LastNo_CA;

                    r.StartingSerial = series + 1;

                    r.EndingSerial = series + QuantityPerBooklet.RegularCommercialPre;

                    branch.LastNo_CA += QuantityPerBooklet.RegularCommercialPre;

                    branch.IfChanges = 1;
                });
            }//END IF

            if(sortedList.CheckOnePersonal.Count > 0)
            {
                sortedList.CheckOnePersonal = sortedList.CheckOnePersonal.OrderBy(r => r.BRSTN).ThenBy(r => r.AccountNo).ToList();

                sortedList.CheckOnePersonal.ForEach(r =>
                {
                    var branch = branchList.FirstOrDefault(c => c.BRSTN == r.BRSTN);

                    var series = branch.LastNo_CheckOne_PA;

                    r.StartingSerial = series + 1;

                    r.EndingSerial = series + QuantityPerBooklet.CheckOnePersonal;

                    branch.LastNo_CheckOne_PA += QuantityPerBooklet.CheckOnePersonal;

                    branch.IfChanges = 1;
                });
            }//END IF

            if(sortedList.CheckOneCommerical.Count > 0)
            {
                sortedList.CheckOneCommerical = sortedList.CheckOneCommerical.OrderBy(r => r.BRSTN).ThenBy(r => r.AccountNo).ToList();

                sortedList.CheckOneCommerical.ForEach(r =>
                {
                    var branch = branchList.FirstOrDefault(c => c.BRSTN == r.BRSTN);

                    var series = branch.LastNo_CheckOne_CA;

                    r.StartingSerial = series + 1;

                    r.EndingSerial = series + QuantityPerBooklet.CheckOneCommercial;

                    branch.LastNo_CheckOne_CA += QuantityPerBooklet.CheckOneCommercial;

                    branch.IfChanges = 1;
                });
            }//END IF

            if(sortedList.CheckPowerPersonal.Count > 0)
            {
                sortedList.CheckPowerPersonal = sortedList.CheckPowerPersonal.OrderBy(r => r.BRSTN).ThenBy(r => r.AccountNo).ToList();

                sortedList.CheckPowerPersonal.ForEach(r =>
                {
                    var branch = branchList.FirstOrDefault(c => c.BRSTN == r.BRSTN);

                    var series = branch.LastNo_Power_PA;

                    r.StartingSerial = series + 1;

                    r.EndingSerial = series + QuantityPerBooklet.CheckPowerPersonal;

                    branch.LastNo_Power_PA += QuantityPerBooklet.CheckPowerPersonal;

                    branch.IfChanges = 1;
                });
            }//END IF

            if(sortedList.CheckPowerCommercial.Count > 0)
            {
                sortedList.CheckPowerCommercial = sortedList.CheckPowerCommercial.OrderBy(r => r.BRSTN).ThenBy(r => r.AccountNo).ToList();

                sortedList.CheckPowerCommercial.ForEach(r =>
                {
                    var branch = branchList.FirstOrDefault(c => c.BRSTN == r.BRSTN);

                    var series = branch.LastNo_Power_CA;

                    r.StartingSerial = series + 1;

                    r.EndingSerial = series + QuantityPerBooklet.CheckPowerCommercial;

                    branch.LastNo_Power_CA += QuantityPerBooklet.CheckPowerCommercial;

                    branch.IfChanges = 1;
                });
            }//END IF

            if(sortedList.ManagersCheck.Count > 0)
            {
                sortedList.ManagersCheck = sortedList.ManagersCheck.OrderBy(r => r.BRSTN).ThenBy(r => r.AccountNo).ToList();

                sortedList.ManagersCheck.ForEach(r =>
                {
                    var branch = branchList.FirstOrDefault(c => c.BRSTN == r.BRSTN);

                    var series = branch.LastNo_MC;

                    r.StartingSerial = series + 1;

                    r.EndingSerial = series + QuantityPerBooklet.RegularPersonal;

                    branch.LastNo_MC += QuantityPerBooklet.RegularPersonal;

                    branch.IfChanges = 1;
                });
            }//END IF

            if(sortedList.ManagersCheckCont.Count > 0)
            {
                sortedList.ManagersCheckCont = sortedList.ManagersCheckCont.OrderBy(r => r.BRSTN).ThenBy(r => r.AccountNo).ToList();

                sortedList.ManagersCheckCont.ForEach(r =>
                {
                    var branch = branchList.FirstOrDefault(c => c.BRSTN == r.BRSTN);

                    var series = branch.LastNo_MC;

                    r.StartingSerial = series + 1;

                    r.EndingSerial = series + QuantityPerBooklet.ManagersCheckCont;

                    branch.LastNo_MC += QuantityPerBooklet.ManagersCheckCont;

                    branch.IfChanges = 1;
                });
            }//END IF
        }

        private List<OrderModel> CheckFiles(string filename)
        {
            string Batch = filename.Substring(0, filename.Length - 4).ToUpper();

            List<OrderModel> returnList = new List<OrderModel>();

            List<OrderModel> tempList = new List<OrderModel>();

            var fileStream = new FileStream(Application.StartupPath + "\\Head\\" + filename, FileMode.Open, FileAccess.Read);

            using (var streamReader = new StreamReader(fileStream, Encoding.UTF8))
            {
                string line;
                while ((line = streamReader.ReadLine()) != null)
                {
                    string temp = line.Substring(82, 2).Trim();

                    if (!int.TryParse(temp, out int test))
                    {
                        errorMessage += "-Error Parsing Quantity of BRSTN=" + line.Substring(1, 9).Trim() + " AccountNo=" + line.Substring(11, 12).Trim() + " FileName=" + filename;
                    }

                    for (int x = 0; x < test; x++)
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

                        order.OrderQuantity = int.Parse(line.Substring(82, 2).Trim());

                        order.FileName = filename;

                        order.Batch = Batch;

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
                            else if ((order.CheckType == "B" && order.FormType == "20" && order.AccountNo.Substring(4, 3) == "212") ||
                                (order.CheckType == "B" && order.FormType == "20" && order.AccountNo.Substring(0, 1) == "9" && order.AccountNo.Substring(5, 7) == "2000022"))
                            {
                                order.CheckType = "GC";
                                order.Name = "";
                            }
                            else if (order.CheckType == "B" && order.FormType == "20" && order.AccountNo.Substring(4, 3) != "212")
                            {
                                order.CheckType = "MC";
                            }

                            if (order.ContCode == "" || order.ContCode == "1")
                            {
                                returnList.Add(order);
                            }
                            else if (order.ContCode == "2")
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
                    }//ENDFOR
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

        private void txtBoxBatchNo_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (Regex.IsMatch(e.KeyChar.ToString(), @"[^0-9^\b]"))
                e.Handled = true;
        }

        private void txtBoxExt_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !(char.IsLetter(e.KeyChar) || e.KeyChar == (char)Keys.Back);
        }

        private void timer1_Tick(object sender, EventArgs e)
        {

        }

        private void lblTotal_Click(object sender, EventArgs e)
        {

        }

        private void dteDeliveryDate_ValueChanged(object sender, EventArgs e)
        {

        }
    }
}
