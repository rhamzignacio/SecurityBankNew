namespace sbtc
{
    partial class frmCustomized
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.lblLastSeries = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.cboBranchName = new System.Windows.Forms.ComboBox();
            this.lblBranchName = new System.Windows.Forms.Label();
            this.txtBRSTN = new System.Windows.Forms.TextBox();
            this.lblBRSTN = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.dteDeliveryDate = new System.Windows.Forms.DateTimePicker();
            this.txtStartingSeries = new System.Windows.Forms.TextBox();
            this.label6 = new System.Windows.Forms.Label();
            this.txtBooks = new System.Windows.Forms.TextBox();
            this.lblBooks = new System.Windows.Forms.Label();
            this.txtName2 = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.txtName1 = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.txtAccountNo = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.cboChequeName = new System.Windows.Forms.ComboBox();
            this.label1 = new System.Windows.Forms.Label();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.ChequeName = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.BRSTN = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.AccountNo = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.AccountName1 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.AccountName2 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Books = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.StartingSerial = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.btnAdd = new System.Windows.Forms.Button();
            this.btnProcess = new System.Windows.Forms.Button();
            this.backgroundWorker1 = new System.ComponentModel.BackgroundWorker();
            this.lblStatus = new System.Windows.Forms.Label();
            this.timer1 = new System.Windows.Forms.Timer(this.components);
            this.checkBoxTestOnly = new System.Windows.Forms.CheckBox();
            this.groupBox1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.SuspendLayout();
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.lblLastSeries);
            this.groupBox1.Controls.Add(this.label5);
            this.groupBox1.Controls.Add(this.cboBranchName);
            this.groupBox1.Controls.Add(this.lblBranchName);
            this.groupBox1.Controls.Add(this.txtBRSTN);
            this.groupBox1.Controls.Add(this.lblBRSTN);
            this.groupBox1.Controls.Add(this.label7);
            this.groupBox1.Controls.Add(this.dteDeliveryDate);
            this.groupBox1.Controls.Add(this.txtStartingSeries);
            this.groupBox1.Controls.Add(this.label6);
            this.groupBox1.Controls.Add(this.txtBooks);
            this.groupBox1.Controls.Add(this.lblBooks);
            this.groupBox1.Controls.Add(this.txtName2);
            this.groupBox1.Controls.Add(this.label4);
            this.groupBox1.Controls.Add(this.txtName1);
            this.groupBox1.Controls.Add(this.label3);
            this.groupBox1.Controls.Add(this.txtAccountNo);
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Controls.Add(this.cboChequeName);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Location = new System.Drawing.Point(12, 0);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(720, 358);
            this.groupBox1.TabIndex = 0;
            this.groupBox1.TabStop = false;
            // 
            // lblLastSeries
            // 
            this.lblLastSeries.AutoSize = true;
            this.lblLastSeries.ForeColor = System.Drawing.Color.White;
            this.lblLastSeries.Location = new System.Drawing.Point(602, 279);
            this.lblLastSeries.Name = "lblLastSeries";
            this.lblLastSeries.Size = new System.Drawing.Size(90, 23);
            this.lblLastSeries.TabIndex = 17;
            this.lblLastSeries.Text = "00000000";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.ForeColor = System.Drawing.Color.White;
            this.label5.Location = new System.Drawing.Point(501, 279);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(98, 23);
            this.label5.TabIndex = 16;
            this.label5.Text = "Last Series:";
            // 
            // cboBranchName
            // 
            this.cboBranchName.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cboBranchName.FormattingEnabled = true;
            this.cboBranchName.Location = new System.Drawing.Point(211, 112);
            this.cboBranchName.Name = "cboBranchName";
            this.cboBranchName.Size = new System.Drawing.Size(504, 31);
            this.cboBranchName.TabIndex = 4;
            this.cboBranchName.Visible = false;
            this.cboBranchName.SelectedIndexChanged += new System.EventHandler(this.cboBranchName_SelectedIndexChanged);
            // 
            // lblBranchName
            // 
            this.lblBranchName.AutoSize = true;
            this.lblBranchName.ForeColor = System.Drawing.Color.White;
            this.lblBranchName.Location = new System.Drawing.Point(6, 111);
            this.lblBranchName.Name = "lblBranchName";
            this.lblBranchName.Size = new System.Drawing.Size(110, 23);
            this.lblBranchName.TabIndex = 15;
            this.lblBranchName.Text = "Branch Name:";
            this.lblBranchName.Visible = false;
            // 
            // txtBRSTN
            // 
            this.txtBRSTN.Location = new System.Drawing.Point(211, 112);
            this.txtBRSTN.MaxLength = 9;
            this.txtBRSTN.Name = "txtBRSTN";
            this.txtBRSTN.Size = new System.Drawing.Size(503, 30);
            this.txtBRSTN.TabIndex = 3;
            this.txtBRSTN.TextChanged += new System.EventHandler(this.txtBRSTN_TextChanged);
            // 
            // lblBRSTN
            // 
            this.lblBRSTN.AutoSize = true;
            this.lblBRSTN.ForeColor = System.Drawing.Color.White;
            this.lblBRSTN.Location = new System.Drawing.Point(6, 111);
            this.lblBRSTN.Name = "lblBRSTN";
            this.lblBRSTN.Size = new System.Drawing.Size(71, 23);
            this.lblBRSTN.TabIndex = 14;
            this.lblBRSTN.Text = "BRSTN:";
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.ForeColor = System.Drawing.Color.White;
            this.label7.Location = new System.Drawing.Point(6, 317);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(118, 23);
            this.label7.TabIndex = 13;
            this.label7.Text = "Delivery Date:";
            // 
            // dteDeliveryDate
            // 
            this.dteDeliveryDate.Location = new System.Drawing.Point(211, 317);
            this.dteDeliveryDate.Name = "dteDeliveryDate";
            this.dteDeliveryDate.Size = new System.Drawing.Size(503, 30);
            this.dteDeliveryDate.TabIndex = 8;
            // 
            // txtStartingSeries
            // 
            this.txtStartingSeries.Location = new System.Drawing.Point(211, 276);
            this.txtStartingSeries.MaxLength = 10;
            this.txtStartingSeries.Name = "txtStartingSeries";
            this.txtStartingSeries.Size = new System.Drawing.Size(284, 30);
            this.txtStartingSeries.TabIndex = 7;
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.ForeColor = System.Drawing.Color.White;
            this.label6.Location = new System.Drawing.Point(5, 276);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(129, 23);
            this.label6.TabIndex = 10;
            this.label6.Text = "Starting Series:";
            // 
            // txtBooks
            // 
            this.txtBooks.Location = new System.Drawing.Point(211, 235);
            this.txtBooks.MaxLength = 3;
            this.txtBooks.Name = "txtBooks";
            this.txtBooks.Size = new System.Drawing.Size(503, 30);
            this.txtBooks.TabIndex = 6;
            // 
            // lblBooks
            // 
            this.lblBooks.AutoSize = true;
            this.lblBooks.ForeColor = System.Drawing.Color.White;
            this.lblBooks.Location = new System.Drawing.Point(5, 235);
            this.lblBooks.Name = "lblBooks";
            this.lblBooks.Size = new System.Drawing.Size(196, 23);
            this.lblBooks.TabIndex = 8;
            this.lblBooks.Text = "Books: (100 pcs per Bkt):";
            // 
            // txtName2
            // 
            this.txtName2.Location = new System.Drawing.Point(211, 194);
            this.txtName2.MaxLength = 50;
            this.txtName2.Name = "txtName2";
            this.txtName2.Size = new System.Drawing.Size(503, 30);
            this.txtName2.TabIndex = 5;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.ForeColor = System.Drawing.Color.White;
            this.label4.Location = new System.Drawing.Point(6, 194);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(129, 23);
            this.label4.TabIndex = 6;
            this.label4.Text = "Account Name2:";
            // 
            // txtName1
            // 
            this.txtName1.Location = new System.Drawing.Point(211, 154);
            this.txtName1.MaxLength = 50;
            this.txtName1.Name = "txtName1";
            this.txtName1.Size = new System.Drawing.Size(503, 30);
            this.txtName1.TabIndex = 4;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.ForeColor = System.Drawing.Color.White;
            this.label3.Location = new System.Drawing.Point(6, 154);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(126, 23);
            this.label3.TabIndex = 4;
            this.label3.Text = "Account Name1:";
            // 
            // txtAccountNo
            // 
            this.txtAccountNo.Location = new System.Drawing.Point(211, 69);
            this.txtAccountNo.MaxLength = 12;
            this.txtAccountNo.Name = "txtAccountNo";
            this.txtAccountNo.Size = new System.Drawing.Size(503, 30);
            this.txtAccountNo.TabIndex = 2;
            this.txtAccountNo.TextChanged += new System.EventHandler(this.txtAccountNo_TextChanged);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.ForeColor = System.Drawing.Color.White;
            this.label2.Location = new System.Drawing.Point(6, 68);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(137, 23);
            this.label2.TabIndex = 2;
            this.label2.Text = "Account Number:";
            // 
            // cboChequeName
            // 
            this.cboChequeName.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cboChequeName.FormattingEnabled = true;
            this.cboChequeName.Location = new System.Drawing.Point(211, 26);
            this.cboChequeName.Name = "cboChequeName";
            this.cboChequeName.Size = new System.Drawing.Size(503, 31);
            this.cboChequeName.TabIndex = 1;
            this.cboChequeName.SelectedIndexChanged += new System.EventHandler(this.cboChequeName_SelectedIndexChanged);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.ForeColor = System.Drawing.Color.White;
            this.label1.Location = new System.Drawing.Point(6, 26);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(110, 23);
            this.label1.TabIndex = 0;
            this.label1.Text = "Cheque Name:";
            // 
            // dataGridView1
            // 
            this.dataGridView1.AllowUserToAddRows = false;
            this.dataGridView1.AllowUserToDeleteRows = false;
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.ChequeName,
            this.BRSTN,
            this.AccountNo,
            this.AccountName1,
            this.AccountName2,
            this.Books,
            this.StartingSerial});
            this.dataGridView1.Location = new System.Drawing.Point(13, 364);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.Size = new System.Drawing.Size(719, 161);
            this.dataGridView1.TabIndex = 1;
            // 
            // ChequeName
            // 
            this.ChequeName.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.ChequeName.HeaderText = "Cheque Name";
            this.ChequeName.Name = "ChequeName";
            this.ChequeName.ReadOnly = true;
            this.ChequeName.Width = 130;
            // 
            // BRSTN
            // 
            this.BRSTN.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.BRSTN.HeaderText = "BRSTN";
            this.BRSTN.Name = "BRSTN";
            this.BRSTN.ReadOnly = true;
            this.BRSTN.Width = 91;
            // 
            // AccountNo
            // 
            this.AccountNo.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.AccountNo.HeaderText = "Account Number";
            this.AccountNo.Name = "AccountNo";
            this.AccountNo.ReadOnly = true;
            this.AccountNo.Width = 157;
            // 
            // AccountName1
            // 
            this.AccountName1.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.AccountName1.HeaderText = "Account Name 1";
            this.AccountName1.Name = "AccountName1";
            this.AccountName1.ReadOnly = true;
            this.AccountName1.Width = 151;
            // 
            // AccountName2
            // 
            this.AccountName2.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.AccountName2.HeaderText = "Account Name 2";
            this.AccountName2.Name = "AccountName2";
            this.AccountName2.ReadOnly = true;
            this.AccountName2.Width = 154;
            // 
            // Books
            // 
            this.Books.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.Books.HeaderText = "Books";
            this.Books.Name = "Books";
            this.Books.ReadOnly = true;
            this.Books.Width = 77;
            // 
            // StartingSerial
            // 
            this.StartingSerial.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.StartingSerial.HeaderText = "Starting Serial";
            this.StartingSerial.Name = "StartingSerial";
            this.StartingSerial.ReadOnly = true;
            this.StartingSerial.Width = 147;
            // 
            // btnAdd
            // 
            this.btnAdd.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(255)))), ((int)(((byte)(255)))));
            this.btnAdd.Location = new System.Drawing.Point(12, 531);
            this.btnAdd.Name = "btnAdd";
            this.btnAdd.Size = new System.Drawing.Size(150, 37);
            this.btnAdd.TabIndex = 9;
            this.btnAdd.Text = "&Add";
            this.btnAdd.UseVisualStyleBackColor = false;
            this.btnAdd.Click += new System.EventHandler(this.btnAdd_Click);
            // 
            // btnProcess
            // 
            this.btnProcess.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(255)))), ((int)(((byte)(255)))));
            this.btnProcess.Enabled = false;
            this.btnProcess.Location = new System.Drawing.Point(582, 531);
            this.btnProcess.Name = "btnProcess";
            this.btnProcess.Size = new System.Drawing.Size(150, 37);
            this.btnProcess.TabIndex = 10;
            this.btnProcess.Text = "&Process";
            this.btnProcess.UseVisualStyleBackColor = false;
            this.btnProcess.Click += new System.EventHandler(this.btnProcess_Click);
            // 
            // backgroundWorker1
            // 
            this.backgroundWorker1.DoWork += new System.ComponentModel.DoWorkEventHandler(this.backgroundWorker1_DoWork);
            this.backgroundWorker1.ProgressChanged += new System.ComponentModel.ProgressChangedEventHandler(this.backgroundWorker1_ProgressChanged);
            this.backgroundWorker1.RunWorkerCompleted += new System.ComponentModel.RunWorkerCompletedEventHandler(this.backgroundWorker1_RunWorkerCompleted);
            // 
            // lblStatus
            // 
            this.lblStatus.AutoSize = true;
            this.lblStatus.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(128)))), ((int)(((byte)(128)))));
            this.lblStatus.Location = new System.Drawing.Point(168, 545);
            this.lblStatus.Name = "lblStatus";
            this.lblStatus.Size = new System.Drawing.Size(86, 23);
            this.lblStatus.TabIndex = 11;
            this.lblStatus.Text = ". . . . . . . . .";
            this.lblStatus.Visible = false;
            // 
            // timer1
            // 
            this.timer1.Tick += new System.EventHandler(this.timer1_Tick);
            // 
            // checkBoxTestOnly
            // 
            this.checkBoxTestOnly.AutoSize = true;
            this.checkBoxTestOnly.Location = new System.Drawing.Point(474, 537);
            this.checkBoxTestOnly.Name = "checkBoxTestOnly";
            this.checkBoxTestOnly.Size = new System.Drawing.Size(102, 27);
            this.checkBoxTestOnly.TabIndex = 13;
            this.checkBoxTestOnly.Text = "Test Only";
            this.checkBoxTestOnly.UseVisualStyleBackColor = true;
            // 
            // frmCustomized
            // 
            this.AcceptButton = this.btnAdd;
            this.AutoScaleDimensions = new System.Drawing.SizeF(10F, 23F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(107)))), ((int)(((byte)(250)))));
            this.ClientSize = new System.Drawing.Size(744, 576);
            this.Controls.Add(this.checkBoxTestOnly);
            this.Controls.Add(this.lblStatus);
            this.Controls.Add(this.btnProcess);
            this.Controls.Add(this.btnAdd);
            this.Controls.Add(this.dataGridView1);
            this.Controls.Add(this.groupBox1);
            this.Font = new System.Drawing.Font("Comic Sans MS", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Margin = new System.Windows.Forms.Padding(5);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "frmCustomized";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Customized";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.frmCustomized_FormClosed);
            this.Load += new System.EventHandler(this.frmCustomized_Load);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.ComboBox cboChequeName;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox txtAccountNo;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox txtName2;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox txtName1;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox txtStartingSeries;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.TextBox txtBooks;
        private System.Windows.Forms.Label lblBooks;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.DateTimePicker dteDeliveryDate;
        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.Button btnAdd;
        private System.Windows.Forms.Button btnProcess;
        private System.Windows.Forms.TextBox txtBRSTN;
        private System.Windows.Forms.Label lblBRSTN;
        private System.ComponentModel.BackgroundWorker backgroundWorker1;
        private System.Windows.Forms.Label lblStatus;
        private System.Windows.Forms.Timer timer1;
        private System.Windows.Forms.ComboBox cboBranchName;
        private System.Windows.Forms.Label lblBranchName;
        private System.Windows.Forms.DataGridViewTextBoxColumn ChequeName;
        private System.Windows.Forms.DataGridViewTextBoxColumn BRSTN;
        private System.Windows.Forms.DataGridViewTextBoxColumn AccountNo;
        private System.Windows.Forms.DataGridViewTextBoxColumn AccountName1;
        private System.Windows.Forms.DataGridViewTextBoxColumn AccountName2;
        private System.Windows.Forms.DataGridViewTextBoxColumn Books;
        private System.Windows.Forms.DataGridViewTextBoxColumn StartingSerial;
        private System.Windows.Forms.Label lblLastSeries;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.CheckBox checkBoxTestOnly;
    }
}