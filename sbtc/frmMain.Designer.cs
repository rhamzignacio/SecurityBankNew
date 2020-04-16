namespace sbtc
{
    partial class frmMain
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmMain));
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.dteDeliveryDate = new System.Windows.Forms.DateTimePicker();
            this.btnCheckFiles = new System.Windows.Forms.Button();
            this.btnEncode = new System.Windows.Forms.Button();
            this.lstFiles = new System.Windows.Forms.ListBox();
            this.backgroundWorker1 = new System.ComponentModel.BackgroundWorker();
            this.lblHashTotal = new System.Windows.Forms.Label();
            this.lblTotal = new System.Windows.Forms.Label();
            this.checkBoxSortRT = new System.Windows.Forms.CheckBox();
            this.txtBoxBatchNo = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.txtBoxExt = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.txtBoxProcessBy = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.statusStrip1 = new System.Windows.Forms.StatusStrip();
            this.lblStatus = new System.Windows.Forms.ToolStripStatusLabel();
            this.lblNotes = new System.Windows.Forms.Label();
            this.groupBox1.SuspendLayout();
            this.statusStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.dteDeliveryDate);
            this.groupBox1.ForeColor = System.Drawing.Color.White;
            this.groupBox1.Location = new System.Drawing.Point(12, 12);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(376, 78);
            this.groupBox1.TabIndex = 0;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Delivery Date:";
            // 
            // dteDeliveryDate
            // 
            this.dteDeliveryDate.Font = new System.Drawing.Font("Comic Sans MS", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.dteDeliveryDate.Location = new System.Drawing.Point(9, 30);
            this.dteDeliveryDate.Name = "dteDeliveryDate";
            this.dteDeliveryDate.Size = new System.Drawing.Size(354, 34);
            this.dteDeliveryDate.TabIndex = 0;
            this.dteDeliveryDate.ValueChanged += new System.EventHandler(this.dteDeliveryDate_ValueChanged);
            // 
            // btnCheckFiles
            // 
            this.btnCheckFiles.BackColor = System.Drawing.Color.Aquamarine;
            this.btnCheckFiles.Location = new System.Drawing.Point(394, 23);
            this.btnCheckFiles.Name = "btnCheckFiles";
            this.btnCheckFiles.Size = new System.Drawing.Size(199, 67);
            this.btnCheckFiles.TabIndex = 1;
            this.btnCheckFiles.Text = "Check Files on Head";
            this.btnCheckFiles.UseVisualStyleBackColor = false;
            this.btnCheckFiles.Click += new System.EventHandler(this.btnCheckFiles_Click);
            // 
            // btnEncode
            // 
            this.btnEncode.BackColor = System.Drawing.Color.White;
            this.btnEncode.Location = new System.Drawing.Point(600, 23);
            this.btnEncode.Name = "btnEncode";
            this.btnEncode.Size = new System.Drawing.Size(132, 67);
            this.btnEncode.TabIndex = 2;
            this.btnEncode.Text = "Encode";
            this.btnEncode.UseVisualStyleBackColor = false;
            this.btnEncode.Click += new System.EventHandler(this.btnEncode_Click);
            // 
            // lstFiles
            // 
            this.lstFiles.FormattingEnabled = true;
            this.lstFiles.ItemHeight = 23;
            this.lstFiles.Location = new System.Drawing.Point(16, 170);
            this.lstFiles.Name = "lstFiles";
            this.lstFiles.Size = new System.Drawing.Size(359, 303);
            this.lstFiles.TabIndex = 3;
            // 
            // backgroundWorker1
            // 
            this.backgroundWorker1.DoWork += new System.ComponentModel.DoWorkEventHandler(this.backgroundWorker1_DoWork);
            // 
            // lblHashTotal
            // 
            this.lblHashTotal.BackColor = System.Drawing.Color.Red;
            this.lblHashTotal.Font = new System.Drawing.Font("Comic Sans MS", 27.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblHashTotal.ForeColor = System.Drawing.Color.White;
            this.lblHashTotal.Location = new System.Drawing.Point(12, 476);
            this.lblHashTotal.Name = "lblHashTotal";
            this.lblHashTotal.Size = new System.Drawing.Size(719, 52);
            this.lblHashTotal.TabIndex = 5;
            this.lblHashTotal.Text = "1 Hash Total hasn\'t been Sent yet";
            this.lblHashTotal.Visible = false;
            this.lblHashTotal.Click += new System.EventHandler(this.lblHashTotal_Click);
            this.lblHashTotal.DoubleClick += new System.EventHandler(this.lblHashTotal_DoubleClick);
            // 
            // lblTotal
            // 
            this.lblTotal.ForeColor = System.Drawing.Color.White;
            this.lblTotal.Location = new System.Drawing.Point(381, 170);
            this.lblTotal.Name = "lblTotal";
            this.lblTotal.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.lblTotal.Size = new System.Drawing.Size(318, 305);
            this.lblTotal.TabIndex = 4;
            this.lblTotal.Click += new System.EventHandler(this.lblTotal_Click);
            // 
            // checkBoxSortRT
            // 
            this.checkBoxSortRT.AutoSize = true;
            this.checkBoxSortRT.Font = new System.Drawing.Font("Comic Sans MS", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.checkBoxSortRT.ForeColor = System.Drawing.Color.White;
            this.checkBoxSortRT.Location = new System.Drawing.Point(501, 102);
            this.checkBoxSortRT.Name = "checkBoxSortRT";
            this.checkBoxSortRT.Size = new System.Drawing.Size(231, 31);
            this.checkBoxSortRT.TabIndex = 8;
            this.checkBoxSortRT.Text = "Generate SortRT File";
            this.checkBoxSortRT.UseVisualStyleBackColor = true;
            // 
            // txtBoxBatchNo
            // 
            this.txtBoxBatchNo.Font = new System.Drawing.Font("Comic Sans MS", 20.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtBoxBatchNo.Location = new System.Drawing.Point(12, 96);
            this.txtBoxBatchNo.MaxLength = 8;
            this.txtBoxBatchNo.Name = "txtBoxBatchNo";
            this.txtBoxBatchNo.Size = new System.Drawing.Size(197, 45);
            this.txtBoxBatchNo.TabIndex = 9;
            this.txtBoxBatchNo.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtBoxBatchNo_KeyPress);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.ForeColor = System.Drawing.Color.White;
            this.label1.Location = new System.Drawing.Point(66, 144);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(77, 23);
            this.label1.TabIndex = 10;
            this.label1.Text = "Batch No";
            // 
            // txtBoxExt
            // 
            this.txtBoxExt.Font = new System.Drawing.Font("Comic Sans MS", 20.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtBoxExt.Location = new System.Drawing.Point(229, 96);
            this.txtBoxExt.MaxLength = 3;
            this.txtBoxExt.Name = "txtBoxExt";
            this.txtBoxExt.Size = new System.Drawing.Size(54, 45);
            this.txtBoxExt.TabIndex = 11;
            this.txtBoxExt.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtBoxExt_KeyPress);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.ForeColor = System.Drawing.Color.White;
            this.label2.Location = new System.Drawing.Point(238, 144);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(36, 23);
            this.label2.TabIndex = 12;
            this.label2.Text = "Ext";
            // 
            // txtBoxProcessBy
            // 
            this.txtBoxProcessBy.Font = new System.Drawing.Font("Comic Sans MS", 20.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtBoxProcessBy.Location = new System.Drawing.Point(298, 96);
            this.txtBoxProcessBy.Name = "txtBoxProcessBy";
            this.txtBoxProcessBy.Size = new System.Drawing.Size(197, 45);
            this.txtBoxProcessBy.TabIndex = 13;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.ForeColor = System.Drawing.Color.White;
            this.label3.Location = new System.Drawing.Point(350, 144);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(88, 23);
            this.label3.TabIndex = 14;
            this.label3.Text = "Process By";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.ForeColor = System.Drawing.Color.White;
            this.label4.Location = new System.Drawing.Point(212, 106);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(17, 23);
            this.label4.TabIndex = 15;
            this.label4.Text = "-";
            // 
            // statusStrip1
            // 
            this.statusStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.lblStatus});
            this.statusStrip1.Location = new System.Drawing.Point(0, 585);
            this.statusStrip1.Name = "statusStrip1";
            this.statusStrip1.Size = new System.Drawing.Size(743, 22);
            this.statusStrip1.TabIndex = 17;
            this.statusStrip1.Text = "statusStrip1";
            // 
            // lblStatus
            // 
            this.lblStatus.AutoSize = false;
            this.lblStatus.BackColor = System.Drawing.Color.White;
            this.lblStatus.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblStatus.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(107)))), ((int)(((byte)(250)))));
            this.lblStatus.Name = "lblStatus";
            this.lblStatus.Size = new System.Drawing.Size(730, 17);
            this.lblStatus.Text = "toolStripStatusLabel1";
            this.lblStatus.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // lblNotes
            // 
            this.lblNotes.ForeColor = System.Drawing.Color.White;
            this.lblNotes.Location = new System.Drawing.Point(16, 531);
            this.lblNotes.Name = "lblNotes";
            this.lblNotes.Size = new System.Drawing.Size(714, 54);
            this.lblNotes.TabIndex = 16;
            this.lblNotes.Text = "Note: Check \"Generate SortRT File for manual process\r\n         Do not use \"0000\" " +
    "for Batch No. for Testing only";
            // 
            // frmMain
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(10F, 23F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(107)))), ((int)(((byte)(250)))));
            this.ClientSize = new System.Drawing.Size(743, 607);
            this.Controls.Add(this.statusStrip1);
            this.Controls.Add(this.lblNotes);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.txtBoxProcessBy);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.txtBoxExt);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.txtBoxBatchNo);
            this.Controls.Add(this.checkBoxSortRT);
            this.Controls.Add(this.lblHashTotal);
            this.Controls.Add(this.lblTotal);
            this.Controls.Add(this.lstFiles);
            this.Controls.Add(this.btnEncode);
            this.Controls.Add(this.btnCheckFiles);
            this.Controls.Add(this.groupBox1);
            this.Font = new System.Drawing.Font("Comic Sans MS", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Margin = new System.Windows.Forms.Padding(5);
            this.MaximizeBox = false;
            this.Name = "frmMain";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Security Bank 2.0";
            this.Load += new System.EventHandler(this.frmMain_Load);
            this.groupBox1.ResumeLayout(false);
            this.statusStrip1.ResumeLayout(false);
            this.statusStrip1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.DateTimePicker dteDeliveryDate;
        private System.Windows.Forms.Button btnCheckFiles;
        private System.Windows.Forms.Button btnEncode;
        private System.Windows.Forms.ListBox lstFiles;
        private System.ComponentModel.BackgroundWorker backgroundWorker1;
        private System.Windows.Forms.Label lblHashTotal;
        private System.Windows.Forms.Label lblTotal;
        private System.Windows.Forms.CheckBox checkBoxSortRT;
        private System.Windows.Forms.TextBox txtBoxBatchNo;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox txtBoxExt;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox txtBoxProcessBy;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.StatusStrip statusStrip1;
        private System.Windows.Forms.ToolStripStatusLabel lblStatus;
        private System.Windows.Forms.Label lblNotes;
    }
}

