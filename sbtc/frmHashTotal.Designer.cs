namespace sbtc
{
    partial class frmHashTotal
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
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.btnSendHashTotal = new System.Windows.Forms.Button();
            this.cboDeliveryDate = new System.Windows.Forms.ComboBox();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.btnSendHashTotal);
            this.groupBox1.Controls.Add(this.cboDeliveryDate);
            this.groupBox1.Font = new System.Drawing.Font("Comic Sans MS", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.groupBox1.ForeColor = System.Drawing.Color.White;
            this.groupBox1.Location = new System.Drawing.Point(12, 12);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(253, 114);
            this.groupBox1.TabIndex = 0;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Select Delivery Date:";
            // 
            // btnSendHashTotal
            // 
            this.btnSendHashTotal.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(255)))), ((int)(((byte)(255)))));
            this.btnSendHashTotal.Font = new System.Drawing.Font("Comic Sans MS", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnSendHashTotal.ForeColor = System.Drawing.Color.Black;
            this.btnSendHashTotal.Location = new System.Drawing.Point(14, 58);
            this.btnSendHashTotal.Name = "btnSendHashTotal";
            this.btnSendHashTotal.Size = new System.Drawing.Size(233, 50);
            this.btnSendHashTotal.TabIndex = 1;
            this.btnSendHashTotal.Text = "Send Hash Total";
            this.btnSendHashTotal.UseVisualStyleBackColor = false;
            this.btnSendHashTotal.Click += new System.EventHandler(this.btnSendHashTotal_Click);
            // 
            // cboDeliveryDate
            // 
            this.cboDeliveryDate.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cboDeliveryDate.Font = new System.Drawing.Font("Comic Sans MS", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cboDeliveryDate.FormattingEnabled = true;
            this.cboDeliveryDate.Location = new System.Drawing.Point(14, 21);
            this.cboDeliveryDate.Name = "cboDeliveryDate";
            this.cboDeliveryDate.Size = new System.Drawing.Size(233, 31);
            this.cboDeliveryDate.TabIndex = 0;
            // 
            // frmHashTotal
            // 
            this.AcceptButton = this.btnSendHashTotal;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(64)))), ((int)(((byte)(64)))));
            this.ClientSize = new System.Drawing.Size(277, 138);
            this.Controls.Add(this.groupBox1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.MinimizeBox = false;
            this.Name = "frmHashTotal";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Hash Total";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.frmHashTotal_FormClosed);
            this.Load += new System.EventHandler(this.frmHashTotal_Load);
            this.groupBox1.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.ComboBox cboDeliveryDate;
        private System.Windows.Forms.Button btnSendHashTotal;
    }
}