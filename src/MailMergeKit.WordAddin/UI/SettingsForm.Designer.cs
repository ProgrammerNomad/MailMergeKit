namespace MailMergeKit.WordAddin.UI
{
    partial class SettingsForm
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
            this.lblTitle = new System.Windows.Forms.Label();
            this.lblRecipientField = new System.Windows.Forms.Label();
            this.cboRecipientField = new System.Windows.Forms.ComboBox();
            this.lblSubject = new System.Windows.Forms.Label();
            this.txtSubject = new System.Windows.Forms.TextBox();
            this.lnkMergeFieldHelp = new System.Windows.Forms.LinkLabel();
            this.lblRecordCount = new System.Windows.Forms.Label();
            this.btnOK = new System.Windows.Forms.Button();
            this.btnCancel = new System.Windows.Forms.Button();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // lblTitle
            // 
            this.lblTitle.AutoSize = true;
            this.lblTitle.Font = new System.Drawing.Font("Segoe UI", 12F, System.Drawing.FontStyle.Bold);
            this.lblTitle.Location = new System.Drawing.Point(12, 9);
            this.lblTitle.Name = "lblTitle";
            this.lblTitle.Size = new System.Drawing.Size(233, 21);
            this.lblTitle.TabIndex = 0;
            this.lblTitle.Text = "MailMergeKit - Send via Outlook";
            // 
            // lblRecipientField
            // 
            this.lblRecipientField.AutoSize = true;
            this.lblRecipientField.Location = new System.Drawing.Point(13, 25);
            this.lblRecipientField.Name = "lblRecipientField";
            this.lblRecipientField.Size = new System.Drawing.Size(144, 15);
            this.lblRecipientField.TabIndex = 1;
            this.lblRecipientField.Text = "Recipient Email Field:";
            // 
            // cboRecipientField
            // 
            this.cboRecipientField.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cboRecipientField.FormattingEnabled = true;
            this.cboRecipientField.Location = new System.Drawing.Point(16, 43);
            this.cboRecipientField.Name = "cboRecipientField";
            this.cboRecipientField.Size = new System.Drawing.Size(430, 23);
            this.cboRecipientField.TabIndex = 2;
            // 
            // lblSubject
            // 
            this.lblSubject.AutoSize = true;
            this.lblSubject.Location = new System.Drawing.Point(13, 81);
            this.lblSubject.Name = "lblSubject";
            this.lblSubject.Size = new System.Drawing.Size(118, 15);
            this.lblSubject.TabIndex = 3;
            this.lblSubject.Text = "Subject Template:";
            // 
            // txtSubject
            // 
            this.txtSubject.Location = new System.Drawing.Point(16, 99);
            this.txtSubject.Name = "txtSubject";
            this.txtSubject.Size = new System.Drawing.Size(430, 23);
            this.txtSubject.TabIndex = 4;
            // 
            // lnkMergeFieldHelp
            // 
            this.lnkMergeFieldHelp.AutoSize = true;
            this.lnkMergeFieldHelp.Location = new System.Drawing.Point(13, 125);
            this.lnkMergeFieldHelp.Name = "lnkMergeFieldHelp";
            this.lnkMergeFieldHelp.Size = new System.Drawing.Size(263, 15);
            this.lnkMergeFieldHelp.TabIndex = 5;
            this.lnkMergeFieldHelp.TabStop = true;
            this.lnkMergeFieldHelp.Text = "How to use merge fields in subject? (Click here)";
            this.lnkMergeFieldHelp.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.lnkMergeFieldHelp_LinkClicked);
            // 
            // lblRecordCount
            // 
            this.lblRecordCount.AutoSize = true;
            this.lblRecordCount.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Bold);
            this.lblRecordCount.Location = new System.Drawing.Point(12, 45);
            this.lblRecordCount.Name = "lblRecordCount";
            this.lblRecordCount.Size = new System.Drawing.Size(93, 15);
            this.lblRecordCount.TabIndex = 6;
            this.lblRecordCount.Text = "Total records: 0";
            // 
            // btnOK
            // 
            this.btnOK.Location = new System.Drawing.Point(290, 251);
            this.btnOK.Name = "btnOK";
            this.btnOK.Size = new System.Drawing.Size(90, 30);
            this.btnOK.TabIndex = 7;
            this.btnOK.Text = "Start Merge";
            this.btnOK.UseVisualStyleBackColor = true;
            this.btnOK.Click += new System.EventHandler(this.btnOK_Click);
            // 
            // btnCancel
            // 
            this.btnCancel.Location = new System.Drawing.Point(386, 251);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(90, 30);
            this.btnCancel.TabIndex = 8;
            this.btnCancel.Text = "Cancel";
            this.btnCancel.UseVisualStyleBackColor = true;
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.lblRecipientField);
            this.groupBox1.Controls.Add(this.cboRecipientField);
            this.groupBox1.Controls.Add(this.lblSubject);
            this.groupBox1.Controls.Add(this.txtSubject);
            this.groupBox1.Controls.Add(this.lnkMergeFieldHelp);
            this.groupBox1.Location = new System.Drawing.Point(12, 73);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(464, 162);
            this.groupBox1.TabIndex = 9;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Merge Settings";
            // 
            // SettingsForm
            // 
            this.AcceptButton = this.btnOK;
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.CancelButton = this.btnCancel;
            this.ClientSize = new System.Drawing.Size(488, 293);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.btnOK);
            this.Controls.Add(this.lblRecordCount);
            this.Controls.Add(this.lblTitle);
            this.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "SettingsForm";
            this.ShowIcon = false;
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "MailMergeKit Settings";
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();
        }

        #endregion

        private System.Windows.Forms.Label lblTitle;
        private System.Windows.Forms.Label lblRecipientField;
        private System.Windows.Forms.ComboBox cboRecipientField;
        private System.Windows.Forms.Label lblSubject;
        private System.Windows.Forms.TextBox txtSubject;
        private System.Windows.Forms.LinkLabel lnkMergeFieldHelp;
        private System.Windows.Forms.Label lblRecordCount;
        private System.Windows.Forms.Button btnOK;
        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.GroupBox groupBox1;
    }
}
