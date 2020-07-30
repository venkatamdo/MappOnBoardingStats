namespace MAPPOnBoardingStats
{
    partial class Form1
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
            this.BtnProcessThisEmailAdr = new System.Windows.Forms.Button();
            this.BtnGetNonSubmittals = new System.Windows.Forms.Button();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.btnExit = new System.Windows.Forms.Button();
            this.textBoxNonSubmittalCount = new System.Windows.Forms.TextBox();
            this.BtnClearSelection = new System.Windows.Forms.Button();
            this.TxtEmail = new System.Windows.Forms.TextBox();
            this.TxtEmailPwd = new System.Windows.Forms.TextBox();
            this.TxtFromEmail = new System.Windows.Forms.TextBox();
            this.TxtContainsText = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.BtnProcessSelectedEmailAdr = new System.Windows.Forms.Button();
            this.ProcessAllRowsBtn = new System.Windows.Forms.Button();
            this.UpdateSheets = new System.Windows.Forms.Button();
            this.SelectedRowEmailAuditBtn = new System.Windows.Forms.Button();
            this.UpdateAccountButton = new System.Windows.Forms.Button();
            this.progressLabel = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.SuspendLayout();
            // 
            // BtnProcessThisEmailAdr
            // 
            this.BtnProcessThisEmailAdr.Location = new System.Drawing.Point(52, 73);
            this.BtnProcessThisEmailAdr.Name = "BtnProcessThisEmailAdr";
            this.BtnProcessThisEmailAdr.Size = new System.Drawing.Size(130, 39);
            this.BtnProcessThisEmailAdr.TabIndex = 5;
            this.BtnProcessThisEmailAdr.Text = "Process This Email Address";
            this.BtnProcessThisEmailAdr.UseVisualStyleBackColor = true;
            this.BtnProcessThisEmailAdr.Click += new System.EventHandler(this.BtnProcessThisEmailAdr_Click);
            // 
            // BtnGetNonSubmittals
            // 
            this.BtnGetNonSubmittals.Location = new System.Drawing.Point(52, 14);
            this.BtnGetNonSubmittals.Name = "BtnGetNonSubmittals";
            this.BtnGetNonSubmittals.Size = new System.Drawing.Size(130, 24);
            this.BtnGetNonSubmittals.TabIndex = 2;
            this.BtnGetNonSubmittals.Text = "Get NonSubmittals";
            this.BtnGetNonSubmittals.UseVisualStyleBackColor = true;
            this.BtnGetNonSubmittals.Click += new System.EventHandler(this.BtnGetNonSubmittals_Click);
            // 
            // dataGridView1
            // 
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Location = new System.Drawing.Point(52, 139);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.RowHeadersWidth = 72;
            this.dataGridView1.Size = new System.Drawing.Size(688, 206);
            this.dataGridView1.TabIndex = 6;
            // 
            // btnExit
            // 
            this.btnExit.Location = new System.Drawing.Point(661, 358);
            this.btnExit.Name = "btnExit";
            this.btnExit.Size = new System.Drawing.Size(79, 37);
            this.btnExit.TabIndex = 10;
            this.btnExit.Text = "E&xit";
            this.btnExit.UseVisualStyleBackColor = true;
            this.btnExit.Click += new System.EventHandler(this.btnExit_Click);
            // 
            // textBoxNonSubmittalCount
            // 
            this.textBoxNonSubmittalCount.Location = new System.Drawing.Point(248, 14);
            this.textBoxNonSubmittalCount.Margin = new System.Windows.Forms.Padding(2);
            this.textBoxNonSubmittalCount.Name = "textBoxNonSubmittalCount";
            this.textBoxNonSubmittalCount.Size = new System.Drawing.Size(343, 20);
            this.textBoxNonSubmittalCount.TabIndex = 13;
            // 
            // BtnClearSelection
            // 
            this.BtnClearSelection.Location = new System.Drawing.Point(549, 358);
            this.BtnClearSelection.Margin = new System.Windows.Forms.Padding(2);
            this.BtnClearSelection.Name = "BtnClearSelection";
            this.BtnClearSelection.Size = new System.Drawing.Size(107, 36);
            this.BtnClearSelection.TabIndex = 9;
            this.BtnClearSelection.Text = "Clear Selection";
            this.BtnClearSelection.UseVisualStyleBackColor = true;
            this.BtnClearSelection.Click += new System.EventHandler(this.BtnClearSelection_Click);
            // 
            // TxtEmail
            // 
            this.TxtEmail.Location = new System.Drawing.Point(248, 59);
            this.TxtEmail.Margin = new System.Windows.Forms.Padding(2);
            this.TxtEmail.Name = "TxtEmail";
            this.TxtEmail.Size = new System.Drawing.Size(188, 20);
            this.TxtEmail.TabIndex = 14;
            // 
            // TxtEmailPwd
            // 
            this.TxtEmailPwd.Location = new System.Drawing.Point(467, 59);
            this.TxtEmailPwd.Margin = new System.Windows.Forms.Padding(2);
            this.TxtEmailPwd.Name = "TxtEmailPwd";
            this.TxtEmailPwd.PasswordChar = '*';
            this.TxtEmailPwd.Size = new System.Drawing.Size(123, 20);
            this.TxtEmailPwd.TabIndex = 15;
            // 
            // TxtFromEmail
            // 
            this.TxtFromEmail.Location = new System.Drawing.Point(248, 108);
            this.TxtFromEmail.Margin = new System.Windows.Forms.Padding(2);
            this.TxtFromEmail.Name = "TxtFromEmail";
            this.TxtFromEmail.Size = new System.Drawing.Size(186, 20);
            this.TxtFromEmail.TabIndex = 16;
            // 
            // TxtContainsText
            // 
            this.TxtContainsText.Location = new System.Drawing.Point(468, 108);
            this.TxtContainsText.Margin = new System.Windows.Forms.Padding(2);
            this.TxtContainsText.Name = "TxtContainsText";
            this.TxtContainsText.Size = new System.Drawing.Size(121, 20);
            this.TxtContainsText.TabIndex = 17;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(250, 43);
            this.label2.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(73, 13);
            this.label2.TabIndex = 18;
            this.label2.Text = "Email Address";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(470, 43);
            this.label3.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(56, 13);
            this.label3.TabIndex = 19;
            this.label3.Text = "Email Pwd";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(249, 91);
            this.label4.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(58, 13);
            this.label4.TabIndex = 20;
            this.label4.Text = "From Email";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(469, 92);
            this.label5.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(71, 13);
            this.label5.TabIndex = 21;
            this.label5.Text = "Search String";
            // 
            // BtnProcessSelectedEmailAdr
            // 
            this.BtnProcessSelectedEmailAdr.Location = new System.Drawing.Point(52, 359);
            this.BtnProcessSelectedEmailAdr.Name = "BtnProcessSelectedEmailAdr";
            this.BtnProcessSelectedEmailAdr.Size = new System.Drawing.Size(130, 37);
            this.BtnProcessSelectedEmailAdr.TabIndex = 22;
            this.BtnProcessSelectedEmailAdr.Text = "Process Selected Group Email Address";
            this.BtnProcessSelectedEmailAdr.UseVisualStyleBackColor = true;
            this.BtnProcessSelectedEmailAdr.Click += new System.EventHandler(this.BtnProcessSelectedEmailAdr_Click);
            // 
            // ProcessAllRowsBtn
            // 
            this.ProcessAllRowsBtn.Location = new System.Drawing.Point(188, 359);
            this.ProcessAllRowsBtn.Name = "ProcessAllRowsBtn";
            this.ProcessAllRowsBtn.Size = new System.Drawing.Size(108, 37);
            this.ProcessAllRowsBtn.TabIndex = 23;
            this.ProcessAllRowsBtn.Text = "Process All Rows";
            this.ProcessAllRowsBtn.UseVisualStyleBackColor = true;
            this.ProcessAllRowsBtn.Click += new System.EventHandler(this.ProcessAllRowsBtn_Click);
            // 
            // UpdateSheets
            // 
            this.UpdateSheets.Location = new System.Drawing.Point(302, 359);
            this.UpdateSheets.Name = "UpdateSheets";
            this.UpdateSheets.Size = new System.Drawing.Size(85, 37);
            this.UpdateSheets.TabIndex = 24;
            this.UpdateSheets.Text = "Update Sheets";
            this.UpdateSheets.UseVisualStyleBackColor = true;
            this.UpdateSheets.Click += new System.EventHandler(this.BtnUpdateSheets);
            // 
            // SelectedRowEmailAuditBtn
            // 
            this.SelectedRowEmailAuditBtn.Location = new System.Drawing.Point(393, 359);
            this.SelectedRowEmailAuditBtn.Name = "SelectedRowEmailAuditBtn";
            this.SelectedRowEmailAuditBtn.Size = new System.Drawing.Size(118, 37);
            this.SelectedRowEmailAuditBtn.TabIndex = 25;
            this.SelectedRowEmailAuditBtn.Text = "Daily Email Audit for Selected Row";
            this.SelectedRowEmailAuditBtn.UseVisualStyleBackColor = true;
            this.SelectedRowEmailAuditBtn.Click += new System.EventHandler(this.SelectedRowEmailAuditBtn_Click);
            // 
            // UpdateAccountButton
            // 
            this.UpdateAccountButton.Location = new System.Drawing.Point(624, 73);
            this.UpdateAccountButton.Name = "UpdateAccountButton";
            this.UpdateAccountButton.Size = new System.Drawing.Size(116, 37);
            this.UpdateAccountButton.TabIndex = 26;
            this.UpdateAccountButton.Text = "Update Account";
            this.UpdateAccountButton.UseVisualStyleBackColor = true;
            this.UpdateAccountButton.Click += new System.EventHandler(this.UpdateAccountButton_Click);
            // 
            // progressLabel
            // 
            this.progressLabel.AutoSize = true;
            this.progressLabel.Location = new System.Drawing.Point(491, 456);
            this.progressLabel.Name = "progressLabel";
            this.progressLabel.Size = new System.Drawing.Size(0, 13);
            this.progressLabel.TabIndex = 27;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(831, 489);
            this.Controls.Add(this.progressLabel);
            this.Controls.Add(this.UpdateAccountButton);
            this.Controls.Add(this.SelectedRowEmailAuditBtn);
            this.Controls.Add(this.UpdateSheets);
            this.Controls.Add(this.ProcessAllRowsBtn);
            this.Controls.Add(this.BtnProcessSelectedEmailAdr);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.TxtContainsText);
            this.Controls.Add(this.TxtFromEmail);
            this.Controls.Add(this.TxtEmailPwd);
            this.Controls.Add(this.TxtEmail);
            this.Controls.Add(this.textBoxNonSubmittalCount);
            this.Controls.Add(this.BtnClearSelection);
            this.Controls.Add(this.btnExit);
            this.Controls.Add(this.dataGridView1);
            this.Controls.Add(this.BtnGetNonSubmittals);
            this.Controls.Add(this.BtnProcessThisEmailAdr);
            this.Name = "Form1";
            this.Text = "Automate MAPP";
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.Button BtnProcessThisEmailAdr;
        private System.Windows.Forms.Button BtnGetNonSubmittals;
        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.Button btnExit;
        private System.Windows.Forms.TextBox textBoxNonSubmittalCount;
        private System.Windows.Forms.Button BtnClearSelection;
        private System.Windows.Forms.TextBox TxtEmail;
        private System.Windows.Forms.TextBox TxtEmailPwd;
        private System.Windows.Forms.TextBox TxtFromEmail;
        private System.Windows.Forms.TextBox TxtContainsText;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Button BtnProcessSelectedEmailAdr;
        private System.Windows.Forms.Button ProcessAllRowsBtn;
        private System.Windows.Forms.Button UpdateSheets;
        private System.Windows.Forms.Button SelectedRowEmailAuditBtn;
        private System.Windows.Forms.Button UpdateAccountButton;
        private System.Windows.Forms.Label progressLabel;
    }
}
