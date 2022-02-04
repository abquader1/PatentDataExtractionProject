namespace Patient_Master
{
    partial class frmMainWindow
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmMainWindow));
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.chkCPCName = new System.Windows.Forms.CheckBox();
            this.chkCPC = new System.Windows.Forms.CheckBox();
            this.chkStatus = new System.Windows.Forms.CheckBox();
            this.chkAssignee = new System.Windows.Forms.CheckBox();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.chkClaimsRequired = new System.Windows.Forms.CheckBox();
            this.chkFilingDate = new System.Windows.Forms.CheckBox();
            this.chkAbstract = new System.Windows.Forms.CheckBox();
            this.chkDescription = new System.Windows.Forms.CheckBox();
            this.chkClaims = new System.Windows.Forms.CheckBox();
            this.label7 = new System.Windows.Forms.Label();
            this.chkWord = new System.Windows.Forms.CheckBox();
            this.chkExcel = new System.Windows.Forms.CheckBox();
            this.label6 = new System.Windows.Forms.Label();
            this.txtNumberOfRecords = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.button2 = new System.Windows.Forms.Button();
            this.btnClear = new System.Windows.Forms.Button();
            this.txtOutputPath = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.btnProcess = new System.Windows.Forms.Button();
            this.btnOutputFolder = new System.Windows.Forms.Button();
            this.btnBrowseWord = new System.Windows.Forms.Button();
            this.btnBrowseExcel = new System.Windows.Forms.Button();
            this.txtWordFile = new System.Windows.Forms.TextBox();
            this.txtExcelFile = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.label4 = new System.Windows.Forms.Label();
            this.lblStatus = new System.Windows.Forms.Label();
            this.groupBox1.SuspendLayout();
            this.groupBox3.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.SuspendLayout();
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.groupBox3);
            this.groupBox1.Controls.Add(this.groupBox2);
            this.groupBox1.Controls.Add(this.label7);
            this.groupBox1.Controls.Add(this.chkWord);
            this.groupBox1.Controls.Add(this.chkExcel);
            this.groupBox1.Controls.Add(this.label6);
            this.groupBox1.Controls.Add(this.txtNumberOfRecords);
            this.groupBox1.Controls.Add(this.label5);
            this.groupBox1.Controls.Add(this.button2);
            this.groupBox1.Controls.Add(this.btnClear);
            this.groupBox1.Controls.Add(this.txtOutputPath);
            this.groupBox1.Controls.Add(this.label3);
            this.groupBox1.Controls.Add(this.btnProcess);
            this.groupBox1.Controls.Add(this.btnOutputFolder);
            this.groupBox1.Controls.Add(this.btnBrowseWord);
            this.groupBox1.Controls.Add(this.btnBrowseExcel);
            this.groupBox1.Controls.Add(this.txtWordFile);
            this.groupBox1.Controls.Add(this.txtExcelFile);
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Location = new System.Drawing.Point(12, 12);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(564, 320);
            this.groupBox1.TabIndex = 0;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Input parameters";
            // 
            // groupBox3
            // 
            this.groupBox3.Controls.Add(this.chkCPCName);
            this.groupBox3.Controls.Add(this.chkCPC);
            this.groupBox3.Controls.Add(this.chkStatus);
            this.groupBox3.Controls.Add(this.chkAssignee);
            this.groupBox3.Location = new System.Drawing.Point(282, 193);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new System.Drawing.Size(181, 93);
            this.groupBox3.TabIndex = 11;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "Metadata";
            // 
            // chkCPCName
            // 
            this.chkCPCName.AutoSize = true;
            this.chkCPCName.Location = new System.Drawing.Point(92, 48);
            this.chkCPCName.Name = "chkCPCName";
            this.chkCPCName.Size = new System.Drawing.Size(78, 17);
            this.chkCPCName.TabIndex = 0;
            this.chkCPCName.Text = "CPC Name";
            this.chkCPCName.UseVisualStyleBackColor = true;
            // 
            // chkCPC
            // 
            this.chkCPC.AutoSize = true;
            this.chkCPC.Location = new System.Drawing.Point(92, 22);
            this.chkCPC.Name = "chkCPC";
            this.chkCPC.Size = new System.Drawing.Size(47, 17);
            this.chkCPC.TabIndex = 0;
            this.chkCPC.Text = "CPC";
            this.chkCPC.UseVisualStyleBackColor = true;
            // 
            // chkStatus
            // 
            this.chkStatus.AutoSize = true;
            this.chkStatus.Location = new System.Drawing.Point(9, 48);
            this.chkStatus.Name = "chkStatus";
            this.chkStatus.Size = new System.Drawing.Size(56, 17);
            this.chkStatus.TabIndex = 0;
            this.chkStatus.Text = "Status";
            this.chkStatus.UseVisualStyleBackColor = true;
            // 
            // chkAssignee
            // 
            this.chkAssignee.AutoSize = true;
            this.chkAssignee.Location = new System.Drawing.Point(9, 22);
            this.chkAssignee.Name = "chkAssignee";
            this.chkAssignee.Size = new System.Drawing.Size(69, 17);
            this.chkAssignee.TabIndex = 0;
            this.chkAssignee.Text = "Assignee";
            this.chkAssignee.UseVisualStyleBackColor = true;
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.chkClaimsRequired);
            this.groupBox2.Controls.Add(this.chkFilingDate);
            this.groupBox2.Controls.Add(this.chkAbstract);
            this.groupBox2.Controls.Add(this.chkDescription);
            this.groupBox2.Controls.Add(this.chkClaims);
            this.groupBox2.Location = new System.Drawing.Point(18, 193);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(258, 93);
            this.groupBox2.TabIndex = 11;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Patent Text";
            // 
            // chkClaimsRequired
            // 
            this.chkClaimsRequired.AutoSize = true;
            this.chkClaimsRequired.Location = new System.Drawing.Point(9, 70);
            this.chkClaimsRequired.Name = "chkClaimsRequired";
            this.chkClaimsRequired.Size = new System.Drawing.Size(161, 17);
            this.chkClaimsRequired.TabIndex = 1;
            this.chkClaimsRequired.Text = "Only 1st Claim in Word Doc?";
            this.chkClaimsRequired.UseVisualStyleBackColor = true;
            // 
            // chkFilingDate
            // 
            this.chkFilingDate.AutoSize = true;
            this.chkFilingDate.Location = new System.Drawing.Point(95, 48);
            this.chkFilingDate.Name = "chkFilingDate";
            this.chkFilingDate.Size = new System.Drawing.Size(76, 17);
            this.chkFilingDate.TabIndex = 0;
            this.chkFilingDate.Text = "Filing Date";
            this.chkFilingDate.UseVisualStyleBackColor = true;
            // 
            // chkAbstract
            // 
            this.chkAbstract.AutoSize = true;
            this.chkAbstract.Location = new System.Drawing.Point(95, 22);
            this.chkAbstract.Name = "chkAbstract";
            this.chkAbstract.Size = new System.Drawing.Size(65, 17);
            this.chkAbstract.TabIndex = 0;
            this.chkAbstract.Text = "Abstract";
            this.chkAbstract.UseVisualStyleBackColor = true;
            // 
            // chkDescription
            // 
            this.chkDescription.AutoSize = true;
            this.chkDescription.Location = new System.Drawing.Point(9, 46);
            this.chkDescription.Name = "chkDescription";
            this.chkDescription.Size = new System.Drawing.Size(79, 17);
            this.chkDescription.TabIndex = 0;
            this.chkDescription.Text = "Description";
            this.chkDescription.UseVisualStyleBackColor = true;
            // 
            // chkClaims
            // 
            this.chkClaims.AutoSize = true;
            this.chkClaims.Location = new System.Drawing.Point(9, 22);
            this.chkClaims.Name = "chkClaims";
            this.chkClaims.Size = new System.Drawing.Size(56, 17);
            this.chkClaims.TabIndex = 0;
            this.chkClaims.Text = "Claims";
            this.chkClaims.UseVisualStyleBackColor = true;
            this.chkClaims.CheckedChanged += new System.EventHandler(this.chkClaims_CheckedChanged);
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(15, 297);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(410, 13);
            this.label7.TabIndex = 10;
            this.label7.Text = "Note: if total records count is greater than available in source, all records wil" +
    "l be pulled";
            // 
            // chkWord
            // 
            this.chkWord.AutoSize = true;
            this.chkWord.Location = new System.Drawing.Point(142, 159);
            this.chkWord.Name = "chkWord";
            this.chkWord.Size = new System.Drawing.Size(52, 17);
            this.chkWord.TabIndex = 1;
            this.chkWord.Text = "Word";
            this.chkWord.UseVisualStyleBackColor = true;
            // 
            // chkExcel
            // 
            this.chkExcel.AutoSize = true;
            this.chkExcel.Location = new System.Drawing.Point(201, 159);
            this.chkExcel.Name = "chkExcel";
            this.chkExcel.Size = new System.Drawing.Size(52, 17);
            this.chkExcel.TabIndex = 2;
            this.chkExcel.Text = "Excel";
            this.chkExcel.UseVisualStyleBackColor = true;
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(15, 163);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(108, 13);
            this.label6.TabIndex = 9;
            this.label6.Text = "Required Output type";
            // 
            // txtNumberOfRecords
            // 
            this.txtNumberOfRecords.Location = new System.Drawing.Point(142, 128);
            this.txtNumberOfRecords.Name = "txtNumberOfRecords";
            this.txtNumberOfRecords.Size = new System.Drawing.Size(166, 20);
            this.txtNumberOfRecords.TabIndex = 0;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(15, 132);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(107, 13);
            this.label5.TabIndex = 7;
            this.label5.Text = "Extract total record(s)";
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(469, 256);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(89, 35);
            this.button2.TabIndex = 8;
            this.button2.Text = "&Cancel";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // btnClear
            // 
            this.btnClear.Location = new System.Drawing.Point(469, 206);
            this.btnClear.Name = "btnClear";
            this.btnClear.Size = new System.Drawing.Size(89, 35);
            this.btnClear.TabIndex = 7;
            this.btnClear.Text = "C&lear";
            this.btnClear.UseVisualStyleBackColor = true;
            this.btnClear.Click += new System.EventHandler(this.btnClear_Click);
            // 
            // txtOutputPath
            // 
            this.txtOutputPath.Location = new System.Drawing.Point(142, 93);
            this.txtOutputPath.Name = "txtOutputPath";
            this.txtOutputPath.Size = new System.Drawing.Size(328, 20);
            this.txtOutputPath.TabIndex = 4;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(15, 97);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(64, 13);
            this.label3.TabIndex = 5;
            this.label3.Text = "Output Path";
            // 
            // btnProcess
            // 
            this.btnProcess.Location = new System.Drawing.Point(469, 157);
            this.btnProcess.Name = "btnProcess";
            this.btnProcess.Size = new System.Drawing.Size(89, 35);
            this.btnProcess.TabIndex = 6;
            this.btnProcess.Text = "&Process";
            this.btnProcess.UseVisualStyleBackColor = true;
            this.btnProcess.Click += new System.EventHandler(this.btnProcess_Click);
            // 
            // btnOutputFolder
            // 
            this.btnOutputFolder.Location = new System.Drawing.Point(476, 91);
            this.btnOutputFolder.Name = "btnOutputFolder";
            this.btnOutputFolder.Size = new System.Drawing.Size(75, 23);
            this.btnOutputFolder.TabIndex = 5;
            this.btnOutputFolder.Text = "Browse...";
            this.btnOutputFolder.UseVisualStyleBackColor = true;
            this.btnOutputFolder.Click += new System.EventHandler(this.btnOutputFolder_Click);
            // 
            // btnBrowseWord
            // 
            this.btnBrowseWord.Location = new System.Drawing.Point(476, 58);
            this.btnBrowseWord.Name = "btnBrowseWord";
            this.btnBrowseWord.Size = new System.Drawing.Size(75, 23);
            this.btnBrowseWord.TabIndex = 3;
            this.btnBrowseWord.Text = "Browse...";
            this.btnBrowseWord.UseVisualStyleBackColor = true;
            this.btnBrowseWord.Click += new System.EventHandler(this.btnBrowseWord_Click);
            // 
            // btnBrowseExcel
            // 
            this.btnBrowseExcel.Location = new System.Drawing.Point(476, 25);
            this.btnBrowseExcel.Name = "btnBrowseExcel";
            this.btnBrowseExcel.Size = new System.Drawing.Size(75, 23);
            this.btnBrowseExcel.TabIndex = 1;
            this.btnBrowseExcel.Text = "Browse...";
            this.btnBrowseExcel.UseVisualStyleBackColor = true;
            this.btnBrowseExcel.Click += new System.EventHandler(this.btnBrowseExcel_Click);
            // 
            // txtWordFile
            // 
            this.txtWordFile.Location = new System.Drawing.Point(142, 59);
            this.txtWordFile.Name = "txtWordFile";
            this.txtWordFile.Size = new System.Drawing.Size(328, 20);
            this.txtWordFile.TabIndex = 2;
            // 
            // txtExcelFile
            // 
            this.txtExcelFile.Location = new System.Drawing.Point(142, 25);
            this.txtExcelFile.Name = "txtExcelFile";
            this.txtExcelFile.Size = new System.Drawing.Size(328, 20);
            this.txtExcelFile.TabIndex = 0;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(15, 61);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(113, 13);
            this.label2.TabIndex = 0;
            this.label2.Text = "Select Word Template";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(15, 25);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(99, 13);
            this.label1.TabIndex = 0;
            this.label1.Text = "Select Result Excel";
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(27, 341);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(40, 13);
            this.label4.TabIndex = 5;
            this.label4.Text = "Status:";
            // 
            // lblStatus
            // 
            this.lblStatus.AutoSize = true;
            this.lblStatus.Location = new System.Drawing.Point(73, 341);
            this.lblStatus.Name = "lblStatus";
            this.lblStatus.Size = new System.Drawing.Size(40, 13);
            this.lblStatus.TabIndex = 5;
            this.lblStatus.Text = "Status:";
            // 
            // frmMainWindow
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(587, 363);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.lblStatus);
            this.Controls.Add(this.label4);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "frmMainWindow";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Master Window";
            this.Load += new System.EventHandler(this.frmMainWindow_Load);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.groupBox3.ResumeLayout(false);
            this.groupBox3.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button btnProcess;
        private System.Windows.Forms.Button btnBrowseWord;
        private System.Windows.Forms.Button btnBrowseExcel;
        private System.Windows.Forms.TextBox txtWordFile;
        private System.Windows.Forms.TextBox txtExcelFile;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.Button btnClear;
        private System.Windows.Forms.TextBox txtOutputPath;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Button btnOutputFolder;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label lblStatus;
        private System.Windows.Forms.TextBox txtNumberOfRecords;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.CheckBox chkWord;
        private System.Windows.Forms.CheckBox chkExcel;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.GroupBox groupBox3;
        private System.Windows.Forms.CheckBox chkCPCName;
        private System.Windows.Forms.CheckBox chkCPC;
        private System.Windows.Forms.CheckBox chkStatus;
        private System.Windows.Forms.CheckBox chkAssignee;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.CheckBox chkAbstract;
        private System.Windows.Forms.CheckBox chkDescription;
        private System.Windows.Forms.CheckBox chkClaims;
        private System.Windows.Forms.CheckBox chkFilingDate;
        private System.Windows.Forms.CheckBox chkClaimsRequired;
    }
}

