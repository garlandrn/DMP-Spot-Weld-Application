namespace DMP_Spot_Weld_Application
{
    partial class DMP_Spot_Weld_Login
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(DMP_Spot_Weld_Login));
            this.EmployeeName_TextBox = new System.Windows.Forms.TextBox();
            this.OperatorLogin_Button = new System.Windows.Forms.Button();
            this.ListBox = new System.Windows.Forms.ListBox();
            this.Clock = new System.Windows.Forms.Timer(this.components);
            this.LoginGridView = new System.Windows.Forms.DataGridView();
            this.EmployeeName_Label = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.DMPID_TextBox = new System.Windows.Forms.TextBox();
            this.Clock_TextBox = new System.Windows.Forms.TextBox();
            this.AdminLogin_Button = new System.Windows.Forms.Button();
            this.Exit_Button = new System.Windows.Forms.Button();
            this.JobList_Button = new System.Windows.Forms.Button();
            this.CellControl_Button = new System.Windows.Forms.Button();
            this.ReportView_Button = new System.Windows.Forms.Button();
            this.EmployeeName_ComboBox = new System.Windows.Forms.ComboBox();
            this.OPC_Button = new System.Windows.Forms.Button();
            this.TestForm_Button = new System.Windows.Forms.Button();
            this.Help_Button = new System.Windows.Forms.Button();
            this.ScanOut_Button = new System.Windows.Forms.Button();
            this.Test_GroupBox = new System.Windows.Forms.GroupBox();
            this.TextFile_Button = new System.Windows.Forms.Button();
            this.Login_GroupBox = new System.Windows.Forms.GroupBox();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            ((System.ComponentModel.ISupportInitialize)(this.LoginGridView)).BeginInit();
            this.Test_GroupBox.SuspendLayout();
            this.Login_GroupBox.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.SuspendLayout();
            // 
            // EmployeeName_TextBox
            // 
            this.EmployeeName_TextBox.Font = new System.Drawing.Font("Microsoft Sans Serif", 26.25F);
            this.EmployeeName_TextBox.Location = new System.Drawing.Point(659, 408);
            this.EmployeeName_TextBox.Name = "EmployeeName_TextBox";
            this.EmployeeName_TextBox.Size = new System.Drawing.Size(608, 47);
            this.EmployeeName_TextBox.TabIndex = 0;
            this.EmployeeName_TextBox.Visible = false;
            // 
            // OperatorLogin_Button
            // 
            this.OperatorLogin_Button.Font = new System.Drawing.Font("Microsoft Sans Serif", 25.25F, System.Drawing.FontStyle.Bold);
            this.OperatorLogin_Button.Location = new System.Drawing.Point(1335, 407);
            this.OperatorLogin_Button.Name = "OperatorLogin_Button";
            this.OperatorLogin_Button.Size = new System.Drawing.Size(285, 47);
            this.OperatorLogin_Button.TabIndex = 2;
            this.OperatorLogin_Button.Text = "Operator Login";
            this.OperatorLogin_Button.UseVisualStyleBackColor = true;
            this.OperatorLogin_Button.Click += new System.EventHandler(this.OperatorLogin_Button_Click);
            // 
            // ListBox
            // 
            this.ListBox.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F);
            this.ListBox.FormattingEnabled = true;
            this.ListBox.ItemHeight = 24;
            this.ListBox.Location = new System.Drawing.Point(427, 533);
            this.ListBox.Name = "ListBox";
            this.ListBox.Size = new System.Drawing.Size(840, 292);
            this.ListBox.TabIndex = 10;
            this.ListBox.TabStop = false;
            // 
            // Clock
            // 
            this.Clock.Interval = 250;
            this.Clock.Tick += new System.EventHandler(this.Clock_Tick);
            // 
            // LoginGridView
            // 
            this.LoginGridView.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.LoginGridView.Location = new System.Drawing.Point(11, 9);
            this.LoginGridView.Name = "LoginGridView";
            this.LoginGridView.Size = new System.Drawing.Size(209, 98);
            this.LoginGridView.TabIndex = 19;
            this.LoginGridView.Visible = false;
            // 
            // EmployeeName_Label
            // 
            this.EmployeeName_Label.AutoSize = true;
            this.EmployeeName_Label.Font = new System.Drawing.Font("Microsoft Sans Serif", 27.25F, System.Drawing.FontStyle.Bold);
            this.EmployeeName_Label.Location = new System.Drawing.Point(336, 408);
            this.EmployeeName_Label.Name = "EmployeeName_Label";
            this.EmployeeName_Label.Size = new System.Drawing.Size(317, 42);
            this.EmployeeName_Label.TabIndex = 5;
            this.EmployeeName_Label.Text = "Employee Name:";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 27.25F, System.Drawing.FontStyle.Bold);
            this.label1.Location = new System.Drawing.Point(489, 478);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(164, 42);
            this.label1.TabIndex = 8;
            this.label1.Text = "DMP ID:";
            // 
            // DMPID_TextBox
            // 
            this.DMPID_TextBox.Font = new System.Drawing.Font("Microsoft Sans Serif", 26.25F);
            this.DMPID_TextBox.Location = new System.Drawing.Point(659, 478);
            this.DMPID_TextBox.Name = "DMPID_TextBox";
            this.DMPID_TextBox.Size = new System.Drawing.Size(608, 47);
            this.DMPID_TextBox.TabIndex = 1;
            this.DMPID_TextBox.KeyDown += new System.Windows.Forms.KeyEventHandler(this.DMPID_TextBox_KeyDown);
            // 
            // Clock_TextBox
            // 
            this.Clock_TextBox.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.Clock_TextBox.Font = new System.Drawing.Font("Microsoft Sans Serif", 25.25F);
            this.Clock_TextBox.Location = new System.Drawing.Point(1490, 12);
            this.Clock_TextBox.Name = "Clock_TextBox";
            this.Clock_TextBox.ReadOnly = true;
            this.Clock_TextBox.Size = new System.Drawing.Size(402, 39);
            this.Clock_TextBox.TabIndex = 15;
            this.Clock_TextBox.TabStop = false;
            // 
            // AdminLogin_Button
            // 
            this.AdminLogin_Button.Font = new System.Drawing.Font("Microsoft Sans Serif", 25.25F, System.Drawing.FontStyle.Bold);
            this.AdminLogin_Button.Location = new System.Drawing.Point(6, 26);
            this.AdminLogin_Button.Name = "AdminLogin_Button";
            this.AdminLogin_Button.Size = new System.Drawing.Size(285, 47);
            this.AdminLogin_Button.TabIndex = 3;
            this.AdminLogin_Button.Text = "Admin Login";
            this.AdminLogin_Button.UseVisualStyleBackColor = true;
            this.AdminLogin_Button.Click += new System.EventHandler(this.AdminLogin_Button_Click);
            // 
            // Exit_Button
            // 
            this.Exit_Button.Font = new System.Drawing.Font("Microsoft Sans Serif", 27.25F, System.Drawing.FontStyle.Bold);
            this.Exit_Button.Location = new System.Drawing.Point(1692, 877);
            this.Exit_Button.Name = "Exit_Button";
            this.Exit_Button.Size = new System.Drawing.Size(185, 60);
            this.Exit_Button.TabIndex = 5;
            this.Exit_Button.Text = "Exit";
            this.Exit_Button.UseVisualStyleBackColor = true;
            this.Exit_Button.Click += new System.EventHandler(this.Exit_Button_Click);
            // 
            // JobList_Button
            // 
            this.JobList_Button.Font = new System.Drawing.Font("Microsoft Sans Serif", 25.25F, System.Drawing.FontStyle.Bold);
            this.JobList_Button.Location = new System.Drawing.Point(6, 172);
            this.JobList_Button.Name = "JobList_Button";
            this.JobList_Button.Size = new System.Drawing.Size(285, 47);
            this.JobList_Button.TabIndex = 4;
            this.JobList_Button.Text = "Job List";
            this.JobList_Button.UseVisualStyleBackColor = true;
            this.JobList_Button.Click += new System.EventHandler(this.JobList_Button_Click);
            // 
            // CellControl_Button
            // 
            this.CellControl_Button.Font = new System.Drawing.Font("Microsoft Sans Serif", 25.25F, System.Drawing.FontStyle.Bold);
            this.CellControl_Button.Location = new System.Drawing.Point(6, 245);
            this.CellControl_Button.Name = "CellControl_Button";
            this.CellControl_Button.Size = new System.Drawing.Size(285, 47);
            this.CellControl_Button.TabIndex = 20;
            this.CellControl_Button.Text = "Cell Control";
            this.CellControl_Button.UseVisualStyleBackColor = true;
            this.CellControl_Button.Click += new System.EventHandler(this.CellControl_Button_Click);
            // 
            // ReportView_Button
            // 
            this.ReportView_Button.Font = new System.Drawing.Font("Microsoft Sans Serif", 25.25F, System.Drawing.FontStyle.Bold);
            this.ReportView_Button.Location = new System.Drawing.Point(6, 99);
            this.ReportView_Button.Name = "ReportView_Button";
            this.ReportView_Button.Size = new System.Drawing.Size(285, 47);
            this.ReportView_Button.TabIndex = 21;
            this.ReportView_Button.Text = "Report View";
            this.ReportView_Button.UseVisualStyleBackColor = true;
            this.ReportView_Button.Click += new System.EventHandler(this.ReportView_Button_Click);
            // 
            // EmployeeName_ComboBox
            // 
            this.EmployeeName_ComboBox.Font = new System.Drawing.Font("Microsoft Sans Serif", 25.25F);
            this.EmployeeName_ComboBox.FormattingEnabled = true;
            this.EmployeeName_ComboBox.Location = new System.Drawing.Point(659, 408);
            this.EmployeeName_ComboBox.Name = "EmployeeName_ComboBox";
            this.EmployeeName_ComboBox.Size = new System.Drawing.Size(608, 47);
            this.EmployeeName_ComboBox.TabIndex = 22;
            this.EmployeeName_ComboBox.SelectedIndexChanged += new System.EventHandler(this.EmployeeName_ComboBox_SelectedIndexChanged);
            // 
            // OPC_Button
            // 
            this.OPC_Button.Font = new System.Drawing.Font("Microsoft Sans Serif", 17.25F, System.Drawing.FontStyle.Bold);
            this.OPC_Button.Location = new System.Drawing.Point(6, 69);
            this.OPC_Button.Name = "OPC_Button";
            this.OPC_Button.Size = new System.Drawing.Size(199, 37);
            this.OPC_Button.TabIndex = 23;
            this.OPC_Button.Text = "OPC";
            this.OPC_Button.UseVisualStyleBackColor = true;
            // 
            // TestForm_Button
            // 
            this.TestForm_Button.Font = new System.Drawing.Font("Microsoft Sans Serif", 17.25F, System.Drawing.FontStyle.Bold);
            this.TestForm_Button.Location = new System.Drawing.Point(6, 112);
            this.TestForm_Button.Name = "TestForm_Button";
            this.TestForm_Button.Size = new System.Drawing.Size(199, 37);
            this.TestForm_Button.TabIndex = 24;
            this.TestForm_Button.Text = "Test Form";
            this.TestForm_Button.UseVisualStyleBackColor = true;
            this.TestForm_Button.Click += new System.EventHandler(this.TestForm_Button_Click);
            // 
            // Help_Button
            // 
            this.Help_Button.Font = new System.Drawing.Font("Microsoft Sans Serif", 27.25F, System.Drawing.FontStyle.Bold);
            this.Help_Button.Location = new System.Drawing.Point(1692, 801);
            this.Help_Button.Name = "Help_Button";
            this.Help_Button.Size = new System.Drawing.Size(185, 60);
            this.Help_Button.TabIndex = 25;
            this.Help_Button.Text = "Help";
            this.Help_Button.UseVisualStyleBackColor = true;
            this.Help_Button.Click += new System.EventHandler(this.Help_Button_Click);
            // 
            // ScanOut_Button
            // 
            this.ScanOut_Button.Font = new System.Drawing.Font("Microsoft Sans Serif", 17.25F, System.Drawing.FontStyle.Bold);
            this.ScanOut_Button.Location = new System.Drawing.Point(6, 26);
            this.ScanOut_Button.Name = "ScanOut_Button";
            this.ScanOut_Button.Size = new System.Drawing.Size(199, 37);
            this.ScanOut_Button.TabIndex = 26;
            this.ScanOut_Button.Text = "Scan Out";
            this.ScanOut_Button.UseVisualStyleBackColor = true;
            this.ScanOut_Button.Click += new System.EventHandler(this.ScanOut_Button_Click);
            // 
            // Test_GroupBox
            // 
            this.Test_GroupBox.Controls.Add(this.TextFile_Button);
            this.Test_GroupBox.Controls.Add(this.ScanOut_Button);
            this.Test_GroupBox.Controls.Add(this.OPC_Button);
            this.Test_GroupBox.Controls.Add(this.TestForm_Button);
            this.Test_GroupBox.Location = new System.Drawing.Point(201, 533);
            this.Test_GroupBox.Name = "Test_GroupBox";
            this.Test_GroupBox.Size = new System.Drawing.Size(211, 206);
            this.Test_GroupBox.TabIndex = 27;
            this.Test_GroupBox.TabStop = false;
            this.Test_GroupBox.Visible = false;
            // 
            // TextFile_Button
            // 
            this.TextFile_Button.Font = new System.Drawing.Font("Microsoft Sans Serif", 17.25F, System.Drawing.FontStyle.Bold);
            this.TextFile_Button.Location = new System.Drawing.Point(6, 155);
            this.TextFile_Button.Name = "TextFile_Button";
            this.TextFile_Button.Size = new System.Drawing.Size(199, 37);
            this.TextFile_Button.TabIndex = 27;
            this.TextFile_Button.Text = "Text File";
            this.TextFile_Button.UseVisualStyleBackColor = true;
            this.TextFile_Button.Click += new System.EventHandler(this.TextFile_Button_Click);
            // 
            // Login_GroupBox
            // 
            this.Login_GroupBox.Controls.Add(this.AdminLogin_Button);
            this.Login_GroupBox.Controls.Add(this.JobList_Button);
            this.Login_GroupBox.Controls.Add(this.CellControl_Button);
            this.Login_GroupBox.Controls.Add(this.ReportView_Button);
            this.Login_GroupBox.Location = new System.Drawing.Point(1329, 478);
            this.Login_GroupBox.Name = "Login_GroupBox";
            this.Login_GroupBox.Size = new System.Drawing.Size(297, 308);
            this.Login_GroupBox.TabIndex = 28;
            this.Login_GroupBox.TabStop = false;
            this.Login_GroupBox.Visible = false;
            // 
            // pictureBox1
            // 
            this.pictureBox1.Image = global::DMP_Spot_Weld_Application.Properties.Resources.DMP_Logo;
            this.pictureBox1.Location = new System.Drawing.Point(278, 73);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(1348, 306);
            this.pictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.pictureBox1.TabIndex = 3;
            this.pictureBox1.TabStop = false;
            // 
            // DMP_Spot_Weld_Login
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1904, 1042);
            this.ControlBox = false;
            this.Controls.Add(this.Login_GroupBox);
            this.Controls.Add(this.Test_GroupBox);
            this.Controls.Add(this.Help_Button);
            this.Controls.Add(this.Exit_Button);
            this.Controls.Add(this.Clock_TextBox);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.DMPID_TextBox);
            this.Controls.Add(this.EmployeeName_Label);
            this.Controls.Add(this.LoginGridView);
            this.Controls.Add(this.pictureBox1);
            this.Controls.Add(this.ListBox);
            this.Controls.Add(this.OperatorLogin_Button);
            this.Controls.Add(this.EmployeeName_TextBox);
            this.Controls.Add(this.EmployeeName_ComboBox);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximumSize = new System.Drawing.Size(1920, 1080);
            this.MinimumSize = new System.Drawing.Size(1918, 1038);
            this.Name = "DMP_Spot_Weld_Login";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "DMP Spot Weld Login";
            this.WindowState = System.Windows.Forms.FormWindowState.Maximized;
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.DMP_Spot_Weld_Login_FormClosing);
            this.Load += new System.EventHandler(this.DMP_Spot_Weld_Login_Load);
            ((System.ComponentModel.ISupportInitialize)(this.LoginGridView)).EndInit();
            this.Test_GroupBox.ResumeLayout(false);
            this.Login_GroupBox.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.Button OperatorLogin_Button;
        private System.Windows.Forms.ListBox ListBox;
        private System.Windows.Forms.Timer Clock;
        private System.Windows.Forms.PictureBox pictureBox1;
        private System.Windows.Forms.DataGridView LoginGridView;
        private System.Windows.Forms.Label EmployeeName_Label;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox DMPID_TextBox;
        private System.Windows.Forms.TextBox Clock_TextBox;
        private System.Windows.Forms.Button AdminLogin_Button;
        private System.Windows.Forms.Button Exit_Button;
        private System.Windows.Forms.Button JobList_Button;
        private System.Windows.Forms.Button CellControl_Button;
        private System.Windows.Forms.Button ReportView_Button;
        private System.Windows.Forms.ComboBox EmployeeName_ComboBox;
        private System.Windows.Forms.Button OPC_Button;
        public System.Windows.Forms.TextBox EmployeeName_TextBox;
        private System.Windows.Forms.Button TestForm_Button;
        private System.Windows.Forms.Button Help_Button;
        private System.Windows.Forms.Button ScanOut_Button;
        private System.Windows.Forms.GroupBox Test_GroupBox;
        private System.Windows.Forms.GroupBox Login_GroupBox;
        private System.Windows.Forms.Button TextFile_Button;
    }
}

