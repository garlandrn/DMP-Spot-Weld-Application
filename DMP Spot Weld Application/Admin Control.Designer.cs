namespace DMP_Spot_Weld_Application
{
    partial class Admin_Control
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Admin_Control));
            this.AddUser_Button = new System.Windows.Forms.Button();
            this.RemoveUser_Button = new System.Windows.Forms.Button();
            this.EditUser_Button = new System.Windows.Forms.Button();
            this.LogOff_Button = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.DMPID_TextBox = new System.Windows.Forms.TextBox();
            this.EmployeeName_TextBox = new System.Windows.Forms.TextBox();
            this.EmployeePassword_TextBox = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.AdminGridView = new System.Windows.Forms.DataGridView();
            this.Clock = new System.Windows.Forms.Timer(this.components);
            this.UserNumber_TextBox = new System.Windows.Forms.TextBox();
            this.Clock_TextBox = new System.Windows.Forms.TextBox();
            this.User_TextBox = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.Cancel_Button = new System.Windows.Forms.Button();
            this.Confirm_Button = new System.Windows.Forms.Button();
            this.Search_Button = new System.Windows.Forms.Button();
            this.SearchName_CheckBox = new System.Windows.Forms.CheckBox();
            this.SearchDMPID_CheckBox = new System.Windows.Forms.CheckBox();
            this.Clear_Button = new System.Windows.Forms.Button();
            this.LogInDataGridView = new System.Windows.Forms.DataGridView();
            ((System.ComponentModel.ISupportInitialize)(this.AdminGridView)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.LogInDataGridView)).BeginInit();
            this.SuspendLayout();
            // 
            // AddUser_Button
            // 
            this.AddUser_Button.Font = new System.Drawing.Font("Microsoft Sans Serif", 16F, System.Drawing.FontStyle.Bold);
            this.AddUser_Button.Location = new System.Drawing.Point(1737, 444);
            this.AddUser_Button.Name = "AddUser_Button";
            this.AddUser_Button.Size = new System.Drawing.Size(148, 40);
            this.AddUser_Button.TabIndex = 0;
            this.AddUser_Button.Text = "Add User";
            this.AddUser_Button.UseVisualStyleBackColor = true;
            this.AddUser_Button.Click += new System.EventHandler(this.AddUser_Button_Click);
            // 
            // RemoveUser_Button
            // 
            this.RemoveUser_Button.Font = new System.Drawing.Font("Microsoft Sans Serif", 16F, System.Drawing.FontStyle.Bold);
            this.RemoveUser_Button.Location = new System.Drawing.Point(1737, 536);
            this.RemoveUser_Button.Name = "RemoveUser_Button";
            this.RemoveUser_Button.Size = new System.Drawing.Size(148, 40);
            this.RemoveUser_Button.TabIndex = 1;
            this.RemoveUser_Button.Text = "Remove User";
            this.RemoveUser_Button.UseVisualStyleBackColor = true;
            this.RemoveUser_Button.Click += new System.EventHandler(this.RemoveUser_Button_Click);
            // 
            // EditUser_Button
            // 
            this.EditUser_Button.Font = new System.Drawing.Font("Microsoft Sans Serif", 16F, System.Drawing.FontStyle.Bold);
            this.EditUser_Button.Location = new System.Drawing.Point(1737, 490);
            this.EditUser_Button.Name = "EditUser_Button";
            this.EditUser_Button.Size = new System.Drawing.Size(148, 40);
            this.EditUser_Button.TabIndex = 2;
            this.EditUser_Button.Text = "Edit User";
            this.EditUser_Button.UseVisualStyleBackColor = true;
            this.EditUser_Button.Click += new System.EventHandler(this.EditUser_Button_Click);
            // 
            // LogOff_Button
            // 
            this.LogOff_Button.Font = new System.Drawing.Font("Microsoft Sans Serif", 16F, System.Drawing.FontStyle.Bold);
            this.LogOff_Button.Location = new System.Drawing.Point(1772, 130);
            this.LogOff_Button.Name = "LogOff_Button";
            this.LogOff_Button.Size = new System.Drawing.Size(120, 40);
            this.LogOff_Button.TabIndex = 7;
            this.LogOff_Button.Text = "Log Off";
            this.LogOff_Button.UseVisualStyleBackColor = true;
            this.LogOff_Button.Click += new System.EventHandler(this.LogOff_Button_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 16F, System.Drawing.FontStyle.Bold);
            this.label1.Location = new System.Drawing.Point(1291, 321);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(102, 26);
            this.label1.TabIndex = 8;
            this.label1.Text = "DMP ID:";
            // 
            // DMPID_TextBox
            // 
            this.DMPID_TextBox.Font = new System.Drawing.Font("Microsoft Sans Serif", 15F);
            this.DMPID_TextBox.Location = new System.Drawing.Point(1291, 350);
            this.DMPID_TextBox.Name = "DMPID_TextBox";
            this.DMPID_TextBox.ReadOnly = true;
            this.DMPID_TextBox.Size = new System.Drawing.Size(294, 30);
            this.DMPID_TextBox.TabIndex = 3;
            // 
            // EmployeeName_TextBox
            // 
            this.EmployeeName_TextBox.Font = new System.Drawing.Font("Microsoft Sans Serif", 15F);
            this.EmployeeName_TextBox.Location = new System.Drawing.Point(1291, 288);
            this.EmployeeName_TextBox.Name = "EmployeeName_TextBox";
            this.EmployeeName_TextBox.Size = new System.Drawing.Size(294, 30);
            this.EmployeeName_TextBox.TabIndex = 1;
            // 
            // EmployeePassword_TextBox
            // 
            this.EmployeePassword_TextBox.Font = new System.Drawing.Font("Microsoft Sans Serif", 15F);
            this.EmployeePassword_TextBox.Location = new System.Drawing.Point(1598, 288);
            this.EmployeePassword_TextBox.Name = "EmployeePassword_TextBox";
            this.EmployeePassword_TextBox.Size = new System.Drawing.Size(294, 30);
            this.EmployeePassword_TextBox.TabIndex = 2;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 16F, System.Drawing.FontStyle.Bold);
            this.label2.Location = new System.Drawing.Point(1286, 259);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(195, 26);
            this.label2.TabIndex = 12;
            this.label2.Text = "Employee Name:";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 16F, System.Drawing.FontStyle.Bold);
            this.label3.Location = new System.Drawing.Point(1586, 259);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(236, 26);
            this.label3.TabIndex = 13;
            this.label3.Text = "Employee Password:";
            // 
            // AdminGridView
            // 
            this.AdminGridView.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.AdminGridView.Location = new System.Drawing.Point(1291, 444);
            this.AdminGridView.Name = "AdminGridView";
            this.AdminGridView.ReadOnly = true;
            this.AdminGridView.Size = new System.Drawing.Size(398, 286);
            this.AdminGridView.TabIndex = 14;
            this.AdminGridView.CellClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.AdminGridView_CellClick);
            // 
            // Clock
            // 
            this.Clock.Interval = 150;
            this.Clock.Tick += new System.EventHandler(this.Clock_Tick);
            // 
            // UserNumber_TextBox
            // 
            this.UserNumber_TextBox.Font = new System.Drawing.Font("Microsoft Sans Serif", 16F);
            this.UserNumber_TextBox.Location = new System.Drawing.Point(1615, 52);
            this.UserNumber_TextBox.Name = "UserNumber_TextBox";
            this.UserNumber_TextBox.ReadOnly = true;
            this.UserNumber_TextBox.Size = new System.Drawing.Size(277, 32);
            this.UserNumber_TextBox.TabIndex = 72;
            this.UserNumber_TextBox.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // Clock_TextBox
            // 
            this.Clock_TextBox.Font = new System.Drawing.Font("Microsoft Sans Serif", 16F);
            this.Clock_TextBox.Location = new System.Drawing.Point(1615, 92);
            this.Clock_TextBox.Name = "Clock_TextBox";
            this.Clock_TextBox.ReadOnly = true;
            this.Clock_TextBox.Size = new System.Drawing.Size(277, 32);
            this.Clock_TextBox.TabIndex = 71;
            this.Clock_TextBox.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // User_TextBox
            // 
            this.User_TextBox.Font = new System.Drawing.Font("Microsoft Sans Serif", 16F);
            this.User_TextBox.Location = new System.Drawing.Point(1615, 12);
            this.User_TextBox.Name = "User_TextBox";
            this.User_TextBox.ReadOnly = true;
            this.User_TextBox.Size = new System.Drawing.Size(277, 32);
            this.User_TextBox.TabIndex = 70;
            this.User_TextBox.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("Microsoft Sans Serif", 16F, System.Drawing.FontStyle.Bold);
            this.label4.Location = new System.Drawing.Point(1454, 12);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(155, 26);
            this.label4.TabIndex = 70;
            this.label4.Text = "Current User:";
            // 
            // Cancel_Button
            // 
            this.Cancel_Button.Font = new System.Drawing.Font("Microsoft Sans Serif", 16F, System.Drawing.FontStyle.Bold);
            this.Cancel_Button.Location = new System.Drawing.Point(1737, 690);
            this.Cancel_Button.Name = "Cancel_Button";
            this.Cancel_Button.Size = new System.Drawing.Size(148, 40);
            this.Cancel_Button.TabIndex = 16;
            this.Cancel_Button.Text = "Cancel";
            this.Cancel_Button.UseVisualStyleBackColor = true;
            this.Cancel_Button.Visible = false;
            this.Cancel_Button.Click += new System.EventHandler(this.Cancel_Button_Click);
            // 
            // Confirm_Button
            // 
            this.Confirm_Button.Font = new System.Drawing.Font("Microsoft Sans Serif", 16F, System.Drawing.FontStyle.Bold);
            this.Confirm_Button.Location = new System.Drawing.Point(1737, 644);
            this.Confirm_Button.Name = "Confirm_Button";
            this.Confirm_Button.Size = new System.Drawing.Size(148, 40);
            this.Confirm_Button.TabIndex = 4;
            this.Confirm_Button.Text = "Confirm";
            this.Confirm_Button.UseVisualStyleBackColor = true;
            this.Confirm_Button.Visible = false;
            this.Confirm_Button.Click += new System.EventHandler(this.Confirm_Button_Click);
            // 
            // Search_Button
            // 
            this.Search_Button.Font = new System.Drawing.Font("Microsoft Sans Serif", 16F, System.Drawing.FontStyle.Bold);
            this.Search_Button.Location = new System.Drawing.Point(1421, 398);
            this.Search_Button.Name = "Search_Button";
            this.Search_Button.Size = new System.Drawing.Size(124, 40);
            this.Search_Button.TabIndex = 75;
            this.Search_Button.Text = "Search";
            this.Search_Button.UseVisualStyleBackColor = true;
            this.Search_Button.Click += new System.EventHandler(this.Search_Button_Click);
            // 
            // SearchName_CheckBox
            // 
            this.SearchName_CheckBox.AutoSize = true;
            this.SearchName_CheckBox.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.25F);
            this.SearchName_CheckBox.Location = new System.Drawing.Point(1591, 355);
            this.SearchName_CheckBox.Name = "SearchName_CheckBox";
            this.SearchName_CheckBox.Size = new System.Drawing.Size(133, 21);
            this.SearchName_CheckBox.TabIndex = 76;
            this.SearchName_CheckBox.Text = "Search By Name";
            this.SearchName_CheckBox.UseVisualStyleBackColor = true;
            this.SearchName_CheckBox.CheckedChanged += new System.EventHandler(this.SearchName_CheckBox_CheckedChanged);
            // 
            // SearchDMPID_CheckBox
            // 
            this.SearchDMPID_CheckBox.AutoSize = true;
            this.SearchDMPID_CheckBox.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.25F);
            this.SearchDMPID_CheckBox.Location = new System.Drawing.Point(1730, 355);
            this.SearchDMPID_CheckBox.Name = "SearchDMPID_CheckBox";
            this.SearchDMPID_CheckBox.Size = new System.Drawing.Size(143, 21);
            this.SearchDMPID_CheckBox.TabIndex = 77;
            this.SearchDMPID_CheckBox.Text = "Search By DMP ID";
            this.SearchDMPID_CheckBox.UseVisualStyleBackColor = true;
            this.SearchDMPID_CheckBox.CheckedChanged += new System.EventHandler(this.SearchDMPID_CheckBox_CheckedChanged);
            // 
            // Clear_Button
            // 
            this.Clear_Button.Font = new System.Drawing.Font("Microsoft Sans Serif", 16F, System.Drawing.FontStyle.Bold);
            this.Clear_Button.Location = new System.Drawing.Point(1291, 398);
            this.Clear_Button.Name = "Clear_Button";
            this.Clear_Button.Size = new System.Drawing.Size(124, 40);
            this.Clear_Button.TabIndex = 78;
            this.Clear_Button.Text = "Clear";
            this.Clear_Button.UseVisualStyleBackColor = true;
            this.Clear_Button.Click += new System.EventHandler(this.Clear_Button_Click);
            // 
            // LogInDataGridView
            // 
            this.LogInDataGridView.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.LogInDataGridView.Location = new System.Drawing.Point(12, 152);
            this.LogInDataGridView.Name = "LogInDataGridView";
            this.LogInDataGridView.ReadOnly = true;
            this.LogInDataGridView.Size = new System.Drawing.Size(1071, 578);
            this.LogInDataGridView.TabIndex = 79;
            // 
            // Admin_Control
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1904, 1054);
            this.ControlBox = false;
            this.Controls.Add(this.LogInDataGridView);
            this.Controls.Add(this.Clear_Button);
            this.Controls.Add(this.SearchDMPID_CheckBox);
            this.Controls.Add(this.SearchName_CheckBox);
            this.Controls.Add(this.Search_Button);
            this.Controls.Add(this.Confirm_Button);
            this.Controls.Add(this.Cancel_Button);
            this.Controls.Add(this.UserNumber_TextBox);
            this.Controls.Add(this.Clock_TextBox);
            this.Controls.Add(this.User_TextBox);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.AdminGridView);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.EmployeePassword_TextBox);
            this.Controls.Add(this.EmployeeName_TextBox);
            this.Controls.Add(this.DMPID_TextBox);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.LogOff_Button);
            this.Controls.Add(this.EditUser_Button);
            this.Controls.Add(this.RemoveUser_Button);
            this.Controls.Add(this.AddUser_Button);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "Admin_Control";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Admin Control";
            this.Load += new System.EventHandler(this.Admin_Control_Load);
            ((System.ComponentModel.ISupportInitialize)(this.AdminGridView)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.LogInDataGridView)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button AddUser_Button;
        private System.Windows.Forms.Button RemoveUser_Button;
        private System.Windows.Forms.Button EditUser_Button;
        private System.Windows.Forms.Button LogOff_Button;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox DMPID_TextBox;
        private System.Windows.Forms.TextBox EmployeeName_TextBox;
        private System.Windows.Forms.TextBox EmployeePassword_TextBox;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;


        private System.Windows.Forms.DataGridView AdminGridView;

        private System.Windows.Forms.Timer Clock;
        public System.Windows.Forms.TextBox UserNumber_TextBox;
        public System.Windows.Forms.TextBox Clock_TextBox;
        public System.Windows.Forms.TextBox User_TextBox;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Button Cancel_Button;
        private System.Windows.Forms.Button Confirm_Button;
        private System.Windows.Forms.Button Search_Button;
        private System.Windows.Forms.CheckBox SearchName_CheckBox;
        private System.Windows.Forms.CheckBox SearchDMPID_CheckBox;
        private System.Windows.Forms.Button Clear_Button;
        private System.Windows.Forms.DataGridView LogInDataGridView;
    }
}