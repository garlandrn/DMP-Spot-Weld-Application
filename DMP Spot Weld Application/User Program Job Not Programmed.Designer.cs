namespace DMP_Spot_Weld_Application
{
    partial class User_Program_Job_Not_Programmed
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(User_Program_Job_Not_Programmed));
            this.label1 = new System.Windows.Forms.Label();
            this.OK_Button = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 88.25F, System.Drawing.FontStyle.Bold);
            this.label1.Location = new System.Drawing.Point(113, 9);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(750, 402);
            this.label1.TabIndex = 0;
            this.label1.Text = "Job Not\r\nProgrammed\r\nOn PLC";
            this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // OK_Button
            // 
            this.OK_Button.BackColor = System.Drawing.Color.Red;
            this.OK_Button.Font = new System.Drawing.Font("Microsoft Sans Serif", 34.75F, System.Drawing.FontStyle.Bold);
            this.OK_Button.Location = new System.Drawing.Point(827, 333);
            this.OK_Button.Name = "OK_Button";
            this.OK_Button.Size = new System.Drawing.Size(146, 79);
            this.OK_Button.TabIndex = 82;
            this.OK_Button.Text = "OK";
            this.OK_Button.UseVisualStyleBackColor = false;
            this.OK_Button.Click += new System.EventHandler(this.OK_Button_Click);
            // 
            // User_Program_Job_Not_Programmed
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.Red;
            this.ClientSize = new System.Drawing.Size(985, 424);
            this.ControlBox = false;
            this.Controls.Add(this.OK_Button);
            this.Controls.Add(this.label1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "User_Program_Job_Not_Programmed";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button OK_Button;
    }
}