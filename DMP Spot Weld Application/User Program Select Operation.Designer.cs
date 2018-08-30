namespace DMP_Spot_Weld_Application
{
    partial class User_Program_Select_Operation
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(User_Program_Select_Operation));
            this.Operation_3_Button = new System.Windows.Forms.Button();
            this.Operation_4_Button = new System.Windows.Forms.Button();
            this.Operation_2_Button = new System.Windows.Forms.Button();
            this.Operation_1_Button = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // Operation_3_Button
            // 
            this.Operation_3_Button.Font = new System.Drawing.Font("Microsoft Sans Serif", 38.75F, System.Drawing.FontStyle.Bold);
            this.Operation_3_Button.Location = new System.Drawing.Point(651, 16);
            this.Operation_3_Button.Name = "Operation_3_Button";
            this.Operation_3_Button.Size = new System.Drawing.Size(275, 171);
            this.Operation_3_Button.TabIndex = 233;
            this.Operation_3_Button.Text = "Operation #3";
            this.Operation_3_Button.UseVisualStyleBackColor = true;
            this.Operation_3_Button.Click += new System.EventHandler(this.Operation_3_Button_Click);
            // 
            // Operation_4_Button
            // 
            this.Operation_4_Button.Font = new System.Drawing.Font("Microsoft Sans Serif", 38.75F, System.Drawing.FontStyle.Bold);
            this.Operation_4_Button.Location = new System.Drawing.Point(970, 16);
            this.Operation_4_Button.Name = "Operation_4_Button";
            this.Operation_4_Button.Size = new System.Drawing.Size(275, 171);
            this.Operation_4_Button.TabIndex = 232;
            this.Operation_4_Button.Text = "Operation #4";
            this.Operation_4_Button.UseVisualStyleBackColor = true;
            this.Operation_4_Button.Click += new System.EventHandler(this.Operation_4_Button_Click);
            // 
            // Operation_2_Button
            // 
            this.Operation_2_Button.Font = new System.Drawing.Font("Microsoft Sans Serif", 38.75F, System.Drawing.FontStyle.Bold);
            this.Operation_2_Button.Location = new System.Drawing.Point(332, 16);
            this.Operation_2_Button.Name = "Operation_2_Button";
            this.Operation_2_Button.Size = new System.Drawing.Size(275, 171);
            this.Operation_2_Button.TabIndex = 231;
            this.Operation_2_Button.Text = "Operation #2";
            this.Operation_2_Button.UseVisualStyleBackColor = true;
            this.Operation_2_Button.Click += new System.EventHandler(this.Operation_2_Button_Click);
            // 
            // Operation_1_Button
            // 
            this.Operation_1_Button.Font = new System.Drawing.Font("Microsoft Sans Serif", 38.75F, System.Drawing.FontStyle.Bold);
            this.Operation_1_Button.Location = new System.Drawing.Point(13, 16);
            this.Operation_1_Button.Name = "Operation_1_Button";
            this.Operation_1_Button.Size = new System.Drawing.Size(275, 171);
            this.Operation_1_Button.TabIndex = 230;
            this.Operation_1_Button.Text = "Operation #1";
            this.Operation_1_Button.UseVisualStyleBackColor = true;
            this.Operation_1_Button.Click += new System.EventHandler(this.Operation_1_Button_Click);
            // 
            // User_Program_Select_Operation
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1257, 199);
            this.ControlBox = false;
            this.Controls.Add(this.Operation_3_Button);
            this.Controls.Add(this.Operation_4_Button);
            this.Controls.Add(this.Operation_2_Button);
            this.Controls.Add(this.Operation_1_Button);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "User_Program_Select_Operation";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.User_Program_Select_Operation_FormClosing);
            this.Load += new System.EventHandler(this.User_Program_Select_Operation_Load);
            this.ResumeLayout(false);

        }

        #endregion

        public System.Windows.Forms.Button Operation_3_Button;
        public System.Windows.Forms.Button Operation_4_Button;
        public System.Windows.Forms.Button Operation_2_Button;
        public System.Windows.Forms.Button Operation_1_Button;
    }
}