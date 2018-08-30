namespace DMP_Spot_Weld_Application
{
    partial class ExcelTest
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
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.Connect_Button = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.SuspendLayout();
            // 
            // dataGridView1
            // 
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Location = new System.Drawing.Point(13, 13);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.Size = new System.Drawing.Size(1372, 425);
            this.dataGridView1.TabIndex = 0;
            // 
            // Connect_Button
            // 
            this.Connect_Button.Font = new System.Drawing.Font("Microsoft Sans Serif", 16F, System.Drawing.FontStyle.Bold);
            this.Connect_Button.Location = new System.Drawing.Point(244, 444);
            this.Connect_Button.Name = "Connect_Button";
            this.Connect_Button.Size = new System.Drawing.Size(239, 37);
            this.Connect_Button.TabIndex = 205;
            this.Connect_Button.Text = "Connect";
            this.Connect_Button.UseVisualStyleBackColor = true;
            this.Connect_Button.Visible = false;
            this.Connect_Button.Click += new System.EventHandler(this.Connect_Button_Click);
            // 
            // ExcelTest
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1397, 576);
            this.Controls.Add(this.Connect_Button);
            this.Controls.Add(this.dataGridView1);
            this.Name = "ExcelTest";
            this.Text = "ExcelTest";
            this.Load += new System.EventHandler(this.ExcelTest_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.Button Connect_Button;
    }
}