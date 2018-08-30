namespace DMP_Spot_Weld_Application
{
    partial class View_PDF
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(View_PDF));
            this.Close_Button = new System.Windows.Forms.Button();
            this.AcroPDF = new AxAcroPDFLib.AxAcroPDF();
            this.ExcelWebBrowser = new System.Windows.Forms.WebBrowser();
            ((System.ComponentModel.ISupportInitialize)(this.AcroPDF)).BeginInit();
            this.SuspendLayout();
            // 
            // Close_Button
            // 
            this.Close_Button.Font = new System.Drawing.Font("Microsoft Sans Serif", 40.25F);
            this.Close_Button.Location = new System.Drawing.Point(0, 900);
            this.Close_Button.Name = "Close_Button";
            this.Close_Button.Size = new System.Drawing.Size(1905, 75);
            this.Close_Button.TabIndex = 1;
            this.Close_Button.Text = "Close";
            this.Close_Button.UseVisualStyleBackColor = true;
            this.Close_Button.Click += new System.EventHandler(this.Close_Button_Click);
            // 
            // AcroPDF
            // 
            this.AcroPDF.Enabled = true;
            this.AcroPDF.Location = new System.Drawing.Point(0, 0);
            this.AcroPDF.Name = "AcroPDF";
            this.AcroPDF.OcxState = ((System.Windows.Forms.AxHost.State)(resources.GetObject("AcroPDF.OcxState")));
            this.AcroPDF.Size = new System.Drawing.Size(1905, 899);
            this.AcroPDF.TabIndex = 2;
            // 
            // ExcelWebBrowser
            // 
            this.ExcelWebBrowser.Location = new System.Drawing.Point(0, 0);
            this.ExcelWebBrowser.MinimumSize = new System.Drawing.Size(20, 20);
            this.ExcelWebBrowser.Name = "ExcelWebBrowser";
            this.ExcelWebBrowser.Size = new System.Drawing.Size(1905, 899);
            this.ExcelWebBrowser.TabIndex = 3;
            // 
            // View_PDF
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1904, 1042);
            this.ControlBox = false;
            this.Controls.Add(this.Close_Button);
            this.Controls.Add(this.ExcelWebBrowser);
            this.Controls.Add(this.AcroPDF);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MinimumSize = new System.Drawing.Size(1598, 38);
            this.Name = "View_PDF";
            this.Text = "View PDF";
            this.WindowState = System.Windows.Forms.FormWindowState.Maximized;
            ((System.ComponentModel.ISupportInitialize)(this.AcroPDF)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion
        private System.Windows.Forms.Button Close_Button;
        protected internal AxAcroPDFLib.AxAcroPDF AcroPDF;
        public System.Windows.Forms.WebBrowser ExcelWebBrowser;
    }
}