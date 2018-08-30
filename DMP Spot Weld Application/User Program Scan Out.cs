using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Net;

namespace DMP_Spot_Weld_Application
{
    public partial class User_Program_Scan_Out : Form
    {
        User_Program Owner = null;
        public User_Program_Scan_Out()
        {
            InitializeComponent();
            this.ShowInTaskbar = false;
            Completed_Button.DialogResult = DialogResult.Yes;
            Cancel_Button.DialogResult = DialogResult.No;
        }

        private static string LoginUsername = "";
        private static string LoginPassword = "WIPINV1";
        private static string LoginConfig = "OH";
        private static string DMPResID = "";
        private static bool FormCompleted = false;

        private static int LogOffValue = 0;

        private void User_Program_Scan_Out_Load(object sender, EventArgs e)
        {
            SpotWeldID();
            Close_Timer.Start();
        }

        /*********************************************************************************************************************
        * 
        * Buttons Region Start
        * -- Total: 3
        * 
        * - Completed Button Click
        * - Submit Button Click
        * - Cancel Button Click
        * 
        *********************************************************************************************************************/
        #region

        private void Completed_Button_Click(object sender, EventArgs e)
        {
            FormCompleted = false;
            User_Program.UserProgram.Enabled = true;
            this.Close();
        }
        
        private void Submit_Button_Click(object sender, EventArgs e)
        {
            HtmlElement EmpNum_TextBox = ScanOutBrowser.Document.GetElementById("EmpNum");
            HtmlElement JobNum_TextBox = ScanOutBrowser.Document.GetElementById("JobNum");
            HtmlElement OperNum_TextBox = ScanOutBrowser.Document.GetElementById("OperNum");
            HtmlElement DMPResID_TextBox = ScanOutBrowser.Document.GetElementById("DMPResID");
            HtmlElement TcQtuQtyComp_TextBox = ScanOutBrowser.Document.GetElementById("TcQtuQtyComp");
            HtmlElement TcQtuQtyScrap_TextBox = ScanOutBrowser.Document.GetElementById("TcQtuQtyScrap");
            HtmlElement TcQtuQtyMove_TextBox = ScanOutBrowser.Document.GetElementById("TcQtuQtyMove");
            HtmlElement Complete_TextBox = ScanOutBrowser.Document.GetElementById("Complete");
            HtmlElement Close_TextBox = ScanOutBrowser.Document.GetElementById("Close");
            if (JobNum_TextBox.GetAttribute("value") != "" && TcQtuQtyComp_TextBox.GetAttribute("value") != "" && TcQtuQtyScrap_TextBox.GetAttribute("value") != "" && TcQtuQtyMove_TextBox.GetAttribute("value") != "" && Complete_TextBox.GetAttribute("value") != "" && Close_TextBox.GetAttribute("value") != "")
            {
                ScanOutBrowser.Focus();
                TcQtuQtyScrap_TextBox.Focus();
                SendKeys.Send("{ENTER}");
                Submit_Button.Visible = false;
                Completed_Button.Visible = true;
                FormCompleted = true;
            }
        }

        private void Cancel_Button_Click(object sender, EventArgs e)
        {
            try
            {
                User_Program.UserProgram.Enabled = true;
            }
            catch
            {

            }
            this.Close();
        }

        /*********************************************************************************************************************
        * 
        * Buttons Region End
        * 
        *********************************************************************************************************************/
        #endregion

        /*********************************************************************************************************************
        * 
        * WebBrowser Region Start
        * -- Total: 1
        * 
        * - ScanOutBrowser_DocumentCompleted
        * 
        *********************************************************************************************************************/
        #region

        private void ScanOutBrowser_DocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs e)
        {
            LogOffValue = 0;
            if (ScanOutBrowser.Url.AbsoluteUri == ("http://ohsenslu803/fsdatacollection/Login.asp"))
            {
                HtmlElement UserID_TextBox = ScanOutBrowser.Document.GetElementById("UserID");
                HtmlElement Password_TextBox = ScanOutBrowser.Document.GetElementById("Password");
                HtmlElement Config_TextBox = ScanOutBrowser.Document.GetElementById("Config");
                UserID_TextBox.SetAttribute("value", LoginUsername);
                Password_TextBox.SetAttribute("value", LoginPassword);
                Config_TextBox.SetAttribute("value", LoginConfig);
                foreach(HtmlElement Button in this.ScanOutBrowser.Document.GetElementsByTagName("input"))
                {
                    if (Button.GetAttribute("value").Equals("Start"))
                    {
                        Button.InvokeMember("click");
                    }
                }
            }
            if(ScanOutBrowser.Url.AbsoluteUri == ("http://ohsenslu803/fsdatacollection/default.asp"))
            {
                HtmlElement MenuSelect_TextBox = ScanOutBrowser.Document.GetElementById("MenuSelect");
                MenuSelect_TextBox.SetAttribute("value", "1");
            }
            if (ScanOutBrowser.Url.AbsoluteUri == ("http://ohsenslu803/fsdatacollection/DMPJobMove.asp") && FormCompleted == true)
            {
                HtmlElement EmpNum_TextBox = ScanOutBrowser.Document.GetElementById("EmpNum");
                HtmlElement Complete_TextBox = ScanOutBrowser.Document.GetElementById("Complete");
                try
                {
                    Complete_TextBox.Focus();
                    SendKeys.Send("{ENTER}");
                }
                catch
                {

                } //
                if(EmpNum_TextBox.GetAttribute("value") == "" && FormCompleted == true)
                {
                    FormCompleted = false;
                }
            }
            if (ScanOutBrowser.Url.AbsoluteUri == ("http://ohsenslu803/fsdatacollection/DMPJobMove.asp") && FormCompleted == false)
            {
                HtmlElement EmpNum_TextBox = ScanOutBrowser.Document.GetElementById("EmpNum");
                HtmlElement JobNum_TextBox = ScanOutBrowser.Document.GetElementById("JobNum");
                HtmlElement OperNum_TextBox = ScanOutBrowser.Document.GetElementById("OperNum");
                HtmlElement DMPResID_TextBox = ScanOutBrowser.Document.GetElementById("DMPResID");
                HtmlElement TcQtuQtyComp_TextBox = ScanOutBrowser.Document.GetElementById("TcQtuQtyComp");
                HtmlElement TcQtuQtyScrap_TextBox = ScanOutBrowser.Document.GetElementById("TcQtuQtyScrap");
                HtmlElement TcQtuQtyMove_TextBox = ScanOutBrowser.Document.GetElementById("TcQtuQtyMove");
                HtmlElement Complete_TextBox = ScanOutBrowser.Document.GetElementById("Complete");
                HtmlElement Close_TextBox = ScanOutBrowser.Document.GetElementById("Close");
                try
                {
                    EmpNum_TextBox.SetAttribute("value", EmployeeNumber_TextBox.Text);
                    if(JobNum_TextBox.GetAttribute("value") == "")
                    {
                        JobNum_TextBox.SetAttribute("value", JobNumber_TextBox.Text);
                    }
                    if (TcQtuQtyComp_TextBox.GetAttribute("value") == "")
                    {
                        TcQtuQtyComp_TextBox.SetAttribute("value", TotalCountQtuQtyComp_TextBox.Text);
                    }
                    DMPResID_TextBox.SetAttribute("value", DMPResID);

                    if (JobNum_TextBox.GetAttribute("value") != "" && TcQtuQtyComp_TextBox.GetAttribute("value") != "" && TcQtuQtyScrap_TextBox.GetAttribute("value") != "" && TcQtuQtyMove_TextBox.GetAttribute("value") != "" && Complete_TextBox.GetAttribute("value") != "" && Close_TextBox.GetAttribute("value") != "" && FormCompleted == false)
                    {
                        //if(TcQtuQtyScrap_TextBox != true)
                        //Close_TextBox.Focus();
                        Submit_Button.Visible = true;
                        FormCompleted = true;
                    }
                }
                catch(Exception ex)
                {
                    //MessageBox.Show(ex.ToString());
                }
            }
        }

        /*********************************************************************************************************************
        * 
        * WebBrowser Region End
        * 
        *********************************************************************************************************************/
        #endregion

        /*********************************************************************************************************************
        * 
        * Method Region Start
        * -- Total: 1
        * 
        * - SpotWeldID
        * 
        *********************************************************************************************************************/
        #region

        private void SpotWeldID()
        {
            string SpotWeldComputerID = System.Environment.MachineName;

            // CAT
            if (SpotWeldComputerID == "123R") // Spot Weld 123R
            {
                LoginUsername = "DC123R";
                DMPResID = "SW-123R";
            }
            if (SpotWeldComputerID == "1088") // Spot Weld 1088
            {
                LoginUsername = "DC1088";
                DMPResID = "SW-1088";
            }  
            // John Deere
            if (SpotWeldComputerID == "108R") // Spot Weld 108R
            {
                LoginUsername = "DC108R";
                DMPResID = "SW-108R";
            }
            if (SpotWeldComputerID == "150R") // Spot Weld 150R
            {
                LoginUsername = "DC150R";
                DMPResID = "SW-150R";
            }
            // Navistar
            if (SpotWeldComputerID == "104R") // Spot Weld 104R
            {
                LoginUsername = "DC104R";
                DMPResID = "SW-104R";
            }
            if (SpotWeldComputerID == "OHN7149") // Spot Weld 121R
            {
                LoginUsername = "DC121R";
                DMPResID = "SW-121R";
            }
            if (SpotWeldComputerID == "OHN7111") // Spot Weld 154R
            {
                LoginUsername = "DC154R";
                DMPResID = "SW-154R";
            }
            // Paccar
            if (SpotWeldComputerID == "OHN7124") // Spot Weld 153R
            {
                LoginUsername = "DC153R";
                DMPResID = "SW-153R";
            }
            if (SpotWeldComputerID == "OHN7123") // Spot Weld 155R
            {
                LoginUsername = "DC155R";
                DMPResID = "SW-155R";
            }
            // My Computer For Testing
            if (SpotWeldComputerID == "OHN7047NL") // My Computer
            {
                LoginUsername = "DC154R";
                DMPResID = "SW-154R";
            }
            if (SpotWeldComputerID == "OHN7017") //  bp 1107
            {
                LoginUsername = "DC155R";
                DMPResID = "SW-155R";
            }
        }

        /*********************************************************************************************************************
        * 
        * Method Region End
        * 
        *********************************************************************************************************************/
        #endregion

        private void Close_Timer_Tick(object sender, EventArgs e)
        {
            LogOffValue = 1 + LogOffValue;
            if (LogOffValue > 300)
            {
                this.Close();
            }
        }
    }
}
