using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Windows.Forms;
using System.Xml;

/*
 * Program: DMP Brake Press Application
 * Form: DMPBrakePressLogin
 * Created By: Ryan Garland
 * Last Updated on 1/31/18
 * 
 * Form Sections:
 *  - User Interface
 *  --- Buttons
 *  --- ComboBox
 *  --- CheckBox
 *  --- GridView
 *  --- PictureBox
 *  
 *  - SQL DataBase Methods
 *  - Methods
 *  
 * - Clock 
 * 
 * 
 * 
 */

namespace DMP_Spot_Weld_Application
{
    public partial class DMP_Spot_Weld_Login : Form
    {
        public static Form Current;

        public DMP_Spot_Weld_Login()
        {
            InitializeComponent();
            Current = this;
        }

        /********************************************************************************************************************
        * 
        * Global Variables 
        * 
        * 
        * 
        ********************************************************************************************************************/
        /********************************************************************************************************************
        * 
        * Form Load Variables 
        * 
        ********************************************************************************************************************/

        string SQL_Connection = @"Data Source = OHN7009,49172; Initial Catalog = Spot_Weld_Data; Integrated Security = True; Connect Timeout = 15;";

        // Clock_Tick();
        private static int ClockHour;
        private static int ClockMinute;
        private static int ClockSecond;

        /********************************************************************************************************************
        * 
        * Variables In Testing Start
        * 
        ********************************************************************************************************************/
        

        /********************************************************************************************************************
        * 
        * Variables In Testing End
        * 
        *********************************************************************************************************************
        *********************************************************************************************************************
        * 
        * DMP Spot Weld Login Start
        * 
        ********************************************************************************************************************/

        private void DMP_Spot_Weld_Login_Load(object sender, EventArgs e)
        {
            Clock.Enabled = true;
            SpotWeldID();
                 
            SqlConnection connection = new SqlConnection(SQL_Connection);
            string LoginData = "SELECT * FROM [dbo].[Employee] ORDER BY EmployeeName ASC";
            SqlDataAdapter DataAdapter = new SqlDataAdapter(LoginData, connection);
            SqlCommandBuilder commandBuilder = new SqlCommandBuilder(DataAdapter);
            DataSet Data = new DataSet();
            DataAdapter.Fill(Data);
            LoginGridView.DataSource = Data.Tables[0];

            int rows = 0;
            string LoginCount = "SELECT COUNT(*) FROM [dbo].[Employee]";
            SqlConnection count = new SqlConnection(SQL_Connection);
            SqlCommand countRows = new SqlCommand(LoginCount, count);
            count.Open();
            rows = (int)countRows.ExecuteScalar();
            count.Close();

            foreach (DataGridViewRow row in LoginGridView.Rows)
            {
                if (row.Index < rows)
                {
                    EmployeeName_ComboBox.Items.Add(row.Cells[1].Value.ToString());
                }
            }
        }
        
        /********************************************************************************************************************
        * 
        * Buttons Region Start 
        * -- Total Buttons: 14
        * 
        * --- Buttons GroupBox Buttons
        * --- Total: 5
        * - Operator Login Click
        * - AdminLogin Click
        * - ReportView Click
        * - JobList Click
        * - CellControl Click
        * 
        * - Exit Click
        * 
        ********************************************************************************************************************/
        #region

        private void OperatorLogin_Button_Click(object sender, EventArgs e)
        {
            OperatorLogin();
        }

        private void AdminLogin_Button_Click(object sender, EventArgs e)
        {
            AdminLogin();
        }

        private void ReportView_Button_Click(object sender, EventArgs e)
        {
            ReportViewLogin();
        }

        private void JobList_Button_Click(object sender, EventArgs e)
        {
            JobListLogin();
        }

        private void CellControl_Button_Click(object sender, EventArgs e)
        {
            CellControlLogin();
        }

        private void Help_Button_Click(object sender, EventArgs e)
        {
            View_PDF PDFViewer = new View_PDF();
            PDFViewer.AcroPDF.src = @"\\insidedmp.com\Corporate\OH\OH Common\Engineering\Spotweld\Spotweld Application Files\Spot Weld User Manual.pdf";
            PDFViewer.AcroPDF.BringToFront();
            PDFViewer.Show();
            PDFViewer.BringToFront();
            PDFViewer.AcroPDF.setCurrentPage(2);
            PDFViewer.AcroPDF.setZoom(100);
        }

        private void Exit_Button_Click(object sender, EventArgs e)
        {
            Application.Exit();
            this.Close();
        }

        /********************************************************************************************************************
        * 
        * Buttons Region End
        * 
        ********************************************************************************************************************/
        #endregion

        /********************************************************************************************************************
        * 
        * Methods Region Start
        * -- Total: 8
        * 
        * - OperatorLogin
        * - AdminLogin
        * - ReportViewLogin
        * - JobListLogin
        * - CellControlLogin
        * - BrakePressID
        * - OpenNewForm
        * 
        ********************************************************************************************************************/
        #region

        private void OperatorLogin()
        {
            if (EmployeeName_TextBox.Text == "" && DMPID_TextBox.Text == "")
            {
                ListBox.Items.Add("Please Enter Your Name and DMP ID to Login");
            }
            else if (EmployeeName_TextBox.Text == "")
            {
                ListBox.Items.Add("Enter Employee Name to Login");
            }
            else if (DMPID_TextBox.Text == "")
            {
                ListBox.Items.Add("Enter DMP ID to Login");
            }
            else if (EmployeeName_TextBox.Text != "" && DMPID_TextBox.Text != "")
            {
                try
                {
                    SqlConnection OperatorLogin = new SqlConnection(SQL_Connection);
                    SqlCommand Login_Command = new SqlCommand("Select * from [dbo].[Employee] where EmployeeName=@EmployeeName and EmployeePassword=@EmployeePassword", OperatorLogin);
                    Login_Command.Parameters.AddWithValue("@EmployeeName", EmployeeName_ComboBox.Text);
                    Login_Command.Parameters.AddWithValue("@EmployeePassword", DMPID_TextBox.Text);
                    OperatorLogin.Open();
                    SqlDataAdapter adapt = new SqlDataAdapter(Login_Command);
                    DataSet ds = new DataSet();
                    adapt.Fill(ds);
                    LoginGridView.DataSource = ds.Tables[0];
                    OperatorLogin.Close();
                    int count = ds.Tables[0].Rows.Count;
                    if (count == 1)
                    {
                        User_Program UserProgram = new User_Program();
                        UserProgram.User_TextBox.Text = EmployeeName_TextBox.Text;
                        UserProgram.Clock_TextBox.Text = Clock_TextBox.Text;
                        UserProgram.DMPID_TextBox.Text = DMPID_TextBox.Text;
                        DMPID_TextBox.Clear();
                        EmployeeName_TextBox.Clear();
                        EmployeeName_ComboBox.Text = "";
                        ListBox.Items.Clear();
                        UserProgram.Show();
                        UserProgram.Focus();
                        OpenNewForm();
                    }
                    else
                    {
                        ListBox.Items.Add("Please Check DMP ID Entered");
                    }
                }
                catch (Exception e)
                {
                    //ListBox.Items.Add("Log In Failed. Employee Not Found.\n" + e.ToString());
                }
            }            
        }

        private void AdminLogin()
        {
            if (EmployeeName_TextBox.Text == "" && DMPID_TextBox.Text == "")
            {
                ListBox.Items.Add("Please Enter Your Name and DMP ID to Login");
            }
            else if (EmployeeName_TextBox.Text == "")
            {
                ListBox.Items.Add("Enter Employee Name to Login");
            }
            else if (DMPID_TextBox.Text == "")
            {
                ListBox.Items.Add("Enter DMP ID to Login");
            }
            else if (EmployeeName_TextBox.Text != "" && DMPID_TextBox.Text != "")
            {
                try
                {
                    SqlConnection AdminConnection = new SqlConnection(SQL_Connection);
                    SqlCommand SQL_Admin = new SqlCommand("Select * from [dbo].[Admin] where EmployeeName=@EmployeeName and EmployeePassword=@EmployeePassword", AdminConnection);
                    SQL_Admin.Parameters.AddWithValue("@EmployeeName", EmployeeName_TextBox.Text);
                    SQL_Admin.Parameters.AddWithValue("@EmployeePassword", DMPID_TextBox.Text);
                    AdminConnection.Open();
                    SqlDataAdapter adapt = new SqlDataAdapter(SQL_Admin);
                    DataSet ds = new DataSet();
                    adapt.Fill(ds);
                    LoginGridView.DataSource = ds.Tables[0];
                    AdminConnection.Close();
                    int count = ds.Tables[0].Rows.Count;
                    if (count == 1)
                    {
                        Admin_Control Admin = new Admin_Control();
                        Admin.User_TextBox.Text = EmployeeName_TextBox.Text;
                        Admin.Clock_TextBox.Text = Clock_TextBox.Text;
                        Admin.UserNumber_TextBox.Text = DMPID_TextBox.Text;
                        ListBox.Items.Clear();
                        Admin.Show();
                        Admin.Focus();
                        OpenNewForm();
                    }
                }
                catch (Exception)
                {
                    ListBox.Items.Add("Access to Admin Form Denied.");
                }
            }
        }

        private void JobListLogin()
        {
            if (EmployeeName_TextBox.Text == "" && DMPID_TextBox.Text == "")
            {
                ListBox.Items.Add("Please Enter Your Name and DMP ID to Login");
            }
            else if (EmployeeName_TextBox.Text == "")
            {
                ListBox.Items.Add("Enter Employee Name to Login");
            }
            else if (DMPID_TextBox.Text == "")
            {
                ListBox.Items.Add("Enter DMP ID to Login");
            }
            else if (EmployeeName_TextBox.Text != "" && DMPID_TextBox.Text != "")
            {
                try
                {
                    SqlConnection JobListConnection = new SqlConnection(SQL_Connection);
                    SqlCommand SQL_JobList = new SqlCommand("Select * from [dbo].[Admin] where EmployeeName=@EmployeeName and EmployeePassword=@EmployeePassword", JobListConnection);
                    SQL_JobList.Parameters.AddWithValue("@EmployeeName", EmployeeName_TextBox.Text);
                    SQL_JobList.Parameters.AddWithValue("@EmployeePassword", DMPID_TextBox.Text);
                    JobListConnection.Open();
                    SqlDataAdapter adapt = new SqlDataAdapter(SQL_JobList);
                    DataSet ds = new DataSet();
                    adapt.Fill(ds);
                    LoginGridView.DataSource = ds.Tables[0];
                    JobListConnection.Close();
                    int count = ds.Tables[0].Rows.Count;
                    if (count == 1)
                    {
                        JobList jobList = new JobList();
                        jobList.User_TextBox.Text = EmployeeName_TextBox.Text;
                        jobList.Clock_TextBox.Text = Clock_TextBox.Text;
                        jobList.DMPID_TextBox.Text = DMPID_TextBox.Text;
                        ListBox.Items.Clear();
                        jobList.Show();
                        jobList.Focus();
                        OpenNewForm();
                    }
                }
                catch (Exception)
                {
                    ListBox.Items.Add("Access to Report View Denied.");
                }
            }
        }

        private void CellControlLogin()
        {
            if (EmployeeName_TextBox.Text == "" && DMPID_TextBox.Text == "")
            {
                ListBox.Items.Add("Please Enter Your Name and DMP ID to Login");
            }
            else if (EmployeeName_TextBox.Text == "")
            {
                ListBox.Items.Add("Enter Employee Name to Login");
            }
            else if (DMPID_TextBox.Text == "")
            {
                ListBox.Items.Add("Enter DMP ID to Login");
            }
            else if (EmployeeName_TextBox.Text != "" && DMPID_TextBox.Text != "")
            {
                try
                {
                    SqlConnection CellManagerConnection = new SqlConnection(SQL_Connection);
                    SqlCommand SQL_CellManager = new SqlCommand("Select * from [dbo].[Admin] where EmployeeName=@EmployeeName and EmployeePassword=@EmployeePassword", CellManagerConnection);
                    SQL_CellManager.Parameters.AddWithValue("@EmployeeName", EmployeeName_TextBox.Text);
                    SQL_CellManager.Parameters.AddWithValue("@EmployeePassword", DMPID_TextBox.Text);
                    CellManagerConnection.Open();
                    SqlDataAdapter adapt = new SqlDataAdapter(SQL_CellManager);
                    DataSet ds = new DataSet();
                    adapt.Fill(ds);
                    LoginGridView.DataSource = ds.Tables[0];
                    CellManagerConnection.Close();
                    int count = ds.Tables[0].Rows.Count;
                    if (count == 1)
                    {
                        Cell_Manager CellManager = new Cell_Manager();
                        CellManager.User_TextBox.Text = EmployeeName_TextBox.Text;
                        CellManager.Clock_TextBox.Text = Clock_TextBox.Text;
                        CellManager.DMPID_TextBox.Text = DMPID_TextBox.Text;
                        ListBox.Items.Clear();
                        CellManager.Show();
                        CellManager.Focus();
                        OpenNewForm();
                    }
                }
                catch (Exception)
                {
                    ListBox.Items.Add("Access to Cell Control Denied");
                }
            }
        }

        private void ReportViewLogin()
        {
            if (EmployeeName_TextBox.Text == "" && DMPID_TextBox.Text == "")
            {
                ListBox.Items.Add("Please Enter Your Name and DMP ID to Login");
            }
            else if (EmployeeName_TextBox.Text == "")
            {
                ListBox.Items.Add("Enter Employee Name to Login");
            }
            else if (DMPID_TextBox.Text == "")
            {
                ListBox.Items.Add("Enter DMP ID to Login");
            }
            else if (EmployeeName_TextBox.Text != "" && DMPID_TextBox.Text != "")
            {
                try
                {
                    SqlConnection ReportLogin = new SqlConnection(SQL_Connection);
                    SqlCommand Login_Command = new SqlCommand("Select * from [dbo].[Admin] where EmployeeName=@EmployeeName and DMPID=@DMPID", ReportLogin);
                    Login_Command.Parameters.AddWithValue("@EmployeeName", EmployeeName_TextBox.Text);
                    Login_Command.Parameters.AddWithValue("@DMPID", DMPID_TextBox.Text);
                    ReportLogin.Open();
                    SqlDataAdapter adapt = new SqlDataAdapter(Login_Command);
                    DataSet ds = new DataSet();
                    adapt.Fill(ds);
                    LoginGridView.DataSource = ds.Tables[0];
                    ReportLogin.Close();
                    int count = ds.Tables[0].Rows.Count;
                    if (count == 1)
                    {
                        Report_View ReportView = new Report_View();
                        ReportView.User_TextBox.Text = EmployeeName_TextBox.Text;
                        ReportView.Clock_TextBox.Text = Clock_TextBox.Text;
                        ReportView.DMPID_TextBox.Text = DMPID_TextBox.Text;
                        ListBox.Items.Clear();
                        ReportView.Show();
                        ReportView.Focus();
                        OpenNewForm();
                    }
                }
                catch (Exception)
                {
                    ListBox.Items.Add("Access to Report View Denied.");
                }
            }
        }

        private void OpenNewForm()
        {
            Current.WindowState = FormWindowState.Minimized;
            Current.Enabled = false;
            Current.ShowInTaskbar = false;
        }

        /********************************************************************************************************************
        * 
        * Methods Region End
        * 
        ********************************************************************************************************************/
        #endregion

        private void Clock_Tick(object sender, EventArgs e)
        {
            string AMPM = "";
            string Date = DateTime.Today.ToShortDateString();
            string Time = "";

            ClockHour   = DateTime.Now.Hour;
            ClockMinute = DateTime.Now.Minute;
            ClockSecond = DateTime.Now.Second;

            if (ClockHour > 12)
            {
                ClockHour = ClockHour - 12;
                Time += ClockHour;
                AMPM = "PM";
            }
            else if (ClockHour == 12)
            {
                Time += ClockHour;
                AMPM = "PM";
            }
            else if (ClockHour >= 10 && ClockHour <= 11)
            {
                Time += ClockHour;
                AMPM = "AM";
            }
            else if (ClockHour == 0)
            {
                Time += ClockHour + 12;
                AMPM = "AM";
            }
            else if (ClockHour < 10)
            {
                Time += "0" + ClockHour;
                AMPM = "AM";
            }
            Time += ":";

            if (ClockMinute < 10)
            {
                Time += "0" + ClockMinute;
            }
            else
            {
                Time += ClockMinute;
            }
            Time += ":";
            if (ClockSecond < 10)
            {
                Time += "0" + ClockSecond;
            }
            else
            {
                Time += ClockSecond;
            }
            Time += " " + AMPM;
            Time += "   " + Date;
            
            Clock_TextBox.Text = Time;
        }

        private void OPC_Button_Click(object sender, EventArgs e)
        {
            //OPC_Data opc = new OPC_Data();
            //opc.Show();
        }

        private void EmployeeName_ComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            EmployeeName_TextBox.Clear();
            DMPID_TextBox.Clear();
            if (EmployeeName_ComboBox.Text == "Ryan Garland" || EmployeeName_ComboBox.Text == "Dale Worline" || EmployeeName_ComboBox.Text == "Sean Pleiman" || EmployeeName_ComboBox.Text == "Nick Gill")
            {
                Login_GroupBox.Visible = true;
                Test_GroupBox.Visible = true;
                EmployeeName_TextBox.Text = EmployeeName_ComboBox.Text;
            }
            else if (EmployeeName_ComboBox.Text != "Ryan Garland" || EmployeeName_ComboBox.Text != "Dale Worline" || EmployeeName_ComboBox.Text != "Sean Pleiman" || EmployeeName_ComboBox.Text != "Nick Gill")
            {
                Login_GroupBox.Visible = false;
                Test_GroupBox.Visible = false;
                EmployeeName_TextBox.Text = EmployeeName_ComboBox.Text;
            }
            
        }

        private void DMP_Spot_Weld_Login_FormClosing(object sender, FormClosingEventArgs e)
        {
            Application.Exit();
        }

        private void TestForm_Button_Click(object sender, EventArgs e)
        {
            TestForm tf = new TestForm();
            tf.Show();
        }

        private void SpotWeldID()
        {
            string SpotWeldComputerID = System.Environment.MachineName;
            if (SpotWeldComputerID == "OHN7047NL")
            {
                EmployeeName_ComboBox.Text = "Ryan Garland";
                EmployeeName_TextBox.Text = "Ryan Garland";
                DMPID_TextBox.Text = "10078";
                OperatorLogin_Button.Focus();
                Login_GroupBox.Visible = true;
                Test_GroupBox.Visible = true;
            }
            if (SpotWeldComputerID == "OHN7125NL")
            {
                EmployeeName_ComboBox.Text = "Dale Worline";
                EmployeeName_TextBox.Text = "Dale Worline";
                DMPID_TextBox.Text = "4418";
                OperatorLogin_Button.Focus();
                Login_GroupBox.Visible = true;
                Test_GroupBox.Visible = true;
            }
            if (SpotWeldComputerID == "OHNSean")
            {
                EmployeeName_ComboBox.Text = "Sean Pleiman";
                EmployeeName_TextBox.Text = "Sean Pleiman";
                DMPID_TextBox.Text = "8064";
                OperatorLogin_Button.Focus();
                Login_GroupBox.Visible = true;
                Test_GroupBox.Visible = true;
            }
        }

        private void DMPID_TextBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter && DMPID_TextBox.Focus() == true)
            {
                OperatorLogin_Button_Click(null, null);
            }
        }

        private void ScanOut_Button_Click(object sender, EventArgs e)
        {
            User_Program_Scan_Out ScanOut = new User_Program_Scan_Out();
            ScanOut.EmployeeNumber_TextBox.Text = "10078";
            ScanOut.JobNumber_TextBox.Text = "J000909782-0";
            ScanOut.TotalCountQtuQtyComp_TextBox.Text = "20";
            ScanOut.Show();
        }

        private void TextFile_Button_Click(object sender, EventArgs e)
        {
            Text_Test tt = new Text_Test();
            tt.Show();
        }
        


        /********************************************************************************************************************
        * 
        * Methods End
        * 
        ********************************************************************************************************************/

    }
}
