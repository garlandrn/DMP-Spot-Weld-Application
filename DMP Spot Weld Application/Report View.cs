using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using PdfSharp;
using PdfSharp.Drawing;
using PdfSharp.Pdf;
using System.Diagnostics;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;
using System.Windows.Forms.DataVisualization.Charting;

/*
 * 
 * Program: DMP Spot Weld Application
 * Form: Report View
 * Created By: Ryan Garland
 * Last Updated on 9/4/18
 * 
 */

namespace DMP_Spot_Weld_Application
{
    public partial class Report_View : Form
    {
        // Background Workers
        BackgroundWorker PartImage;
        BackgroundWorker CreateExcel;
        public Report_View()
        {
            InitializeComponent();
            
            // Find The Part Image on the Network and Display it in the Form
            PartImage = new BackgroundWorker();
            PartImage.DoWork += new DoWorkEventHandler(FindItemImage);
            PartImage.RunWorkerCompleted += new RunWorkerCompletedEventHandler(PartImage_RunWorkerCompleted);

            // Create a Report in an Excel File
            CreateExcel = new BackgroundWorker();
            CreateExcel.DoWork += new DoWorkEventHandler(CreateExcelFile);
            CreateExcel.RunWorkerCompleted += new RunWorkerCompletedEventHandler(CreateExcel_RunWorkerComplete);
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

        // SQL Data
        string SQL_Source = @"Data Source=OHN7009,49172;Initial Catalog=Spot_Weld_Data;Integrated Security=True;Connect Timeout=15;";
        private string LoginForm = "Report Viewer";
        private string LoginTime = "";

        //string SQL_Source = @"Data Source=OHN7009,49172;Initial Catalog=Brake_Press_Data;Integrated Security=True;Connect Timeout=15;";

        // Clock_Tick();
        private static int ClockHour;
        private static int ClockMinute;
        private static int ClockSecond;

        // Search_Button_Click()
        private static string SearchCommand;
        private static string ReportItemID;
        private static string ReportSpotWelder;
        private static string ReportEmployee;
        private static string ReportDate;

        // GetTime()
        private static string PlannedOperationTime;
        private static string ActualOperationTime;
        private static double PlannedOperationConvert;
        private static double ActualOperationConvert;

        // PDF FileLocation
        private static string PDFFileLocation;

        // Excel File Creation
        private static Excel._Workbook ReportWB;
        private static Excel.Application ReportApp;
        private static Excel._Worksheet ReportWS;
        private static Excel.Range ReportRange;
        private static string ExcelFileLocation;
        private DataSet ReportDataSet;

        // FindItemImage();
        private static string ItemImagePath = "";
        private static string ItemID = "";
        private static string[] ItemIDSplit = ItemID.Split('-');
        private static double ItemID_Three;
        private static double ItemID_Five;


        private static string[] itemID;
        private static int[] partCount;


        private string[] SpotWeldID = { "153R", "155R" };
        private string[] CATSpotWelders = { "123R", "1088" };
        private string[] JohnDeereSpotWelders = { "108R", "150R" };
        private string[] NavistarSpotWelders = { "104R", "121R", "154R" };
        private string[] PaccarSpotWelders = { "153R", "155R" };


        /********************************************************************************************************************
        * 
        * ReportViewer Start
        * 
        ********************************************************************************************************************/

        private void Report_View_Load(object sender, EventArgs e)
        {
            // Write a login Report to SQL Server
            SqlConnection ReportLogin = new SqlConnection(SQL_Source);
            SqlCommand Login = new SqlCommand();
            Login.CommandType = System.Data.CommandType.Text;
            Login.CommandText = "INSERT INTO [dbo].[LoginData] (EmployeeName,DMPID,LoginDateTime,LoginForm) VALUES (@EmployeeName,@DMPID,@LoginDateTime,@LoginForm)";
            Login.Connection = ReportLogin;
            Login.Parameters.AddWithValue("@LoginDateTime", Clock_TextBox.Text);
            Login.Parameters.AddWithValue("@EmployeeName", User_TextBox.Text);
            Login.Parameters.AddWithValue("@DMPID", DMPID_TextBox.Text);
            Login.Parameters.AddWithValue("@LoginForm", LoginForm.ToString());
            ReportLogin.Open();
            Login.ExecuteNonQuery();
            ReportLogin.Close();

            Clock.Enabled = true;
            LoginTime = Clock_TextBox.Text;

            // Connect to SQL Server and Get Employee information
            SqlConnection connection = new SqlConnection(SQL_Source);
            string EmployeeData = "SELECT * FROM [dbo].[Employee]";
            SqlDataAdapter dataAdapter = new SqlDataAdapter(EmployeeData, connection);
            SqlCommandBuilder commandBuilder = new SqlCommandBuilder(dataAdapter);
            DataSet Data = new DataSet();
            dataAdapter.Fill(Data);
            LoginGridView.DataSource = Data.Tables[0];
            for(int i = 0 ;i < Data.Tables[0].Rows.Count; i++)
            {
                // Add Employee ID to ComboBox
                DMPID_ComboBox.Items.Add(Data.Tables[0].Rows[i][0]);
            }

            // Add Spot Welders to the Spotweld ComboBox for searching
            Spotweld_ComboBox.Items.AddRange(CATSpotWelders);
            Spotweld_ComboBox.Items.AddRange(JohnDeereSpotWelders);
            Spotweld_ComboBox.Items.AddRange(NavistarSpotWelders);
            Spotweld_ComboBox.Items.AddRange(PaccarSpotWelders);            
        }

        /********************************************************************************************************************
        * 
        * Buttons Region Start 
        * -- Total Buttons: 5
        * 
        * - Clear Button Click
        * - Create Button Click
        * - Search Button Click
        * - Excel Button Click
        * - LogOff Button Click
        * 
        ********************************************************************************************************************/
        #region

            // Call the Clear Method
        private void Clear_Button_Click(object sender, EventArgs e)
        {
            Clear();
        }

        // Creates a PDF File From the Search Commands Specified
        // if no Search Commands have been Specified we remind the user to do so
        private void Create_Button_Click(object sender, EventArgs e)
        {
            if (SearchCommand == null)
            {
                MessageBox.Show("Please Create a DataTable Before Creating a PDF");
            }
            else
            {
                PDFFileCreate();
            }
        }

        // Call the CommandCreator Method
        private void Search_Button_Click(object sender, EventArgs e)
        {
            CommandCreator();
        }

        // Create an Excel File from the Data that the Search Commands Specified
        private void Excel_Button_Click(object sender, EventArgs e)
        {
            if (SearchCommand == null)
            {
                MessageBox.Show("Please Create a DataTable Before Creating an Excel File");
            }
            else
            {  string ReportName = DateTime.Now.ToShortDateString() + "_" + DateTime.Now.Millisecond;
                ReportName = ReportName.Replace("/", "_");
                ReportName = ReportName.Replace(":", "_");
                SaveFileDialog saveFile = new SaveFileDialog();
                saveFile.Filter = "Excel Files (*.xlsx)|*.xlsx|All files (*.*)|*.*";
                string ReportPDFName = "Spotweld_Report_" + ReportName + ".xls";
                saveFile.FileName = ReportPDFName;
                if (saveFile.ShowDialog() == DialogResult.OK)
                {
                    ExcelFileLocation = saveFile.FileName;
                    //CreateExcelFile();
                    CreateExcel.RunWorkerAsync();
                }
            }
        }

        private void LogOff_Button_Click(object sender, EventArgs e)
        {
            EmployeeLogOff();
            DMP_Spot_Weld_Login.Current.Focus();
            DMP_Spot_Weld_Login.Current.Enabled = true;
            DMP_Spot_Weld_Login.Current.WindowState = FormWindowState.Maximized;
            DMP_Spot_Weld_Login.Current.ShowInTaskbar = true;
            this.Close();
        }

        /********************************************************************************************************************
        * 
        * Buttons Region End
        * 
        ********************************************************************************************************************/
        #endregion

        /*********************************************************************************************************************
        * 
        * ComboBox Region Start && DateTimePicker
        * -- Total: 2
        * 
        * - Spotweld ComboBox SelectedIndexChanged
        * - DMPID ComboBox SelectedIndexChanged
        * - DateStartPicker DropDown
        * - DateEndPicker DropDown
        * 
        *********************************************************************************************************************/
        #region

        private void Spotweld_ComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            Spotweld_TextBox.Text = Spotweld_ComboBox.Text;
        }

        private void DMPID_ComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            {
                if (DMPID_ComboBox.Text == "Puck Garland")
                {
                    SearchDMPID_TextBox.Text = "1111";
                }
                else if (DMPID_ComboBox.Text == "Edward Bennett")
                {
                    SearchDMPID_TextBox.Text = "1246";
                }
                else if (DMPID_ComboBox.Text == "Jeff Brandt")
                {
                    SearchDMPID_TextBox.Text = "1344";
                }
                else if (DMPID_ComboBox.Text == "Tim Lothamer")
                {
                    SearchDMPID_TextBox.Text = "1395";
                }
                else if (DMPID_ComboBox.Text == "Lee Decker")
                {
                    SearchDMPID_TextBox.Text = "1733";
                }
                else if (DMPID_ComboBox.Text == "Felix Garland")
                {
                    SearchDMPID_TextBox.Text = "2222";
                }
                else if (DMPID_ComboBox.Text == "Zac Lawson")
                {
                    SearchDMPID_TextBox.Text = "2604";
                }
                else if (DMPID_ComboBox.Text == "Joe Nichols")
                {
                    SearchDMPID_TextBox.Text = "2955";
                }
                else if (DMPID_ComboBox.Text == "Corey Shirk")
                {
                    SearchDMPID_TextBox.Text = "2970";
                }
                else if (DMPID_ComboBox.Text == "Rich Pepper")
                {
                    SearchDMPID_TextBox.Text = "3040";
                }
                else if (DMPID_ComboBox.Text == "Stephen Reeder")
                {
                    SearchDMPID_TextBox.Text = "3267";
                }
                else if (DMPID_ComboBox.Text == "Felicity Garland")
                {
                    SearchDMPID_TextBox.Text = "3333";
                }
                else if (DMPID_ComboBox.Text == "Paxton Garland")
                {
                    SearchDMPID_TextBox.Text = "4444";
                }
                else if (DMPID_ComboBox.Text == "Ryan Garland")
                {
                    SearchDMPID_TextBox.Text = "10078";
                }
                SearchDMPID_TextBox.Text = DMPID_ComboBox.Text;
            }
        }

        private void DateStartPicker_DropDown(object sender, EventArgs e)
        {
            DateStartPicker.Checked = true;
            DateStartPicker.Size = new System.Drawing.Size(357, 30);
            DateStartPicker.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
        }

        private void DateEndPicker_DropDown(object sender, EventArgs e)
        {
            DateEndPicker.Checked = true;
            DateEndPicker.Size = new System.Drawing.Size(357, 30);
            DateEndPicker.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
        }

        /*********************************************************************************************************************
        * 
        * ComboBox Region End
        * 
        *********************************************************************************************************************/
        #endregion

        /*********************************************************************************************************************
        * 
        * Methods Region Start 
        * -- Total: 
        * 
        * - Clear
        * - CommandCreator
        * - ConvertPlannedTime
        * - ConvertActualTime
        * - CreateReport
        * - PDFFileCreate
        * - CreateExcelFile
        * - CreateExcel RunWorker Complete
        * - EmployeeLogOff
        * - FindItemImage
        * - PartImage DoWork
        * - PartImage RunWorkerCompleted
        * 
        **********************************************************************************************************************/
        #region

        // Clear Method
        private void Clear()
        {
            Spotweld_TextBox.Clear();
            Spotweld_ComboBox.Text = "";
            DateStartPicker.Checked = false;
            DateStartPicker.Size = new System.Drawing.Size(323, 30);
            DateStartPicker.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            DateStartPicker.ResetText();
            DateEndPicker.Size = new System.Drawing.Size(323, 30);
            DateEndPicker.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            DateEndPicker.ResetText();
            DMPID_ComboBox.Text = "";
            ItemID_TextBox.Clear();
            OperationID_TextBox.Clear();
            ReportGridView.DataSource = null;
            SearchDMPID_TextBox.Clear();
        }

        // the SQL Command is determined by which search variables are entered or selected
        // we join the OperationOEE and ItemOperationData tables 
        // as long as data is entered we then call CreateReport method. This takes the SearchCommand and loads the returned data to the datatable
        private void CommandCreator()
        {
            // #1
            // Item ID: Yes
            // Brake Press: Yes
            // DMP: Yes
            // Date: Yes
            if (ItemID_TextBox.Text != "" && Spotweld_TextBox.Text != "" && SearchDMPID_TextBox.Text != "" && DateStartPicker.Checked == true)
            {
                SearchCommand = "SELECT D.ItemID as ItemID, D.ReferenceNumber as ReferenceNumber, O.RunDateTime as RunDateTime, O.OperationTime as OperationTime, O.PlannedTime as PlannedTime, D.PartsManufactured as PartsManufactured, D.PartsPerMinute as PartsPerMinute, O.Efficiency as Efficiency, O.OEE as OEE, D.SpotWelder as SpotWelder, D.EmployeeName as EmployeeName, D.DMPID as DMPID, D.OperationID as OperationID FROM [dbo].[OperationOEE] as O INNER JOIN [dbo].[ItemOperationData] as D ON O.OperationID = D.OperationID WHERE O.ItemID='" + ItemID_TextBox.Text + "' AND O.SpotWelder='" + Spotweld_TextBox.Text + "' AND O.DMPID='" + SearchDMPID_TextBox.Text + "' AND O.RunDateTime BETWEEN '" + DateStartPicker.Value.Date.ToShortDateString() + "' AND '" + DateEndPicker.Value.Date.ToShortDateString() + "'";
                //SearchCommand = "SELECT D.ItemID as ItemID, D.OperationID as OperationID, O.RunDateTime as RunDateTime, O.OperationTime as OperationTime, O.PlannedTime as PlannedTime, O.Efficiency as Efficiency, O.OEE as OEE, D.SpotWelder as SpotWelder, D.PartsManufactured as PartsManufactured, D.PartsPerMinute as PartsPerMinute, D.EmployeeName as EmployeeName, D.DMPID as DMPID, D.ReferenceNumber as ReferenceNumber FROM [dbo].[OperationOEE] as O INNER JOIN [dbo].[ItemOperationData] as D ON O.OperationID = D.OperationID WHERE O.ItemID='" + ItemID_TextBox.Text + "' AND O.SpotWelder='" + Spotweld_TextBox.Text + "' AND O.DMPID='" + SearchDMPID_TextBox.Text + "' AND O.RunDateTime BETWEEN '" + DateStartPicker.Value.Date.ToShortDateString() + "' AND '" + DateEndPicker.Value.Date.ToShortDateString() + "'";
                ReportItemID = "Item ID: " + ItemID_TextBox.Text;
                ReportSpotWelder = "Spotweld: " + Spotweld_TextBox.Text;
                ReportEmployee = "DMP ID: " + SearchDMPID_TextBox.Text;
                ReportDate = "Date: " + DateStartPicker.Value.ToShortDateString() + " - " + DateEndPicker.Value.ToShortDateString();
                CreateReport();
            }
            // #2
            // Item ID: Yes
            // Brake Press: Yes
            // DMP: Yes
            // Date: No
            else if (ItemID_TextBox.Text != "" && Spotweld_TextBox.Text != "" && SearchDMPID_TextBox.Text != "" && DateStartPicker.Checked == false)
            {
                SearchCommand = "SELECT D.ItemID as ItemID, D.ReferenceNumber as ReferenceNumber, O.RunDateTime as RunDateTime, O.OperationTime as OperationTime, O.PlannedTime as PlannedTime, D.PartsManufactured as PartsManufactured, D.PartsPerMinute as PartsPerMinute, O.Efficiency as Efficiency, O.OEE as OEE, D.SpotWelder as SpotWelder, D.EmployeeName as EmployeeName, D.DMPID as DMPID, D.OperationID as OperationID FROM [dbo].[OperationOEE] as O INNER JOIN [dbo].[ItemOperationData] as D ON O.OperationID = D.OperationID WHERE O.ItemID='" + ItemID_TextBox.Text + "' AND O.SpotWelder='" + Spotweld_TextBox.Text + "' AND O.DMPID='" + SearchDMPID_TextBox.Text + "'";
                ReportItemID = "Item ID: " + ItemID_TextBox.Text;
                ReportSpotWelder = "Spotweld: " + Spotweld_TextBox.Text;
                ReportEmployee = "DMP ID: " + SearchDMPID_TextBox.Text;
                ReportDate = "Date: All";
                CreateReport();
            }
            // #3
            // Item ID: Yes
            // Brake Press: Yes
            // DMP: No
            // Date: Yes
            else if (ItemID_TextBox.Text != "" && Spotweld_TextBox.Text != "" && SearchDMPID_TextBox.Text == "" && DateStartPicker.Checked == true)
            {
                SearchCommand = "SELECT D.ItemID as ItemID, D.ReferenceNumber as ReferenceNumber, O.RunDateTime as RunDateTime, O.OperationTime as OperationTime, O.PlannedTime as PlannedTime, D.PartsManufactured as PartsManufactured, D.PartsPerMinute as PartsPerMinute, O.Efficiency as Efficiency, O.OEE as OEE, D.SpotWelder as SpotWelder, D.EmployeeName as EmployeeName, D.DMPID as DMPID, D.OperationID as OperationID FROM [dbo].[OperationOEE] as O INNER JOIN [dbo].[ItemOperationData] as D ON O.OperationID = D.OperationID WHERE O.SpotWelder='" + Spotweld_TextBox.Text + "' AND O.ItemID='" + ItemID_TextBox.Text + "' AND O.RunDateTime BETWEEN '" + DateStartPicker.Value.Date.ToShortDateString() + "' AND '" + DateEndPicker.Value.Date.ToShortDateString() + "'";
                ReportItemID = "Item ID: " + ItemID_TextBox.Text;
                ReportSpotWelder = "Spotweld: " + Spotweld_TextBox.Text;
                ReportEmployee = "DMP ID: All";
                ReportDate = "Date: " + DateStartPicker.Value.ToShortDateString() + " - " + DateEndPicker.Value.ToShortDateString();
                CreateReport();
            }
            // #4
            // Item ID: Yes
            // Brake Press: Yes
            // DMP: No
            // Date: No
            else if (ItemID_TextBox.Text != "" && Spotweld_TextBox.Text != "" && SearchDMPID_TextBox.Text == "" && DateStartPicker.Checked == false)
            {
                SearchCommand = "SELECT D.ItemID as ItemID, D.ReferenceNumber as ReferenceNumber, O.RunDateTime as RunDateTime, O.OperationTime as OperationTime, O.PlannedTime as PlannedTime, D.PartsManufactured as PartsManufactured, D.PartsPerMinute as PartsPerMinute, O.Efficiency as Efficiency, O.OEE as OEE, D.SpotWelder as SpotWelder, D.EmployeeName as EmployeeName, D.DMPID as DMPID, D.OperationID as OperationID FROM [dbo].[OperationOEE] as O INNER JOIN [dbo].[ItemOperationData] as D ON O.OperationID = D.OperationID WHERE O.ItemID='" + ItemID_TextBox.Text + "' AND O.SpotWelder='" + Spotweld_TextBox.Text + "'";
                ReportItemID = "Item ID: " + ItemID_TextBox.Text;
                ReportSpotWelder = "Spotweld: " + Spotweld_TextBox.Text;
                ReportEmployee = "DMP ID: All";
                ReportDate = "Date: All";
                CreateReport();
            }
            // #5
            // Item ID: Yes
            // Brake Press: No
            // DMP: Yes
            // Date: Yes
            else if (ItemID_TextBox.Text != "" && Spotweld_TextBox.Text == "" && SearchDMPID_TextBox.Text != "" && DateStartPicker.Checked == true)
            {
                SearchCommand = "SELECT D.ItemID as ItemID, D.ReferenceNumber as ReferenceNumber, O.RunDateTime as RunDateTime, O.OperationTime as OperationTime, O.PlannedTime as PlannedTime, D.PartsManufactured as PartsManufactured, D.PartsPerMinute as PartsPerMinute, O.Efficiency as Efficiency, O.OEE as OEE, D.SpotWelder as SpotWelder, D.EmployeeName as EmployeeName, D.DMPID as DMPID, D.OperationID as OperationID FROM [dbo].[OperationOEE] as O INNER JOIN [dbo].[ItemOperationData] as D ON O.OperationID = D.OperationID WHERE O.ItemID='" + ItemID_TextBox.Text + "' AND O.DMPID='" + SearchDMPID_TextBox.Text + "'AND O.RunDateTime BETWEEN '" + DateStartPicker.Value.Date.ToShortDateString() + "' AND '" + DateEndPicker.Value.Date.ToShortDateString() + "'";
                ReportItemID = "Item ID: " + ItemID_TextBox.Text;
                ReportSpotWelder = "Spotweld: All";
                ReportEmployee = "DMP ID: " + SearchDMPID_TextBox.Text;
                ReportDate = "Date: " + DateStartPicker.Value.ToShortDateString() + " - " + DateEndPicker.Value.ToShortDateString();
                CreateReport();
            }
            // #6
            // Item ID: Yes
            // Brake Press: No
            // DMP: Yes
            // Date: No
            else if (ItemID_TextBox.Text != "" && Spotweld_TextBox.Text == "" && SearchDMPID_TextBox.Text != "" && DateStartPicker.Checked == false)
            {
                SearchCommand = "SELECT D.ItemID as ItemID, D.ReferenceNumber as ReferenceNumber, O.RunDateTime as RunDateTime, O.OperationTime as OperationTime, O.PlannedTime as PlannedTime, D.PartsManufactured as PartsManufactured, D.PartsPerMinute as PartsPerMinute, O.Efficiency as Efficiency, O.OEE as OEE, D.SpotWelder as SpotWelder, D.EmployeeName as EmployeeName, D.DMPID as DMPID, D.OperationID as OperationID FROM [dbo].[OperationOEE] as O INNER JOIN [dbo].[ItemOperationData] as D ON O.OperationID = D.OperationID WHERE O.ItemID='" + ItemID_TextBox.Text + "' AND O.DMPID='" + SearchDMPID_TextBox.Text + "'";
                ReportItemID = "Item ID: " + ItemID_TextBox.Text;
                ReportSpotWelder = "Spotweld: All";
                ReportEmployee = "DMP ID: " + SearchDMPID_TextBox.Text;
                ReportDate = "Date: All";
                CreateReport();
            }
            // #7
            // Item ID: No
            // Brake Press: Yes
            // DMP: Yes
            // Date: Yes
            else if (ItemID_TextBox.Text == "" && Spotweld_TextBox.Text != "" && SearchDMPID_TextBox.Text != "" && DateStartPicker.Checked == true)
            {
                SearchCommand = "SELECT D.ItemID as ItemID, D.ReferenceNumber as ReferenceNumber, O.RunDateTime as RunDateTime, O.OperationTime as OperationTime, O.PlannedTime as PlannedTime, D.PartsManufactured as PartsManufactured, D.PartsPerMinute as PartsPerMinute, O.Efficiency as Efficiency, O.OEE as OEE, D.SpotWelder as SpotWelder, D.EmployeeName as EmployeeName, D.DMPID as DMPID, D.OperationID as OperationID FROM [dbo].[OperationOEE] as O INNER JOIN [dbo].[ItemOperationData] as D ON O.OperationID = D.OperationID WHERE O.SpotWelder='" + Spotweld_TextBox.Text + "' AND O.DMPID='" + SearchDMPID_TextBox.Text + "' AND O.RunDateTime BETWEEN '" + DateStartPicker.Value.Date.ToShortDateString() + "' AND '" + DateEndPicker.Value.Date.ToShortDateString() + "'";
                ReportItemID = "Item ID: All";
                ReportSpotWelder = "Spotweld: " + Spotweld_TextBox.Text;
                ReportEmployee = "DMP ID: " + SearchDMPID_TextBox.Text;
                ReportDate = "Date: " + DateStartPicker.Value.ToShortDateString() + " - " + DateEndPicker.Value.ToShortDateString();
                CreateReport();
            }
            // #8
            // Item ID: No
            // Brake Press: Yes
            // DMP: Yes
            // Date: No
            else if (ItemID_TextBox.Text == "" && Spotweld_TextBox.Text != "" && SearchDMPID_TextBox.Text != "" && DateStartPicker.Checked == false)
            {
                SearchCommand = "SELECT D.ItemID as ItemID, D.ReferenceNumber as ReferenceNumber, O.RunDateTime as RunDateTime, O.OperationTime as OperationTime, O.PlannedTime as PlannedTime, D.PartsManufactured as PartsManufactured, D.PartsPerMinute as PartsPerMinute, O.Efficiency as Efficiency, O.OEE as OEE, D.SpotWelder as SpotWelder, D.EmployeeName as EmployeeName, D.DMPID as DMPID, D.OperationID as OperationID FROM [dbo].[OperationOEE] as O INNER JOIN [dbo].[ItemOperationData] as D ON O.OperationID = D.OperationID WHERE O.DMPID='" + SearchDMPID_TextBox.Text + "' AND O.SpotWelder='" + Spotweld_TextBox.Text + "'";
                ReportItemID = "Item ID: All";
                ReportSpotWelder = "Spotweld: " + Spotweld_TextBox.Text;
                ReportEmployee = "DMP ID: " + SearchDMPID_TextBox.Text;
                ReportDate = "Date: All";
                CreateReport();
            }
            // #9
            // Item ID: No
            // Brake Press: Yes
            // DMP: No
            // Date: Yes
            if (ItemID_TextBox.Text == "" && Spotweld_TextBox.Text != "" && SearchDMPID_TextBox.Text == "" && DateStartPicker.Checked == true)
            {
                SearchCommand = "SELECT D.ItemID as ItemID, D.ReferenceNumber as ReferenceNumber, O.RunDateTime as RunDateTime, O.OperationTime as OperationTime, O.PlannedTime as PlannedTime, D.PartsManufactured as PartsManufactured, D.PartsPerMinute as PartsPerMinute, O.Efficiency as Efficiency, O.OEE as OEE, D.SpotWelder as SpotWelder, D.EmployeeName as EmployeeName, D.DMPID as DMPID, D.OperationID as OperationID FROM [dbo].[OperationOEE] as O INNER JOIN [dbo].[ItemOperationData] as D ON O.OperationID = D.OperationID WHERE O.SpotWelder='" + Spotweld_TextBox.Text + "' AND O.RunDateTime BETWEEN '" + DateStartPicker.Value.Date.ToShortDateString() + "' AND '" + DateEndPicker.Value.Date.ToShortDateString() + "'";
                ReportItemID = "Item ID: All";
                ReportSpotWelder = "Spotwelder: " + Spotweld_TextBox.Text;
                ReportEmployee = "DMP ID: All";
                ReportDate = "Date: " + DateStartPicker.Value.ToShortDateString() + " - " + DateEndPicker.Value.ToShortDateString();
                CreateReport();
            }
            // #10
            // Item ID: No
            // Brake Press: Yes
            // DMP: No
            // Date: No
            else if (ItemID_TextBox.Text == "" && Spotweld_TextBox.Text != "" && SearchDMPID_TextBox.Text == "" && DateStartPicker.Checked == false)
            {
                SearchCommand = "SELECT D.ItemID as ItemID, D.ReferenceNumber as ReferenceNumber, O.RunDateTime as RunDateTime, O.OperationTime as OperationTime, O.PlannedTime as PlannedTime, D.PartsManufactured as PartsManufactured, D.PartsPerMinute as PartsPerMinute, O.Efficiency as Efficiency, O.OEE as OEE, D.SpotWelder as SpotWelder, D.EmployeeName as EmployeeName, D.DMPID as DMPID, D.OperationID as OperationID FROM [dbo].[OperationOEE] as O INNER JOIN [dbo].[ItemOperationData] as D ON O.OperationID = D.OperationID WHERE O.SpotWelder='" + Spotweld_TextBox.Text + "'";
                ReportItemID = "Item ID: All";
                ReportSpotWelder = "Spotwelder: " + Spotweld_TextBox.Text;
                ReportEmployee = "DMP ID: All";
                ReportDate = "Date: All";
                CreateReport();
            }
            // #11
            // Item ID: Yes
            // Brake Press: No
            // DMP: No
            // Date: Yes
            else if (ItemID_TextBox.Text != "" && Spotweld_TextBox.Text == "" && SearchDMPID_TextBox.Text == "" && DateStartPicker.Checked == true)
            {
                SearchCommand = "SELECT D.ItemID as ItemID, D.ReferenceNumber as ReferenceNumber, O.RunDateTime as RunDateTime, O.OperationTime as OperationTime, O.PlannedTime as PlannedTime, D.PartsManufactured as PartsManufactured, D.PartsPerMinute as PartsPerMinute, O.Efficiency as Efficiency, O.OEE as OEE, D.SpotWelder as SpotWelder, D.EmployeeName as EmployeeName, D.DMPID as DMPID, D.OperationID as OperationID FROM [dbo].[OperationOEE] as O INNER JOIN [dbo].[ItemOperationData] as D ON O.OperationID = D.OperationID WHERE O.ItemID='" + ItemID_TextBox.Text + "' AND O.RunDateTime BETWEEN '" + DateStartPicker.Value.Date.ToShortDateString() + "' AND '" + DateEndPicker.Value.Date.ToShortDateString() + "'";
                ReportItemID = "Item ID: " + ItemID_TextBox.Text;
                ReportSpotWelder = "Spotweld: All";
                ReportEmployee = "DMP ID: All";
                ReportDate = "Date: " + DateStartPicker.Value.ToShortDateString() + " - " + DateEndPicker.Value.ToShortDateString();
                CreateReport();
            }
            // #12
            // Item ID: Yes
            // Brake Press: No
            // DMP: No
            // Date: No
            else if (ItemID_TextBox.Text != "" && Spotweld_TextBox.Text == "" && SearchDMPID_TextBox.Text == "" && DateStartPicker.Checked == false)
            {
                SearchCommand = "SELECT D.ItemID as ItemID, D.ReferenceNumber as ReferenceNumber, O.RunDateTime as RunDateTime, O.OperationTime as OperationTime, O.PlannedTime as PlannedTime, D.PartsManufactured as PartsManufactured, D.PartsPerMinute as PartsPerMinute, O.Efficiency as Efficiency, O.OEE as OEE, D.SpotWelder as SpotWelder, D.EmployeeName as EmployeeName, D.DMPID as DMPID, D.OperationID as OperationID FROM [dbo].[ItemOperationData] as D INNER JOIN [dbo].[OperationOEE] as O ON D.OperationID = O.OperationID WHERE D.ItemID='" + ItemID_TextBox.Text + "'";
                //SearchCommand = "SELECT D.ItemID as ItemID, D.OperationID as OperationID, O.RunDateTime as RunDateTime, O.OperationTime as OperationTime, O.PlannedTime as PlannedTime, O.Efficiency as Efficiency, O.OEE as OEE, D.SpotWelder as SpotWelder, D.PartsManufactured as PartsManufactured, D.PartsPerMinute as PartsPerMinute, D.EmployeeName as EmployeeName, D.DMPID as DMPID, D.ReferenceNumber as ReferenceNumber FROM [dbo].[OperationOEE] as O  INNER JOIN [dbo].[ItemOperationData] as D ON O.OperationID = D.OperationID WHERE O.ItemID='" + ItemID_TextBox.Text + "'";
                ReportItemID = "Item ID: " + ItemID_TextBox.Text;
                ReportSpotWelder = "Spotweld: All";
                ReportEmployee = "DMP ID: All";
                ReportDate = "Date: All";
                CreateReport();
            }           
            // #12 b
            // Item ID: No
            // Reference Number: Yes
            // Brake Press: No
            // DMP: No
            // Date: No
            else if (ItemID_TextBox.Text == "" && ReferenceNumber_TextBox.Text != "" && Spotweld_TextBox.Text == "" && SearchDMPID_TextBox.Text == "" && DateStartPicker.Checked == false)
            {
                SearchCommand = "SELECT D.ItemID as ItemID, D.ReferenceNumber as ReferenceNumber, O.RunDateTime as RunDateTime, O.OperationTime as OperationTime, O.PlannedTime as PlannedTime, D.PartsManufactured as PartsManufactured, D.PartsPerMinute as PartsPerMinute, O.Efficiency as Efficiency, O.OEE as OEE, D.SpotWelder as SpotWelder, D.EmployeeName as EmployeeName, D.DMPID as DMPID, D.OperationID as OperationID FROM [dbo].[ItemOperationData] as D INNER JOIN [dbo].[OperationOEE] as O ON D.OperationID = O.OperationID WHERE D.ReferenceNumber='" + ReferenceNumber_TextBox.Text + "'";
                //SearchCommand = "SELECT D.ItemID as ItemID, D.OperationID as OperationID, O.RunDateTime as RunDateTime, O.OperationTime as OperationTime, O.PlannedTime as PlannedTime, O.Efficiency as Efficiency, O.OEE as OEE, D.SpotWelder as SpotWelder, D.PartsManufactured as PartsManufactured, D.PartsPerMinute as PartsPerMinute, D.EmployeeName as EmployeeName, D.DMPID as DMPID, D.ReferenceNumber as ReferenceNumber FROM [dbo].[OperationOEE] as O  INNER JOIN [dbo].[ItemOperationData] as D ON O.OperationID = D.OperationID WHERE O.ItemID='" + ItemID_TextBox.Text + "'";
                ReportItemID = "Item ID: " + ItemID_TextBox.Text;
                ReportSpotWelder = "Spotweld: All";
                ReportEmployee = "DMP ID: All";
                ReportDate = "Date: All";
                CreateReport();
            }
            // #13
            // Item ID: No
            // Brake Press: No
            // DMP: Yes
            // Date: Yes
            else if (ItemID_TextBox.Text == "" && Spotweld_TextBox.Text == "" && SearchDMPID_TextBox.Text != "" && DateStartPicker.Checked == true)
            {
                SearchCommand = "SELECT D.ItemID as ItemID, D.ReferenceNumber as ReferenceNumber, O.RunDateTime as RunDateTime, O.OperationTime as OperationTime, O.PlannedTime as PlannedTime, D.PartsManufactured as PartsManufactured, D.PartsPerMinute as PartsPerMinute, O.Efficiency as Efficiency, O.OEE as OEE, D.SpotWelder as SpotWelder, D.EmployeeName as EmployeeName, D.DMPID as DMPID, D.OperationID as OperationID FROM [dbo].[OperationOEE] as O INNER JOIN [dbo].[ItemOperationData] as D ON O.OperationID = D.OperationID WHERE O.DMPID='" + SearchDMPID_TextBox.Text + "' AND O.RunDateTime BETWEEN '" + DateStartPicker.Value.Date.ToShortDateString() + "' AND '" + DateEndPicker.Value.Date.ToShortDateString() + "'";
                ReportItemID = "Item ID: All";
                ReportSpotWelder = "Spotweld: All";
                ReportEmployee = "DMP ID: " + SearchDMPID_TextBox.Text;
                ReportDate = "Date: " + DateStartPicker.Value.ToShortDateString() + " - " + DateEndPicker.Value.ToShortDateString();
                CreateReport();
            }
            // #14
            // Item ID: No
            // Brake Press: No
            // DMP: Yes
            // Date: No
            else if (ItemID_TextBox.Text == "" && Spotweld_TextBox.Text == "" && SearchDMPID_TextBox.Text != "" && DateStartPicker.Checked == false)
            {
                SearchCommand = "SELECT D.ItemID as ItemID, D.ReferenceNumber as ReferenceNumber, O.RunDateTime as RunDateTime, O.OperationTime as OperationTime, O.PlannedTime as PlannedTime, D.PartsManufactured as PartsManufactured, D.PartsPerMinute as PartsPerMinute, O.Efficiency as Efficiency, O.OEE as OEE, D.SpotWelder as SpotWelder, D.EmployeeName as EmployeeName, D.DMPID as DMPID, D.OperationID as OperationID FROM [dbo].[OperationOEE] as O  INNER JOIN [dbo].[ItemOperationData] as D ON O.OperationID = D.OperationID WHERE O.DMPID='" + SearchDMPID_TextBox.Text + "'";
                ReportItemID = "Item ID: All";
                ReportSpotWelder = "Spotweld: All";
                ReportEmployee = "DMP ID: " + SearchDMPID_TextBox.Text;
                ReportDate = "Date: All";
                CreateReport();
            }
            // #15
            // Item ID: No
            // Brake Press: No
            // DMP: No
            // Date: Yes
            else if (ItemID_TextBox.Text == "" && Spotweld_TextBox.Text == "" && SearchDMPID_TextBox.Text == "" && DateStartPicker.Checked == true)
            {
                SearchCommand = "SELECT D.ItemID as ItemID, D.ReferenceNumber as ReferenceNumber, O.RunDateTime as RunDateTime, O.OperationTime as OperationTime, O.PlannedTime as PlannedTime, D.PartsManufactured as PartsManufactured, D.PartsPerMinute as PartsPerMinute, O.Efficiency as Efficiency, O.OEE as OEE, D.SpotWelder as SpotWelder, D.EmployeeName as EmployeeName, D.DMPID as DMPID, D.OperationID as OperationID FROM [dbo].[OperationOEE] as O INNER JOIN [dbo].[ItemOperationData] as D ON O.OperationID = D.OperationID WHERE O.RunDateTime BETWEEN '" + DateStartPicker.Value.Date.ToShortDateString() + "' AND '" + DateEndPicker.Value.Date.ToShortDateString() + "'";
                ReportItemID = "Item ID: All";
                ReportSpotWelder = "Spotweld: All";
                ReportEmployee = "DMP ID: All";
                ReportDate = "Date: " + DateStartPicker.Value.ToShortDateString() + " - " + DateEndPicker.Value.ToShortDateString();
                CreateReport();
            }
            // Last
            // Excecutes When Item ID, Operation ID, DMP ID, and DateStartPicker are all Empty
            else if (ItemID_TextBox.Text == "" && OperationID_TextBox.Text == "" && Spotweld_TextBox.Text == "" && SearchDMPID_TextBox.Text == "" && DateStartPicker.Checked == false)
            {
                MessageBox.Show("Please Select a Date, DMP ID, or Item Number to Search Data");
            }
        }

        // Convert the planned time for job run
        private void ConvertPlannedTime()
        {
            PlannedOperationConvert = double.Parse(PlannedOperationTime);
            double PlannedHours = 0;
            double PlannedMinutes = 0;
            double TotalPlannedMinutes = (PlannedOperationConvert * 60);
            string PlannedTime = "";

            if (TotalPlannedMinutes < 60)
            {
                PlannedMinutes = TotalPlannedMinutes;
                PlannedMinutes = Math.Round(PlannedMinutes, MidpointRounding.ToEven);
                if (PlannedMinutes == 1)
                {
                    PlannedTime = PlannedMinutes + " Minute".ToString();
                }
                else if (PlannedMinutes == 0)
                {
                    PlannedTime = "N/A";
                }
                else
                {
                    PlannedTime = PlannedMinutes + " Minutes".ToString();
                }
            }
            else if (120 > TotalPlannedMinutes && TotalPlannedMinutes >= 60)
            {
                TotalPlannedMinutes = TotalPlannedMinutes - 60;
                PlannedHours = 1;
                PlannedMinutes = TotalPlannedMinutes;
                PlannedMinutes = Math.Round(PlannedMinutes, MidpointRounding.ToEven);
                PlannedTime = PlannedHours + " Hour " + PlannedMinutes + " Minutes".ToString();

            }
            else if (180 > TotalPlannedMinutes && TotalPlannedMinutes >= 120)
            {
                TotalPlannedMinutes = TotalPlannedMinutes - 120;
                PlannedHours = 2;
                PlannedMinutes = TotalPlannedMinutes;
                PlannedMinutes = Math.Round(PlannedMinutes, MidpointRounding.ToEven);
                PlannedTime = PlannedHours + " Hours " + PlannedMinutes + " Minutes".ToString();
            }
            else if (240 > TotalPlannedMinutes && TotalPlannedMinutes >= 180)
            {
                TotalPlannedMinutes = TotalPlannedMinutes - 180;
                PlannedHours = 3;
                PlannedMinutes = TotalPlannedMinutes;
                PlannedMinutes = Math.Round(PlannedMinutes, MidpointRounding.ToEven);
                PlannedTime = PlannedHours + " Hours " + PlannedMinutes + " Minutes".ToString();
            }
            else if (300 > TotalPlannedMinutes && TotalPlannedMinutes >= 240)
            {
                TotalPlannedMinutes = TotalPlannedMinutes - 240;
                PlannedHours = 4;
                PlannedMinutes = TotalPlannedMinutes;
                PlannedMinutes = Math.Round(PlannedMinutes, MidpointRounding.ToEven);
                PlannedTime = PlannedHours + " Hours " + PlannedMinutes + " Minutes".ToString();
            }
            else if (360 > TotalPlannedMinutes && TotalPlannedMinutes >= 300)
            {
                TotalPlannedMinutes = TotalPlannedMinutes - 300;
                PlannedHours = 5;
                PlannedMinutes = TotalPlannedMinutes;
                PlannedMinutes = Math.Round(PlannedMinutes, MidpointRounding.ToEven);
                PlannedTime = PlannedHours + " Hours " + PlannedMinutes + " Minutes".ToString();
            }
            else if (420 > TotalPlannedMinutes && TotalPlannedMinutes >= 360)
            {
                TotalPlannedMinutes = TotalPlannedMinutes - 360;
                PlannedHours = 6;
                PlannedMinutes = TotalPlannedMinutes;
                PlannedMinutes = Math.Round(PlannedMinutes, MidpointRounding.ToEven);
                PlannedTime = PlannedHours + " Hours " + PlannedMinutes + " Minutes".ToString();
            }
            else if (480 > TotalPlannedMinutes && TotalPlannedMinutes >= 420)
            {
                TotalPlannedMinutes = TotalPlannedMinutes - 420;
                PlannedHours = 7;
                PlannedMinutes = TotalPlannedMinutes;
                PlannedMinutes = Math.Round(PlannedMinutes, MidpointRounding.ToEven);
                PlannedTime = PlannedHours + " Hours " + PlannedMinutes + " Minutes".ToString();
            }
            else if (540 > TotalPlannedMinutes && TotalPlannedMinutes >= 480)
            {
                TotalPlannedMinutes = TotalPlannedMinutes - 480;
                PlannedHours = 8;
                PlannedMinutes = TotalPlannedMinutes;
                PlannedMinutes = Math.Round(PlannedMinutes, MidpointRounding.ToEven);
                PlannedTime = PlannedHours + " Hours " + PlannedMinutes + " Minutes".ToString();
            }
            else if (600 > TotalPlannedMinutes && TotalPlannedMinutes >= 540)
            {
                TotalPlannedMinutes = TotalPlannedMinutes - 540;
                PlannedHours = 9;
                PlannedMinutes = TotalPlannedMinutes;
                PlannedMinutes = Math.Round(PlannedMinutes, MidpointRounding.ToEven);
                PlannedTime = PlannedHours + " Hours " + PlannedMinutes + " Minutes".ToString();
            }
            else if (660 > TotalPlannedMinutes && TotalPlannedMinutes >= 600)
            {
                TotalPlannedMinutes = TotalPlannedMinutes - 600;
                PlannedHours = 10;
                PlannedMinutes = TotalPlannedMinutes;
                PlannedMinutes = Math.Round(PlannedMinutes, MidpointRounding.ToEven);
                PlannedTime = PlannedHours + " Hours " + PlannedMinutes + " Minutes".ToString();
            }
            else if (720 > TotalPlannedMinutes && TotalPlannedMinutes >= 660) // 11 Hours
            {
                TotalPlannedMinutes = TotalPlannedMinutes - 660;
                PlannedHours = 11;
                PlannedMinutes = TotalPlannedMinutes;
                PlannedMinutes = Math.Round(PlannedMinutes, MidpointRounding.ToEven);
                PlannedTime = PlannedHours + " Hours " + PlannedMinutes + " Minutes".ToString();
            }
            else if (780 > TotalPlannedMinutes && TotalPlannedMinutes >= 720) // 12 Hours
            {
                TotalPlannedMinutes = TotalPlannedMinutes - 720;
                PlannedHours = 12;
                PlannedMinutes = TotalPlannedMinutes;
                PlannedMinutes = Math.Round(PlannedMinutes, MidpointRounding.ToEven);
                PlannedTime = PlannedHours + " Hours " + PlannedMinutes + " Minutes".ToString();
            }
            else if (840 > TotalPlannedMinutes && TotalPlannedMinutes >= 780) // 13 Hours
            {
                TotalPlannedMinutes = TotalPlannedMinutes - 780;
                PlannedHours = 13;
                PlannedMinutes = TotalPlannedMinutes;
                PlannedMinutes = Math.Round(PlannedMinutes, MidpointRounding.ToEven);
                PlannedTime = PlannedHours + " Hours " + PlannedMinutes + " Minutes".ToString();
            }
            PlannedTimeResults_TextBox.Text = PlannedTime;
        }

        // Convert the actual time for job run
        private void ConvertActualTime()
        {
            ActualOperationConvert = double.Parse(ActualOperationTime);
            double ActualHours = 0;
            double ActualMinutes = 0;
            double TotalActualMinutes = (ActualOperationConvert * 60);
            string ActualTime = "";

            if (TotalActualMinutes < 60) // Less Than 1 Hour
            {
                ActualMinutes = TotalActualMinutes;
                ActualMinutes = Math.Round(ActualMinutes, MidpointRounding.AwayFromZero);
                if (ActualMinutes == 1)
                {
                    ActualTime = ActualMinutes + " Minute".ToString();
                }
                else if (ActualMinutes == 0)
                {
                    ActualTime = "N/A";
                }
                else
                {
                    ActualTime = ActualMinutes + " Minutes".ToString();
                }
            }
            else if (120 > TotalActualMinutes && TotalActualMinutes >= 60) // 1 Hour
            {
                TotalActualMinutes = TotalActualMinutes - 60;
                ActualHours = 1;
                ActualMinutes = TotalActualMinutes;
                ActualMinutes = Math.Round(ActualMinutes, MidpointRounding.AwayFromZero);
                ActualTime = ActualHours + " Hour " + ActualMinutes + " Minutes".ToString();

            }
            else if (180 > TotalActualMinutes && TotalActualMinutes >= 120) // 2 Hours
            {
                TotalActualMinutes = TotalActualMinutes - 120;
                ActualHours = 2;
                ActualMinutes = TotalActualMinutes;
                ActualMinutes = Math.Round(ActualMinutes, MidpointRounding.AwayFromZero);
                ActualTime = ActualHours + " Hours " + ActualMinutes + " Minutes".ToString();
            }
            else if (240 > TotalActualMinutes && TotalActualMinutes >= 180) // 3 Hours
            {
                TotalActualMinutes = TotalActualMinutes - 180;
                ActualHours = 3;
                ActualMinutes = TotalActualMinutes;
                ActualMinutes = Math.Round(ActualMinutes, MidpointRounding.AwayFromZero);
                ActualTime = ActualHours + " Hours " + ActualMinutes + " Minutes".ToString();
            }
            else if (300 > TotalActualMinutes && TotalActualMinutes >= 240) // 4 Hours
            {
                TotalActualMinutes = TotalActualMinutes - 240;
                ActualHours = 4;
                ActualMinutes = TotalActualMinutes;
                ActualMinutes = Math.Round(ActualMinutes, MidpointRounding.AwayFromZero);
                ActualTime = ActualHours + " Hours " + ActualMinutes + " Minutes".ToString();
            }
            else if (360 > TotalActualMinutes && TotalActualMinutes >= 300) // 5 Hours
            {
                TotalActualMinutes = TotalActualMinutes - 300;
                ActualHours = 5;
                ActualMinutes = TotalActualMinutes;
                ActualMinutes = Math.Round(ActualMinutes, MidpointRounding.AwayFromZero);
                ActualTime = ActualHours + " Hours " + ActualMinutes + " Minutes".ToString();
            }
            else if (420 > TotalActualMinutes && TotalActualMinutes >= 360) // 6 Hours
            {
                TotalActualMinutes = TotalActualMinutes - 360;
                ActualHours = 6;
                ActualMinutes = TotalActualMinutes;
                ActualMinutes = Math.Round(ActualMinutes, MidpointRounding.AwayFromZero);
                ActualTime = ActualHours + " Hours " + ActualMinutes + " Minutes".ToString();
            }
            else if (480 > TotalActualMinutes && TotalActualMinutes >= 420) // 7 Hours
            {
                TotalActualMinutes = TotalActualMinutes - 420;
                ActualHours = 7;
                ActualMinutes = TotalActualMinutes;
                ActualMinutes = Math.Round(ActualMinutes, MidpointRounding.AwayFromZero);
                ActualTime = ActualHours + " Hours " + ActualMinutes + " Minutes".ToString();
            }
            else if (540 > TotalActualMinutes && TotalActualMinutes >= 480) // 8 Hours
            {
                TotalActualMinutes = TotalActualMinutes - 480;
                ActualHours = 8;
                ActualMinutes = TotalActualMinutes;
                ActualMinutes = Math.Round(ActualMinutes, MidpointRounding.AwayFromZero);
                ActualTime = ActualHours + " Hours " + ActualMinutes + " Minutes".ToString();
            }
            else if (600 > TotalActualMinutes && TotalActualMinutes >= 540) // 9 Hours
            {
                TotalActualMinutes = TotalActualMinutes - 540;
                ActualHours = 9;
                ActualMinutes = TotalActualMinutes;
                ActualMinutes = Math.Round(ActualMinutes, MidpointRounding.AwayFromZero);
                ActualTime = ActualHours + " Hours " + ActualMinutes + " Minutes".ToString();
            }
            else if (660 > TotalActualMinutes && TotalActualMinutes >= 600) // 10 Hours
            {
                TotalActualMinutes = TotalActualMinutes - 600;
                ActualHours = 10;
                ActualMinutes = TotalActualMinutes;
                ActualMinutes = Math.Round(ActualMinutes, MidpointRounding.AwayFromZero);
                ActualTime = ActualHours + " Hours " + ActualMinutes + " Minutes".ToString();
            }
            else if (720 > TotalActualMinutes && TotalActualMinutes >= 660) // 11 Hours
            {
                TotalActualMinutes = TotalActualMinutes - 660;
                ActualHours = 11;
                ActualMinutes = TotalActualMinutes;
                ActualMinutes = Math.Round(ActualMinutes, MidpointRounding.AwayFromZero);
                ActualTime = ActualHours + " Hours " + ActualMinutes + " Minutes".ToString();
            }
            else if (780 > TotalActualMinutes && TotalActualMinutes >= 720) // 12 Hours
            {
                TotalActualMinutes = TotalActualMinutes - 720;
                ActualHours = 12;
                ActualMinutes = TotalActualMinutes;
                ActualMinutes = Math.Round(ActualMinutes, MidpointRounding.AwayFromZero);
                ActualTime = ActualHours + " Hours " + ActualMinutes + " Minutes".ToString();
            }
            else if (840 > TotalActualMinutes && TotalActualMinutes >= 780) // 13 Hours
            {
                TotalActualMinutes = TotalActualMinutes - 780;
                ActualHours = 13;
                ActualMinutes = TotalActualMinutes;
                ActualMinutes = Math.Round(ActualMinutes, MidpointRounding.AwayFromZero);
                ActualTime = ActualHours + " Hours " + ActualMinutes + " Minutes".ToString();
            }
            OperationTimeResults_TextBox.Text = ActualTime;
        }

        // Return the Data that the SearchCommand finds
        private void CreateReport()
        {
            try
            {
                SqlConnection ReportCreate = new SqlConnection(SQL_Source);
                string ReportCommand = SearchCommand;
                SqlDataAdapter ReportAdapter = new SqlDataAdapter(ReportCommand, ReportCreate);
                SqlCommandBuilder CommandBuilder = new SqlCommandBuilder(ReportAdapter);
                ReportDataSet = new DataSet();
                ReportAdapter.Fill(ReportDataSet);
                ReportGridView.DataSource = ReportDataSet.Tables[0];
            }
            catch (SqlException ExceptionValue)
            {
                int ErrorNumber = ExceptionValue.Number;
                if (ErrorNumber.Equals(2627))
                {
                    MessageBox.Show("This DMP ID Belongs to Another User");
                }
                else if (ErrorNumber.Equals(245))
                {
                    MessageBox.Show("DMP ID Can Only Contain Numbers");
                }
                else
                {
                    MessageBox.Show("Unable to Add User. Please Try Again." + "\n" + "Error Code: " + ErrorNumber.ToString());
                }
            }
        }

        private void PDFFileCreate()
        {
            // New PDF Document
            PdfDocument BrakePressReport = new PdfDocument();
            PdfPage ReportPage = BrakePressReport.AddPage();
            ReportPage.Size = PdfSharp.PageSize.Letter;
            ReportPage.Orientation = PdfSharp.PageOrientation.Landscape;
            ReportPage.Rotate = 0;
            XGraphics ReportGraph = XGraphics.FromPdfPage(ReportPage);

            // Fonts
            XFont ReportDataHeader = new XFont("Verdana", 12, XFontStyle.Bold);
            XFont ColumnHeader = new XFont("Verdana", 8, XFontStyle.Bold | XFontStyle.Underline);
            XFont RowFont = new XFont("Verdana", 6, XFontStyle.Regular);
            XFont PageFooterFont = new XFont("Verdana", 5, XFontStyle.Regular);

            int PointY = 0;
            int CurrentRow = 0;

            // PDF Report Name

            string ReportName = DateTime.Now.ToShortDateString() + "_" + DateTime.Now.Millisecond;
            ReportName = ReportName.Replace("/", "_");
            ReportName = ReportName.Replace(":", "_");
            string ReportFooter = " | Report Created On: " + DateTime.Now.ToShortDateString() + " | Created By: " + User_TextBox.Text;

            // Set the Image and General Report Information
            ReportGraph.DrawImage(XImage.FromFile(@"\\OHN66FS01\BPprogs\Brake Press Vision\Applications\DMPLogo700.jpg"), 35, 5);
            ReportGraph.DrawString(ReportItemID, ReportDataHeader, XBrushes.Black, new XRect(400, 15, ReportPage.Width.Point, ReportPage.Height.Point), XStringFormats.TopLeft);
            ReportGraph.DrawString(ReportSpotWelder, ReportDataHeader, XBrushes.Black, new XRect(400, 33, ReportPage.Width.Point, ReportPage.Height.Point), XStringFormats.TopLeft);
            ReportGraph.DrawString(ReportEmployee, ReportDataHeader, XBrushes.Black, new XRect(400, 51, ReportPage.Width.Point, ReportPage.Height.Point), XStringFormats.TopLeft);
            ReportGraph.DrawString(ReportDate, ReportDataHeader, XBrushes.Black, new XRect(400, 69, ReportPage.Width.Point, ReportPage.Height.Point), XStringFormats.TopLeft);
            PointY = PointY + 100;
            
            // Set the Headers for each page
            ReportGraph.DrawString("Item ID", ColumnHeader, XBrushes.Black, new XRect(10, PointY, ReportPage.Width.Point, ReportPage.Height.Point), XStringFormats.TopLeft);
            ReportGraph.DrawString("Operation", ColumnHeader, XBrushes.Black, new XRect(60, PointY, ReportPage.Width.Point, ReportPage.Height.Point), XStringFormats.TopLeft);
            ReportGraph.DrawString("Date", ColumnHeader, XBrushes.Black, new XRect(121, PointY, ReportPage.Width.Point, ReportPage.Height.Point), XStringFormats.TopLeft);
            ReportGraph.DrawString("Run Time", ColumnHeader, XBrushes.Black, new XRect(165, PointY, ReportPage.Width.Point, ReportPage.Height.Point), XStringFormats.TopLeft);
            ReportGraph.DrawString("Planned", ColumnHeader, XBrushes.Black, new XRect(224, PointY, ReportPage.Width.Point, ReportPage.Height.Point), XStringFormats.TopLeft);
            ReportGraph.DrawString("Efficiency", ColumnHeader, XBrushes.Black, new XRect(269, PointY, ReportPage.Width.Point, ReportPage.Height.Point), XStringFormats.TopLeft);
            ReportGraph.DrawString("Parts Manufactured", ColumnHeader, XBrushes.Black, new XRect(330, PointY, ReportPage.Width.Point, ReportPage.Height.Point), XStringFormats.TopLeft);
            ReportGraph.DrawString("PPM", ColumnHeader, XBrushes.Black, new XRect(450, PointY, ReportPage.Width.Point, ReportPage.Height.Point), XStringFormats.TopLeft);
            ReportGraph.DrawString("Spot Welder", ColumnHeader, XBrushes.Black, new XRect(500, PointY, ReportPage.Width.Point, ReportPage.Height.Point), XStringFormats.TopLeft);
            ReportGraph.DrawString("Employee", ColumnHeader, XBrushes.Black, new XRect(600, PointY, ReportPage.Width.Point, ReportPage.Height.Point), XStringFormats.TopLeft);
            ReportGraph.DrawString("DMP ID", ColumnHeader, XBrushes.Black, new XRect(700, PointY, ReportPage.Width.Point, ReportPage.Height.Point), XStringFormats.TopLeft);


            // Report Footer 
            string PageNumber = "Page: " + BrakePressReport.PageCount + ReportFooter;
            ReportGraph.DrawString(PageNumber, PageFooterFont, XBrushes.Black, new XRect(5, 585, ReportPage.Width.Point, ReportPage.Height.Point), XStringFormats.TopLeft);

            PointY = PointY + 25;
            try
            {
                SqlConnection CreatePDF = new SqlConnection(SQL_Source);
                string PDFCommand = SearchCommand;
                SqlDataAdapter PDFAdapter = new SqlDataAdapter(PDFCommand, CreatePDF);
                DataSet PDFData = new DataSet();
                PDFAdapter.Fill(PDFData);

                for (int i = 0; i <= PDFData.Tables[0].Rows.Count - 1; i++)
                {
                    string ItemIDResults = PDFData.Tables[0].Rows[i].ItemArray[0].ToString();
                    string OperationIDResults = PDFData.Tables[0].Rows[i].ItemArray[1].ToString();
                    string RunDateResults = PDFData.Tables[0].Rows[i].ItemArray[2].ToString();
                    string OperationTimeResults = PDFData.Tables[0].Rows[i].ItemArray[3].ToString();
                    string PlannedTimeResults = PDFData.Tables[0].Rows[i].ItemArray[4].ToString();
                    string PartsManufacturedResults = PDFData.Tables[0].Rows[i].ItemArray[5].ToString();
                    string PPMResults = PDFData.Tables[0].Rows[i].ItemArray[6].ToString();
                    string EfficiencyResults = PDFData.Tables[0].Rows[i].ItemArray[7].ToString();
                    string OEEResults = PDFData.Tables[0].Rows[i].ItemArray[8].ToString();
                    string SpotweldResults = PDFData.Tables[0].Rows[i].ItemArray[9].ToString();
                    string EmployeeResults = PDFData.Tables[0].Rows[i].ItemArray[10].ToString();
                    string DMPIDResults = PDFData.Tables[0].Rows[i].ItemArray[11].ToString();
                    string JobReferenceNumber = PDFData.Tables[0].Rows[i].ItemArray[12].ToString();

                    RunDateResults = RunDateResults.Replace("12:00:00 AM", "");
                    
                    // Report Row Data 
                    ReportGraph.DrawString(ItemIDResults, RowFont, XBrushes.Black, new XRect(12, PointY, ReportPage.Width.Point, ReportPage.Height.Point), XStringFormats.TopLeft);
                    ReportGraph.DrawString(OperationIDResults, RowFont, XBrushes.Black, new XRect(62, PointY, ReportPage.Width.Point, ReportPage.Height.Point), XStringFormats.TopLeft);
                    ReportGraph.DrawString(RunDateResults, RowFont, XBrushes.Black, new XRect(117, PointY, ReportPage.Width.Point, ReportPage.Height.Point), XStringFormats.TopLeft);
                    ReportGraph.DrawString(OperationTimeResults, RowFont, XBrushes.Black, new XRect(174, PointY, ReportPage.Width.Point, ReportPage.Height.Point), XStringFormats.TopLeft);
                    ReportGraph.DrawString(PlannedTimeResults, RowFont, XBrushes.Black, new XRect(229, PointY, ReportPage.Width.Point, ReportPage.Height.Point), XStringFormats.TopLeft);
                    ReportGraph.DrawString(EfficiencyResults, RowFont, XBrushes.Black, new XRect(282, PointY, ReportPage.Width.Point, ReportPage.Height.Point), XStringFormats.TopLeft);
                    ReportGraph.DrawString(OEEResults, RowFont, XBrushes.Black, new XRect(335, PointY, ReportPage.Width.Point, ReportPage.Height.Point), XStringFormats.TopLeft);

                    ReportGraph.DrawString(PartsManufacturedResults, RowFont, XBrushes.Black, new XRect(370, PointY, ReportPage.Width.Point, ReportPage.Height.Point), XStringFormats.TopLeft);
                    ReportGraph.DrawString(PPMResults, RowFont, XBrushes.Black, new XRect(452, PointY, ReportPage.Width.Point, ReportPage.Height.Point), XStringFormats.TopLeft);

                    ReportGraph.DrawString(SpotweldResults, RowFont, XBrushes.Black, new XRect(520, PointY, ReportPage.Width.Point, ReportPage.Height.Point), XStringFormats.TopLeft);
                    ReportGraph.DrawString(EmployeeResults, RowFont, XBrushes.Black, new XRect(600, PointY, ReportPage.Width.Point, ReportPage.Height.Point), XStringFormats.TopLeft);
                    ReportGraph.DrawString(DMPIDResults, RowFont, XBrushes.Black, new XRect(705, PointY, ReportPage.Width.Point, ReportPage.Height.Point), XStringFormats.TopLeft);
                    PointY = PointY + 20;
                    CurrentRow = CurrentRow + 1;

                    // Report Creates Adds Another Page If Data is Larger than 22 Rows
                    // Only 22 Entries on Page One Due To Report Header
                    if (CurrentRow == 22 && BrakePressReport.PageCount == 1)
                    {
                        PointY = 0;
                        ReportPage = BrakePressReport.AddPage();
                        ReportGraph = XGraphics.FromPdfPage(ReportPage);
                        ReportPage.Size = PdfSharp.PageSize.Letter;
                        ReportPage.Orientation = PdfSharp.PageOrientation.Landscape;
                        ReportPage.Rotate = 0;
                        PointY = PointY + 50;
                                                
                        ReportGraph.DrawString("Item ID", ColumnHeader, XBrushes.Black, new XRect(10, PointY, ReportPage.Width.Point, ReportPage.Height.Point), XStringFormats.TopLeft);
                        ReportGraph.DrawString("Operation", ColumnHeader, XBrushes.Black, new XRect(60, PointY, ReportPage.Width.Point, ReportPage.Height.Point), XStringFormats.TopLeft);
                        ReportGraph.DrawString("Date", ColumnHeader, XBrushes.Black, new XRect(121, PointY, ReportPage.Width.Point, ReportPage.Height.Point), XStringFormats.TopLeft);
                        ReportGraph.DrawString("Run Time", ColumnHeader, XBrushes.Black, new XRect(165, PointY, ReportPage.Width.Point, ReportPage.Height.Point), XStringFormats.TopLeft);
                        ReportGraph.DrawString("Planned", ColumnHeader, XBrushes.Black, new XRect(224, PointY, ReportPage.Width.Point, ReportPage.Height.Point), XStringFormats.TopLeft);
                        ReportGraph.DrawString("Efficiency", ColumnHeader, XBrushes.Black, new XRect(269, PointY, ReportPage.Width.Point, ReportPage.Height.Point), XStringFormats.TopLeft);
                        ReportGraph.DrawString("Parts Manufactured", ColumnHeader, XBrushes.Black, new XRect(330, PointY, ReportPage.Width.Point, ReportPage.Height.Point), XStringFormats.TopLeft);
                        ReportGraph.DrawString("PPM", ColumnHeader, XBrushes.Black, new XRect(450, PointY, ReportPage.Width.Point, ReportPage.Height.Point), XStringFormats.TopLeft);
                        ReportGraph.DrawString("Spot Welder", ColumnHeader, XBrushes.Black, new XRect(500, PointY, ReportPage.Width.Point, ReportPage.Height.Point), XStringFormats.TopLeft);
                        ReportGraph.DrawString("Employee", ColumnHeader, XBrushes.Black, new XRect(600, PointY, ReportPage.Width.Point, ReportPage.Height.Point), XStringFormats.TopLeft);
                        ReportGraph.DrawString("DMP ID", ColumnHeader, XBrushes.Black, new XRect(700, PointY, ReportPage.Width.Point, ReportPage.Height.Point), XStringFormats.TopLeft);
                        PageNumber = "Page: " + BrakePressReport.PageCount + ReportFooter;
                        ReportGraph.DrawString(PageNumber, PageFooterFont, XBrushes.Black, new XRect(5, 585, ReportPage.Width.Point, ReportPage.Height.Point), XStringFormats.TopLeft); PointY = PointY + 25;
                        CurrentRow = 0;

                    }
                    else if (CurrentRow == 25 && BrakePressReport.PageCount >= 2)
                    {
                        PointY = 0;
                        ReportPage = BrakePressReport.AddPage();
                        ReportGraph = XGraphics.FromPdfPage(ReportPage);
                        ReportPage.Size = PdfSharp.PageSize.Letter;
                        ReportPage.Orientation = PdfSharp.PageOrientation.Landscape;
                        ReportPage.Rotate = 0;
                        PointY = PointY + 50;
                        
                        ReportGraph.DrawString("Item ID", ColumnHeader, XBrushes.Black, new XRect(10, PointY, ReportPage.Width.Point, ReportPage.Height.Point), XStringFormats.TopLeft);
                        ReportGraph.DrawString("Operation", ColumnHeader, XBrushes.Black, new XRect(60, PointY, ReportPage.Width.Point, ReportPage.Height.Point), XStringFormats.TopLeft);
                        ReportGraph.DrawString("Date", ColumnHeader, XBrushes.Black, new XRect(121, PointY, ReportPage.Width.Point, ReportPage.Height.Point), XStringFormats.TopLeft);
                        ReportGraph.DrawString("Run Time", ColumnHeader, XBrushes.Black, new XRect(165, PointY, ReportPage.Width.Point, ReportPage.Height.Point), XStringFormats.TopLeft);
                        ReportGraph.DrawString("Planned", ColumnHeader, XBrushes.Black, new XRect(224, PointY, ReportPage.Width.Point, ReportPage.Height.Point), XStringFormats.TopLeft);
                        ReportGraph.DrawString("Efficiency", ColumnHeader, XBrushes.Black, new XRect(269, PointY, ReportPage.Width.Point, ReportPage.Height.Point), XStringFormats.TopLeft);
                        ReportGraph.DrawString("Parts Manufactured", ColumnHeader, XBrushes.Black, new XRect(330, PointY, ReportPage.Width.Point, ReportPage.Height.Point), XStringFormats.TopLeft);
                        ReportGraph.DrawString("PPM", ColumnHeader, XBrushes.Black, new XRect(450, PointY, ReportPage.Width.Point, ReportPage.Height.Point), XStringFormats.TopLeft);
                        ReportGraph.DrawString("Spot Welder", ColumnHeader, XBrushes.Black, new XRect(500, PointY, ReportPage.Width.Point, ReportPage.Height.Point), XStringFormats.TopLeft);
                        ReportGraph.DrawString("Employee", ColumnHeader, XBrushes.Black, new XRect(600, PointY, ReportPage.Width.Point, ReportPage.Height.Point), XStringFormats.TopLeft);
                        ReportGraph.DrawString("DMP ID", ColumnHeader, XBrushes.Black, new XRect(700, PointY, ReportPage.Width.Point, ReportPage.Height.Point), XStringFormats.TopLeft);
                        PageNumber = "Page: " + BrakePressReport.PageCount + ReportFooter;
                        ReportGraph.DrawString(PageNumber, PageFooterFont, XBrushes.Black, new XRect(5, 585, ReportPage.Width.Point, ReportPage.Height.Point), XStringFormats.TopLeft); PointY = PointY + 25;
                        CurrentRow = 0;
                    }
                }

                // Chart Data
                ReportPage = BrakePressReport.AddPage();
                ReportGraph = XGraphics.FromPdfPage(ReportPage);
                ReportPage.Size = PdfSharp.PageSize.Letter;
                ReportPage.Orientation = PdfSharp.PageOrientation.Landscape;
                ReportPage.Rotate = 0;
                chart1.Size = new System.Drawing.Size(1050, 575);

                chart1.ChartAreas[0].AxisX.Interval = 1;

                foreach (DataGridViewRow r in ReportGridView.Rows)
                {
                    chart1.Series["Series1"].Points.AddXY(r.Cells[0].Value.ToString(), r.Cells[5].Value.ToString());
                }
                chart1.SaveImage(@"C:\Users\rgarland\Desktop\"+ ReportName+".jpg", ChartImageFormat.Jpeg);
                XImage img = XImage.FromFile(@"C:\Users\rgarland\Desktop\" + ReportName + ".jpg");
                ReportGraph.DrawImage(img, 0, 0);


                //ReportGraph.DrawImage(chart1, 0);

                string ReportPDFName = "Spot_Weld_Report_" + ReportName + ".pdf";
                SaveFileDialog saveFile = new SaveFileDialog();
                saveFile.Filter = "PDF Files (*.pdf)|*.pdf|All files (*.*)|*.*";
                saveFile.FileName = ReportPDFName;
                if (saveFile.ShowDialog() == DialogResult.OK)
                {
                    PDFFileLocation = saveFile.FileName;
                    BrakePressReport.Save(PDFFileLocation);
                    View_PDF PDFViewer = new View_PDF();
                    PDFViewer.AcroPDF.src = PDFFileLocation;
                    PDFViewer.AcroPDF.BringToFront();
                    PDFViewer.Show();
                    PDFViewer.BringToFront();
                    PDFViewer.AcroPDF.setZoom(95);
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void CreateExcelFile(object sender, EventArgs e)
        {
            int picture = 0;
            string ReportName = DateTime.Now.ToShortDateString() + "_" + DateTime.Now.Millisecond;
            ReportName = ReportName.Replace("/", "_");
            ReportName = ReportName.Replace(":", "_");
            //SaveFileDialog saveFile = new SaveFileDialog();
            //saveFile.Filter = "Excel Files (*.xls)|*.xls|All files (*.*)|*.*";
            //string ReportPDFName = "Spotweld_Report_" + ReportName + ".xls";
            //saveFile.FileName = ReportPDFName;
            //if (saveFile.ShowDialog() == DialogResult.OK)
            //{
                // Excel Initialize
                ReportApp = new Excel.Application();
                ReportApp.Visible = false;
                ReportWB = (Excel._Workbook)(ReportApp.Workbooks.Add(""));
                ReportWS = (Excel._Worksheet)ReportWB.ActiveSheet;

                /*
                                ReportItemID = "Item ID: " + ItemID_TextBox.Text;
                    ReportSpotWelder = "Spotweld: " + Spotweld_TextBox.Text;
                    ReportEmployee = "DMP ID: " + SearchDMPID_TextBox.Text;
                    ReportDate
                    */
                ReportRange = ReportWS.get_Range("A1", "E1");
                ReportRange.get_Range("A1", "E1").Merge();
                ReportRange.get_Range("A2", "E2").Merge();
                ReportRange.get_Range("A3", "E3").Merge();
                ReportRange.get_Range("A4", "E4").Merge();
                ReportWS.Shapes.AddPicture(@"\\OHN66FS01\BPprogs\Brake Press Vision\Applications\DMPLogo700.jpg", Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, 0, 0, 325, 60);
                //ReportWS.Range["F1", "H1"].Merge();
                string Name = User_TextBox.Text;
                ReportWS.Cells[1, 6] = ReportItemID;
                ReportWS.Cells[2, 6] = ReportSpotWelder;
                ReportWS.Cells[3, 6] = ReportEmployee;
                ReportWS.Cells[4, 6] = ReportDate;
                ReportWS.get_Range("F1", "F4").Font.Bold = true;
                ReportWS.get_Range("F1", "F4").Font.Size = 14;
                ReportRange.EntireColumn.AutoFit();

                string[] ColumnNames = new string[ReportGridView.Columns.Count];
                int ExcelColumns = 1;
                foreach (DataGridViewColumn dc in ReportGridView.Columns)
                {
                    ReportWS.Cells[5, ExcelColumns] = dc.Name;
                    ExcelColumns++;
                }
                ReportWS.get_Range("A5", "L5").Font.Bold = true;
                ReportRange = ReportWS.get_Range("A5", "L5");
                ReportRange.EntireColumn.AutoFit();

                for (int i = 0; i < ReportDataSet.Tables[0].Rows.Count; i++)
                {
                picture++;
                    // to do: format datetime values before printing
                    for (int j = 0; j < ReportDataSet.Tables[0].Columns.Count; j++)
                    {
                        ReportWS.Cells[(i + 6), (j + 1)] = ReportDataSet.Tables[0].Rows[i][j];
                    }
            }
            ReportRange = ReportWS.get_Range("A" + picture.ToString(), "L" + (picture + 25).ToString());
            chart1.ChartAreas[0].AxisX.Interval = 1;

            foreach (DataGridViewRow r in ReportGridView.Rows)
            {
                chart1.Series["Series1"].Points.AddXY(r.Cells[0].Value.ToString(), r.Cells[5].Value.ToString());
                //chart1.Series["Series1"].Points.Add(Convert.ToDouble(r.Cells[5].Value));
            }
            chart1.SaveImage(@"C:\Users\rgarland\Desktop\" + ReportName + ".jpg", ChartImageFormat.Jpeg);
            picture = (picture + 6) * 15; 
            //ReportWS.Shapes.AddPicture(@"C:\Users\rgarland\Desktop\" + ReportName + ".jpg", Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, 0, picture, 650, 350);
            ReportRange = ReportWS.get_Range("A6", "L6");
                ReportRange.EntireColumn.AutoFit();
                //string ReportPDFName = "Spotweld_Report_" + ReportName + ".xls";
                /*
                SaveFileDialog saveFile = new SaveFileDialog();
                saveFile.Filter = "Excel Files (*.xls)|*.xls|All files (*.*)|*.*";
                saveFile.FileName = ReportPDFName;
                if (saveFile.ShowDialog() == DialogResult.OK)
                {
                    ExcelFileLocation = saveFile.FileName;
                    ReportWS.SaveAs(ExcelFileLocation);
                    //ReportWB.Close();
                }
                */
                //ExcelFileLocation = saveFile.FileName;
                ReportWS.SaveAs(ExcelFileLocation);
                ReportWB.Close();
            //}
            //ReportWB.Close();
            /*
            ReportWS.SaveAs(@"C:\Users\rgarland\Desktop\ExcelTest.xls", Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing,
                false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange,
                Type.Missing, Type.Missing);
            */

        }

        private void CreateExcel_RunWorkerComplete(object sender, RunWorkerCompletedEventArgs e)
        {
            Process.Start(ExcelFileLocation);
        }

        private void EmployeeLogOff()
        {
            SqlConnection ReportLogoff = new SqlConnection(SQL_Source);
            SqlCommand Logoff = new SqlCommand();
            Logoff.CommandType = System.Data.CommandType.Text;
            Logoff.CommandText = "UPDATE [dbo].[LoginData] SET LogoutDateTime=@LogoutDateTime WHERE LoginDateTime=@LoginDateTime";
            Logoff.Connection = ReportLogoff;
            Logoff.Parameters.AddWithValue("@LoginDateTime", LoginTime.ToString());
            Logoff.Parameters.AddWithValue("@LogoutDateTime", Clock_TextBox.Text);
            ReportLogoff.Open();
            Logoff.ExecuteNonQuery();
            ReportLogoff.Close();
        }

        private void FindItemImage(object sender, EventArgs e)
        {
            ItemID = ItemIDResults_TextBox.Text;
            ItemIDSplit = ItemID.Split('-');
            ItemID_Three = double.Parse(ItemIDSplit[0]);
            ItemID_Five = double.Parse(ItemIDSplit[1]);

            if (ItemID_Five >= 1 && ItemID_Five <= 10000)
            {
                ItemImagePath = @"\\insidedmp.com\Corporate\OH\OH Common\Part Pictures\1-9999\";
            }
            else if (ItemID_Five >= 10000 && ItemID_Five <= 14999)
            {
                ItemImagePath = @"\\insidedmp.com\Corporate\OH\OH Common\Part Pictures\10000-14999\";
            }
            else if (ItemID_Five >= 15000 && ItemID_Five <= 19999)
            {
                ItemImagePath = @"\\insidedmp.com\Corporate\OH\OH Common\Part Pictures\15000-19999\";
            }
            else if (ItemID_Five >= 20000 && ItemID_Five <= 24999)
            {
                ItemImagePath = @"\\insidedmp.com\Corporate\OH\OH Common\Part Pictures\20000-24999\";
            }
            else if (ItemID_Five >= 25000 && ItemID_Five <= 29999)
            {
                ItemImagePath = @"\\insidedmp.com\Corporate\OH\OH Common\Part Pictures\25000-29999\";
            }
            else if (ItemID_Five >= 30000 && ItemID_Five <= 34999)
            {
                ItemImagePath = @"\\insidedmp.com\Corporate\OH\OH Common\Part Pictures\30000-34999\";
            }
            else if (ItemID_Five >= 35000 && ItemID_Five <= 39999)
            {
                ItemImagePath = @"\\insidedmp.com\Corporate\OH\OH Common\Part Pictures\35000-39999\";
            }
            else if (ItemID_Five >= 40000 && ItemID_Five <= 44999)
            {
                ItemImagePath = @"\\insidedmp.com\Corporate\OH\OH Common\Part Pictures\40000-44999\";
            }
            else if (ItemID_Five >= 45000 && ItemID_Five <= 49999)
            {
                ItemImagePath = @"\\insidedmp.com\Corporate\OH\OH Common\Part Pictures\45000-49999\";
            }
            else if (ItemID_Five >= 50000 && ItemID_Five <= 54999)
            {
                ItemImagePath = @"\\insidedmp.com\Corporate\OH\OH Common\Part Pictures\50000-54999\";
            }
            else if (ItemID_Five >= 55000 && ItemID_Five <= 59999)
            {
                ItemImagePath = @"\\insidedmp.com\Corporate\OH\OH Common\Part Pictures\55000-59999\";
            }
            else if (ItemID_Five >= 60000 && ItemID_Five <= 69999)
            {
                ItemImagePath = @"\\insidedmp.com\Corporate\OH\OH Common\Part Pictures\60000-69999\";
            }
            else if (ItemID_Five >= 70000 && ItemID_Five <= 79999)
            {
                ItemImagePath = @"\\insidedmp.com\Corporate\OH\OH Common\Part Pictures\70000-79999\";
            }

            string ImageName = ItemID + ".JPG";

            try
            {
                if (File.Exists(ItemImagePath + ImageName))
                {
                    Part_PictureBox.Image = Image.FromFile(ItemImagePath + ImageName);
                    Part_PictureBox.SizeMode = PictureBoxSizeMode.StretchImage;
                }
                else if (File.Exists(ItemImagePath + ItemIDSplit[1] + ".JPG"))
                {
                    Part_PictureBox.Image = Image.FromFile(ItemImagePath + ItemIDSplit[1] + ".JPG");
                    Part_PictureBox.SizeMode = PictureBoxSizeMode.StretchImage;

                }
                else if (File.Exists(ItemImagePath + ItemIDSplit[0] + '-' + ItemIDSplit[1] + ".JPG"))
                {
                    Part_PictureBox.Image = Image.FromFile(ItemImagePath + ItemIDSplit[0] + '-' + ItemIDSplit[1] + ".JPG");
                    Part_PictureBox.SizeMode = PictureBoxSizeMode.StretchImage;

                }
                else
                {
                    Part_PictureBox.Image = null;
                }
            }
            catch
            {
                MessageBox.Show("Error Finding Image");
            }
        }
        
        private void PartImage_DoWork(object sender, DoWorkEventArgs e)
        {
            if (PartImage.IsBusy != true)
            {
                FindItemImage(null, null);
            }
        }        

        private void PartImage_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {

        }
        
        /*********************************************************************************************************************
        * 
        * Methods Region End
        * 
        **********************************************************************************************************************/
        #endregion

        /*********************************************************************************************************************
        * 
        * Events Region Start
        * -- Inner Regions: 3
        * 
        * - Timer Region
        * - TextBox Method Region
        * - TextBox Enter Region 
        * 
        *********************************************************************************************************************/
        #region

        private void Clock_Tick(object sender, EventArgs e)
        {
            string AMPM = "";
            string Date = DateTime.Today.ToShortDateString();
            string Time = "";

            ClockHour = DateTime.Now.Hour;
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

        private void ReportGridView_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
            {
                DataGridViewRow Row = ReportGridView.Rows[e.RowIndex];
                ItemIDResults_TextBox.Text = Row.Cells[0].Value.ToString();
                ReferenceNumberResults_TextBox.Text = Row.Cells[1].Value.ToString();
                RunDateResults_TextBox.Text = Row.Cells[2].Value.ToString();
                OperationTimeResults_TextBox.Text = Row.Cells[3].Value.ToString();
                ActualOperationTime = Row.Cells[3].Value.ToString();
                PlannedTimeResults_TextBox.Text = Row.Cells[4].Value.ToString();
                PlannedOperationTime = Row.Cells[4].Value.ToString();
                PartsManufacturedResults_TextBox.Text = Row.Cells[5].Value.ToString();
                PPMResults_TextBox.Text = Row.Cells[6].Value.ToString();
                EfficiencyResults_TextBox.Text = Row.Cells[7].Value.ToString();
                OEEResults_TextBox.Text = Row.Cells[8].Value.ToString();
                SpotWeldResults_TextBox.Text = Row.Cells[9].Value.ToString();
                EmployeeResults_TextBox.Text = Row.Cells[10].Value.ToString();
                DMPIDResults_TextBox.Text = Row.Cells[11].Value.ToString();
                OperationIDResults_TextBox.Text = Row.Cells[12].Value.ToString();
                if (PartImage.IsBusy != true)
                {
                    PartImage.RunWorkerAsync();
                }
            }
            ConvertPlannedTime();
            ConvertActualTime();
        }

        /*********************************************************************************************************************
        * TextBox Enter Region Start
        * 
        * -- Total TextBox: 16
        * 
        *********************************************************************************************************************/
        #region

        private void ItemIDResults_TextBox_Enter(object sender, EventArgs e)
        {
            this.ActiveControl = null;
        }

        private void EmployeeResults_TextBox_Enter(object sender, EventArgs e)
        {
            this.ActiveControl = null;
        }

        private void OperationIDResults_TextBox_Enter(object sender, EventArgs e)
        {
            this.ActiveControl = null;
        }

        private void ReferenceNumberResults_TextBox_Enter(object sender, EventArgs e)
        {
            this.ActiveControl = null;
        }

        private void DMPIDResults_TextBox_Enter(object sender, EventArgs e)
        {
            this.ActiveControl = null;
        }

        private void RunDateResults_TextBox_Enter(object sender, EventArgs e)
        {
            this.ActiveControl = null;
        }

        private void PlannedTimeResults_TextBox_Enter(object sender, EventArgs e)
        {
            this.ActiveControl = null;
        }

        private void OperationTimeResults_TextBox_Enter(object sender, EventArgs e)
        {
            this.ActiveControl = null;
        }

        private void PPMResults_TextBox_Enter(object sender, EventArgs e)
        {
            this.ActiveControl = null;
        }

        private void EfficiencyResults_TextBox_Enter(object sender, EventArgs e)
        {
            this.ActiveControl = null;
        }

        private void OEEResults_TextBox_Enter(object sender, EventArgs e)
        {
            this.ActiveControl = null;
        }

        private void SpotWeldResults_TextBox_Enter(object sender, EventArgs e)
        {
            this.ActiveControl = null;
        }

        private void PartsManufacturedResults_TextBox_Enter(object sender, EventArgs e)
        {
            this.ActiveControl = null;
        }

        private void User_TextBox_Enter(object sender, EventArgs e)
        {
            this.ActiveControl = null;
        }

        private void Clock_TextBox_Enter(object sender, EventArgs e)
        {
            this.ActiveControl = null;
        }

        private void DMPID_TextBox_Enter(object sender, EventArgs e)
        {
            this.ActiveControl = null;
        }

        /*********************************************************************************************************************
        * 
        * TextBox Enter Region End
        * 
        *********************************************************************************************************************/
        #endregion

        /*********************************************************************************************************************
        * 
        * Events Region End
        * 
        *********************************************************************************************************************/
        #endregion

        /*********************************************************************************************************************
        * 
        * Testing Region Start
        * 
        *********************************************************************************************************************/
        #region

        private void Total_Button_Click(object sender, EventArgs e)
        {
            ReportTotals();
        }

        private void ReportTotals()
        {
            double rows = 0;
            double SumOperationTime = 0;
            double SumPlannedTime = 0;
            int SumPartsManufactured = 0;
            double AveragePPM = 0;
            foreach (DataGridViewRow row in ReportGridView.Rows)
            {
                SumOperationTime += Convert.ToDouble(row.Cells[3].Value);
                SumPlannedTime += Convert.ToDouble(row.Cells[4].Value);
                SumPartsManufactured += Convert.ToInt32(row.Cells[8].Value);
                AveragePPM += Convert.ToDouble(row.Cells[9].Value);
                rows++;
            }
            OperationTimeTotal_TextBox.Text = SumOperationTime.ToString("0.0000");
            PartsMTotal_TextBox.Text = SumPartsManufactured.ToString();
            PlannedTimeTotal_TextBox.Text = SumPlannedTime.ToString("0.0000");
            PPMAverageTotal_TextBox.Text = (AveragePPM / (rows)).ToString("0.00");
            Efficiency_TextBox.Text = (SumPlannedTime / SumOperationTime).ToString("0.00" + " %");
        }

        /*********************************************************************************************************************
        * 
        * Testing Region End
        * 
        *********************************************************************************************************************/
        #endregion

        private void chart1_Click(object sender, EventArgs e)
        {
        }
    }
}
