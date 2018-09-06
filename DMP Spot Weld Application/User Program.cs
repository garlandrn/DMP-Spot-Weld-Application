using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Runtime.InteropServices;
using System.IO;
using System.Threading;
using System.Xml;
using System.Diagnostics;
using System.IO.Ports;
using Opc.Da;
using System.ComponentModel;

/*
 * 
 * Program: DMP Spot Weld Application
 * Form: User Program
 * Created By: Ryan Garland
 * Last Updated on 8/30/18
 * 
 */

namespace DMP_Spot_Weld_Application
{
    public partial class User_Program : Form
    {
        public static Form UserProgram;
        BackgroundWorker ConnectToOPC;
        BackgroundWorker PartCounterOPC;
        public User_Program()
        {
            InitializeComponent();
            //BackgroundWorker ConnectToOPC;

            // Connect to Kepware Server
            ConnectToOPC = new BackgroundWorker();
            ConnectToOPC.DoWork += new DoWorkEventHandler(ConnectToServer_OPC);
            ConnectToOPC.RunWorkerCompleted += new RunWorkerCompletedEventHandler(ConnectToServer_OPC_RunWorkerCompleted);

            // 
            PartCounterOPC = new BackgroundWorker();
            PartCounterOPC.DoWork += new DoWorkEventHandler(PartsCompleted_OPC);
            PartCounterOPC.RunWorkerCompleted += new RunWorkerCompletedEventHandler(PartsCompleted_RunWorkerCompleted);
            UserProgram = this;            
        }

        /********************************************************************************************************************
        * 
        * Global Variables
        *  
        ********************************************************************************************************************/

        /********************************************************************************************************************
        * 
        * Barcode Scanner Variables 
        * 
        ********************************************************************************************************************/

        /********************************************************************************************************************
        * 
        * OPC Tag Variables 
        * 
        ********************************************************************************************************************/

        // ConnectToServer_OPC Method
        private Opc.Da.Server OPCServer;
        private OpcCom.Factory OPCFactory = new OpcCom.Factory();

        private Opc.Da.Subscription GroupRead;
        private Opc.Da.SubscriptionState GroupStateRead;

        // OPC Writing Methods

        // StartNewJob_OPC Method
        private Opc.Da.Subscription StartNewJob_Write;
        private Opc.Da.SubscriptionState StartNewJob_StateWrite;

        // ItemIDWrite_OPC Method
        private Opc.Da.Subscription ItemID_Write;
        private Opc.Da.SubscriptionState ItemID_StateWrite;

        // OperationSelect_OPC Method
        private Opc.Da.Subscription OperationSelection_Write;
        private Opc.Da.SubscriptionState OperationSelection_StateWrite;

        // TeachSensor_OPC Method
        private Opc.Da.Subscription TeachSensor_Write;
        private Opc.Da.SubscriptionState TeachSensor_StateWrite;

        // SystemInRunMode_OPC Method 
        private Opc.Da.Subscription RunMode_Write;
        private Opc.Da.SubscriptionState RunMode_WriteState;

        // SystemInSetupMode_OPC Method
        private Opc.Da.Subscription SetupMode_Write;
        private Opc.Da.SubscriptionState SetupMode_WriteState;

        // OPC Reading Methods
        // PartsCompleted_OPC Method
        private Opc.Da.Subscription PartComplete_GroupRead;
        private Opc.Da.SubscriptionState PartComplete_StateRead;

        // PLC_JobIDRead_OPC Method
        private Opc.Da.Subscription Part_JobID_GroupRead;
        private Opc.Da.SubscriptionState Part_JobID_StateRead;

        // Currently Not Used
        private Opc.Da.Subscription StartNewJobComplete_GroupWrite;
        private Opc.Da.SubscriptionState StartNewJobComplete_StateWrite;

        // PartCompleted_GroupRead_DataChanged
        private string Job_Order_Counter_ACC;
        private string Part_Complete_Counter_ACC;
        private string Part_Complete_Counter_PRE;
        private string HMI_Part_Complete_VALUE;
        private string Setup_Mode_Set_VALUE = "";
        private string Fault_Value = "";
        private string Item_Number_PLC = "";

        // SearchItemIDOperation Method
        private static string Spot_Weld_Operation_Selection;

        // SpotWeldID Method
        private static string SpotWeld_TagID;

        // Run Mode Input Value
        private static int OPC_RunModeIntValue;

        // Not Used
        //private List<Item> ItemID_OPCList = new List<Item>();
        //private List<Item> NewJob_List = new List<Item>();
        //private List<Item> scanList = new List<Item>();

        //BackgroundWorker ConnectToOPC = new System.ComponentModel.BackgroundWorker();
        //ConnectToOPC = new BackgroundWorker();
        //ConnectToOPC.DoWork += new DoWorkEventHandler();
        //ConnectToOPC.RunWorkerCompleted += new RunWorkerCompletedEventHandler(ConnectToOPC_RunWorkerCompleted);

        /********************************************************************************************************************
        * 
        * Form Load Variables 
        * 
        ********************************************************************************************************************/

        // SQL Data Source String
        private string SQL_Source = @"Data Source = OHN7009,49172; Initial Catalog = Spot_Weld_Data; Integrated Security = True; Connect Timeout = 15;";

        // SQL Login Data Info
        private string LoginForm = "User Program";
        private string LoginTime = "";

        // SQL ItemOperationData
        // ItemRunCounter()
        private static int ItemRunCount = 0;
        // OperationIDCounter()
        private static int OperationsID = 0;

        private string Customer = "";
        private string CustomerPartNumber = "";

        private float Efficiency;
        private float Utilization;
        private float OEE;

        private float TimePlanned;
        private float TimeActual;

        // 
        private string JobStartTime = "";
        private string JobEndTime = "";
        
        // 
        private static bool JobFound = false;

        // Clock_Timer();
        private static int ClockHour;
        private static int ClockMinute;
        private static int ClockSecond;

        // FindTotalRunTime()
        TimeSpan TimeOfOperation = new TimeSpan();
        private float ItemSetupTime;
        
        // ShowResult();
        public static string CurrentItemID = "";
        public static int CurrentItemIDOperation;

        // Get The Current Jobs Sequence Number 
        List<string> OperationValue = new List<string>();
        private static int OperationRows_NA = 0;

        // RunningStatistics();
        private static double CurrentParts;
        private double PartsRemaining;

        // ItemOperationCalculation();
        private static string AveragePPM_String = "";
        private static string PartsManufacturedTotal_String = "";
        private static double TotalItemPartsManufactured = 0;
        private static double PartsManufacturedTotal_Double = 0;
        private float AveragePPM = 0;


        // ViewPrint_Button_Click();
        // PDF File Paths For Setup Cards and Prints
        private string PDFSetupPath = "";
        private string SpotWeldSetupCardPDFPath = @"\\insidedmp.com\Corporate\OH\OH Common\Engineering\Spotweld\Spotweld Application Files\Spotweld Setup Cards\";
        private string ItemPrintPDFPath = @"\\insidedmp.com\Corporate\OH\OH Common\Engineering\Spotweld\Spotweld Application Files\Spotweld Prints\";

        //private string PDFPrintPath = "";
        //private string PDFImagePath = "";
        //private string CurrentSetupCardID = "";
        //private string BrakePressNamePDFPath = "";
        //private string ItemIDBrakePressPDFPath = "";
        //private string ItemIDPrintPDFPath = "";

        // Custom Colors For Buttons
        Color ScanRunModeColor = ColorTranslator.FromHtml("#32EB00");
        Color SetupModeColor = SystemColors.Control;

        // SearchForItemID Method
        // Find the Value of The Components For Current Job
        public static string Component_1 = "";
        public static string Component_2 = "";
        public static string Component_3 = "";
        public static string Component_4 = "";

        /********************************************************************************************************************
        * 
        * Variables In Testing Start
        * 
        ********************************************************************************************************************/

        private Stopwatch RunModeStopWatch = new Stopwatch();
        private Stopwatch SetupModeStopWatch = new Stopwatch();

        // Not Used
        private Stopwatch PartCompleteStopWatch = new Stopwatch();
        private Stopwatch GoToSetupMode = new Stopwatch();
        
        // OEE and PPM Calculation Run Mode
        private static string PastItemPPM_String = ""; // Gets Value From Paccar_Item_Data When Scanning Job
        private static double PastItemPPM;             // Gets Value After Converting to double PastItemPPM_String
        private static float PastPPM;

        // bool Values to Indicate the Current Button Clicked
        private static bool RunMode_Button_Clicked = false;
        private static bool SetupMode_Button_Clicked = false;
        private static bool TeachSensor_Button_Clicked = false;
        private static bool PartsComplete_Activate = false;


        private static string ProgramListUpdate_String = "";
        public static string RefreshUpdate_String = "";
        public static string ScanOutComputer = "";
        
        private string[] CustomerCell = { "CAT", "John Deere", "Navistar", "Paccar" };
        private string[] CATSpotWelders = { "123R", "1088" };
        private string[] JohnDeereSpotWelders = { "108R", "150R" };
        private string[] NavistarSpotWelders = {"104R", "121R", "154R"};
        private string[] PaccarSpotWelders = { "153R", "155R" };      

        // Not in Use
        private static bool Component1Found = false;
        private static bool Component2Found = false;
        private static bool Component3Found = false;
        private static bool ToolingFound = false;
        private static bool FixtureFound = false;
        private static bool AllComponentsFound = false;
        private static int Comp;
        private static string CurrentFixture = "";
        private static bool FixtureScanned = false;
        private static bool FixtureForItem = false;
        private string Job_Order_Counter = "";
        private string Part_Complete_ACC = "";
        private string Part_Complete_PRE = "";
        private string HMI_Part_Complete = "";
        private string Setup_Mode_Set = "";
        private string[] SpotweldNames = { "121R", "154R", "153R", "155R" };

        User_Program_Part_Not_Completed NotCompleted = new User_Program_Part_Not_Completed();

        /********************************************************************************************************************
        * 
        * Variables In Testing End
        * 
        *********************************************************************************************************************
        *********************************************************************************************************************
        * 
        * User_Program Start
        * 
        ********************************************************************************************************************/

        private void User_Program_Load(object sender, EventArgs e)
        {
            Company_ComboBox.Items.AddRange(CustomerCell);
            SpotWeldID(); // Find the Computer and SpotWeld Name

            // Add Employee Login To SQL Data
            SqlConnection UserLogin = new SqlConnection(SQL_Source);
            SqlCommand Login = new SqlCommand();
            Login.CommandType = System.Data.CommandType.Text;
            Login.CommandText = "INSERT INTO [dbo].[LoginData] (EmployeeName,DMPID,LoginDateTime,LoginForm,SpotWelder) VALUES (@EmployeeName,@DMPID,@LoginDateTime,@LoginForm,@SpotWelder)";
            Login.Connection = UserLogin;
            Login.Parameters.AddWithValue("@LoginDateTime", Clock_TextBox.Text);
            Login.Parameters.AddWithValue("@EmployeeName", User_TextBox.Text);
            Login.Parameters.AddWithValue("@DMPID", DMPID_TextBox.Text);
            Login.Parameters.AddWithValue("@LoginForm", LoginForm.ToString());
            Login.Parameters.AddWithValue("@SpotWelder", SpotWeld_ComboBox.Text);
            UserLogin.Open();
            Login.ExecuteNonQuery();
            UserLogin.Close();

            Clock.Enabled = true;
            LoginTime = Clock_TextBox.Text;
            StartNewJob_Button.Focus();
            //ConnectToServer_OPC();
            ConnectToOPC.RunWorkerAsync(); // Setup Connection To Kepware OPC Server
            OPCStatus_Timer.Start();
        }

        /********************************************************************************************************************
        * 
        * Buttons Region Start 
        * -- Total Buttons: 15
        * 
        * --- User Program Form Buttons
        * --- Total: 5 
        * - Help Click
        * - LogOff Click
        * - Sytelin Scan Out Click
        * - ViewSchedule Click
        * - HideSchedule Click
        * - ReportError Click
        * 
        * --- ItemData GroupBox Buttons
        * --- Total: 3
        * - ViewSetupCard Click
        * - ViewPrint Click
        * - CheckCardData Click
        * 
        * --- JobData GroupBox Button
        * --- Total: 1
        * - ResetWeldCount Click 
        * 
        * --- OPC Buttons GroupBox Buttons
        * --- Total: 5
        * - ScanNewJob Click
        * - RunMode Click
        * - SetupMode Click
        * - CancelRun Click
        * - JobEnd Click
        * 
        ********************************************************************************************************************/
        #region

        /*********************************************************************************************************************
        * 
        * User Program Form Buttons
        * 
        ********************************************************************************************************************/

        // Open View PDF Form and Display the User Manual
        private void Help_Button_Click(object sender, EventArgs e)
        {
            View_PDF PDFViewer = new View_PDF();
            PDFViewer.AcroPDF.src = @"\\insidedmp.com\Corporate\OH\OH Common\Engineering\Spotweld\Spotweld Application Files\Spot Weld User Manual.pdf";
            PDFViewer.AcroPDF.setZoom(100);
            PDFViewer.AcroPDF.setCurrentPage(4);
            PDFViewer.AcroPDF.BringToFront();
            PDFViewer.BringToFront();
            PDFViewer.Show();
            UserProgram.Enabled = false;
        }
        
        // Open User Program Scan Out Form and Display 
        private void SytelineScanOut_Button_Click(object sender, EventArgs e)
        {
            User_Program_Scan_Out ScanOut = new User_Program_Scan_Out();
            ScanOut.EmployeeNumber_TextBox.Text = DMPID_TextBox.Text;
            UserProgram.Enabled = false;
            if (ScanOut.ShowDialog(this) == DialogResult.Yes)
            {
                ScanOutComputer = "Yes";
                ScanOutCompletion_SQL();
            }
            else if (ScanOut.DialogResult == DialogResult.No)
            {
                ScanOutComputer = "No";
                ScanOutCompletion_SQL();
            }
        }

        // Updates the SQL Table with Logoff Time
        private void LogOff_Button_Click(object sender, EventArgs e)
        {
            EmployeeLogOff_SQL();
            OPCStatus_Timer.Enabled = false;
            DMP_Spot_Weld_Login.Current.Focus();
            DMP_Spot_Weld_Login.Current.Enabled = true;
            DMP_Spot_Weld_Login.Current.WindowState = FormWindowState.Maximized;
            DMP_Spot_Weld_Login.Current.ShowInTaskbar = true;
            this.Close();
        }

        private void ViewSchedule_Button_Click(object sender, EventArgs e)
        {

        }

        private void ReportError_Button_Click(object sender, EventArgs e)
        {

        }

        private void HideSchedule_Button_Click(object sender, EventArgs e)
        {

        }
        
        /*********************************************************************************************************************
        * 
        * ItemData GroupBox Buttons
        * 
        ********************************************************************************************************************/

        // Open Up PDF Form and Load Item Print
        private void ViewPrint_Button_Click(object sender, EventArgs e)
        {
            CurrentItemID = CurrentItemID.Replace("-", "");
            string CompletePDFPath = Path.Combine(ItemPrintPDFPath, CurrentItemID);
            string ViewPrintPDFFile = CompletePDFPath + "_Print.pdf";
            View_PDF PDFViewer = new View_PDF();
            PDFViewer.AcroPDF.src = ViewPrintPDFFile;
            PDFViewer.AcroPDF.BringToFront();
            PDFViewer.Show();
            UserProgram.Enabled = false;
        }

        // Open Up PDF Form and Load Item Setup Card
        private void ViewSetupCard_Button_Click(object sender, EventArgs e)
        {
            PDFSetupPath = SpotWeldSetupCardPDFPath + CurrentItemID;
            string ViewSetupPDFFile = PDFSetupPath + ".pdf";
            View_PDF PDFViewer = new View_PDF();
            PDFViewer.AcroPDF.src = ViewSetupPDFFile;
            PDFViewer.AcroPDF.BringToFront();
            PDFViewer.Show();
            PDFViewer.BringToFront();
            PDFViewer.AcroPDF.setZoom(85);
            UserProgram.Enabled = false;
        }

        // Open Check Card Data Form
        private void CheckCardData_Button_Click(object sender, EventArgs e)
        {
            User_Program_Check_Card_Data_Enter CheckCardData = new User_Program_Check_Card_Data_Enter();
            CheckCardData.Show();
            CheckCardData.DateTime_TextBox.Text = Clock_TextBox.Text;
            CheckCardData.ItemID_TextBox.Text = ItemID_TextBox.Text;
            CheckCardData.Sequence_TextBox.Text = Sequence_TextBox.Text;
            CheckCardData.OperatorName_TextBox.Text = User_TextBox.Text;
            CheckCardData.OperationID_TextBox.Text = OperationsID.ToString();
            CheckCardData.Customer_TextBox.Text = Customer;
            CheckCardData.CustomerPartNumber_TextBox.Text = CustomerPartNumber;
            UserProgram.Enabled = false;
        }

        /*********************************************************************************************************************
        * 
        * JobData GroupBox Buttons
        * 
        ********************************************************************************************************************/

        private void ResetWeldCount_Button_Click(object sender, EventArgs e)
        {
            User_Program_Reset_Weld_Count_Dialog ResetDialog = new User_Program_Reset_Weld_Count_Dialog();
            ResetDialog.Show();
            UserProgram.Enabled = false;
        }
        
        /*********************************************************************************************************************
        * 
        * OPC Buttons GroupBox Buttons 
        * 
        ********************************************************************************************************************/

        private void StartNewJob_Button_Click(object sender, EventArgs e)
        {
            StartNewJob_OPC();  // Set Inputs in PLC to Start New Job
            ClearForm();        // Clear TextBoxes
            ItemID_TextBox.ReadOnly = false;
            ItemID_TextBox.Focus();
            StartNewJob_Button.BackColor = ScanRunModeColor; 
            RunMode_Button.BackColor = Color.Transparent;
            SetupMode_Button.BackColor = Color.Transparent;
        }

        private void TeachSensor_Button_Click(object sender, EventArgs e)
        {
            TeachSensor_OPC();

            UserProgram.Enabled = false;
            User_Program_Teach_Senson TS = new User_Program_Teach_Senson(this);
            if (TS.ShowDialog(this) == DialogResult.Yes)
            {
                OperationInitialize(); // SQL and Job Data Begins to Collect
                SetupMode_Button.PerformClick();
                SetupMode_Button.Focus();
                PartsComplete_Activate = true; // Data Change Event Starts
                PartCounterOPC.RunWorkerAsync(); // Activate Data Change Event
            }
        }

        private void RunMode_Button_Click(object sender, EventArgs e)
        {
            SetupModeSet_TextBox.Clear();

            // Put System into Run Mode
            OPC_RunModeIntValue = 1;
            SystemInRunMode_OPC();
            
            // Button Actions
            StartNewJob_Button.Enabled = false; // Disabled Until End Job is Clicked
            TeachSensor_Button.Enabled = true; // Can Be Clicked 
            JobEnd_Button.Enabled = true;      // Replace CancelRun Button 
            JobEnd_Button.Visible = true;
            CancelRun_Button.Enabled = false;
            CancelRun_Button.Visible = false;
            ResetWeldCount_Button.Enabled = true;
            ResetWeldCount_Button.Visible = true;
            StartNewJob_Button.BackColor = Color.Transparent;
            RunMode_Button.BackColor = ScanRunModeColor;
            SetupMode_Button.BackColor = Color.Transparent;
            PartsRunProgressBar.Visible = true; // Show Progress Bar


            SetupModeStopWatch.Stop();
            RunModeStopWatch.Start();
            Timer.Enabled = true; // RunningStatistics() Start

            string StartingTime = Clock_TextBox.Text;
            string ShortDate = DateTime.Today.ToShortDateString();
                        
            if (RunMode_Button_Clicked == false)
            {
                JobStartTime_TextBox.Text = StartingTime.Replace("   " + ShortDate, "");
                //RunMode_Button_Clicked = true;
            }
            RunMode_Button_Clicked = true; // 
            PartsComplete_Activate = true; // Data Change Event Starts
        }
        

        public void SetupMode_Button_Click(object sender, EventArgs e)
        {
            SystemInSetupMode_OPC();
            // Not Sure? 
            OPC_RunModeIntValue = 0;
            //

            ResetWeldCount_Button.Visible = true;
            RunMode_Button.Enabled = true;
            StartNewJob_Button.BackColor = Color.Transparent;
            SetupMode_Button.BackColor = Color.Yellow;
            RunMode_Button.BackColor = Color.Transparent;

            // Not Sure
            SetupMode_Button_Clicked = true;

            Timer.Enabled = false;  // RunningStatistics() Stop
            RunModeStopWatch.Stop();
            SetupModeStopWatch.Start();
            
        }

        private void CancelRun_Button_Click(object sender, EventArgs e)
        {
            ClearForm();

            // Button Settings
            RunMode_Button.BackColor = Color.Transparent;
            SetupMode_Button.BackColor = Color.Transparent;
            StartNewJob_Button.Enabled = true;
            TeachSensor_Button.Enabled = false;
            RunMode_Button.Enabled = false;
            SetupMode_Button.Enabled = false;
            CancelRun_Button.Enabled = false;
            CancelRun_Button.Visible = false;
        }

        private void JobEnd_Button_Click(object sender, EventArgs e)
        {
            // OPC_Timer.Enabled = false; no longer in use

            PartsComplete_Activate = false; // Stop Data Read

            // Turn off Run Mode
            OPC_RunModeIntValue = 0;
            SystemInRunMode_OPC();

            // Button Settings
            StartNewJob_Button.Enabled = true;
            TeachSensor_Button.Enabled = false;
            RunMode_Button.Enabled = false;
            SetupMode_Button.Enabled = false;
            JobEnd_Button.Visible = false;
            JobEnd_Button.Enabled = false;
            TeachSensor_Button_Clicked = false;
            ResetWeldCount_Button.Visible = false;
            RunMode_Button.BackColor = Color.Transparent;
            SetupMode_Button.BackColor = Color.Transparent;

            RunMode_Button_Clicked = false;
            SetupMode_Button_Clicked = false;

            SetupModeStopWatch.Stop();
            JobEndTime = Clock_TextBox.Text;

            ItemOperationCalculation(); // Calculate
            FindTotalRunTime();
            OperationOEECalculation();
            OperationOEEData_SQL();
            ItemOperationDataEnd_SQL();
            OperationDataEnd_SQL();
            ProgramListUpdate_SQL();
            UserProgram.Focus();

            User_Program_Scan_Out ScanOut = new User_Program_Scan_Out();
            ScanOut.EmployeeNumber_TextBox.Text = DMPID_TextBox.Text;
            ScanOut.JobNumber_TextBox.Text = ReferenceNumber_TextBox.Text;
            ScanOut.TotalCountQtuQtyComp_TextBox.Text = PartsFormed_TextBox.Text;
            UserProgram.Enabled = false;
            if (ScanOut.ShowDialog(this) == DialogResult.Yes)
            {
                ScanOutComputer = "Yes";
                ScanOutCompletion_SQL();
            }
            else if (ScanOut.DialogResult == DialogResult.No)
            {
                ScanOutComputer = "No";
                ScanOutCompletion_SQL();
            }

            StartNewJob_Button.Enabled = true;
            Timer.Enabled = false;
            PartsRunProgressBar.Visible = false;
            // HMI_NotActive_TextBox.Visible = true;            
            RefreshItemData_SQL();
        }

        /*********************************************************************************************************************
        * 
        * Buttons Region End
        * 
        *********************************************************************************************************************/
        #endregion

        /*********************************************************************************************************************
        * 
        * ComboBox Region Start
        * -- Total: 2
        * 
        * - Company ComboBox SelectedIndexChanged
        * - SpotWeld ComboBox SelectedIndexChanged
        * 
        *********************************************************************************************************************/
        #region

        private void Company_ComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            SpotWeld_ComboBox.Items.Clear();
            SpotWeld_ComboBox.Text = "";
                                    
            if (Company_ComboBox.Text == "CAT")
            {
                ProgramListUpdate_String = "UPDATE [dbo].[CAT_Item_Data] SET PartsManufactured=@PartsManufactured,PartsPerMinute=@PartsPerMinute,TotalRuns=@TotalRuns WHERE ItemID=@ItemID AND Sequence=@Sequence";
                RefreshUpdate_String = "SELECT * FROM [dbo].[CAT_Item_Data]";

                SpotWeld_ComboBox.Items.AddRange(CATSpotWelders);
                SqlConnection connection = new SqlConnection(SQL_Source);
                string CAT = "SELECT * FROM [dbo].[CAT_Item_Data]";
                SqlDataAdapter dataAdapter = new SqlDataAdapter(CAT, connection);
                SqlCommandBuilder commandBuilder = new SqlCommandBuilder(dataAdapter);
                DataSet Data = new DataSet();
                dataAdapter.Fill(Data);
                UserProgramGridView.DataSource = Data.Tables[0];

                int rows = 0;
                string CATCount = "SELECT COUNT(*) FROM [dbo].[CAT_Item_Data]";
                SqlConnection CATConnect = new SqlConnection(SQL_Source);
                SqlCommand countRows = new SqlCommand(CATCount, CATConnect);
                CATConnect.Open();
                rows = (int)countRows.ExecuteScalar();
                CATConnect.Close();
            }
            else if (Company_ComboBox.Text == "John Deere")
            {
                ProgramListUpdate_String = "UPDATE [dbo].[JohnDeere_Item_Data] SET PartsManufactured=@PartsManufactured,PartsPerMinute=@PartsPerMinute,TotalRuns=@TotalRuns WHERE ItemID=@ItemID AND Sequence=@Sequence";
                RefreshUpdate_String = "SELECT * FROM [dbo].[JohnDeere_Item_Data]";

                SpotWeld_ComboBox.Items.AddRange(JohnDeereSpotWelders);
                SqlConnection connection = new SqlConnection(SQL_Source);
                string JohnDeere = "SELECT * FROM [dbo].[JohnDeere_Item_Data]";
                SqlDataAdapter dataAdapter = new SqlDataAdapter(JohnDeere, connection);
                SqlCommandBuilder commandBuilder = new SqlCommandBuilder(dataAdapter);
                DataSet Data = new DataSet();
                dataAdapter.Fill(Data);
                UserProgramGridView.DataSource = Data.Tables[0];

                int rows = 0;
                string JDCount = "SELECT COUNT(*) FROM [dbo].[JohnDeere_Item_Data]";
                SqlConnection JDConnect = new SqlConnection(SQL_Source);
                SqlCommand countRows = new SqlCommand(JDCount, JDConnect);
                JDConnect.Open();
                rows = (int)countRows.ExecuteScalar();
                JDConnect.Close();
            }
            else if (Company_ComboBox.Text == "Navistar")
            {
                ProgramListUpdate_String = "UPDATE [dbo].[Navistar_Item_Data] SET PartsManufactured=@PartsManufactured,PartsPerMinute=@PartsPerMinute,TotalRuns=@TotalRuns WHERE ItemID=@ItemID AND Sequence=@Sequence";
                RefreshUpdate_String = "SELECT * FROM [dbo].[Navistar_Item_Data]";

                SpotWeld_ComboBox.Items.AddRange(NavistarSpotWelders);
                SqlConnection connection = new SqlConnection(SQL_Source);
                string Navistar = "SELECT * FROM [dbo].[Navistar_Item_Data]";
                SqlDataAdapter dataAdapter = new SqlDataAdapter(Navistar, connection);
                SqlCommandBuilder commandBuilder = new SqlCommandBuilder(dataAdapter);
                DataSet Data = new DataSet();
                dataAdapter.Fill(Data);
                UserProgramGridView.DataSource = Data.Tables[0];

                int rows = 0;
                string NavistarCount = "SELECT COUNT(*) FROM [dbo].[Navistar_Item_Data]";
                SqlConnection NavistarConnect = new SqlConnection(SQL_Source);
                SqlCommand countRows = new SqlCommand(NavistarCount, NavistarConnect);
                NavistarConnect.Open();
                rows = (int)countRows.ExecuteScalar();
                NavistarConnect.Close();
            }
            else if (Company_ComboBox.Text == "Paccar")
            {
                ProgramListUpdate_String = "UPDATE [dbo].[Paccar_Item_Data] SET PartsManufactured=@PartsManufactured,PartsPerMinute=@PartsPerMinute,TotalRuns=@TotalRuns WHERE ItemID=@ItemID AND Sequence=@Sequence";
                RefreshUpdate_String = "SELECT * FROM [dbo].[Paccar_Item_Data]";

                SpotWeld_ComboBox.Items.AddRange(PaccarSpotWelders);
                SqlConnection connection = new SqlConnection(SQL_Source);
                string Paccar = "SELECT * FROM [dbo].[Paccar_Item_Data]";
                SqlDataAdapter dataAdapter = new SqlDataAdapter(Paccar, connection);
                SqlCommandBuilder commandBuilder = new SqlCommandBuilder(dataAdapter);
                DataSet Data = new DataSet();
                dataAdapter.Fill(Data);
                UserProgramGridView.DataSource = Data.Tables[0];

                int rows = 0;
                string PaccarCount = "SELECT COUNT(*) FROM [dbo].[Paccar_Item_Data]";
                SqlConnection PaccarConnect = new SqlConnection(SQL_Source);
                SqlCommand countRows = new SqlCommand(PaccarCount, PaccarConnect);
                PaccarConnect.Open();
                rows = (int)countRows.ExecuteScalar();
                PaccarConnect.Close();            }

        }

        // Not Used
        private void SpotWeld_ComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            
        }

        /*********************************************************************************************************************
        * 
        * ComboBox Region End
        * 
        *********************************************************************************************************************/
        #endregion

        /********************************************************************************************************************
        *  
        *  SQL Data Region Start
        *  -- Total: 11
        *  
        *  - ItemOperationDataStart_SQL
        *  - ItemOperationDataEnd_SQL
        *  - ItemRunCounter_SQL
        *  - OperationIDCounter_SQL
        *  - OperationDataStart_SQL
        *  - OperationDataEnd_SQL
        *  - OperationOEEData_SQL
        *  - ProgramListUpdate_SQL
        *  - RefreshItemData_SQL
        *  - ScanOutCompletion_SQL
        *  - EmployeeLogOff_SQL  
        * 
        *********************************************************************************************************************/
        #region

        private void ItemOperationDataStart_SQL()
        {
            SqlConnection OperationStart = new SqlConnection(SQL_Source);
            SqlCommand StartOperation = new SqlCommand();
            StartOperation.CommandType = System.Data.CommandType.Text;
            StartOperation.CommandText = "INSERT INTO [dbo].[ItemOperationData] (ItemID, Sequence, OperationID, ItemRunCount, StartDateTime, EmployeeName, DMPID, SpotWelder, ReferenceNumber) VALUES (@ItemID,@Sequence,@OperationID,@ItemRunCount,@StartDateTime,@EmployeeName,@DMPID,@SpotWelder,@ReferenceNumber)";
            StartOperation.Connection = OperationStart;
            StartOperation.Parameters.AddWithValue("@ItemID", CurrentItemID);
            StartOperation.Parameters.AddWithValue("@Sequence", CurrentItemIDOperation);
            StartOperation.Parameters.AddWithValue("@OperationID", OperationsID.ToString());
            StartOperation.Parameters.AddWithValue("@ItemRunCount", ItemRunCount.ToString());
            StartOperation.Parameters.AddWithValue("@StartDateTime", Clock_TextBox.Text);
            StartOperation.Parameters.AddWithValue("@EmployeeName", User_TextBox.Text);
            StartOperation.Parameters.AddWithValue("@DMPID", DMPID_TextBox.Text);
            StartOperation.Parameters.AddWithValue("@SpotWelder", SpotWeld_ComboBox.Text);
            StartOperation.Parameters.AddWithValue("@ReferenceNumber", ReferenceNumber_TextBox.Text);
            OperationStart.Open();
            StartOperation.ExecuteNonQuery();
            OperationStart.Close();
        }

        private void ItemOperationDataEnd_SQL()
        {
            SqlConnection OperationEnd = new SqlConnection(SQL_Source);
            SqlCommand EndOperation = new SqlCommand();
            EndOperation.CommandType = System.Data.CommandType.Text;
            EndOperation.CommandText = "UPDATE [dbo].[ItemOperationData] SET EndDateTime=@EndDateTime,PartsManufactured=@PartsManufactured, PartsPerMinute=@PartsPerMinute WHERE OperationID=@OperationID";
            EndOperation.Connection = OperationEnd;
            EndOperation.Parameters.AddWithValue("@OperationID", OperationsID);
            EndOperation.Parameters.AddWithValue("@EndDateTime", Clock_TextBox.Text);
            EndOperation.Parameters.AddWithValue("@PartsManufactured", PartsFormed_TextBox.Text);
            EndOperation.Parameters.AddWithValue("@PartsPerMinute", CurrentPPM_TextBox.Text);
            OperationEnd.Open();
            EndOperation.ExecuteNonQuery();
            OperationEnd.Close();
        }

        private void ItemRunCounter_SQL()
        {
            try
            {
                string CountOperations = "SELECT COUNT(ItemID) FROM [dbo].[ItemOperationData] WHERE ItemID='" + CurrentItemID + "' AND Sequence='" + CurrentItemIDOperation + "'";
                SqlConnection OperationCount = new SqlConnection(SQL_Source);
                SqlCommand CountItemRun = new SqlCommand(CountOperations, OperationCount);
                OperationCount.Open();
                int OperationRunCount = (int)CountItemRun.ExecuteScalar();
                OperationCount.Close();
                ItemRunCount = OperationRunCount + 1;
                //ItemRunCount_TextBox.Text = ItemRunCount.ToString();
            }
            catch (Exception e)
            {
                MessageBox.Show(e.ToString());
            }
        }

        private void OperationIDCounter_SQL()
        {
            try
            {
                string CountOperations = "SELECT COUNT(*) FROM [dbo].[OperationData]";
                SqlConnection OperationCount = new SqlConnection(SQL_Source);
                SqlCommand CountOperation = new SqlCommand(CountOperations, OperationCount);
                OperationCount.Open();
                int OperationCountID = (int)CountOperation.ExecuteScalar();
                OperationCount.Close();
                OperationsID = OperationCountID + 1;
            }
            catch (Exception e)
            {
                MessageBox.Show(e.ToString());
            }
        }

        private void OperationDataStart_SQL()
        {
            try
            {
                SqlConnection OperationStart = new SqlConnection(SQL_Source);
                SqlCommand StartOperation = new SqlCommand();
                StartOperation.CommandType = System.Data.CommandType.Text;
                StartOperation.CommandText = "INSERT INTO [dbo].[OperationData] (ItemID, Sequence, OperationID, RunDateTime, EmployeeName, DMPID, SpotWelder) VALUES (@ItemID,@Sequence,@OperationID,@RunDateTime,@EmployeeName,@DMPID,@SpotWelder)";
                StartOperation.Connection = OperationStart;
                StartOperation.Parameters.AddWithValue("@ItemID", CurrentItemID);
                StartOperation.Parameters.AddWithValue("@Sequence", CurrentItemIDOperation);
                StartOperation.Parameters.AddWithValue("@OperationID", OperationsID.ToString());
                StartOperation.Parameters.AddWithValue("@RunDateTime", Clock_TextBox.Text);
                StartOperation.Parameters.AddWithValue("@EmployeeName", User_TextBox.Text);
                StartOperation.Parameters.AddWithValue("@DMPID", DMPID_TextBox.Text);
                StartOperation.Parameters.AddWithValue("@SpotWelder", SpotWeld_ComboBox.Text);
                OperationStart.Open();
                StartOperation.ExecuteNonQuery();
                OperationStart.Close();
            }
            catch (Exception e)
            {
                MessageBox.Show(e.ToString());
            }
        }

        private void OperationDataEnd_SQL()
        {
            try
            {
                SqlConnection OperationEnd = new SqlConnection(SQL_Source);
                SqlCommand EndOperation = new SqlCommand();
                EndOperation.CommandType = System.Data.CommandType.Text;
                EndOperation.CommandText = "UPDATE [dbo].[OperationData] SET OperationTime=@OperationTime, PartsManufactured=@PartsManufactured,PartsPerMinute=@PartsPerMinute,SetupTime=@SetupTime WHERE OperationID=@OperationID";
                EndOperation.Connection = OperationEnd;
                EndOperation.Parameters.AddWithValue("@OperationID", OperationsID.ToString());
                EndOperation.Parameters.AddWithValue("@OperationTime", TimeActual.ToString());
                EndOperation.Parameters.AddWithValue("@PartsManufactured", PartsFormed_TextBox.Text);
                EndOperation.Parameters.AddWithValue("@PartsPerMinute", CurrentPPM_TextBox.Text);
                EndOperation.Parameters.AddWithValue("@SetupTime", ItemSetupTime.ToString());
                OperationEnd.Open();
                EndOperation.ExecuteNonQuery();
                OperationEnd.Close();
            }
            catch (Exception e)
            {
                MessageBox.Show(e.ToString());
            }
        }

        private void OperationOEEData_SQL()
        {
            try
            {
                SqlConnection OperationOEEReport = new SqlConnection(SQL_Source);
                SqlCommand OEEData = new SqlCommand();
                OEEData.CommandType = System.Data.CommandType.Text;
                OEEData.CommandText = "INSERT INTO [dbo].[OperationOEE] (ItemID, Sequence, OperationID, RunDateTime, OperationTime, PlannedTime, Efficiency, EmployeeName, DMPID, SpotWelder) VALUES (@ItemID,@Sequence,@OperationID,@RunDateTime,@OperationTime,@PlannedTime,@Efficiency,@EmployeeName,@DMPID,@SpotWelder)";
                OEEData.Connection = OperationOEEReport;
                OEEData.Parameters.AddWithValue("@ItemID", CurrentItemID);
                OEEData.Parameters.AddWithValue("@Sequence", CurrentItemIDOperation);
                OEEData.Parameters.AddWithValue("@OperationID", OperationsID.ToString());
                OEEData.Parameters.AddWithValue("@RunDateTime", DateTime.Today.ToShortDateString());
                OEEData.Parameters.AddWithValue("@OperationTime", TimeActual.ToString());
                OEEData.Parameters.AddWithValue("@PlannedTime", TimePlanned.ToString());
                OEEData.Parameters.AddWithValue("@Efficiency", Efficiency.ToString());
                OEEData.Parameters.AddWithValue("@EmployeeName", User_TextBox.Text);
                OEEData.Parameters.AddWithValue("@DMPID", DMPID_TextBox.Text);
                OEEData.Parameters.AddWithValue("@SpotWelder", SpotWeld_ComboBox.Text);
                OperationOEEReport.Open();
                OEEData.ExecuteNonQuery();
                OperationOEEReport.Close();
            }
            catch (Exception e)
            {
                MessageBox.Show(e.ToString());
            }
        }

        private void ProgramListUpdate_SQL()
        {
            try
            {
                SqlConnection ProgramUpdate = new SqlConnection(SQL_Source);
                SqlCommand UpdateItem = new SqlCommand();
                UpdateItem.CommandType = System.Data.CommandType.Text;
                UpdateItem.CommandText = ProgramListUpdate_String;
                UpdateItem.Connection = ProgramUpdate;
                UpdateItem.Parameters.AddWithValue("@ItemID", CurrentItemID);
                UpdateItem.Parameters.AddWithValue("@TotalRuns", ItemRunCount);
                UpdateItem.Parameters.AddWithValue("@PartsManufactured", TotalItemPartsManufactured.ToString());
                UpdateItem.Parameters.AddWithValue("@PartsPerMinute", AveragePPM);
                UpdateItem.Parameters.AddWithValue("@Sequence", Sequence_TextBox.Text);
                ProgramUpdate.Open();
                UpdateItem.ExecuteNonQuery();
                ProgramUpdate.Close();
            }
            catch (Exception e)
            {
                MessageBox.Show(e.ToString());
            }
        }

        private void RefreshItemData_SQL()
        {
            try
            {
                SqlConnection connection = new SqlConnection(SQL_Source);
                string BP1176 = RefreshUpdate_String;
                SqlDataAdapter dataAdapter = new SqlDataAdapter(BP1176, connection);
                SqlCommandBuilder commandBuilder = new SqlCommandBuilder(dataAdapter);
                DataSet Data = new DataSet();
                dataAdapter.Fill(Data);
                UserProgramGridView.DataSource = Data.Tables[0];
            }
            catch (Exception e)
            {
                MessageBox.Show(e.ToString());
            }
            //ItemID_ComboBox.SelectedItem = null;
        }

        private void ScanOutCompletion_SQL()
        {
            SqlConnection ScanOutCompletion = new SqlConnection(SQL_Source);
            SqlCommand CompletionScanOut = new SqlCommand();
            CompletionScanOut.CommandType = System.Data.CommandType.Text;
            CompletionScanOut.CommandText = "UPDATE [dbo].[ItemOperationData] SET ScanOutComputer=@ScanOutComputer WHERE OperationID=@OperationID";
            CompletionScanOut.Connection = ScanOutCompletion;
            CompletionScanOut.Parameters.AddWithValue("@OperationID", OperationsID);
            CompletionScanOut.Parameters.AddWithValue("@ScanOutComputer", ScanOutComputer);
            ScanOutCompletion.Open();
            CompletionScanOut.ExecuteNonQuery();
            ScanOutCompletion.Close();
        }

        private void EmployeeLogOff_SQL()
        {
            SqlConnection UserLogoff = new SqlConnection(SQL_Source);
            SqlCommand Logoff = new SqlCommand();
            Logoff.CommandType = System.Data.CommandType.Text;
            Logoff.CommandText = "UPDATE [dbo].[LoginData] SET LogoutDateTime=@LogoutDateTime WHERE LoginDateTime=@LoginDateTime";
            Logoff.Connection = UserLogoff;
            Logoff.Parameters.AddWithValue("@LoginDateTime", LoginTime.ToString());
            Logoff.Parameters.AddWithValue("@LogoutDateTime", Clock_TextBox.Text);
            UserLogoff.Open();
            Logoff.ExecuteNonQuery();
            UserLogoff.Close();
        }

        /********************************************************************************************************************
        *  
        *  SQL Data Region End
        * 
        *********************************************************************************************************************/
        #endregion

        /********************************************************************************************************************
        *  
        *  OPC Region Start
        *  -- Total: 14
        *  
        *  - ConnectToServer_OPC
        *  - ConnectToServer_OPC_RunWorkerCompleted
        *  - StartNewJob_OPC
        *  - ItemIDWrite_OPC
        *  - OperationSelect_OPC
        *  - TeachSensor_OPC
        *  - SystemInRunMode_OPC
        *  - SystemInSetupMode_OPC
        *  - PartsCompleted_OPC
        *  - PartCompleted_GroupRead_DataChanged
        *  - PartsCompleted_RunWorkerCompleted
        *  - ReadCompleteCallback
        *  - ReadPartNotProgrammedCallback
        *  - WriteCompleteCallback
        *  - Update_UI_OPC
        *  - SetupModeTimeOut_OPC
        *  - HMIPartComplete_OPC
        *  
        ********************************************************************************************************************/
        #region

        private void ConnectToServer_OPC(object sender, EventArgs e)
        {
            try
            {
                // OPC Server
                OPCServer = new Opc.Da.Server(OPCFactory, null);
                OPCServer.Url = new Opc.URL("opcda://OHN66OPC/Kepware.KEPServerEX.V6");
                OPCServer.Connect();

                // OPC Read Groups
                GroupStateRead = new Opc.Da.SubscriptionState();
                GroupStateRead.Name = SpotWeld_TextBox.Text + "_Spotweld";
                GroupStateRead.UpdateRate = 1000;
                GroupStateRead.Active = true;
                GroupRead = (Opc.Da.Subscription)OPCServer.CreateSubscription(GroupStateRead);

                Part_JobID_StateRead = new Opc.Da.SubscriptionState();
                Part_JobID_StateRead.Name = SpotWeld_TextBox.Text + "_JobID";
                Part_JobID_StateRead.UpdateRate = 1000;
                Part_JobID_StateRead.Active = true;
                Part_JobID_GroupRead = (Opc.Da.Subscription)OPCServer.CreateSubscription(Part_JobID_StateRead);

                // RunModePartsCompleted_OPC()
                PartComplete_StateRead = new Opc.Da.SubscriptionState();
                PartComplete_StateRead.Name = SpotWeld_TextBox.Text + "Part Complete Read";
                PartComplete_StateRead.UpdateRate = 1000;
                PartComplete_StateRead.Active = true;
                PartComplete_GroupRead = (Opc.Da.Subscription)OPCServer.CreateSubscription(PartComplete_StateRead);
                PartComplete_GroupRead.DataChanged += new Opc.Da.DataChangedEventHandler(PartCompleted_GroupRead_DataChanged);
                
                // OPC Write Groups

                // StartNewJob_OPC()
                StartNewJob_StateWrite = new Opc.Da.SubscriptionState();
                StartNewJob_StateWrite.Name = SpotWeld_TextBox.Text + "StartNewJob_WriteGroup";
                StartNewJob_StateWrite.UpdateRate = 500;
                StartNewJob_StateWrite.Active = true;
                StartNewJob_Write = (Opc.Da.Subscription)OPCServer.CreateSubscription(StartNewJob_StateWrite);

                StartNewJobComplete_StateWrite = new Opc.Da.SubscriptionState();
                StartNewJobComplete_StateWrite.Name = SpotWeld_TextBox.Text + "StartNewJobComplete_WriteGroup";
                StartNewJobComplete_StateWrite.UpdateRate = 500;
                StartNewJobComplete_StateWrite.Active = true;
                StartNewJobComplete_GroupWrite = (Opc.Da.Subscription)OPCServer.CreateSubscription(StartNewJobComplete_StateWrite);

                // ItemIDWriteTo_OPC()
                ItemID_StateWrite = new Opc.Da.SubscriptionState();
                ItemID_StateWrite.Name = SpotWeld_TextBox.Text + "NewJob_WriteGroup";
                ItemID_StateWrite.Active = false;
                ItemID_Write = (Opc.Da.Subscription)OPCServer.CreateSubscription(ItemID_StateWrite);

                //TeachSensor_OPC()
                TeachSensor_StateWrite = new Opc.Da.SubscriptionState();
                TeachSensor_StateWrite.Name = SpotWeld_TextBox.Text + "TeachSensor_WriteGroup";
                TeachSensor_StateWrite.Active = false;
                TeachSensor_Write = (Opc.Da.Subscription)OPCServer.CreateSubscription(TeachSensor_StateWrite);

                // SystemInRunMode_OPC()
                RunMode_WriteState = new Opc.Da.SubscriptionState();
                RunMode_WriteState.Name = SpotWeld_TextBox.Text + "RunMode_Group";
                RunMode_WriteState.Active = false;
                RunMode_Write = (Opc.Da.Subscription)OPCServer.CreateSubscription(RunMode_WriteState);

                // SystemInSetupMode_OPC()
                SetupMode_WriteState = new Opc.Da.SubscriptionState();
                SetupMode_WriteState.Name = SpotWeld_TextBox.Text + "SetupMode_Group";
                SetupMode_WriteState.Active = false;
                SetupMode_Write = (Opc.Da.Subscription)OPCServer.CreateSubscription(SetupMode_WriteState);

                // OperationSelect_OPC()
                OperationSelection_StateWrite = new Opc.Da.SubscriptionState();
                OperationSelection_StateWrite.Name = SpotWeld_TextBox.Text + "OperationOne_WriteGroup";
                OperationSelection_StateWrite.Active = false;
                OperationSelection_Write = (Opc.Da.Subscription)OPCServer.CreateSubscription(OperationSelection_StateWrite);

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void ConnectToServer_OPC_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            StartNewJob_Button.Enabled = true;
        }

        // Set the Tag Values and Turn on the PLC Inputs to Scan a New Job
        private void StartNewJob_OPC()
        {
            try
            {                
                Opc.Da.Item[] OPC_StartNewJob = new Opc.Da.Item[8];
                OPC_StartNewJob[0] = new Opc.Da.Item();
                OPC_StartNewJob[0].ItemName = SpotWeld_TagID + "HMI_PB_New_Job";            // Turn On
                OPC_StartNewJob[1] = new Opc.Da.Item();
                OPC_StartNewJob[1].ItemName = SpotWeld_TagID + "HMI_Operation_One_PB";      // Turn Off
                OPC_StartNewJob[2] = new Opc.Da.Item();
                OPC_StartNewJob[2].ItemName = SpotWeld_TagID + "HMI_Operation_Two_PB";      // Turn Off
                OPC_StartNewJob[3] = new Opc.Da.Item();
                OPC_StartNewJob[3].ItemName = SpotWeld_TagID + "HMI_Operation_Three_PB";    // Turn Off
                OPC_StartNewJob[4] = new Opc.Da.Item();
                OPC_StartNewJob[4].ItemName = SpotWeld_TagID + "HMI_Operation_Four_PB";     // Turn Off
                OPC_StartNewJob[5] = new Opc.Da.Item();
                OPC_StartNewJob[5].ItemName = SpotWeld_TagID + "HMI_PB_SCAN_NEW_PART";      // Turn On
                OPC_StartNewJob[6] = new Opc.Da.Item();
                OPC_StartNewJob[6].ItemName = SpotWeld_TagID + "RUN_TEACH_MODE_TOGGLE_BIT"; // Turn On
                OPC_StartNewJob[7] = new Opc.Da.Item();
                OPC_StartNewJob[7].ItemName = SpotWeld_TagID + "SYSTEM_IN_RUN_MODE";        // Turn Off
                OPC_StartNewJob = StartNewJob_Write.AddItems(OPC_StartNewJob);
                
                Opc.Da.ItemValue[] OPC_StartNewJobValue = new Opc.Da.ItemValue[8];
                OPC_StartNewJobValue[0] = new Opc.Da.ItemValue();
                OPC_StartNewJobValue[0].ServerHandle = StartNewJob_Write.Items[0].ServerHandle;
                OPC_StartNewJobValue[0].Value = 1;
                OPC_StartNewJobValue[1] = new Opc.Da.ItemValue();
                OPC_StartNewJobValue[1].ServerHandle = StartNewJob_Write.Items[1].ServerHandle;
                OPC_StartNewJobValue[1].Value = 0;
                OPC_StartNewJobValue[2] = new Opc.Da.ItemValue();
                OPC_StartNewJobValue[2].ServerHandle = StartNewJob_Write.Items[2].ServerHandle;
                OPC_StartNewJobValue[2].Value = 0;
                OPC_StartNewJobValue[3] = new Opc.Da.ItemValue();
                OPC_StartNewJobValue[3].ServerHandle = StartNewJob_Write.Items[3].ServerHandle;
                OPC_StartNewJobValue[3].Value = 0;
                OPC_StartNewJobValue[4] = new Opc.Da.ItemValue();
                OPC_StartNewJobValue[4].ServerHandle = StartNewJob_Write.Items[4].ServerHandle;
                OPC_StartNewJobValue[4].Value = 0;
                OPC_StartNewJobValue[5] = new Opc.Da.ItemValue();
                OPC_StartNewJobValue[5].ServerHandle = StartNewJob_Write.Items[5].ServerHandle;
                OPC_StartNewJobValue[5].Value = 1;
                OPC_StartNewJobValue[6] = new Opc.Da.ItemValue();
                OPC_StartNewJobValue[6].ServerHandle = StartNewJob_Write.Items[6].ServerHandle;
                OPC_StartNewJobValue[6].Value = 1;
                OPC_StartNewJobValue[7] = new Opc.Da.ItemValue();
                OPC_StartNewJobValue[7].ServerHandle = StartNewJob_Write.Items[7].ServerHandle;
                OPC_StartNewJobValue[7].Value = 0;

                Opc.IRequest OPCRequest;
                StartNewJob_Write.Write(OPC_StartNewJobValue, 123, new Opc.Da.WriteCompleteEventHandler(WriteCompleteCallback), out OPCRequest);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        
        private void ItemIDWrite_OPC()
        {
            try
            {
                Opc.Da.Item[] OPC_ItemID = new Opc.Da.Item[2];
                OPC_ItemID[0] = new Opc.Da.Item();
                OPC_ItemID[0].ItemName = SpotWeld_TagID + "DM8050_READ_RESULTS.DATA";
                OPC_ItemID[1] = new Opc.Da.Item();
                OPC_ItemID[1].ItemName = SpotWeld_TagID + "HMI_PB_New_Job";
                OPC_ItemID = ItemID_Write.AddItems(OPC_ItemID);
                
                Opc.Da.ItemValue[] OPC_ItemIDValue = new Opc.Da.ItemValue[2];
                OPC_ItemIDValue[0] = new Opc.Da.ItemValue();
                OPC_ItemIDValue[1] = new Opc.Da.ItemValue();
                OPC_ItemIDValue[0].ServerHandle = ItemID_Write.Items[0].ServerHandle;
                OPC_ItemIDValue[0].Value = CurrentItemID;
                OPC_ItemIDValue[1].ServerHandle = ItemID_Write.Items[1].ServerHandle;
                OPC_ItemIDValue[1].Value = 0;

                Opc.IRequest OPCRequest;
                ItemID_Write.Write(OPC_ItemIDValue, 123, new Opc.Da.WriteCompleteEventHandler(WriteCompleteCallback), out OPCRequest);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        
        private void OperationSelect_OPC()
        {
            try
            {
                Opc.Da.Item[] OPC_Operation = new Opc.Da.Item[4];
                OPC_Operation[0] = new Opc.Da.Item();
                OPC_Operation[0].ItemName = SpotWeld_TagID + "HMI_Operation_One_PB";
                OPC_Operation[1] = new Opc.Da.Item();
                OPC_Operation[1].ItemName = SpotWeld_TagID + "HMI_Operation_Two_PB";
                OPC_Operation[2] = new Opc.Da.Item();
                OPC_Operation[2].ItemName = SpotWeld_TagID + "HMI_Operation_Three_PB";
                OPC_Operation[3] = new Opc.Da.Item();
                OPC_Operation[3].ItemName = SpotWeld_TagID + "HMI_Operation_Four_PB";
                OPC_Operation = OperationSelection_Write.AddItems(OPC_Operation);
                
                Opc.Da.ItemValue[] OPC_OperationOneValue = new Opc.Da.ItemValue[4];
                OPC_OperationOneValue[0] = new Opc.Da.ItemValue();
                OPC_OperationOneValue[0].ServerHandle = OperationSelection_Write.Items[0].ServerHandle;
                OPC_OperationOneValue[0].Value = 1;
                OPC_OperationOneValue[1] = new Opc.Da.ItemValue();
                OPC_OperationOneValue[1].ServerHandle = OperationSelection_Write.Items[1].ServerHandle;
                OPC_OperationOneValue[1].Value = 0;
                OPC_OperationOneValue[2] = new Opc.Da.ItemValue();
                OPC_OperationOneValue[2].ServerHandle = OperationSelection_Write.Items[2].ServerHandle;
                OPC_OperationOneValue[2].Value = 0;
                OPC_OperationOneValue[3] = new Opc.Da.ItemValue();
                OPC_OperationOneValue[3].ServerHandle = OperationSelection_Write.Items[3].ServerHandle;
                OPC_OperationOneValue[3].Value = 0;

                Opc.IRequest WriteRequest;
                OperationSelection_Write.Write(OPC_OperationOneValue, 123, new Opc.Da.WriteCompleteEventHandler(WriteCompleteCallback), out WriteRequest);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void TeachSensor_OPC()
        {
            try
            {
                Opc.Da.Item[] OPC_SensorWrite = new Opc.Da.Item[2];
                OPC_SensorWrite[0] = new Opc.Da.Item();
                OPC_SensorWrite[0].ItemName = SpotWeld_TagID + "HMI_PB_TEACH_SENSOR";
                OPC_SensorWrite[1] = new Opc.Da.Item();
                OPC_SensorWrite[1].ItemName = SpotWeld_TagID + "HMI_PB_TEACH_SENSOR";
                OPC_SensorWrite = TeachSensor_Write.AddItems(OPC_SensorWrite);
                
                Opc.Da.ItemValue[] OPC_SensorValue = new Opc.Da.ItemValue[2];
                OPC_SensorValue[0] = new Opc.Da.ItemValue();
                OPC_SensorValue[0].ServerHandle = TeachSensor_Write.Items[0].ServerHandle;
                OPC_SensorValue[0].Value = 1;
                OPC_SensorValue[1] = new Opc.Da.ItemValue();
                OPC_SensorValue[1].ServerHandle = TeachSensor_Write.Items[1].ServerHandle;
                OPC_SensorValue[1].Value = 0;

                Opc.IRequest WriteRequest;
                TeachSensor_Write.Write(OPC_SensorValue, 123, new Opc.Da.WriteCompleteEventHandler(WriteCompleteCallback), out WriteRequest);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void SystemInRunMode_OPC()
        {
            try
            {
                Opc.Da.Item[] OPC_RunMode = new Opc.Da.Item[3];
                OPC_RunMode[0] = new Opc.Da.Item();
                OPC_RunMode[0].ItemName = SpotWeld_TagID + "SYSTEM_IN_RUN_MODE";
                OPC_RunMode[1] = new Opc.Da.Item();
                OPC_RunMode[1].ItemName = SpotWeld_TagID + "RUN_TEACH_MODE_TOGGLE_BIT";
                OPC_RunMode[2] = new Opc.Da.Item();
                OPC_RunMode[2].ItemName = SpotWeld_TagID + "GOOD_WELD";
                OPC_RunMode = RunMode_Write.AddItems(OPC_RunMode);
                
                Opc.Da.ItemValue[] OPC_RunModeValue = new Opc.Da.ItemValue[3];
                OPC_RunModeValue[0] = new Opc.Da.ItemValue();
                OPC_RunModeValue[0].ServerHandle = RunMode_Write.Items[0].ServerHandle;
                OPC_RunModeValue[0].Value = OPC_RunModeIntValue;
                OPC_RunModeValue[1] = new Opc.Da.ItemValue();
                OPC_RunModeValue[1].ServerHandle = RunMode_Write.Items[1].ServerHandle;
                OPC_RunModeValue[1].Value = 0;
                OPC_RunModeValue[2] = new Opc.Da.ItemValue();
                OPC_RunModeValue[2].ServerHandle = RunMode_Write.Items[2].ServerHandle;
                OPC_RunModeValue[2].Value = 1;

                Opc.IRequest OPCRequest;
                RunMode_Write.Write(OPC_RunModeValue, 123, new Opc.Da.WriteCompleteEventHandler(WriteCompleteCallback), out OPCRequest);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void SystemInSetupMode_OPC()
        {
            try
            {
                Opc.Da.Item[] OPC_SetupMode = new Opc.Da.Item[2];
                OPC_SetupMode[0] = new Opc.Da.Item();
                OPC_SetupMode[0].ItemName = SpotWeld_TagID + "SYSTEM_IN_RUN_MODE";
                OPC_SetupMode[1] = new Opc.Da.Item();
                OPC_SetupMode[1].ItemName = SpotWeld_TagID + "RUN_TEACH_MODE_TOGGLE_BIT";
                OPC_SetupMode = SetupMode_Write.AddItems(OPC_SetupMode);
                
                Opc.Da.ItemValue[] OPC_SetupModeValue = new Opc.Da.ItemValue[2];
                OPC_SetupModeValue[0] = new Opc.Da.ItemValue();
                OPC_SetupModeValue[0].ServerHandle = SetupMode_Write.Items[0].ServerHandle;
                OPC_SetupModeValue[0].Value = 0;
                OPC_SetupModeValue[1] = new Opc.Da.ItemValue();
                OPC_SetupModeValue[1].ServerHandle = SetupMode_Write.Items[1].ServerHandle;
                OPC_SetupModeValue[1].Value = 1;

                Opc.IRequest OPCRequest;
                SetupMode_Write.Write(OPC_SetupModeValue, 123, new Opc.Da.WriteCompleteEventHandler(WriteCompleteCallback), out OPCRequest);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        // 
        private void PartsCompleted_OPC(object sender, EventArgs e)
        {
            try
            {
                List<Item> OPC_PartsCompleteRead = new List<Item>();
                Opc.Da.Item[] OPC_PartsCompleteValue = new Opc.Da.Item[7];
                OPC_PartsCompleteValue[0] = new Opc.Da.Item();
                OPC_PartsCompleteValue[0].ItemName = SpotWeld_TagID + "JOB_ORDER_COUNTER.ACC";
                OPC_PartsCompleteRead.Add(OPC_PartsCompleteValue[0]);
                OPC_PartsCompleteValue[1] = new Opc.Da.Item();
                OPC_PartsCompleteValue[1].ItemName = SpotWeld_TagID + "PART_COMPLETE_COUNTER.ACC";
                OPC_PartsCompleteRead.Add(OPC_PartsCompleteValue[1]);
                OPC_PartsCompleteValue[2] = new Opc.Da.Item();
                OPC_PartsCompleteValue[2].ItemName = SpotWeld_TagID + "PART_COMPLETE_COUNTER.PRE";
                OPC_PartsCompleteRead.Add(OPC_PartsCompleteValue[2]);
                OPC_PartsCompleteValue[3] = new Opc.Da.Item();
                OPC_PartsCompleteValue[3].ItemName = SpotWeld_TagID + "HMI_PART_COMPLETE";
                OPC_PartsCompleteRead.Add(OPC_PartsCompleteValue[3]);
                OPC_PartsCompleteValue[4] = new Opc.Da.Item();
                OPC_PartsCompleteValue[4].ItemName = SpotWeld_TagID + "SETUP_MODE_SET";
                OPC_PartsCompleteRead.Add(OPC_PartsCompleteValue[4]);
                OPC_PartsCompleteValue[5] = new Opc.Da.Item();
                OPC_PartsCompleteValue[5].ItemName = SpotWeld_TagID + "Fault";
                OPC_PartsCompleteRead.Add(OPC_PartsCompleteValue[5]);
                OPC_PartsCompleteValue[6] = new Opc.Da.Item();
                OPC_PartsCompleteValue[6].ItemName = SpotWeld_TagID + "Item_Number_Compare_Value";
                OPC_PartsCompleteRead.Add(OPC_PartsCompleteValue[6]);
                PartComplete_GroupRead.AddItems(OPC_PartsCompleteRead.ToArray());
                
                Opc.IRequest ReadRequest;
                PartComplete_GroupRead.Read(PartComplete_GroupRead.Items, 123, new Opc.Da.ReadCompleteEventHandler(ReadCompleteCallback), out ReadRequest);
            }
            catch (Exception ex)
            {
                // OPC_Timer.Enabled = false; no longer in use
                Timer.Enabled = false;
                SetupMode_Button_Click(null, null);
                MessageBox.Show("Please Select End Job");
            }
        }

        // 
        void PartCompleted_GroupRead_DataChanged(object subscriptionHandle, object requestHandle, ItemValueResult[] values)
        {
            if(PartsComplete_Activate == true)
            {
                // CAT Spot Welders
                if (System.Environment.MachineName == "123R") // CAT - 123R
                {
                    foreach (ItemValueResult itemValue in values)
                    {
                        switch (itemValue.ItemName)
                        {
                            case "OHN66OPC.Spot_Weld_123R.Global.JOB_ORDER_COUNTER.ACC":
                                Job_Order_Counter_ACC = Convert.ToString(itemValue.Value);
                                break;

                            case "OHN66OPC.Spot_Weld_123R.Global.PART_COMPLETE_COUNTER.ACC":
                                Part_Complete_Counter_ACC = Convert.ToString(itemValue.Value);
                                break;

                            case "OHN66OPC.Spot_Weld_123R.Global.PART_COMPLETE_COUNTER.PRE":
                                Part_Complete_Counter_PRE = Convert.ToString(itemValue.Value);
                                break;

                            case "OHN66OPC.Spot_Weld_123R.Global.HMI_PART_COMPLETE":
                                HMI_Part_Complete_VALUE = Convert.ToString(itemValue.Value);
                                break;

                            case "OHN66OPC.Spot_Weld_123R.Global.SETUP_MODE_SET":
                                Setup_Mode_Set_VALUE = Convert.ToString(itemValue.Value);
                                break;

                            case "OHN66OPC.Spot_Weld_123R.Global.Fault":
                                Fault_Value = Convert.ToString(itemValue.Value);
                                break;

                            case "OHN66OPC.Spot_Weld_123R.Global.Item_Number_Compare_Value":
                                Item_Number_PLC = Convert.ToString(itemValue.Value);
                                break;
                        }
                    }
                }
                else if (System.Environment.MachineName == "1088") // CAT - 1088
                {
                    foreach (ItemValueResult itemValue in values)
                    {
                        switch (itemValue.ItemName)
                        {
                            case "OHN66OPC.Spot_Weld_1088.Global.JOB_ORDER_COUNTER.ACC":
                                Job_Order_Counter_ACC = Convert.ToString(itemValue.Value);
                                break;

                            case "OHN66OPC.Spot_Weld_1088.Global.PART_COMPLETE_COUNTER.ACC":
                                Part_Complete_Counter_ACC = Convert.ToString(itemValue.Value);
                                break;

                            case "OHN66OPC.Spot_Weld_1088.Global.PART_COMPLETE_COUNTER.PRE":
                                Part_Complete_Counter_PRE = Convert.ToString(itemValue.Value);
                                break;

                            case "OHN66OPC.Spot_Weld_1088.Global.HMI_PART_COMPLETE":
                                HMI_Part_Complete_VALUE = Convert.ToString(itemValue.Value);
                                break;

                            case "OHN66OPC.Spot_Weld_1088.Global.SETUP_MODE_SET":
                                Setup_Mode_Set_VALUE = Convert.ToString(itemValue.Value);
                                break;

                            case "OHN66OPC.Spot_Weld_1088.Global.Fault":
                                Fault_Value = Convert.ToString(itemValue.Value);
                                break;

                            case "OHN66OPC.Spot_Weld_1088.Global.Item_Number_Compare_Value":
                                Item_Number_PLC = Convert.ToString(itemValue.Value);
                                break;
                        }
                    }
                }
                // John Deere Spot Welders
                else if (System.Environment.MachineName == "108R") // John Deere - 108R
                {
                    foreach (ItemValueResult itemValue in values)
                    {
                        switch (itemValue.ItemName)
                        {
                            case "OHN66OPC.Spot_Weld_108R.Global.JOB_ORDER_COUNTER.ACC":
                                Job_Order_Counter_ACC = Convert.ToString(itemValue.Value);
                                break;

                            case "OHN66OPC.Spot_Weld_108R.Global.PART_COMPLETE_COUNTER.ACC":
                                Part_Complete_Counter_ACC = Convert.ToString(itemValue.Value);
                                break;

                            case "OHN66OPC.Spot_Weld_108R.Global.PART_COMPLETE_COUNTER.PRE":
                                Part_Complete_Counter_PRE = Convert.ToString(itemValue.Value);
                                break;

                            case "OHN66OPC.Spot_Weld_108R.Global.HMI_PART_COMPLETE":
                                HMI_Part_Complete_VALUE = Convert.ToString(itemValue.Value);
                                break;

                            case "OHN66OPC.Spot_Weld_108R.Global.SETUP_MODE_SET":
                                Setup_Mode_Set_VALUE = Convert.ToString(itemValue.Value);
                                break;

                            case "OHN66OPC.Spot_Weld_108R.Global.Fault":
                                Fault_Value = Convert.ToString(itemValue.Value);
                                break;

                            case "OHN66OPC.Spot_Weld_108R.Global.Item_Number_Compare_Value":
                                Item_Number_PLC = Convert.ToString(itemValue.Value);
                                break;
                        }
                    }
                }
                else if (System.Environment.MachineName == "150R") // John Deere - 150R
                {
                    foreach (ItemValueResult itemValue in values)
                    {
                        switch (itemValue.ItemName)
                        {
                            case "OHN66OPC.Spot_Weld_150R.Global.JOB_ORDER_COUNTER.ACC":
                                Job_Order_Counter_ACC = Convert.ToString(itemValue.Value);
                                break;

                            case "OHN66OPC.Spot_Weld_150R.Global.PART_COMPLETE_COUNTER.ACC":
                                Part_Complete_Counter_ACC = Convert.ToString(itemValue.Value);
                                break;

                            case "OHN66OPC.Spot_Weld_150R.Global.PART_COMPLETE_COUNTER.PRE":
                                Part_Complete_Counter_PRE = Convert.ToString(itemValue.Value);
                                break;

                            case "OHN66OPC.Spot_Weld_150R.Global.HMI_PART_COMPLETE":
                                HMI_Part_Complete_VALUE = Convert.ToString(itemValue.Value);
                                break;

                            case "OHN66OPC.Spot_Weld_150R.Global.SETUP_MODE_SET":
                                Setup_Mode_Set_VALUE = Convert.ToString(itemValue.Value);
                                break;

                            case "OHN66OPC.Spot_Weld_150R.Global.Fault":
                                Fault_Value = Convert.ToString(itemValue.Value);
                                break;

                            case "OHN66OPC.Spot_Weld_150R.Global.Item_Number_Compare_Value":
                                Item_Number_PLC = Convert.ToString(itemValue.Value);
                                break;
                        }
                    }
                }
                // Navistar Spot Welders
                else if (System.Environment.MachineName == "OHN7149") // Navistar - 121R
                {
                    foreach (ItemValueResult itemValue in values)
                    {
                        switch (itemValue.ItemName)
                        {
                            case "OHN66OPC.Spot_Weld_121R.Global.SW121R_JOB_ORDER_COUNTER.ACC":
                                Job_Order_Counter_ACC = Convert.ToString(itemValue.Value);
                                break;

                            case "OHN66OPC.Spot_Weld_121R.Global.SW121R_PART_COMPLETE_COUNTER.ACC":
                                Part_Complete_Counter_ACC = Convert.ToString(itemValue.Value);
                                break;

                            case "OHN66OPC.Spot_Weld_121R.Global.SW121R_PART_COMPLETE_COUNTER.PRE":
                                Part_Complete_Counter_PRE = Convert.ToString(itemValue.Value);
                                break;

                            case "OHN66OPC.Spot_Weld_121R.Global.SW121R_HMI_PART_COMPLETE":
                                HMI_Part_Complete_VALUE = Convert.ToString(itemValue.Value);
                                break;

                            case "OHN66OPC.Spot_Weld_121R.Global.SW121R_SETUP_MODE_SET":
                                Setup_Mode_Set_VALUE = Convert.ToString(itemValue.Value);
                                break;

                            case "OHN66OPC.Spot_Weld_121R.Global.SW121R_Fault":
                                Fault_Value = Convert.ToString(itemValue.Value);
                                break;

                            case "OHN66OPC.Spot_Weld_121R.Global.SW121R_Item_Number_Compare_Value":
                                Item_Number_PLC = Convert.ToString(itemValue.Value);
                                break;
                        }
                    }
                }
                else if (System.Environment.MachineName == "OHN7111") // Navistar - 154R
                {
                    foreach (ItemValueResult itemValue in values) // 154R
                    {
                        switch (itemValue.ItemName)
                        {
                            case "OHN66OPC.Spot_Weld_154R.Global.JOB_ORDER_COUNTER.ACC":
                                Job_Order_Counter_ACC = Convert.ToString(itemValue.Value);
                                break;

                            case "OHN66OPC.Spot_Weld_154R.Global.PART_COMPLETE_COUNTER.ACC":
                                Part_Complete_Counter_ACC = Convert.ToString(itemValue.Value);
                                break;

                            case "OHN66OPC.Spot_Weld_154R.Global.PART_COMPLETE_COUNTER.PRE":
                                Part_Complete_Counter_PRE = Convert.ToString(itemValue.Value);
                                break;

                            case "OHN66OPC.Spot_Weld_154R.Global.HMI_PART_COMPLETE":
                                HMI_Part_Complete_VALUE = Convert.ToString(itemValue.Value);
                                break;

                            case "OHN66OPC.Spot_Weld_154R.Global.SETUP_MODE_SET":
                                Setup_Mode_Set_VALUE = Convert.ToString(itemValue.Value);
                                break;

                            case "OHN66OPC.Spot_Weld_154R.Global.Fault":
                                Fault_Value = Convert.ToString(itemValue.Value);
                                break;

                            case "OHN66OPC.Spot_Weld_154R.Global.Item_Number_Compare_Value":
                                Item_Number_PLC = Convert.ToString(itemValue.Value);
                                break;
                        }
                    }
                }
                // Paccar Spot Welders
                else if (System.Environment.MachineName == "OHN7124") // Paccar - 153R
                {
                    foreach (ItemValueResult itemValue in values) // 153R
                    {
                        switch (itemValue.ItemName)
                        {
                            case "OHN66OPC.Spot_Weld_153R.Global.JOB_ORDER_COUNTER.ACC":
                                Job_Order_Counter_ACC = Convert.ToString(itemValue.Value);
                                break;

                            case "OHN66OPC.Spot_Weld_153R.Global.PART_COMPLETE_COUNTER.ACC":
                                Part_Complete_Counter_ACC = Convert.ToString(itemValue.Value);
                                break;

                            case "OHN66OPC.Spot_Weld_153R.Global.PART_COMPLETE_COUNTER.PRE":
                                Part_Complete_Counter_PRE = Convert.ToString(itemValue.Value);
                                break;

                            case "OHN66OPC.Spot_Weld_153R.Global.HMI_PART_COMPLETE":
                                HMI_Part_Complete_VALUE = Convert.ToString(itemValue.Value);
                                break;

                            case "OHN66OPC.Spot_Weld_153R.Global.SETUP_MODE_SET":
                                Setup_Mode_Set_VALUE = Convert.ToString(itemValue.Value);
                                break;

                            case "OHN66OPC.Spot_Weld_153R.Global.Fault":
                                Fault_Value = Convert.ToString(itemValue.Value);
                                break;

                            case "OHN66OPC.Spot_Weld_153R.Global.Item_Number_Compare_Value":
                                Item_Number_PLC = Convert.ToString(itemValue.Value);
                                break;
                        }
                    }
                }
                else if (System.Environment.MachineName == "OHN7123") // Paccar - 155R
                {
                    foreach (ItemValueResult itemValue in values)
                    {
                        switch (itemValue.ItemName)
                        {
                            case "OHN66OPC.Spot_Weld_155R.Global.JOB_ORDER_COUNTER.ACC":
                                Job_Order_Counter_ACC = Convert.ToString(itemValue.Value);
                                break;

                            case "OHN66OPC.Spot_Weld_155R.Global.PART_COMPLETE_COUNTER.ACC":
                                Part_Complete_Counter_ACC = Convert.ToString(itemValue.Value);
                                break;

                            case "OHN66OPC.Spot_Weld_155R.Global.PART_COMPLETE_COUNTER.PRE":
                                Part_Complete_Counter_PRE = Convert.ToString(itemValue.Value);
                                break;

                            case "OHN66OPC.Spot_Weld_155R.Global.HMI_PART_COMPLETE":
                                HMI_Part_Complete_VALUE = Convert.ToString(itemValue.Value);
                                break;

                            case "OHN66OPC.Spot_Weld_155R.Global.SETUP_MODE_SET":
                                Setup_Mode_Set_VALUE = Convert.ToString(itemValue.Value);
                                break;

                            case "OHN66OPC.Spot_Weld_155R.Global.Fault":
                                Fault_Value = Convert.ToString(itemValue.Value);
                                break;

                            case "OHN66OPC.Spot_Weld_155R.Global.Item_Number_Compare_Value":
                                Item_Number_PLC = Convert.ToString(itemValue.Value);
                                break;
                        }
                    }
                }
                // My Computer For Testing
                else if (System.Environment.MachineName == "OHN7047NL") // My Computer For Testing
                {
                    foreach (ItemValueResult itemValue in values)
                    {
                        switch (itemValue.ItemName)
                        {
                            case "OHN66OPC.Spot_Weld_121R.Global.SW121R_JOB_ORDER_COUNTER.ACC":
                                Job_Order_Counter_ACC = Convert.ToString(itemValue.Value);
                                break;

                            case "OHN66OPC.Spot_Weld_121R.Global.SW121R_PART_COMPLETE_COUNTER.ACC":
                                Part_Complete_Counter_ACC = Convert.ToString(itemValue.Value);
                                break;

                            case "OHN66OPC.Spot_Weld_121R.Global.SW121R_PART_COMPLETE_COUNTER.PRE":
                                Part_Complete_Counter_PRE = Convert.ToString(itemValue.Value);
                                break;

                            case "OHN66OPC.Spot_Weld_121R.Global.SW121R_HMI_PART_COMPLETE":
                                HMI_Part_Complete_VALUE = Convert.ToString(itemValue.Value);
                                break;

                            case "OHN66OPC.Spot_Weld_121R.Global.SW121R_SETUP_MODE_SET":
                                Setup_Mode_Set_VALUE = Convert.ToString(itemValue.Value);
                                break;

                            case "OHN66OPC.Spot_Weld_121R.Global.SW121R_Fault":
                                Fault_Value = Convert.ToString(itemValue.Value);
                                break;

                            case "OHN66OPC.Spot_Weld_121R.Global.SW121R_Item_Number_Compare_Value":
                                Item_Number_PLC = Convert.ToString(itemValue.Value);
                                break;
                        }
                    }
                }
                else // Default 
                {
                    foreach (ItemValueResult itemValue in values)
                    {
                        switch (itemValue.ItemName)
                        {
                            case "OHN66OPC.Spot_Weld_121R.Global.SW121R_JOB_ORDER_COUNTER.ACC":
                                Job_Order_Counter_ACC = Convert.ToString(itemValue.Value);
                                break;

                            case "OHN66OPC.Spot_Weld_121R.Global.SW121R_PART_COMPLETE_COUNTER.ACC":
                                Part_Complete_Counter_ACC = Convert.ToString(itemValue.Value);
                                break;

                            case "OHN66OPC.Spot_Weld_121R.Global.SW121R_PART_COMPLETE_COUNTER.PRE":
                                Part_Complete_Counter_PRE = Convert.ToString(itemValue.Value);
                                break;

                            case "OHN66OPC.Spot_Weld_121R.Global.SW121R_HMI_PART_COMPLETE":
                                HMI_Part_Complete_VALUE = Convert.ToString(itemValue.Value);
                                break;

                            case "OHN66OPC.Spot_Weld_121R.Global.SW121R_SETUP_MODE_SET":
                                Setup_Mode_Set_VALUE = Convert.ToString(itemValue.Value);
                                break;

                            case "OHN66OPC.Spot_Weld_121R.Global.SW121R_Fault":
                                Fault_Value = Convert.ToString(itemValue.Value);
                                break;

                            case "OHN66OPC.Spot_Weld_121R.Global.SW121R_Item_Number_Compare_Value":
                                Item_Number_PLC = Convert.ToString(itemValue.Value);
                                break;
                        }
                    }
                }
                PartsFormed_TextBox.Invoke(new EventHandler(delegate { PartsFormed_TextBox.Text = Job_Order_Counter_ACC; }));
                CurrentWeld_TextBox.Invoke(new EventHandler(delegate { CurrentWeld_TextBox.Text = Part_Complete_Counter_ACC; }));
                TotalWeldsNeeded_TextBox.Invoke(new EventHandler(delegate { TotalWeldsNeeded_TextBox.Text = Part_Complete_Counter_PRE; }));
                HMIPartComplete_TextBox.Invoke(new EventHandler(delegate { HMIPartComplete_TextBox.Text = HMI_Part_Complete_VALUE; }));
                SetupModeSet_TextBox.Invoke(new EventHandler(delegate { SetupModeSet_TextBox.Text = Setup_Mode_Set_VALUE; }));
                Fault_TextBox.Invoke(new EventHandler(delegate { Fault_TextBox.Text = Fault_Value; }));
                JobID_TextBox.Invoke(new EventHandler(delegate { JobID_TextBox.Text = Item_Number_PLC; }));

                if (HMIPartComplete_TextBox.Text == "True")
                {
                    PartCompleted_TextBox.Invoke(new EventHandler(delegate { PartCompleted_TextBox.Visible = true; }));
                    PartCompleted_TextBox.Invoke(new EventHandler(delegate { PartCompleted_TextBox.BringToFront(); }));
                    //PartCompleted_TextBox.Visible = true;
                    //PartCompleted_TextBox.BringToFront();
                }
                else if (HMIPartComplete_TextBox.Text == "False")
                {
                    PartCompleted_TextBox.Invoke(new EventHandler(delegate { PartCompleted_TextBox.Visible = false; }));
                    PartCompleted_TextBox.Invoke(new EventHandler(delegate { PartCompleted_TextBox.SendToBack(); }));
                    //PartCompleted_TextBox.Visible = false;
                    //PartCompleted_TextBox.SendToBack();
                }
                if (SetupModeSet_TextBox.Text == "True")
                {
                    SetupMode_Button.Invoke(new EventHandler(delegate { SetupMode_Button.PerformClick(); }));
                    //SetupMode_Button_Click(null, null);
                }
                if (Fault_TextBox.Text == "1")
                {
                    /*
                    FaultMessage_TextBox.Invoke(new EventHandler(delegate {
                        FaultMessage_TextBox.Location = new System.Drawing.Point(444, 967);
                        FaultMessage_TextBox.Size = new System.Drawing.Size(1880, 450); }));
                        */
                    UserProgram.Invoke(new EventHandler(delegate { Update_UI_OPC(); }));
                }
            }
            

        }

        private void PartsCompleted_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
        }

        private void ReadCompleteCallback(object clientHandle, Opc.Da.ItemValueResult[] results)
        {
            /*
            PartsFormed_TextBox.Invoke(new EventHandler(delegate { PartsFormed_TextBox.Text = (results[0].Value).ToString(); }));
            CurrentWeld_TextBox.Invoke(new EventHandler(delegate { CurrentWeld_TextBox.Text = (results[1].Value).ToString(); }));
            TotalWeldsNeeded_TextBox.Invoke(new EventHandler(delegate { TotalWeldsNeeded_TextBox.Text = (results[2].Value).ToString(); }));
            HMIPartComplete_TextBox.Invoke(new EventHandler(delegate { HMIPartComplete_TextBox.Text = (results[3].Value).ToString(); }));
            SetupModeSet_TextBox.Invoke(new EventHandler(delegate { SetupModeSet_TextBox.Text = (results[4].Value).ToString(); }));
            Fault_TextBox.Invoke(new EventHandler(delegate { Fault_TextBox.Text = (results[5].Value).ToString(); }));
            JobID_TextBox.Invoke(new EventHandler(delegate { JobID_TextBox.Text = (results[6].Value).ToString(); }));
            */
        }

        private void ReadPartNotProgrammedCallback(object clientHandle, Opc.Da.ItemValueResult[] results)
        {
            PartNotProgrammed_TextBox.Invoke(new EventHandler(delegate { PartNotProgrammed_TextBox.Text = (results[0].Value).ToString(); }));
        }

        private void WriteCompleteCallback(object clientHandle, Opc.IdentifiedResult[] results)
        {
            foreach (Opc.IdentifiedResult writeResult in results)
            {
                Console.WriteLine("\t{0} write result: {1}", writeResult.ItemName, writeResult.ResultID);
            }
        }

        private void Update_UI_OPC()
        {/*
            PartsFormed_TextBox.Text = Job_Order_Counter_ACC;
            CurrentWeld_TextBox.Text = Part_Complete_Counter_ACC;
            TotalWeldsNeeded_TextBox.Text = Part_Complete_Counter_PRE;
            HMIPartComplete_TextBox.Text = HMI_Part_Complete_VALUE;
            SetupModeSet_TextBox.Text = Setup_Mode_Set_VALUE;
            Fault_TextBox.Text = Fault_Value;
            JobID_TextBox.Text = Item_Number_PLC;

            if (HMIPartComplete_TextBox.Text == "True")
            {
                PartCompleted_TextBox.Visible = true;
                PartCompleted_TextBox.BringToFront();
            }
            else if (HMIPartComplete_TextBox.Text == "False")
            {
                PartCompleted_TextBox.Visible = false;
                PartCompleted_TextBox.SendToBack();
            }
            if (SetupModeSet_TextBox.Text == "True")
            {
                SetupMode_Button_Click(null, null);
            }
            if (Fault_TextBox.Text == "1")
            {
                OPC_Timer.Stop();
                User_Program_Part_Not_Completed NotCompleted = new User_Program_Part_Not_Completed();
                UserProgram.Enabled = false;
                if (NotCompleted.ShowDialog(this) == DialogResult.Yes)
                {
                    Fault_Value = "0";
                    OPC_Timer.Start();
                }
            }
            */
            User_Program_Part_Not_Completed NotCompleted = new User_Program_Part_Not_Completed();
            if (NotCompleted.ShowDialog(this) == DialogResult.Yes)
            {
                Fault_Value = "0";
            }
        }

        /*
        private void SetupModeTimeOut_OPC()
        {
            if (SetupModeSet_TextBox.Text == "True")
            {
                SetupMode_Button_Click(null, null);
            }
        }
        */

        private void HMIPartComplete_OPC()
        {
            if (HMIPartComplete_TextBox.Text == "True")
            {
                PartCompleted_TextBox.Visible = true;
                PartCompleted_TextBox.BringToFront();
            }
            else if (HMIPartComplete_TextBox.Text == "False")
            {
                PartCompleted_TextBox.Visible = false;
                PartCompleted_TextBox.SendToBack();
            }
        }

        /********************************************************************************************************************
        *  
        *  OPC Region End
        *  
        ********************************************************************************************************************/
        #endregion

        /*********************************************************************************************************************
        * 
        * Methods Region Start 
        * -- Total: 15
        * 
        * - SearchForItemID
        * - SearchItemIDOperation
        * - ShowComponents
        * - OperationInitialize
        * - OperationOEECalculation
        * - FindTotalRunTime
        * - ItemOperationCalculation
        * - RunningStatistics
        * - LivePPMOEECalculation
        * - PDFFileCheck
        * - SpotWeldID
        * - ClearForm
        * - PassOperationValue
        * - PassReferenceNumber
        * - PassValue
        * 
        **********************************************************************************************************************/
        #region

        private void SearchForItemID()
        {
            string SearchValue = CurrentItemID;
            string OperationValue = CurrentItemIDOperation.ToString();
            UserProgramGridView.SelectionMode = DataGridViewSelectionMode.FullRowSelect;

            try
            {
                foreach (DataGridViewRow Row in UserProgramGridView.Rows)
                {
                    Row.Selected = false;
                    if (Row.Cells[0].Value.ToString().Equals(SearchValue) && Row.Cells[4].Value.ToString().Equals(OperationValue))
                    {
                        Row.Selected = true;
                        ItemID_TextBox.Text = Row.Cells[0].Value.ToString();
                        Customer = Row.Cells[1].Value.ToString();
                        CustomerPartNumber = Row.Cells[2].Value.ToString();
                        //JobID_TextBox.Text = Row.Cells[3].Value.ToString();
                        Sequence_TextBox.Text = Row.Cells[4].Value.ToString();
                        Fixture_TextBox.Text = Row.Cells[5].Value.ToString();
                        FixtureLocation_TextBox.Text = Row.Cells[6].Value.ToString();
                        Comp1_TextBox.Text = Row.Cells[7].Value.ToString();
                        Component_1 = Row.Cells[7].Value.ToString();
                        Quantity1_TextBox.Text = Row.Cells[8].Value.ToString();
                        Comp2_TextBox.Text = Row.Cells[9].Value.ToString();
                        Component_2 = Row.Cells[9].Value.ToString();
                        Quantity2_TextBox.Text = Row.Cells[10].Value.ToString();
                        Comp3_TextBox.Text = Row.Cells[11].Value.ToString();
                        Component_3 = Row.Cells[11].Value.ToString();
                        Quantity3_TextBox.Text = Row.Cells[12].Value.ToString();
                        Comp4_TextBox.Text = Row.Cells[13].Value.ToString();
                        Component_4 = Row.Cells[13].Value.ToString();
                        Quantity4_TextBox.Text = Row.Cells[14].Value.ToString();
                        //TotalRuns_TextBox.Text = Row.Cells[15].Value.ToString();
                        PartsManufacturedTotal_String = Row.Cells[16].Value.ToString();
                        PastItemPPM_String = Row.Cells[17].Value.ToString();
                        AveragePPM_String = Row.Cells[17].Value.ToString();
                        UserProgramGridView.FirstDisplayedScrollingRowIndex = UserProgramGridView.SelectedRows[0].Index;
                        JobFound = true;
                        break;
                    }
                }
            }
            catch (Exception)
            {
                User_Program_Job_Not_Programmed JobNotProgrammed = new User_Program_Job_Not_Programmed();
                JobNotProgrammed.Show();
            }
        }

        private void SearchItemIDOperation()
        {
            string SearchValue = CurrentItemID;
            int OperationRows = 0;
            if (Company_ComboBox.Text == "CAT")
            {
                Spot_Weld_Operation_Selection = "SELECT COUNT(*) FROM [dbo].[CAT_Item_Data] WHERE ItemID = '" + SearchValue + "'";
            }
            else if (Company_ComboBox.Text == "John Deere")
            {
                Spot_Weld_Operation_Selection = "SELECT COUNT(*) FROM [dbo].[JohnDeere_Item_Data] WHERE ItemID = '" + SearchValue + "'";
            }
            else if (Company_ComboBox.Text == "Navistar")
            {
                Spot_Weld_Operation_Selection = "SELECT COUNT(*) FROM [dbo].[Navistar_Item_Data] WHERE ItemID = '" + SearchValue + "'";
            }
            else if (Company_ComboBox.Text == "Paccar")
            {
                Spot_Weld_Operation_Selection = "SELECT COUNT(*) FROM [dbo].[Paccar_Item_Data] WHERE ItemID = '" + SearchValue + "'";
            }
            string OperationCount = Spot_Weld_Operation_Selection;
            SqlConnection Count = new SqlConnection(SQL_Source);
            SqlCommand CountRows = new SqlCommand(OperationCount, Count);
            Count.Open();
            OperationRows = (int)CountRows.ExecuteScalar();
            OperationRows_NA = (int)CountRows.ExecuteScalar();
            Count.Close();

            if (OperationRows >= 2)
            {
                User_Program_Select_Operation OperationSelect = new User_Program_Select_Operation(this);
                if (OperationRows == 2)
                {
                    OperationSelect.Operation_1_Button.Location = new System.Drawing.Point(13, 16);
                    OperationSelect.Operation_2_Button.Location = new System.Drawing.Point(332, 16);
                    OperationSelect.Operation_3_Button.Hide();
                    OperationSelect.Operation_4_Button.Hide();
                    OperationSelect.ClientSize = new System.Drawing.Size(625, 205);
                }
                else if (OperationRows == 3)
                {
                    OperationSelect.Operation_1_Button.Location = new System.Drawing.Point(13, 16);
                    OperationSelect.Operation_2_Button.Location = new System.Drawing.Point(332, 16);
                    OperationSelect.Operation_3_Button.Location = new System.Drawing.Point(651, 16);
                    OperationSelect.Operation_4_Button.Hide();
                    OperationSelect.ClientSize = new System.Drawing.Size(940, 205);
                }
                if (OperationSelect.ShowDialog(this) == DialogResult.Yes)
                {
                    SearchForItemID();
                }
            }
            else
            {
                OperationSelect_OPC();
                CurrentItemIDOperation = 1;
                SearchForItemID();
            }
        }
        
        private void ShowComponents()
        {
            if (Comp1_TextBox.TextLength == 9)
            {
                Comp1_Label.Show();
                Comp1_TextBox.Show();
                Quantity1_Label.Show();
                Quantity1_TextBox.Show();
            }
            if (Comp2_TextBox.TextLength == 9)
            {
                Comp2_Label.Show();
                Comp2_TextBox.Show();
                Quantity2_Label.Show();
                Quantity2_TextBox.Show();
            }
            if (Comp3_TextBox.TextLength == 9)
            {
                Comp3_Label.Show();
                Comp3_TextBox.Show();
                Quantity3_Label.Show();
                Quantity3_TextBox.Show();
            }
        }

        public void OperationInitialize()
        {
            //OPC_Timer.Enabled = true;
            StartNewJob_Button.Enabled = false;
            RunMode_Button.Enabled = true;
            SetupMode_Button.Enabled = true;
            CancelRun_Button.Enabled = true;
            CancelRun_Button.Visible = true;
            PartsRunProgressBar.Visible = true;
            JobStartTime = Clock_TextBox.Text;
            RunModeStopWatch.Reset();
            SetupModeStopWatch.Reset();
            TimeOfOperation.Equals(0);
            PDFFileCheck();
            ShowComponents();
            ItemRunCounter_SQL();
            OperationIDCounter_SQL();
            ItemOperationDataStart_SQL();
            OperationDataStart_SQL();
            //Timer.Enabled = true;            
            string StartingTime = Clock_TextBox.Text;
            string ReplaceTime = DateTime.Today.ToShortDateString();
            JobStartTime_TextBox.Text = StartingTime.Replace("   " + ReplaceTime, "");
            //PartsFormed_TextBox.Text = 0.ToString();
            //SetupMode_Button_Click(null, null);
        }

        private void OperationOEECalculation()
        {
            /*********************************************************************************************************************
            *                                         OEE = Efficiency * Utilization * Quality                                   *   
            **********************************************************************************************************************                                         
            *                                       |                                         |                                  |     
            *                       Planned         |                                         |                                  |   
            *                   Operation Time      |                    Operation Time       |                Good Parts        |   
            *    Efficiency = ------------------    |    Utilization = -------------------    |    Quality = ---------------     |                 
            *                   Operation Time      |                    Available Hours      |                  Total           |  
            *                                       |                                         |               Formed Parts       |  
            *                                       |                                         |                                  |                                    
            *_______________________________________|_________________________________________|__________________________________|
            *                                       |                                         |                                  |
            *                                       |                                         |                                  |
            *                   Planned Minutes     |                   Operation Time        |                Good Parts        |
            * Efficiency = ------------------------ | Utilization = ------------------------  | Quality = ---------------------  |
            *                   Actual Minutes      |                   Available Hours       |            Total Formed Parts    |     
            *                                       |                                         |                                  |
            *                                       |                                         |                                  |
            **********************************************************************************************************************/

            string PartsOnOrder = PartsNeeded_TextBox.Text;
            double DoublePartsOnOrder = double.Parse(PartsOnOrder);
            double RunMinutes = TimeOfOperation.TotalMinutes;
            double PlannedTime;
            if (PastPPM == 0)
            {
                PlannedTime = RunMinutes;
            }
            else
            {
                string PFTest_String = PartsFormed_TextBox.Text;
                double PFTest = double.Parse(PFTest_String);
                //PlannedTime = (DoublePartsOnOrder / PastPPM);
                PlannedTime = (PFTest / PastPPM);
            }
            try
            {
                double PlannedMinutes = PlannedTime / 60;
                double ActualMinutes = RunMinutes / 60;
                Efficiency = (float)(PlannedMinutes / ActualMinutes);
                Efficiency = Efficiency * 100;
                Efficiency = (float)Math.Round(Efficiency, 2);
                TimePlanned = (float)PlannedMinutes;
                TimePlanned = (float)Math.Round(TimePlanned, 5);
                TimeActual = (float)ActualMinutes;
                TimeActual = (float)Math.Round(TimeActual, 5);
            }
            catch (Exception)
            {
                MessageBox.Show("Unable to Calculate OEE Results");
            }

        }

        private void FindTotalRunTime()
        {
            TimeOfOperation = DateTime.Parse(JobEndTime).Subtract(DateTime.Parse(JobStartTime));
            LiveRunTime_TextBox.Text = TimeOfOperation.ToString();
            TimeSpan SetupMode = SetupModeStopWatch.Elapsed;
            double ItemSetupTime_Double = SetupMode.TotalMinutes;
            ItemSetupTime_Double = ItemSetupTime_Double / 60;
            ItemSetupTime = (float)ItemSetupTime_Double;
            ItemSetupTime = (float)Math.Round(ItemSetupTime, 5);
        }

        // Called From End Job
        private void ItemOperationCalculation()
        {
            /*********************************************************************************************************************
            *                                |                                       |
            *                                |                                       |
            *            Parts Formed        |                                       |                  Parts Formed
            *     PPM = ----------------     |     Parts Formed = Run Time * PPM     |     Run Time = ----------------
            *              Run Time          |                                       |                      PPM
            *                                |                                       |
            *                                |                                       |
            *________________________________|_______________________________________|__________________________________________
            *                                |                                       |
            *                                |                                       |
            *        OverallFormedParts      |      Overall                          |                  OverallFormedParts
            * PPM = --------------------     |       Parts   =  Run Time * PPM       |     Run Time = --------------------
            *         OverallRunTime         |      Formed                           |                     AveragePPM
            *                                |                                       |
            *                                |                                       |
            *********************************************************************************************************************/

            // Convert The Current Parts Formed to double
            string CurrentPartsFormed_String = PartsFormed_TextBox.Text;
            double CurrentPartsFormed_Double = double.Parse(CurrentPartsFormed_String);

            // add the total parts of this item to the current run
            if (PartsManufacturedTotal_String != "")
            {
                PartsManufacturedTotal_Double = double.Parse(PartsManufacturedTotal_String);
            }            
            TotalItemPartsManufactured = CurrentPartsFormed_Double + PartsManufacturedTotal_Double;

            // Current PPM
            string CurrentPPM_String = CurrentPPM_TextBox.Text;
            float CurrentPartsFormed_Float = (float)CurrentPartsFormed_Double;
            float CurrentPPM = float.Parse(CurrentPPM_String);
            CurrentParts = Math.Round(CurrentParts, 2);
            float CurrentRunTime = CurrentPartsFormed_Float / CurrentPPM;

            //Past PPM
            if (AveragePPM_String == "")
            {
                AveragePPM = CurrentPPM;
            }
            else if (AveragePPM_String != "")
            {
                PastPPM = float.Parse(AveragePPM_String);
            }
            if (PastPPM == 0)
            {
                AveragePPM = CurrentPPM;
            }
            else if (PastPPM != 0)
            {
                float PartsManufacturedTotal_Float = (float)PartsManufacturedTotal_Double;
                float PreviousRunTime = PartsManufacturedTotal_Float / PastPPM;

                float OverallRunTime = PreviousRunTime + CurrentRunTime;
                float OverallFormedParts = PartsManufacturedTotal_Float + CurrentPartsFormed_Float;
                AveragePPM = OverallFormedParts / OverallRunTime;
                AveragePPM = (float)Math.Round(AveragePPM, 2);
                if(AveragePPM.ToString() == "NaN")
                {
                    AveragePPM = 0;
                }
            }
        }

        private void RunningStatistics()
        {
            double HoursRemaining = 0;
            double MinutesRemaining = 0;
            double PartsNeeded = double.Parse(PartsNeeded_TextBox.Text);
            //string PPMString = CurrentPPM_TextBox.Text;
            string RemainingTime = "";
            string CurrentParts_String = PartsFormed_TextBox.Text;
            //CurrentParts = double.Parse(PartsFormed_TextBox.Text);
            int CurrentParts_int = Int32.Parse(CurrentParts_String);
            CurrentParts = (CurrentParts_int);
            LivePPMOEECalculation();
            string PPMString = CurrentPPM_TextBox.Text;
            if (PartsNeeded == CurrentParts)
            {
                //SetupMode_Button_Click(null, null);
                PartsNeeded = PartsNeeded + 5;
                PartsNeeded_TextBox.Text = PartsNeeded.ToString();
            }
            else
            {
                PartsRemaining = PartsNeeded - CurrentParts;

                PartsRemaining_TextBox.Text = PartsRemaining.ToString();
            }
            //LivePPMOEECalculation();
            double CurrentPPM = double.Parse(PPMString);
            double TimeRemaining = (PartsRemaining / CurrentPPM);

            if (TimeRemaining < 60)
            {
                MinutesRemaining = TimeRemaining;
                MinutesRemaining = Math.Round(MinutesRemaining, MidpointRounding.AwayFromZero);
                if (MinutesRemaining == 1)
                {
                    RemainingTime = MinutesRemaining + " Minute".ToString();
                }
                else
                {
                    RemainingTime = MinutesRemaining + " Minutes".ToString();
                }
            }
            else if (120 > TimeRemaining && TimeRemaining >= 60)
            {
                TimeRemaining = TimeRemaining - 60;
                HoursRemaining = 1;
                MinutesRemaining = TimeRemaining;
                MinutesRemaining = Math.Round(MinutesRemaining, MidpointRounding.AwayFromZero);
                RemainingTime = HoursRemaining + " Hour " + MinutesRemaining + " Minutes".ToString();

            }
            else if (180 > TimeRemaining && TimeRemaining >= 120)
            {
                TimeRemaining = TimeRemaining - 120;
                HoursRemaining = 2;
                MinutesRemaining = TimeRemaining;
                MinutesRemaining = Math.Round(MinutesRemaining, MidpointRounding.AwayFromZero);
                RemainingTime = HoursRemaining + " Hours " + MinutesRemaining + " Minutes".ToString();
            }
            else if (240 > TimeRemaining && TimeRemaining >= 180)
            {
                TimeRemaining = TimeRemaining - 180;
                HoursRemaining = 3;
                MinutesRemaining = TimeRemaining;
                MinutesRemaining = Math.Round(MinutesRemaining, MidpointRounding.AwayFromZero);
                RemainingTime = HoursRemaining + " Hours " + MinutesRemaining + " Minutes".ToString();
            }
            else if (300 > TimeRemaining && TimeRemaining >= 240)
            {
                TimeRemaining = TimeRemaining - 240;
                HoursRemaining = 4;
                MinutesRemaining = TimeRemaining;
                MinutesRemaining = Math.Round(MinutesRemaining, MidpointRounding.AwayFromZero);
                RemainingTime = HoursRemaining + " Hours " + MinutesRemaining + " Minutes".ToString();
            }
            else if (360 > TimeRemaining && TimeRemaining >= 300)
            {
                TimeRemaining = TimeRemaining - 300;
                HoursRemaining = 5;
                MinutesRemaining = TimeRemaining;
                MinutesRemaining = Math.Round(MinutesRemaining, MidpointRounding.AwayFromZero);
                RemainingTime = HoursRemaining + " Hours " + MinutesRemaining + " Minutes".ToString();
            }
            else if (420 > TimeRemaining && TimeRemaining >= 360)
            {
                TimeRemaining = TimeRemaining - 360;
                HoursRemaining = 6;
                MinutesRemaining = TimeRemaining;
                MinutesRemaining = Math.Round(MinutesRemaining, MidpointRounding.AwayFromZero);
                RemainingTime = HoursRemaining + " Hours " + MinutesRemaining + " Minutes".ToString();
            }
            else if (480 > TimeRemaining && TimeRemaining >= 420)
            {
                TimeRemaining = TimeRemaining - 420;
                HoursRemaining = 7;
                MinutesRemaining = TimeRemaining;
                MinutesRemaining = Math.Round(MinutesRemaining, MidpointRounding.AwayFromZero);
                RemainingTime = HoursRemaining + " Hours " + MinutesRemaining + " Minutes".ToString();
            }
            else if (540 > TimeRemaining && TimeRemaining >= 480)
            {
                TimeRemaining = TimeRemaining - 480;
                HoursRemaining = 8;
                MinutesRemaining = TimeRemaining;
                MinutesRemaining = Math.Round(MinutesRemaining, MidpointRounding.AwayFromZero);
                RemainingTime = HoursRemaining + " Hours " + MinutesRemaining + " Minutes".ToString();
            }
            else if (600 > TimeRemaining && TimeRemaining >= 540)
            {
                TimeRemaining = TimeRemaining - 540;
                HoursRemaining = 9;
                MinutesRemaining = TimeRemaining;
                MinutesRemaining = Math.Round(MinutesRemaining, MidpointRounding.AwayFromZero);
                RemainingTime = HoursRemaining + " Hours " + MinutesRemaining + " Minutes".ToString();
            }
            else if (660 > TimeRemaining && TimeRemaining >= 600)
            {
                TimeRemaining = TimeRemaining - 600;
                HoursRemaining = 10;
                MinutesRemaining = TimeRemaining;
                MinutesRemaining = Math.Round(MinutesRemaining, MidpointRounding.AwayFromZero);
                RemainingTime = HoursRemaining + " Hours " + MinutesRemaining + " Minutes".ToString();
            }
            TimeRemaining_TextBox.Text = RemainingTime;
            PartsRunProgressBar.Maximum = int.Parse(PartsNeeded_TextBox.Text);
            PartsRunProgressBar.Value = int.Parse(PartsFormed_TextBox.Text);

        }
            
        private void LivePPMOEECalculation()
        {
            /*
            * LivePPMCalculation
            * 
            * Current Parts Formed / TotalTime in Run Mode
            * 
            */
            PastItemPPM = double.Parse(PastItemPPM_String);
            double TotalElapsedTime = RunModeStopWatch.Elapsed.TotalMinutes;
            double CurrentRunPPM = (Math.Round((CurrentParts / TotalElapsedTime), 2));
            CurrentRunPPM = System.Math.Ceiling(CurrentRunPPM * 100) / 100;
            CurrentPPM_TextBox.Text = CurrentRunPPM.ToString("0.00");
            double LiveOEE = Math.Round(CurrentRunPPM / PastItemPPM, 3);
            LiveOEE = LiveOEE * 100;
            LiveOEE_TextBox.Text = LiveOEE.ToString() + "%";

        }

        private void PDFFileCheck()
        {
            //string CheckCurrentItemID = CurrentItemID.Replace("-", "");
            string CheckCurrentItemID = CurrentItemID;
            string SpotWeldSetupCheck = SpotWeldSetupCardPDFPath + CheckCurrentItemID + ".pdf";
            string SpotWeldPrintCheck = ItemPrintPDFPath + CheckCurrentItemID + "_Print.pdf";
            if (File.Exists(SpotWeldSetupCheck))
            {
                ViewSetupCard_Button.Enabled = true;
            }
            if (File.Exists(SpotWeldPrintCheck))
            {
                ViewPrint_Button.Enabled = true;
            }
        }

        // Get the Computer Name and assign OPC Variables
        private void SpotWeldID()
        {
            string SpotWeldComputerID = System.Environment.MachineName;

            // CAT Spot Weld
            if (SpotWeldComputerID == "123R") // 123R
            {
                //SpotWeld_TagID = "SpotWeld_TagID_123RR:SPOTWELD_123RR:";
                SpotWeld_TagID = "OHN66OPC.Spot_Weld_123R.Global.";
                Company_ComboBox.Text = "CAT";
                Company_TextBox.Text = "CAT";
                SpotWeld_ComboBox.Text = "123R";
                SpotWeld_TextBox.Text = "123R";
            }
            if (SpotWeldComputerID == "1088") // 1088
            {
                //SpotWeld_TagID = "SpotWeld_TagID_1088:SPOTWELD_1088:";
                SpotWeld_TagID = "OHN66OPC.Spot_Weld_1088.Global.";
                Company_ComboBox.Text = "CAT";
                Company_TextBox.Text = "CAT";
                SpotWeld_ComboBox.Text = "1088";
                SpotWeld_TextBox.Text = "1088";
            }  
            // John Deere Spot Weld
            if (SpotWeldComputerID == "108R") // 108R
            {
                //SpotWeld_TagID = "SpotWeld_TagID_108R:SPOTWELD_108R:";
                SpotWeld_TagID = "OHN66OPC.Spot_Weld_108R.Global.";
                Company_ComboBox.Text = "John Deere";
                Company_TextBox.Text = "John Deere";
                SpotWeld_ComboBox.Text = "108R";
                SpotWeld_TextBox.Text = "108R";
            }
            if (SpotWeldComputerID == "150R") // 150R
            {
                //SpotWeld_TagID = "SpotWeld_TagID_150R:SPOTWELD_150R:";
                SpotWeld_TagID = "OHN66OPC.Spot_Weld_150R.Global.";
                Company_ComboBox.Text = "John Deere";
                Company_TextBox.Text = "John Deere";
                SpotWeld_ComboBox.Text = "150R";
                SpotWeld_TextBox.Text = "150R";
            }
            // Navistar Spot Weld
            if (SpotWeldComputerID == "104R") // 104R
            {
                //SpotWeld_TagID = "SpotWeld_TagID_104R:SPOTWELD_104R:";
                SpotWeld_TagID = "OHN66OPC.Spot_Weld_104R.Global.";
                Company_ComboBox.Text = "Navistar";
                Company_TextBox.Text = "Navistar";
                SpotWeld_ComboBox.Text = "104R";
                SpotWeld_TextBox.Text = "104R";
            }
            if (SpotWeldComputerID == "OHN7149") // 121R
            {
                //SpotWeld_TagID = "SpotWeld_TagID_121R:SPOTWELD_121R:";
                SpotWeld_TagID = "OHN66OPC.Spot_Weld_121R.Global.SW121R_";
                Company_ComboBox.Text = "Navistar";
                Company_TextBox.Text = "Navistar";
                SpotWeld_ComboBox.Text = "121R";
                SpotWeld_TextBox.Text = "121R";
            }
            if (SpotWeldComputerID == "OHN7111") // 154R
            {
                //SpotWeld_TagID = "SpotWeld_TagID_155R:SPOTWELD_155R:";
                SpotWeld_TagID = "OHN66OPC.Spot_Weld_154R.Global.";
                Company_ComboBox.Text = "Navistar";
                Company_TextBox.Text = "Navistar";
                SpotWeld_ComboBox.Text = "154R";
                SpotWeld_TextBox.Text = "154R";
            }
            // Paccar Spot Weld
            if (SpotWeldComputerID == "OHN7124") // 153R
            {
                //SpotWeld_TagID = "SpotWeld_TagID:SPOTWELD_153R:";
                SpotWeld_TagID = "OHN66OPC.Spot_Weld_153R.Global.";
                Company_ComboBox.Text = "Paccar";
                Company_TextBox.Text = "Paccar";
                SpotWeld_ComboBox.Text = "153R";
                SpotWeld_TextBox.Text = "153R";
            }
            if (SpotWeldComputerID == "OHN7123") // 155R
            {
                //SpotWeld_TagID = "SpotWeld_TagID_155R:SPOTWELD_155R:";
                SpotWeld_TagID = "OHN66OPC.Spot_Weld_155R.Global.";
                Company_ComboBox.Text = "Paccar";
                Company_TextBox.Text = "Paccar";
                SpotWeld_ComboBox.Text = "155R";
                SpotWeld_TextBox.Text = "155R";
            }
            if (SpotWeldComputerID == "OHN7047NL") // My Laptop
            {
                //SpotWeld_TagID = "SpotWeld_TagID_121R:SPOTWELD_121R:";
                SpotWeld_TagID = "OHN66OPC.Spot_Weld_121R.Global.SW121R_";
                Company_ComboBox.Text = "Navistar";
                Company_TextBox.Text = "Navistar";
                SpotWeld_ComboBox.Text = "121R";
                SpotWeld_TextBox.Text = "121R";
            }

            Company_ComboBox.Enabled = false;
            SpotWeld_ComboBox.Enabled = false;
        }

        // Called When StartNewJob_Button is Clicked
        private void ClearForm()
        {
            // Clear Each TextBox
            //ItemID_TextBox, JobID_TextBox, Sequence_TextBox, ReferenceNumber_TextBox, Fixture_TextBox, FixtureLocation_TextBox 
            foreach (Control TextBox_Clear in ItemData_GroupBox.Controls)
            {
                if(TextBox_Clear is TextBox)
                {
                    TextBox_Clear.Text = "";
                }
            }
            // Comp1_TextBox, Comp2_TextBox,Comp3_TextBox, Comp4_TextBox, Quantity1_TextBox, Quantity2_TextBox, Quantity3_TextBox, Quantity4_TextBox
            foreach (Control TextBox_Clear in Component_GroupBox.Controls)
            {
                if (TextBox_Clear is TextBox)
                {
                    TextBox_Clear.Text = "";
                }
            }
            // PartsNeeded_TextBox, PartsFormed_TextBox, PartsRemaining_TextBox, CurrentPPM_TextBox
            // JobStartTime_TextBox, JobEndTime_TextBox, CurrentItemID_TextBox, CurrentWeld_TextBox, TotalWeldsNeeded_TextBox
            foreach (Control TextBox_Clear in JobData_GroupBox.Controls)
            {
                if (TextBox_Clear is TextBox)
                {
                    TextBox_Clear.Text = "";
                }
            }
                        
            JobFound = false;
            ViewSetupCard_Button.Enabled = false;
            PartsRunProgressBar.Value = 0;

            /*
            ItemID_TextBox.Clear();
            JobID_TextBox.Clear();
            Fixture_TextBox.Clear();
            FixtureLocation_TextBox.Clear();
            Sequence_TextBox.Clear();
            ReferenceNumber_TextBox.Clear();
            Comp1_TextBox.Clear();
            Comp2_TextBox.Clear();
            Comp3_TextBox.Clear();
            Comp4_TextBox.Clear();
            Quantity1_TextBox.Clear();
            Quantity2_TextBox.Clear();
            Quantity3_TextBox.Clear();
            Quantity4_TextBox.Clear();
            PartsNeeded_TextBox.Clear();
            PartsFormed_TextBox.Clear();
            PartsRemaining_TextBox.Clear();
            CurrentPPM_TextBox.Clear();
            LiveOEE_TextBox.Clear();
            JobStartTime_TextBox.Clear();
            JobEndTime_TextBox.Clear();
            TimeRemaining_TextBox.Clear();
            CurrentItemID_TextBox.Clear();
            CurrentWeld_TextBox.Clear();
            TotalWeldsNeeded_TextBox.Clear();
            */
        }

        public void PassOperationValue(int SelectedOperation)
        {
            CurrentItemIDOperation = SelectedOperation;
        }

        public void PassReferenceNumber(string RefNumber)
        {
            ReferenceNumber_TextBox.Text = RefNumber;
        }

        public void PassValue(string strValue)
        {
           PartsNeeded_TextBox.Text = strValue;
        }

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

        /*********************************************************************************************************************
        * Timer Region Start
        * -- Total: 3
        * 
        * - Clock Tick - Interval 500ms
        * - Timer Tick - Interval 5000ms - 5s
        * - OPCStatus Timer Tick - Interval 1800000ms - 30min
        * 
        *********************************************************************************************************************/
        #region

        // Date and Time 
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
        
        private void Timer_Tick(object sender, EventArgs e)
        {
            RunningStatistics();
        }

        // Check the Connection of the OPC Server
        // Reconnect if Connection has been lost
        private void OPCStatus_Timer_Tick(object sender, EventArgs e)
        {
            if (OPCServer.IsConnected != true)
            {
                ConnectToServer_OPC(null, null);
            }
        }

 

        /********************************************************************************************************************* 
        * Timer Region End
        *********************************************************************************************************************/
        #endregion

        /*********************************************************************************************************************
        * TextBox Method Region Start
        * 
        * -- Total: 2
        * - ItemID TextBox Enter
        * - ItemID TextBox KeyDown
        * 
        *********************************************************************************************************************/
        #region

        private void ItemID_TextBox_Enter(object sender, EventArgs e)
        {
            if ((ItemID_TextBox.ReadOnly == true) && (StartNewJob_Button.Enabled == true))
            {
                StartNewJob_Button.Focus();
            }
            else if ((ItemID_TextBox.ReadOnly == true) && (StartNewJob_Button.Enabled == false))
            {
                RunMode_Button.Focus();
            }
        }

        private void ItemID_TextBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                CurrentItemID = ItemID_TextBox.Text;
                string per = CurrentItemID.Replace("-", "");
                ItemIDWrite_OPC();
                SearchItemIDOperation();
                if (JobFound == true)
                {
                    //PLC_JobIDRead_OPC();
                    //ShowComponents();
                    UserProgram.Focus();
                    UserProgram.Enabled = false;
                    User_Program_Job_Data JobData = new User_Program_Job_Data(this);
                    JobData.ItemID_TextBox.Text = ItemID_TextBox.Text;
                    JobData.JobID_TextBox.Text = JobID_TextBox.Text;
                    JobData.Comp1_TextBox.Text = Comp1_TextBox.Text;
                    JobData.Comp2_TextBox.Text = Comp2_TextBox.Text;
                    JobData.Comp3_TextBox.Text = Comp3_TextBox.Text;
                    JobData.Comp4_TextBox.Text = Comp4_TextBox.Text;
                    JobData.Fixture_TextBox.Text = Fixture_TextBox.Text;
                    JobData.FixtureLocation_TextBox.Text = FixtureLocation_TextBox.Text;
                    //JobData.ReferenceNumber_TextBox.Focus();
                    if (JobData.ShowDialog(this) == DialogResult.Yes)
                    {
                        ItemID_TextBox.ReadOnly = true;
                        //OperationInitialize();
                        //PDFFileCheck();
                        if (JobID_TextBox.Text == 0.ToString())
                        {
                            CurrentItemID_TextBox.Text = "Job Needs Spot Weld Program";
                        }
                        else if (JobID_TextBox.Text != 0.ToString())
                        {
                            CurrentItemID_TextBox.Text = CurrentItemID;
                        }
                        TeachSensor_Button.Enabled = true;
                        TeachSensor_Button.Focus();
                    }
                    else if (JobData.DialogResult == DialogResult.No)
                    {
                        ClearForm();
                        ItemID_TextBox.Focus();
                        ItemID_TextBox.ReadOnly = false;
                    }
                }
            }
        }

        /*********************************************************************************************************************
        * TextBox Method Region End
        *********************************************************************************************************************/
        #endregion

        /*********************************************************************************************************************
        * TextBox Enter Region Start
        * 
        * -- Total TextBox: 28
        * 
        *********************************************************************************************************************/
        #region

        private void JobID_TextBox_Enter(object sender, EventArgs e)
        {
            this.ActiveControl = null;
        }

        private void Sequence_TextBox_Enter(object sender, EventArgs e)
        {
            this.ActiveControl = null;
        }

        private void Tooling_TextBox_Enter(object sender, EventArgs e)
        {
            this.ActiveControl = null;
        }

        private void ToolingLocation_TextBox_Enter(object sender, EventArgs e)
        {
            this.ActiveControl = null;
        }

        private void Fixture_TextBox_Enter(object sender, EventArgs e)
        {
            this.ActiveControl = null;
        }

        private void FixtureLocation_TextBox_Enter(object sender, EventArgs e)
        {
            this.ActiveControl = null;
        }

        private void Comp1_TextBox_Enter(object sender, EventArgs e)
        {
            this.ActiveControl = null;
        }

        private void Comp2_TextBox_Enter(object sender, EventArgs e)
        {
            this.ActiveControl = null;
        }

        private void Comp3_TextBox_Enter(object sender, EventArgs e)
        {
            this.ActiveControl = null;
        }

        private void Quantity1_TextBox_Enter(object sender, EventArgs e)
        {
            this.ActiveControl = null;
        }

        private void Quantity2_TextBox_Enter(object sender, EventArgs e)
        {
            this.ActiveControl = null;
        }

        private void Quantity3_TextBox_Enter(object sender, EventArgs e)
        {
            this.ActiveControl = null;
        }

        private void PartsNeeded_TextBox_Enter(object sender, EventArgs e)
        {
            this.ActiveControl = null;
        }

        private void PartsFormed_TextBox_Enter(object sender, EventArgs e)
        {
            this.ActiveControl = null;
        }

        private void PartsRemaining_TextBox_Enter(object sender, EventArgs e)
        {
            this.ActiveControl = null;
        }

        private void CurrentPPM_TextBox_Enter(object sender, EventArgs e)
        {
            this.ActiveControl = null;
        }

        private void LiveOEE_TextBox_Enter(object sender, EventArgs e)
        {
            this.ActiveControl = null;
        }

        private void JobStartTime_TextBox_Enter(object sender, EventArgs e)
        {
            this.ActiveControl = null;
        }

        private void JobEndTime_TextBox_Enter(object sender, EventArgs e)
        {
            this.ActiveControl = null;
        }

        private void TimeRemaining_TextBox_Enter(object sender, EventArgs e)
        {
            this.ActiveControl = null;
        }

        private void CurrentItemID_TextBox_Enter(object sender, EventArgs e)
        {
            this.ActiveControl = null;
        }

        private void SpotWeld_TextBox_Enter(object sender, EventArgs e)
        {
            this.ActiveControl = null;
        }

        private void Company_TextBox_Enter(object sender, EventArgs e)
        {
            this.ActiveControl = null;
        }

        private void TotalWeldsNeeded_TextBox_Enter(object sender, EventArgs e)
        {
            this.ActiveControl = null;
        }

        private void CurrentWeld_TextBox_Enter(object sender, EventArgs e)
        {
            this.ActiveControl = null;
        }

        private void PartCompleted_TextBox_Enter(object sender, EventArgs e)
        {
            this.ActiveControl = null;
        }

        private void ReferenceNumber_TextBox_Enter(object sender, EventArgs e)
        {
            this.ActiveControl = null;
        }

        /*********************************************************************************************************************
        * TextBox Enter Region End
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
        * Methods in Testing Start
        * 
        *********************************************************************************************************************/
        #region

        /*********************************************************************************************************************
        * 
        * Methods in Testing End
        * 
        *********************************************************************************************************************/
        #endregion

        /*********************************************************************************************************************
        *
        * Code No Longer Used Start
        * 
        *********************************************************************************************************************/
        #region

        private void OPC_Timer_Tick(object sender, EventArgs e)
        {
            //PartsCompletedRunMode_OPC();
            //if (PartCounterOPC.IsBusy != true)
            //{
            //    PartCounterOPC.RunWorkerAsync();
            //}
            //PartCounterOPC.RunWorkerAsync();

            //Update_UI_OPC();
            //SetupModeTimeOut_OPC();
            //HMIPartComplete_OPC();

            //OPCServer.GetStatus();
        }

        /*********************************************************************************************************************
        *
        * Code No Longer Used End
        * 
        *********************************************************************************************************************/
        #endregion

        /*********************************************************************************************************************
        *
        * User Program End
        * 
        *********************************************************************************************************************/

        private void User_Program_FormClosing(object sender, FormClosingEventArgs e)
        {
            OPCStatus_Timer.Enabled = false;
            DMP_Spot_Weld_Login.Current.Focus();
            DMP_Spot_Weld_Login.Current.Enabled = true;
            DMP_Spot_Weld_Login.Current.WindowState = FormWindowState.Maximized;
            DMP_Spot_Weld_Login.Current.ShowInTaskbar = true;
        }
    }
}
