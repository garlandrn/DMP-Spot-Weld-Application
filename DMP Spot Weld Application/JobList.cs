using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

/*
 * Program: DMP Spot Weld Application
 * Form: JobList
 * Created By: Ryan Garland
 * Last Updated on 1/22/18
 * 
 * Form Sections
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
    public partial class JobList : Form
    {
        public static Form Job_List;
        BackgroundWorker PartImage;
        BackgroundWorker CreateExcel;
        //
        BackgroundWorker ExcelFileImportCreate;
        BackgroundWorker ExcelFileImportSave;

        public JobList()
        {
            InitializeComponent();
            Job_List = this;
            PartImage = new BackgroundWorker();
            PartImage.DoWork += new DoWorkEventHandler(FindItemImage);
            PartImage.RunWorkerCompleted += new RunWorkerCompletedEventHandler(PartImage_RunWorkerCompleted);
            CreateExcel = new BackgroundWorker();
            CreateExcel.WorkerReportsProgress = true;
            CreateExcel.DoWork += new DoWorkEventHandler(CreateExcelFile);
            CreateExcel.RunWorkerCompleted += new RunWorkerCompletedEventHandler(CreateExcelFile_RunWorkerComplete);
            CreateExcel.ProgressChanged += new ProgressChangedEventHandler(CreateExcelFile_ProgressChanged);

            //
            ExcelFileImportCreate = new BackgroundWorker();
            ExcelFileImportCreate.DoWork += new DoWorkEventHandler(ImportUpdateSQL);
            ExcelFileImportCreate.RunWorkerCompleted += new RunWorkerCompletedEventHandler(ImportUpdateSQL_Complete);
            //ExcelFileImportCreate.ProgressChanged += new ProgressChangedEventHandler(ImportUpdateExcel_ProgressChanged);

            //
            ExcelFileImportSave = new BackgroundWorker();
            ExcelFileImportSave.WorkerReportsProgress = true;
            ExcelFileImportSave.DoWork += new DoWorkEventHandler(ImportUpdateExcel);
            ExcelFileImportSave.RunWorkerCompleted += new RunWorkerCompletedEventHandler(ImportUpdateExcel_Complete);
            ExcelFileImportSave.ProgressChanged += new ProgressChangedEventHandler(ImportUpdateExcel_ProgressChanged);
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

        private string LoginTime = "";
        private string LoginForm = "Job List";
        private bool AddJob_ButtonWasClicked = false;
        private bool EditJob_ButtonWasClicked = false;
        private bool RemoveJob_ButtonWasClicked = false;
        string[] Companies = { "CAT", "HINO", "JLG", "John Deere", "Navistar", "Paccar", "" };
        string[] Quantity = { "", "1", "2", "3", "4", "5", "6", "7", "8", "9", "10" };
        string[] ComboSelections = { "Yes", "No", "Not Feasible", "" };

        string SQL_Source = @"Data Source=OHN7009,49172;Initial Catalog=Spot_Weld_Data;Integrated Security=True;Connect Timeout=15;";

        // Clock_Tick();
        private static int ClockHour;
        private static int ClockMinute;
        private static int ClockSecond;

        // Search_Button_Click();
        private static int SearchColumn;
        private static string SearchValue;

        // FindItemImage();
        private static string ItemImagePath = "";
        private static string ItemID = "";
        private static string[] ItemIDSplit = ItemID.Split('-');
        private static double ItemID_Three;
        private static double ItemID_Five;

        private static string SQLAddCommand = "";
        private static string SQLEditCommand = "";
        private static string SQLRemoveCommand = "";
        private static string SQLComponentData = "";
        private static string SQLComponentCount = "";
        private static string Refresh_Data = "";

        // Excel File Creation
        private static Excel._Workbook ReportWB;
        private static Excel.Application ReportApp;
        private static Excel._Worksheet ReportWS;
        private static Excel.Range ReportRange;
        private static string ExcelFileLocation;
        private DataSet ReportDataSet;
        private static string ReportCell = "";
        private static int RowCount;
        private static string RowCountString = "";

        private bool CustomExcelSave = false;

        List<string> ComponentList = new List<string>();
        private static bool Update = false;


        //

        private string CustomerCell = "";

        private int ImportCurrentRows = 0;
        private int ImportUpdatedRows = 0;

        // Starting Arrays
        private static string[] ItemID_CurrentArray = new string[1000];
        private static string[] JobID_CurrentArray = new string[1000];
        private static string[] Sequence_CurrentArray = new string[1000];

        // Update Arrays
        private static string[] ItemID_UpdatedArray = new string[1000];
        private static string[] Customer_UpdatedArray = new string[1000];
        private static string[] CustomerItemID_UpdatedArray = new string[1000];
        private static string[] JobID_UpdatedArray = new string[1000];
        private static string[] Sequence_UpdatedArray = new string[1000];
        private static string[] Fixture_UpdatedArray = new string[1000];
        private static string[] FixtureLocation_UpdatedArray = new string[1000];
        private static string[] Component1_UpdatedArray = new string[1000];
        private static string[] Quantity1_UpdatedArray = new string[1000];
        private static string[] Component2_UpdatedArray = new string[1000];
        private static string[] Quantity2_UpdatedArray = new string[1000];
        private static string[] Component3_UpdatedArray = new string[1000];
        private static string[] Quantity3_UpdatedArray = new string[1000];
        private static string[] Component4_UpdatedArray = new string[1000];
        private static string[] Quantity4_UpdatedArray = new string[1000];
        private static string[] TotalRuns_UpdatedArray = new string[1000];
        private static string[] PartsManufactured_UpdatedArray = new string[1000];
        private static string[] PartsPerMinute_UpdatedArray = new string[1000];
        private static string[] SetupTime_UpdatedArray = new string[1000];

        /********************************************************************************************************************
        * 
        * Variables In Testing Start
        * 
        ********************************************************************************************************************/

        private string AddJobCommandText = "";

        /********************************************************************************************************************
        * 
        * Variables In Testing End
        * 
        *********************************************************************************************************************
        *********************************************************************************************************************
        * 
        * JobList Start
        * 
        ********************************************************************************************************************/

        private void JobList_Load(object sender, EventArgs e)
        {
            SqlConnection JobListLogin = new SqlConnection(SQL_Source);
            SqlCommand LoginJobList = new SqlCommand();
            LoginJobList.CommandType = System.Data.CommandType.Text;
            LoginJobList.CommandText = "INSERT INTO [dbo].[LoginData] (EmployeeName,DMPID,LoginDateTime,LoginForm) VALUES (@EmployeeName,@DMPID,@LoginDateTime,@LoginForm)";
            LoginJobList.Connection = JobListLogin;
            LoginJobList.Parameters.AddWithValue("@LoginDateTime", Clock_TextBox.Text);
            LoginJobList.Parameters.AddWithValue("@EmployeeName", User_TextBox.Text);
            LoginJobList.Parameters.AddWithValue("@DMPID", DMPID_TextBox.Text);
            LoginJobList.Parameters.AddWithValue("@LoginForm", LoginForm);
            JobListLogin.Open();
            LoginJobList.ExecuteNonQuery();
            JobListLogin.Close();

            LoginTime = Clock_TextBox.Text;
            Clock.Enabled = true;
            CompanyCell_ComboBox.Items.Add("Navistar");
            CompanyCell_ComboBox.Items.Add("Paccar");
            //CompanyCell_ComboBox.Items.Add("John Deere");
            Customer_ComboBox.Items.AddRange(Companies);
            Quantity1_ComboBox.Items.AddRange(Quantity);
            Quantity2_ComboBox.Items.AddRange(Quantity);
            Quantity3_ComboBox.Items.AddRange(Quantity);
            Quantity4_ComboBox.Items.AddRange(Quantity);
            SearchItemID_CheckBox.Checked = true;
        }

        /*******************************************************************************************************************
        *  
        *  [User Interface]
        *  
        *  Buttons
        *  
        *  CheckBoxes
        *  
        *  ComboBoxes
        *  
        *  DataGridView
        * 
        ********************************************************************************************************************
        ********************************************************************************************************************
        * [Buttons]
        * 
        * -------------------------------------------------------[Add Job]--------------------------------------------------
        * --Actions
        *   EditJob_Button.Hide();
        *   RemoveJob_Button.Hide();
        *   
        * --Global Variables: 
        *   AddJob_ButtonWasClicked = true;
        *
        * --Methods:
        *   Clear();
        *   JobHideShow();
        * ------------------------------------------------------[Edit Job]--------------------------------------------------
        * --Actions
        *   ItemID_TextBox.ReadOnly = true;
        *   AddJob_Button.Hide();
        *   Exit_Button.Hide();
        *   RemoveJob_Button.Hide(); 
        *
        * --Global Variables:   
        *   EditJob_ButtonWasClicked = true;
        *   
        * --Methods:
        *   JobHideShow();
        *
        * -----------------------------------------------------[Remove Job]-------------------------------------------------
        * --Actions
        *   ItemID_TextBox.ReadOnly = true;
        *   EditJob_Button.Hide();
        *   AddJob_Button.Hide();
        *   Exit_Button.Hide();
        *   
        * --Global Variables:
        *   RemoveJob_ButtonWasClicked = true;
        *
        * --Methods:
        *   JobHideShow();
        *   
        * -------------------------------------------------------[Search]---------------------------------------------------
        * --Global Variables:
        *   SearchValue = ItemID_TextBox.Text;
        *   SearchColumn = 0;
        * 
        * --Methods:
        *   DateParse();
        *   FindItemImage();
        *   
        * -------------------------------------------------------[Clear]----------------------------------------------------
        * --Methods:
        *   Clear();
        * 
        * -------------------------------------------------------[Refresh]--------------------------------------------------
        * --Methods:
        *   RefreshJobs();
        * 
        * -------------------------------------------------------[Confirm]--------------------------------------------------
        * --Global Variables:
        *   JobFound = false;
        * 
        * -------------------------------------------------------[Cancel]---------------------------------------------------
        * --Actions:
        *   Confirm_Button.Hide();
        *   Cancel_Button.Hide();
        *   DateOfCompletionPicker.Hide();
        *   Customer_ComboBox.Hide();
        *   Steps_ComboBox.Hide();
        *   AddJob_Button.Show();
        *   EditJob_Button.Show();
        *   RemoveJob_Button.Show();
        *   Exit_Button.Show();
        *   SearchItemID_CheckBox.Show();
        *   SearchJobID_CheckBox.Show();
        *   CompletionDate_TextBox.Show();
        *   Steps_TextBox.Show();
        *   Customer_TextBox.Show();
        * 
        * --Global Variables:
        *   AddJob_ButtonWasClicked = false;
        *   EditJob_ButtonWasClicked = false;
        *   RemoveJob_ButtonWasClicked = false;
        *
        * --Method Variables:
        *   Clear();
        * 
        * --------------------------------------------------------[Exit]----------------------------------------------------
        * --SQL Method:
        *   JobListLogoff();
        *    
        ********************************************************************************************************************/
        
        private void Search_Button_Click(object sender, EventArgs e)
        {
            bool found = false;
            if (SearchItemID_CheckBox.Checked == true)
            {
                SearchJobID_CheckBox.Checked = false;
                SearchValue = ItemID_TextBox.Text;
                SearchColumn = 0;
            }
            else if (SearchJobID_CheckBox.Checked == true)
            {
                SearchItemID_CheckBox.Checked = false;
                SearchValue = JobID_TextBox.Text;
                SearchColumn = 3;
            }
            JobListGridView.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            try
            {
                foreach (DataGridViewRow Row in JobListGridView.Rows)
                {
                    Row.Selected = false;
                    if (Row.Cells[SearchColumn].Value.ToString().Equals(SearchValue))
                    {
                        Row.Selected = true;
                        ItemID_TextBox.Text = Row.Cells[0].Value.ToString();
                        Customer_TextBox.Text = Row.Cells[1].Value.ToString();
                        Customer_ComboBox.Text = Row.Cells[1].Value.ToString();
                        CustomerItemID_TextBox.Text = Row.Cells[2].Value.ToString();
                        JobID_TextBox.Text = Row.Cells[3].Value.ToString();
                        Sequence_TextBox.Text = Row.Cells[4].Value.ToString();
                        Fixture_TextBox.Text = Row.Cells[5].Value.ToString();
                        FixtureLocation_TextBox.Text = Row.Cells[6].Value.ToString();
                        Comp1_TextBox.Text = Row.Cells[7].Value.ToString();
                        Comp1_ComboBox.Text = Row.Cells[7].Value.ToString();
                        Quantity1_TextBox.Text = Row.Cells[8].Value.ToString();
                        Quantity1_ComboBox.Text = Row.Cells[8].Value.ToString();
                        Comp2_TextBox.Text = Row.Cells[9].Value.ToString();
                        Comp2_ComboBox.Text = Row.Cells[9].Value.ToString();
                        Quantity2_TextBox.Text = Row.Cells[10].Value.ToString();
                        Quantity2_ComboBox.Text = Row.Cells[10].Value.ToString();
                        Comp3_TextBox.Text = Row.Cells[11].Value.ToString();
                        Comp3_ComboBox.Text = Row.Cells[11].Value.ToString();
                        Quantity3_TextBox.Text = Row.Cells[12].Value.ToString();
                        Quantity3_ComboBox.Text = Row.Cells[12].Value.ToString();
                        Comp4_TextBox.Text = Row.Cells[13].Value.ToString();
                        Comp4_ComboBox.Text = Row.Cells[13].Value.ToString();
                        Quantity4_TextBox.Text = Row.Cells[14].Value.ToString();
                        Quantity4_ComboBox.Text = Row.Cells[14].Value.ToString();
                        TotalRuns_TextBox.Text = Row.Cells[15].Value.ToString();
                        PartsManufactured_TextBox.Text = Row.Cells[16].Value.ToString();
                        PPM_TextBox.Text = Row.Cells[17].Value.ToString();
                        SetupTime_TextBox.Text = Row.Cells[18].Value.ToString();
                        DateParse();
                        //FindItemImage();
                        if (PartImage.IsBusy != true)
                        {
                            PartImage.RunWorkerAsync();
                        }
                        JobListGridView.FirstDisplayedScrollingRowIndex = JobListGridView.SelectedRows[0].Index;
                        break;
                    }
                }
            }
            catch (Exception)
            {
                if (found == false)
                {
                    MessageBox.Show("Item ID: " + SearchValue + " Was Not Found");
                }
                else if (SearchColumn == 3)
                {
                    MessageBox.Show("Spot Weld Job ID: " + SearchValue + " Was Not Found");
                }
            }
        }

        private void Clear_Button_Click(object sender, EventArgs e)
        {
            if (CreateExcel.IsBusy == true)
            {
                CreateExcel.CancelAsync();
            }
            ClearMethod();
            GroupBoxControlStart();
        }

        private void AddJob_Button_Click(object sender, EventArgs e)
        {
            JobListGridView.Enabled = false;
            AddJob_ButtonWasClicked = true;
            Search_Button.Enabled = false;
            Clear_Button.Enabled = false;
            EditJob_Button.Hide();
            RemoveJob_Button.Hide();
            GroupBoxControlAddEditJob();
            ClearMethod();
            JobHideShow();

            Fixture_TextBox.Text = "N/A";
            FixtureLocation_TextBox.Text = "N/A";
        }

        private void EditJob_Button_Click(object sender, EventArgs e)
        {
            if (ItemID_TextBox.Text == "")
            {
                MessageBox.Show("Please Select a Job to Edit");
            }
            else
            {
                JobListGridView.Enabled = false;
                ItemID_TextBox.ReadOnly = true;
                EditJob_ButtonWasClicked = true;
                Search_Button.Enabled = false;
                Clear_Button.Enabled = false;
                AddJob_Button.Hide();
                Exit_Button.Hide();
                RemoveJob_Button.Hide();
                Customer_ComboBox.Text = Customer_TextBox.Text;
                Comp1_ComboBox.Text = Comp1_TextBox.Text;
                Quantity1_ComboBox.Text = Quantity1_TextBox.Text;
                Comp2_ComboBox.Text = Comp2_TextBox.Text;
                Quantity2_ComboBox.Text = Quantity2_TextBox.Text;
                Comp3_ComboBox.Text = Comp3_TextBox.Text;
                Quantity3_ComboBox.Text = Quantity3_TextBox.Text;
                Comp4_ComboBox.Text = Comp4_TextBox.Text;
                Quantity4_ComboBox.Text = Quantity4_TextBox.Text;

                GroupBoxControlAddEditJob();
                JobHideShow();
            }
        }

        private void RemoveJob_Button_Click(object sender, EventArgs e)
        {
            if (ItemID_TextBox.Text == "")
            {
                MessageBox.Show("Please Select a Job to Remove");
            }
            else
            {
                JobListGridView.Enabled = false;
                RemoveJob_ButtonWasClicked = true;
                Search_Button.Enabled = false;
                Clear_Button.Enabled = false;
                EditJob_Button.Hide();
                AddJob_Button.Hide();
                Exit_Button.Hide();
                ItemID_TextBox.ReadOnly = true;
                GroupBoxControlInactive();
                JobHideShow();
            }
        }

        private void Confirm_Button_Click(object sender, EventArgs e)
        {
            if (AddJob_ButtonWasClicked == true)
            {
                JobList_Add AddConfirm = new JobList_Add();
                AddConfirm.ItemID_TextBox.Text = ItemID_TextBox.Text;
                AddConfirm.Customer_TextBox.Text = Customer_ComboBox.Text;
                AddConfirm.CustomerItemID_TextBox.Text = CustomerItemID_TextBox.Text;
                AddConfirm.JobID_TextBox.Text = JobID_TextBox.Text;
                AddConfirm.Fixture_TextBox.Text = Fixture_TextBox.Text;
                AddConfirm.FixtureLocation_TextBox.Text = FixtureLocation_TextBox.Text;
                AddConfirm.Sequence_TextBox.Text = Sequence_TextBox.Text;
                AddConfirm.Comp1_TextBox.Text = Comp1_ComboBox.Text;
                AddConfirm.Quantity1_TextBox.Text = Quantity1_ComboBox.Text;
                AddConfirm.Comp2_TextBox.Text = Comp2_ComboBox.Text;
                AddConfirm.Quantity2_TextBox.Text = Quantity2_ComboBox.Text;
                AddConfirm.Comp3_TextBox.Text = Comp3_ComboBox.Text;
                AddConfirm.Quantity3_TextBox.Text = Quantity3_ComboBox.Text;
                AddConfirm.Comp4_TextBox.Text = Comp4_ComboBox.Text;
                AddConfirm.Quantity4_TextBox.Text = Quantity4_ComboBox.Text;

                if (AddConfirm.ShowDialog(this) == DialogResult.Yes)
                {
                    try
                    {
                        SqlConnection Job_Connection = new SqlConnection(SQL_Source);
                        SqlCommand Add_Job = new SqlCommand();
                        Add_Job.CommandType = System.Data.CommandType.Text;
                        Add_Job.CommandText = SQLAddCommand;
                        Add_Job.Connection = Job_Connection;
                        Add_Job.Parameters.AddWithValue("@ItemID", ItemID_TextBox.Text);
                        Add_Job.Parameters.AddWithValue("@Customer", Customer_ComboBox.Text);
                        Add_Job.Parameters.AddWithValue("@CustomerItemID", CustomerItemID_TextBox.Text);
                        Add_Job.Parameters.AddWithValue("@JobID", JobID_TextBox.Text);
                        Add_Job.Parameters.AddWithValue("@Sequence", Sequence_TextBox.Text);
                        Add_Job.Parameters.AddWithValue("@Fixture", Fixture_TextBox.Text);
                        Add_Job.Parameters.AddWithValue("@FixtureLocation", FixtureLocation_TextBox.Text);
                        Add_Job.Parameters.AddWithValue("@Component1", Comp1_ComboBox.Text);
                        Add_Job.Parameters.AddWithValue("@Quantity1", Quantity1_ComboBox.Text);
                        Add_Job.Parameters.AddWithValue("@Component2", Comp2_ComboBox.Text);
                        Add_Job.Parameters.AddWithValue("@Quantity2", Quantity2_ComboBox.Text);
                        Add_Job.Parameters.AddWithValue("@Component3", Comp3_ComboBox.Text);
                        Add_Job.Parameters.AddWithValue("@Quantity3", Quantity3_ComboBox.Text);
                        Add_Job.Parameters.AddWithValue("@Component4", Comp4_ComboBox.Text);
                        Add_Job.Parameters.AddWithValue("@Quantity4", Quantity4_ComboBox.Text);
                        Add_Job.Parameters.AddWithValue("@TotalRuns", 0);
                        Add_Job.Parameters.AddWithValue("@PartsManufactured", 0);
                        Add_Job.Parameters.AddWithValue("@PartsPerMinute", 0);
                        Add_Job.Parameters.AddWithValue("@SetupTime", 0);

                        Job_Connection.Open();
                        Add_Job.ExecuteNonQuery();
                        Job_Connection.Close();
                    }
                    catch (SqlException ExceptionValue)
                    {
                        int ErrorNumber = ExceptionValue.Number;
                        if (ErrorNumber.Equals(2627))
                        {
                            MessageBox.Show("Item ID: " + ItemID_TextBox.Text + " is Already on this List");
                        }
                        else if (ErrorNumber.Equals(245))
                        {
                            MessageBox.Show("Item ID Can Only Contain Numbers");
                        }
                        else
                        {
                            MessageBox.Show("Unable to Add Job. Please Try Again." + "\n" + "Error Code: " + ErrorNumber.ToString());
                        }
                    }
                    ConfirmFinished();
                    RefreshJobs();
                    CreateExcel_Start();
                    if(CreateExcel.IsBusy != true)
                    {
                        CreateExcel.RunWorkerAsync();
                    }
                    ClearMethod();
                    GroupBoxControlStart();
                }
                else
                {
                    ConfirmFinished();
                    ClearMethod();
                    GroupBoxControlStart();
                }
            }
            else if (EditJob_ButtonWasClicked == true)
            {
                JobList_Edit EditConfirm = new JobList_Edit();
                EditConfirm.UpdatedItemID_TextBox.Text = ItemID_TextBox.Text;
                EditConfirm.UpdatedCustomer_TextBox.Text = Customer_ComboBox.Text;
                EditConfirm.UpdatedCustomerItemID_TextBox.Text = CustomerItemID_TextBox.Text;
                EditConfirm.UpdatedJobID_TextBox.Text = JobID_TextBox.Text;
                EditConfirm.UpdatedSequence_TextBox.Text = Sequence_TextBox.Text;
                EditConfirm.UpdatedFixture_TextBox.Text = Fixture_TextBox.Text;
                EditConfirm.UpdatedFixtureLocation_TextBox.Text = FixtureLocation_TextBox.Text;
                EditConfirm.UpdatedComp1_TextBox.Text = Comp1_ComboBox.Text;
                EditConfirm.UpdatedQuantity1_TextBox.Text = Quantity1_ComboBox.Text;
                EditConfirm.UpdatedComp2_TextBox.Text = Comp2_ComboBox.Text;
                EditConfirm.UpdatedQuantity2_TextBox.Text = Quantity2_ComboBox.Text;
                EditConfirm.UpdatedComp3_TextBox.Text = Comp3_ComboBox.Text;
                EditConfirm.UpdatedQuantity3_TextBox.Text = Quantity3_ComboBox.Text;
                EditConfirm.UpdatedComp4_TextBox.Text = Comp4_ComboBox.Text;
                EditConfirm.UpdatedQuantity4_TextBox.Text = Quantity4_ComboBox.Text;
                JobListGridView.SelectionMode = DataGridViewSelectionMode.FullRowSelect;

                try
                {
                    SearchValue = ItemID_TextBox.Text;
                    foreach (DataGridViewRow Row in JobListGridView.Rows)
                    {
                        Row.Selected = false;
                        if (Row.Cells[0].Value.ToString().Equals(SearchValue))
                        {
                            Row.Selected = true;

                            EditConfirm.ItemID_TextBox.Text = Row.Cells[0].Value.ToString();
                            EditConfirm.Customer_TextBox.Text = Row.Cells[1].Value.ToString();
                            EditConfirm.CustomerItemID_TextBox.Text = Row.Cells[2].Value.ToString();
                            EditConfirm.JobID_TextBox.Text = Row.Cells[3].Value.ToString();
                            EditConfirm.Sequence_TextBox.Text = Row.Cells[4].Value.ToString();
                            EditConfirm.Fixture_TextBox.Text = Row.Cells[5].Value.ToString();
                            EditConfirm.FixtureLocation_TextBox.Text = Row.Cells[6].Value.ToString();
                            EditConfirm.Comp1_TextBox.Text = Row.Cells[7].Value.ToString();
                            EditConfirm.Quantity1_TextBox.Text = Row.Cells[8].Value.ToString();
                            EditConfirm.Comp2_TextBox.Text = Row.Cells[9].Value.ToString();
                            EditConfirm.Quantity2_TextBox.Text = Row.Cells[10].Value.ToString();
                            EditConfirm.Comp3_TextBox.Text = Row.Cells[11].Value.ToString();
                            EditConfirm.Quantity3_TextBox.Text = Row.Cells[12].Value.ToString();
                            EditConfirm.Comp4_TextBox.Text = Row.Cells[13].Value.ToString();
                            EditConfirm.Quantity4_TextBox.Text = Row.Cells[14].Value.ToString();
                            break;
                        }
                    }
                }
                catch (Exception)
                {
                    MessageBox.Show("Error Finding Job");
                }
                if (EditConfirm.ShowDialog(this) == DialogResult.Yes)
                {
                    if(CompanyCell_ComboBox.Text == "CAT")
                    {
                        SQLEditCommand = "UPDATE [dbo].[CAT_Item_Data] SET Customer=@Customer,CustomerItemID=@CustomerItemID,JobID=@JobID,Fixture=@Fixture,FixtureLocation=@FixtureLocation,Component1=@Component1,Quantity1=@Quantity1,Component2=@Component2,Quantity2=@Quantity2,Component3=@Component3,Quantity3=@Quantity3,Component4=@Component4,Quantity4=@Quantity4 WHERE ItemID=@ItemID AND Sequence='" + Sequence_TextBox.Text + "'";
                    }
                    else if(CompanyCell_ComboBox.Text == "John Deere")
                    {
                        SQLEditCommand = "UPDATE [dbo].[JohnDeere_Item_Data] SET Customer=@Customer,CustomerItemID=@CustomerItemID,JobID=@JobID,Fixture=@Fixture,FixtureLocation=@FixtureLocation,Component1=@Component1,Quantity1=@Quantity1,Component2=@Component2,Quantity2=@Quantity2,Component3=@Component3,Quantity3=@Quantity3,Component4=@Component4,Quantity4=@Quantity4 WHERE ItemID=@ItemID AND Sequence='" + Sequence_TextBox.Text + "'";
                    }
                    else if (CompanyCell_ComboBox.Text == "Navistar")
                    {
                        SQLEditCommand = "UPDATE [dbo].[Navistar_Item_Data] SET Customer=@Customer,CustomerItemID=@CustomerItemID,JobID=@JobID,Fixture=@Fixture,FixtureLocation=@FixtureLocation,Component1=@Component1,Quantity1=@Quantity1,Component2=@Component2,Quantity2=@Quantity2,Component3=@Component3,Quantity3=@Quantity3,Component4=@Component4,Quantity4=@Quantity4 WHERE ItemID=@ItemID AND Sequence='" + Sequence_TextBox.Text + "'";
                    }
                    else if (CompanyCell_ComboBox.Text == "Paccar")
                    {
                        SQLEditCommand = "UPDATE [dbo].[Paccar_Item_Data] SET Customer=@Customer,CustomerItemID=@CustomerItemID,JobID=@JobID,Fixture=@Fixture,FixtureLocation=@FixtureLocation,Component1=@Component1,Quantity1=@Quantity1,Component2=@Component2,Quantity2=@Quantity2,Component3=@Component3,Quantity3=@Quantity3,Component4=@Component4,Quantity4=@Quantity4 WHERE ItemID=@ItemID AND Sequence='" + Sequence_TextBox.Text + "'";
                    }
                    try
                    {
                        SqlConnection Job_Connection = new SqlConnection(SQL_Source);
                        SqlCommand Edit_Job = new SqlCommand();
                        Edit_Job.CommandType = System.Data.CommandType.Text;
                        Edit_Job.CommandText = SQLEditCommand;
                        Edit_Job.Connection = Job_Connection;
                        Edit_Job.Parameters.AddWithValue("@ItemID", ItemID_TextBox.Text);
                        Edit_Job.Parameters.AddWithValue("@Customer", Customer_ComboBox.Text);
                        Edit_Job.Parameters.AddWithValue("@CustomerItemID", CustomerItemID_TextBox.Text);
                        Edit_Job.Parameters.AddWithValue("@JobID", JobID_TextBox.Text);
                        Edit_Job.Parameters.AddWithValue("@Sequence", Sequence_TextBox.Text);
                        Edit_Job.Parameters.AddWithValue("@Fixture", Fixture_TextBox.Text);
                        Edit_Job.Parameters.AddWithValue("@FixtureLocation", FixtureLocation_TextBox.Text);
                        Edit_Job.Parameters.AddWithValue("@Component1", Comp1_ComboBox.Text);
                        Edit_Job.Parameters.AddWithValue("@Quantity1", Quantity1_ComboBox.Text);
                        Edit_Job.Parameters.AddWithValue("@Component2", Comp2_ComboBox.Text);
                        Edit_Job.Parameters.AddWithValue("@Quantity2", Quantity2_ComboBox.Text);
                        Edit_Job.Parameters.AddWithValue("@Component3", Comp3_ComboBox.Text);
                        Edit_Job.Parameters.AddWithValue("@Quantity3", Quantity3_ComboBox.Text.ToString());
                        Edit_Job.Parameters.AddWithValue("@Component4", Comp4_ComboBox.Text);
                        Edit_Job.Parameters.AddWithValue("@Quantity4", Quantity4_ComboBox.Text);
                        Job_Connection.Open();
                        Edit_Job.ExecuteNonQuery();
                        Job_Connection.Close();
                    }
                    catch (SqlException ExceptionValue)
                    {
                        int ErrorNumber = ExceptionValue.Number;
                        MessageBox.Show("Unable to Edit Job" + "\n" + "Error Code: " + ErrorNumber.ToString());
                    }
                    ConfirmFinished();
                    RefreshJobs();
                    CreateExcel_Start();
                    if (CreateExcel.IsBusy != true)
                    {
                        CreateExcel.RunWorkerAsync();
                    }
                    ClearMethod();
                    GroupBoxControlStart();
                }
                else
                {
                    ConfirmFinished();
                    ClearMethod();
                    GroupBoxControlStart();
                }
            }
            else if (RemoveJob_ButtonWasClicked == true)
            {
                JobList_Remove RemoveConfirm = new JobList_Remove();
                RemoveConfirm.ItemID_TextBox.Text = ItemID_TextBox.Text;
                RemoveConfirm.Customer_TextBox.Text = Customer_ComboBox.Text;
                RemoveConfirm.CustomerItemID_TextBox.Text = CustomerItemID_TextBox.Text;
                RemoveConfirm.JobID_TextBox.Text = JobID_TextBox.Text;
                RemoveConfirm.Sequence_TextBox.Text = Sequence_TextBox.Text;
                RemoveConfirm.Fixture_TextBox.Text = Fixture_TextBox.Text;
                RemoveConfirm.FixtureLocation_TextBox.Text = FixtureLocation_TextBox.Text;
                RemoveConfirm.Comp1_TextBox.Text = Comp1_ComboBox.Text;
                RemoveConfirm.Quantity1_TextBox.Text = Quantity1_ComboBox.Text;
                RemoveConfirm.Comp2_TextBox.Text = Comp2_ComboBox.Text;
                RemoveConfirm.Quantity2_TextBox.Text = Quantity2_ComboBox.Text;
                RemoveConfirm.Comp3_TextBox.Text = Comp3_ComboBox.Text;
                RemoveConfirm.Quantity3_TextBox.Text = Quantity3_ComboBox.Text;
                RemoveConfirm.Comp4_TextBox.Text = Comp4_ComboBox.Text;
                RemoveConfirm.Quantity4_TextBox.Text = Quantity4_ComboBox.Text;

                if (RemoveConfirm.ShowDialog(this) == DialogResult.Yes)
                {
                    try
                    {
                        SqlConnection Job_Connection = new SqlConnection(SQL_Source);
                        SqlCommand Delete_Job = new SqlCommand();
                        Delete_Job.CommandType = System.Data.CommandType.Text;
                        Delete_Job.CommandText = SQLRemoveCommand;
                        Delete_Job.Connection = Job_Connection;
                        Delete_Job.Parameters.AddWithValue("@ItemID", ItemID_TextBox.Text);
                        Job_Connection.Open();
                        Delete_Job.ExecuteNonQuery();
                        Job_Connection.Close();
                        MessageBox.Show("Job Was Successfully Removed");
                    }
                    catch (SqlException ExceptionValue)
                    {
                        int ErrorNumber = ExceptionValue.Number;
                        MessageBox.Show("Error Removing Job" + "\n" + "Error Code: " + ErrorNumber.ToString());
                    }
                    ConfirmFinished();
                    RefreshJobs();
                    CreateExcel_Start();
                    if (CreateExcel.IsBusy != true)
                    {
                        CreateExcel.RunWorkerAsync();
                    }
                    ClearMethod();
                    GroupBoxControlStart();
                }
                else
                {
                    ConfirmFinished();
                    ClearMethod();
                    GroupBoxControlStart();
                }
            }
        }

        private void Cancel_Button_Click(object sender, EventArgs e)
        {
            if (CreateExcel.IsBusy == true)
            {
                CreateExcel.CancelAsync();
            }
            JobListGridView.Enabled = true;
            AddJob_Button.Show();
            EditJob_Button.Show();
            RemoveJob_Button.Show();
            Exit_Button.Show();
            Confirm_Button.Hide();
            Cancel_Button.Hide();
            Search_Button.Enabled = true;
            Clear_Button.Enabled = true;
            Customer_ComboBox.Hide();
            SearchItemID_CheckBox.Show();
            SearchJobID_CheckBox.Show();
            Customer_TextBox.Show();
            ItemID_TextBox.ReadOnly = false;
            AddJob_ButtonWasClicked = false;
            EditJob_ButtonWasClicked = false;
            RemoveJob_ButtonWasClicked = false;
            ClearMethod();
            GroupBoxControlStart();
        }

        private void Refresh_Button_Click(object sender, EventArgs e)
        {
            RefreshJobs();
        }

        private void ViewCheckCard_Button_Click(object sender, EventArgs e)
        {
            Check_Card_Data_Viewer CheckCardViewer = new Check_Card_Data_Viewer();
            CheckCardViewer.ItemID_TextBox.Text = ItemID_TextBox.Text;
            CheckCardViewer.Clock_TextBox.Text = Clock_TextBox.Text;
            CheckCardViewer.User_TextBox.Text = User_TextBox.Text;
            CheckCardViewer.Show();
        }

        private void ChangeCell_Button_Click(object sender, EventArgs e)
        {
            if (CreateExcel.IsBusy == true)
            {
                //CreateExcel.CancelAsync();
            }
            CompanyCell_ComboBox.Enabled = true;
            CompanyCell_ComboBox.Visible = true;
            CompanyCell_TextBox.Visible = false;
            JobListGridView.Enabled = false;
            GroupBoxControlStart();
            ClearMethod();
            ChangeCell_Button.Hide();
            Comp1_ComboBox.Items.Clear();
            Comp2_ComboBox.Items.Clear();
            Comp3_ComboBox.Items.Clear();
            Comp4_ComboBox.Items.Clear();
        }

        private void Exit_Button_Click(object sender, EventArgs e)
        {
            if (CreateExcel.IsBusy != true)
            {
                JobListLogoff();
                DMP_Spot_Weld_Login.Current.Focus();
                DMP_Spot_Weld_Login.Current.Enabled = true;
                DMP_Spot_Weld_Login.Current.WindowState = FormWindowState.Maximized;
                DMP_Spot_Weld_Login.Current.ShowInTaskbar = true;
                this.Close();
            }
        }

        /*********************************************************************************************************************
        * 
        * Buttons End
        * 
        *********************************************************************************************************************/
        /********************************************************************************************************************
        * [CheckBoxes]
        * 
        * -------------------------------------------------------[SearchItemID]----------------------------------------------
        * 
        * Sets SearchItemID to True and SearchJobID to False When Selected
        * 
        * -------------------------------------------------------[SearchJobID]-----------------------------------------------
        * 
        * Sets SearchJobID to True and SearchItemID to False When Selected
        * 
        ********************************************************************************************************************/

        private void SearchItemID_CheckBox_CheckedChanged(object sender, EventArgs e)
        {
            if (SearchItemID_CheckBox.Checked == true)
            {
                SearchJobID_CheckBox.Checked = false;
            }
        }

        private void SearchJobID_CheckBox_CheckedChanged(object sender, EventArgs e)
        {
            if (SearchJobID_CheckBox.Checked == true)
            {
                SearchItemID_CheckBox.Checked = false;
            }
        }

        /********************************************************************************************************************
        * 
        * CheckBoxes End
        * 
        ********************************************************************************************************************/

        /*********************************************************************************************************************
        * 
        * [ComboBoxes]
        * 
        * ---------------------------------------------------[Company Cell]---------------------------------------------------
        * 
        * 
        * ----------------------------------------------------[Spot Welder]---------------------------------------------------
        * 
        * -----------------------------------------------------[Customer]-----------------------------------------------------
        * 
        * 
        * ----------------------------------------------------[Component 1]---------------------------------------------------
        * 
        * ----------------------------------------------------[Component 2]---------------------------------------------------
        * 
        * ----------------------------------------------------[Component 3]---------------------------------------------------        * 
        * 
        * -----------------------------------------------------[Quantity 1]---------------------------------------------------
        * 
        * -----------------------------------------------------[Quantity 2]---------------------------------------------------
        * 
        * -----------------------------------------------------[Quantity 3]--------------------------------------------------- 
        * 
        *********************************************************************************************************************/

        private void CompanyCell_ComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            CompanyCell_ComboBox.Enabled = false;
            CompanyCell_ComboBox.Visible = false;
            CompanyCell_TextBox.Visible = true;
            CompanyCell_TextBox.Text = CompanyCell_ComboBox.Text;
            ChangeCell_Button.Show();
            JobListGridView.Enabled = true;

            if (CompanyCell_ComboBox.Text == "John Deere")
            {/*
                BrakePress_ComboBox.Items.Add("1107");
                BrakePress_ComboBox.Items.Add("1139");
                */
                ReportCell = "JohnDeere_";
                Refresh_Data = "SELECT * FROM [dbo].[JohnDeere_Item_Data] ORDER BY JobID DESC";
                SQLAddCommand = "INSERT INTO [dbo].[JohnDeere_Item_Data] (ItemID,Customer,CustomerItemID,Steps,JobID,StepsUsed,CompletionDate,Sample3D,Ready3D,BendSimulation,Tooling,ToolingLocation,FixtureLocation) VALUES (@ItemID,@Customer,@CustomerItemID,@Steps,@JobID,@StepsUsed,@CompletionDate,@Sample3D,@Ready3D,@BendSimulation,@Tooling,@ToolingLocation,@FixtureLocation)";
                SQLEditCommand = "UPDATE [dbo].[JohnDeere_Item_Data] SET Customer=@Customer,CustomerItemID=@CustomerItemID,Steps=@Steps,JobID=@JobID,StepsUsed=@StepsUsed,CompletionDate=@CompletionDate,Sample3D=@Sample3D,Ready3D=@Ready3D,BendSimulation=@BendSimulation,Tooling=@Tooling,ToolingLocation=@ToolingLocation,FixtureLocation=@FixtureLocation WHERE ItemID=@ItemID";
               // SQLRemoveCommand = "DELETE FROM [dbo].[JohnDeere_Item_Data] WHERE ItemID=@ItemID";
                SqlConnection BrakePressConnect = new SqlConnection(SQL_Source);
                string JohnDeere = "SELECT * FROM [dbo].[JohnDeere_Item_Data] ORDER BY JobID DESC";
                SqlDataAdapter Data1178 = new SqlDataAdapter(JohnDeere, BrakePressConnect);
                SqlCommandBuilder CommandBuilder = new SqlCommandBuilder(Data1178);
                DataSet JohnDeereData = new DataSet();
                Data1178.Fill(JohnDeereData);
                JobListGridView.DataSource = JohnDeereData.Tables[0];
                SpotWelder_ComboBox.Items.Add("1178");

            }
            else if (CompanyCell_ComboBox.Text == "Navistar")
            {
                // Connect to SQL DataTable and Load
                ReportCell = "Navistar_";
                Refresh_Data = "SELECT * FROM [dbo].[Navistar_Item_Data] ORDER BY JobID DESC";
                SQLAddCommand = "INSERT INTO [dbo].[Navistar_Item_Data] (ItemID,Customer,CustomerItemID,JobID,Sequence,Fixture,FixtureLocation,Component1,Quantity1,Component2,Quantity2,Component3,Quantity3,Component4,Quantity4,TotalRuns,PartsManufactured,PartsPerMinute,SetupTime) VALUES (@ItemID,@Customer,@CustomerItemID,@JobID,@Sequence,@Fixture,@FixtureLocation,@Component1,@Quantity1,@Component2,@Quantity2,@Component3,@Quantity3,@Component4,@Quantity4,@TotalRuns,@PartsManufactured,@PartsPerMinute,@SetupTime)";
                //SQLEditCommand = "UPDATE [dbo].[Navistar_Item_Data] SET Customer=@Customer,CustomerItemID=@CustomerItemID,JobID=@JobID,Sequence=@Sequence,Fixture=@Fixture,FixtureLocation=@FixtureLocation,Component1=@Component1,Quantity1=@Quantity1,Component2=@Component2,Quantity2=@Quantity2,Component3=@Component3,Quantity3=@Quantity3 WHERE ItemID=@ItemID";
                SQLRemoveCommand = "DELETE FROM [dbo].[Navistar_Item_Data] WHERE ItemID=@ItemID";
                SQLComponentData = "SELECT * FROM [dbo].[Navistar_Component_Data]";
                SQLComponentCount = "SELECT COUNT(*) FROM [dbo].[Navistar_Component_Data]";
                SqlConnection BrakePressConnect = new SqlConnection(SQL_Source);
                string Navistar1176 = "SELECT * FROM [dbo].[Navistar_Item_Data] ORDER BY JobID DESC";
                SqlDataAdapter Data1176 = new SqlDataAdapter(Navistar1176, BrakePressConnect);
                SqlCommandBuilder CommandBuilder = new SqlCommandBuilder(Data1176);
                DataSet NavistarData = new DataSet();
                Data1176.Fill(NavistarData);
                JobListGridView.DataSource = NavistarData.Tables[0];
                //LoadComponentData();
            }
            else if (CompanyCell_ComboBox.Text == "Paccar")
            {
                // Connect to SQL DataTable and Load
                ReportCell = "Paccar_";
                Refresh_Data = "SELECT * FROM [dbo].[Paccar_Item_Data] ORDER BY JobID DESC";
                SQLAddCommand = "INSERT INTO [dbo].[Paccar_Item_Data] (ItemID,Customer,CustomerItemID,JobID,Sequence,Fixture,FixtureLocation,Component1,Quantity1,Component2,Quantity2,Component3,Quantity3,Component4,Quantity4,TotalRuns,PartsManufactured,PartsPerMinute,SetupTime) VALUES (@ItemID,@Customer,@CustomerItemID,@JobID,@Sequence,@Fixture,@FixtureLocation,@Component1,@Quantity1,@Component2,@Quantity2,@Component3,@Quantity3,@Component4,@Quantity4,@TotalRuns,@PartsManufactured,@PartsPerMinute,@SetupTime)";
                //SQLEditCommand = "UPDATE [dbo].[Paccar_Item_Data] SET Customer=@Customer,CustomerItemID=@CustomerItemID,JobID=@JobID,Sequence=@Sequence,Fixture=@Fixture,FixtureLocation=@FixtureLocation,Component1=@Component1,Quantity1=@Quantity1,Component2=@Component2,Quantity2=@Quantity2,Component3=@Component3,Quantity3=@Quantity3 WHERE ItemID=@ItemID AND Sequence='" + Sequence_TextBox.Text + "'";
                //SQLEditCommand = "UPDATE [dbo].[Paccar_Item_Data] SET Customer=@Customer,CustomerItemID=@CustomerItemID,JobID=@JobID,Fixture=@Fixture,FixtureLocation=@FixtureLocation,Component1=@Component1,Quantity1=@Quantity1,Component2=@Component2,Quantity2=@Quantity2,Component3=@Component3,Quantity3=@Quantity3 WHERE ItemID=@ItemID AND Sequence='" + Sequence_TextBox.Text + "'";
                SQLRemoveCommand = "DELETE FROM [dbo].[Paccar_Item_Data] WHERE ItemID=@ItemID";
                SQLComponentData = "SELECT * FROM [dbo].[Paccar_Component_Data]";
                SQLComponentCount = "SELECT COUNT(*) FROM [dbo].[Paccar_Component_Data]";
                SqlConnection BrakePressConnect = new SqlConnection(SQL_Source);
                string Paccar1176 = "SELECT * FROM [dbo].[Paccar_Item_Data] ORDER BY JobID DESC";
                SqlDataAdapter Data1176 = new SqlDataAdapter(Paccar1176, BrakePressConnect);
                SqlCommandBuilder CommandBuilder = new SqlCommandBuilder(Data1176);
                DataSet PaccarData = new DataSet();
                Data1176.Fill(PaccarData);
                JobListGridView.DataSource = PaccarData.Tables[0];
                //LoadComponentData();
            }
            Row_TotalCount();
            Search_Button.Show();
            Clear_Button.Show();
            JobListGridView.Focus();
            LoadComponentData();
            FirstLoadTest();
        }

        private void SpotWelder_ComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (SpotWelder_ComboBox.Text == "104R")
            {

            }
            else if (SpotWelder_ComboBox.Text == "121R")
            {

            }
            else if (SpotWelder_ComboBox.Text == "153R")
            {

            }
            else if (SpotWelder_ComboBox.Text == "154R")
            {

            }
            else if (SpotWelder_ComboBox.Text == "155R")
            {

            }
        }

        /********************************************************************************************************************
        * 
        * ComboBoxes End
        * 
        ********************************************************************************************************************/

        /********************************************************************************************************************
        * [GridView]                                                                                                        
        *                                                                                                                   
        * JobListGridView_CellClick: When Cell is Clicked the Entire Row is Selected                                                   
        * --Methods:
        *   DateParse();
        *   FindItemImage();                                                           
        *                                                                                                                   
        ********************************************************************************************************************/

        private void JobListGridView_RowEnter(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void JobListGridView_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {
            e.Cancel = true;
        }

        /********************************************************************************************************************
        * 
        * DataGridView End
        * 
        ********************************************************************************************************************/

        /********************************************************************************************************************
        *  
        *  Methods
        *
        *  ---------------------------------------------------(Clock Tick)---------------------------------------------------
        *  --Global Variables:
        *    ClockHour = DateTime.Now.Hour;
        *    ClockMinute = DateTime.Now.Minute;
        *    ClockSecond = DateTime.Now.Second;
        *  
        *  --Method Variables:
        *    string AMPM = "";
        *    string Date = DateTime.Today.ToShortDateString();
        *    string Time = "";
        *    
        *  ---------------------------------------------------(EmployeeLogOff)-----------------------------------------------
        *  --Method SQL Data:
        *    ("@LoginDateTime", LoginTime.ToString());
        *    ("@LogoutDateTime", Clock_TextBox.Text);
        *    
        *  ---------------------------------------------------(FindTotalRunTime)---------------------------------------------
        *  --GlobalVariables:
        *    JobEndTime
        *    JobStartTime
        *    TimeOfOperation = DateTime.Parse(JobEndTime).Subtract(DateTime.Parse(JobStartTime));
        *    
        *  ---------------------------------------------------(ItemOperationCalculation)-------------------------------------
        *  --Global Variables:
        *    AveragePPM = CurrentPPM;
        *    or
        *    AveragePPM = CurrentPPM;
        *    AveragePPM = OverallFormedParts / OverallRunTime;
        *    AveragePPM = (float)Math.Round(AveragePPM, 2);
        *   
        *  --Method Variables:
        *    string CurrentPartsFormed_String = PartsFormed_TextBox.Text;
        *    double CurrentPartsFormed_Double = double.Parse(CurrentPartsFormed_String);
        *    or
        *    PartsManufacturedTotal_Double = double.Parse(PartsManufacturedTotal_String);
        *    TotalItemPartsManufactured = CurrentPartsFormed_Double + PartsManufacturedTotal_Double;
        *    string CurrentPPM_String = CurrentPPM_TextBox.Text;
        *    float CurrentPartsFormed_Float = (float)CurrentPartsFormed_Double;
        *    float CurrentPPM = float.Parse(CurrentPPM_String);
        *    float CurrentRunTime = CurrentPartsFormed_Float / CurrentPPM;
        *    string AveragePPM_String = AveragePPM_TextBox.Text;
        *    PastPPM = float.Parse(AveragePPM_String);
        *    float PartsManufacturedTotal_Float = (float)PartsManufacturedTotal_Double;
        *    float PreviousRunTime = PartsManufacturedTotal_Float / PastPPM;
        *    float OverallRunTime = PreviousRunTime + CurrentRunTime;
        *    float OverallFormedParts = PartsManufacturedTotal_Float + CurrentPartsFormed_Float;
        *    
        *  ---------------------------------------------------(ItemOperationDataStart)---------------------------------------
        *  --Method SQL Data:
        *    ("@ItemID", CurrentItemID);
        *    ("@OperationID", OperationsID.ToString());
        *    ("@ItemRunCount", ItemRunCount.ToString());
        *    ("@StartDateTime", Clock_TextBox.Text);
        *    ("@EmployeeName", User_TextBox.Text);
        *    ("@DMPID", DMPID_TextBox.Text);
        *    ("@BrakePress", BrakePress_ComboBox.Text);
        *  
        *  ---------------------------------------------------(ItemOperationDataEnd)-----------------------------------------
        *  --Method SQL Data:
        *    ("@OperationID", OperationsID);
        *    ("@EndDateTime", Clock_TextBox.Text);
        *    ("@PartsManufactured", PartsFormed_TextBox.Text);
        *    ("@PartsPerMinute", CurrentPPM_TextBox.Text);
        *  
        *  ---------------------------------------------------(ItemRunCounter)-----------------------------------------------
        *  --Global Variables:
        *    ItemRunCount = OperationRunCount + 1;
        *  
        *  --Method Variables:
        *    int OperationRunCount = (int)CountItemRun.ExecuteScalar();
        *    
        *  --Method SQL Data:  
        *    
        *  
        *  ---------------------------------------------------(LoadImage)----------------------------------------------------
        *  --Global Variables:
        *  
        *  
        *  --Method Variables:
        *    PictureBox ImageBox = Part_PictureBox;
        *    int i = 0;
        *    var DataImage = (Byte[])(ImageData.Tables[0].Rows[0][i]);
        *    var ImageStream = new MemoryStream(DataImage);
        *    
        *  ---------------------------------------------------(OperationIDCounter)-------------------------------------------
        *  --Global Variables:
        *    OperationsID = OperationCountID + 1;
        *    
        *  --Method Variables:
        *    int OperationCountID = (int)countOperation.ExecuteScalar();
        *    
        *  ---------------------------------------------------(OperationDataStart)-------------------------------------------
        *  --Global Variables:
        *    JobStartTime;
        *    
        *  --Method SQL Data:
        *    ("@ItemID", CurrentItemID);
        *    ("@OperationID", OperationsID.ToString());
        *    ("@RunDateTime", Clock_TextBox.Text);
        *    ("@EmployeeName", User_TextBox.Text);
        *    ("@DMPID", DMPID_TextBox.Text);
        *    ("@BrakePress", BrakePress_ComboBox.Text);
        *    
        *  ---------------------------------------------------(OperationDataEnd)---------------------------------------------
        *  --Global Variables:
        *  
        *  
        *  --Method Variables:
        *  
        *  --Method SQL Data:
        *    ("@OperationID", OperationsID.ToString());
        *    ("@PartsManufactured", PartsFormed_TextBox.Text);
        *    ("@PartsPerMinute", CurrentPPM_TextBox.Text);
        *    
        *  ---------------------------------------------------(ProgramListUpdate)--------------------------------------------
        *   --Global Variables:
        *  
        *  
        *  --Method Variables:
        *  
        *  --Method SQL Data:
        *    ("@ItemID", CurrentItemID);
        *    ("@PartsManufactured", PartsManufacturedTotal_String);
        *    ("@PartsPerMinute", PPMAverage.ToString());
        *    ("@TotalRuns", ItemRunCount);
        *    
        *  ---------------------------------------------------(RemoveSolution)-----------------------------------------------
        *  --Global Variables:
        *  
        *  
        *  --Method Variables:
        *    string ItemID = BP_Array[0];
        *    
        *  ---------------------------------------------------(RemoveSolution)-----------------------------------------------
        *      
        *  ---------------------------------------------------(RunningStatistics)--------------------------------------------
        *  --Global Variables:
        *    CurrentParts = int.Parse(PartsFormed_TextBox.Text);
        *    PartsRemaining = PartsNeeded - CurrentParts;
        *    
        *  --Method Variables:
        *    double HoursRemaining = 0;
        *    double MinutesRemaining = 0;
        *    string RemainingTime = "";
        *    int PartsNeeded = int.Parse(PartsNeeded_TextBox.Text);
        *    
        *  ---------------------------------------------------(RunMode)------------------------------------------------------
        *  --Global Variables:
        *  
        *  
        *  --Method Variables:
        *  ---------------------------------------------------(SearchForItemID)-----------------------------------------------
        *  --Global Variables:
        *    JobFound = true;
        *  
        *  --Method Variables:
        *    string SearchValue = CurrentItemID;
        *    
        *  --Methods:
        *    
        *    
        *  ---------------------------------------------------(Timer_Tick)---------------------------------------------------
        *  --Methods:
        *    RunMode();
        *    RunningStatistics();
        * 
        ********************************************************************************************************************/
        private void ClearMethod()
        {
            //RefreshJobs();
            ItemID_TextBox.Clear();
            JobID_TextBox.Clear();
            CustomerItemID_TextBox.Clear();
            Sequence_TextBox.Clear();
            Customer_TextBox.Clear();
            Customer_ComboBox.Text = "";
            Comp1_ComboBox.Text = "";
            Comp2_ComboBox.Text = "";
            Comp3_ComboBox.Text = "";
            Comp4_ComboBox.Text = "";
            //Quantity1_ComboBox.Text = "";
            //Quantity2_ComboBox.Text = "";
            //Quantity3_ComboBox.Text = "";
            Comp1_TextBox.Clear();
            Comp2_TextBox.Clear();
            Comp3_TextBox.Clear();
            Comp4_TextBox.Clear();
            Quantity1_TextBox.Clear();
            Quantity2_TextBox.Clear();
            Quantity3_TextBox.Clear();
            Quantity4_TextBox.Clear();
            PartsManufactured_TextBox.Clear();
            TotalRuns_TextBox.Clear();
            SetupTime_TextBox.Clear();
            PPM_TextBox.Clear();
            Fixture_TextBox.Clear();
            FixtureLocation_TextBox.Clear();
            Part_PictureBox.Image = null;
            RefreshJobs();
        }

        private void ConfirmFinished()
        {
            // Show and Hide Info
            // Item Information GroupBox
            // - Buttons
            AddJob_ButtonWasClicked = false;
            EditJob_ButtonWasClicked = false;
            RemoveJob_ButtonWasClicked = false;
            AddJob_Button.Show();
            EditJob_Button.Show();
            RemoveJob_Button.Show();
            Confirm_Button.Hide();
            Cancel_Button.Hide();
            Search_Button.Enabled = true;
            Clear_Button.Enabled = true;

            // - TextBoxes
            ItemID_TextBox.ReadOnly = false;
            JobID_TextBox.ReadOnly = false;
            Customer_TextBox.ReadOnly = true;
            Customer_TextBox.Show();
            CustomerItemID_TextBox.ReadOnly = true;
            Sequence_TextBox.ReadOnly = true;

            // - ComboBoxes
            Customer_TextBox.Show();
            Customer_ComboBox.Hide();

            // - CheckBoxes
            SearchItemID_CheckBox.Show();
            SearchJobID_CheckBox.Show();


            //Component Items GroupBox

            Comp1_ComboBox.Hide();
            Comp2_ComboBox.Hide();
            Comp3_ComboBox.Hide();
            Comp4_ComboBox.Hide();
            Quantity1_ComboBox.Hide();
            Quantity2_ComboBox.Hide();
            Quantity3_ComboBox.Hide();
            Quantity4_ComboBox.Hide();

            Comp1_TextBox.ReadOnly = true;
            Comp2_TextBox.ReadOnly = true;
            Comp3_TextBox.ReadOnly = true;
            Comp4_TextBox.ReadOnly = true;
            Quantity1_TextBox.ReadOnly = true;
            Quantity2_TextBox.ReadOnly = true;
            Quantity3_TextBox.ReadOnly = true;
            Quantity4_TextBox.ReadOnly = true;
            Comp1_TextBox.Show();
            Comp2_TextBox.Show();
            Comp3_TextBox.Show();
            Comp4_TextBox.Show();
            Quantity1_TextBox.Show();
            Quantity2_TextBox.Show();
            Quantity3_TextBox.Show();
            Quantity4_TextBox.Show();

            //Item Statistics GroupBox Always ReadOnly = true            

            //Folder Items GroupBox
            Fixture_TextBox.ReadOnly = true;
            FixtureLocation_TextBox.ReadOnly = true;

            JobListGridView.Enabled = true;
            Exit_Button.Show();
            ClearMethod();
        }

        private void LoadComponentData()
        {
            SqlConnection connection = new SqlConnection(SQL_Source);
            string LoginData = SQLComponentData;
            SqlDataAdapter DataAdapter = new SqlDataAdapter(LoginData, connection);
            SqlCommandBuilder commandBuilder = new SqlCommandBuilder(DataAdapter);
            DataSet Data = new DataSet();
            DataAdapter.Fill(Data);
            Component_DataGridView.DataSource = Data.Tables[0];

            int rows = 0;
            string ComponentCount = SQLComponentCount;
            SqlConnection count = new SqlConnection(SQL_Source);
            SqlCommand countRows = new SqlCommand(ComponentCount, count);
            count.Open();
            rows = (int)countRows.ExecuteScalar();
            count.Close();

            foreach (DataGridViewRow row in Component_DataGridView.Rows)
            {
                if (row.Index < rows)
                {
                    Comp1_ComboBox.Items.Add(row.Cells[0].Value.ToString());
                    Comp2_ComboBox.Items.Add(row.Cells[0].Value.ToString());
                    Comp3_ComboBox.Items.Add(row.Cells[0].Value.ToString());
                    Comp4_ComboBox.Items.Add(row.Cells[0].Value.ToString());
                    ComponentList.Add(row.Cells[0].Value.ToString());
                }
            }
        }

        private void DateParse()
        {
            // string DateValue = CompletionDate_TextBox.Text;
            // DateValue = DateValue.Replace(" 12:00:00 AM", "");
            //  CompletionDate_TextBox.Text = DateValue;
        }

        private void FindItemImage(object sender, EventArgs e)
        {            
            ItemID = ItemID_TextBox.Text;
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

        /*
        private void PartImage_DoWork(object sender, DoWorkEventArgs e)
        {
            if (PartImage.IsBusy != true)
            {
                FindItemImage(null, null);
            }
        }
        */

        private void PartImage_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {

        }

        private void JobHideShow()
        {
            //Item Information GroupBox
            // Show
            Confirm_Button.Show();
            Cancel_Button.Show();
            Customer_ComboBox.Show();


            // Hide
            SearchItemID_CheckBox.Hide();
            SearchJobID_CheckBox.Hide();
            Customer_TextBox.Hide();
            Exit_Button.Hide();


            //Component Items GroupBox
            // Show
            Comp1_ComboBox.Show();
            Comp2_ComboBox.Show();
            Comp3_ComboBox.Show();
            Comp4_ComboBox.Show();
            Quantity1_ComboBox.Show();
            Quantity2_ComboBox.Show();
            Quantity3_ComboBox.Show();
            Quantity4_ComboBox.Show();

            // Hide
            Comp1_TextBox.Hide();
            Comp2_TextBox.Hide();
            Comp3_TextBox.Hide();
            Comp4_TextBox.Hide();
            Quantity1_TextBox.Hide();
            Quantity2_TextBox.Hide();
            Quantity3_TextBox.Hide();
            Quantity4_TextBox.Hide();

            //Item Statistics GroupBox Always ReadOnly = true            
            // Show

            // Hide

            //Folder Items GroupBox
            // Show

            // Hide

        }

        private void JobListLogoff()
        {
            SqlConnection JobListLogoff = new SqlConnection(SQL_Source);
            SqlCommand Logoff = new SqlCommand();
            Logoff.CommandType = System.Data.CommandType.Text;
            Logoff.CommandText = "UPDATE [dbo].[LoginData] SET LogoutDateTime=@LogoutDateTime WHERE LoginDateTime=@LoginDateTime";
            Logoff.Connection = JobListLogoff;
            Logoff.Parameters.AddWithValue("@LoginDateTime", LoginTime.ToString());
            Logoff.Parameters.AddWithValue("@LogoutDateTime", Clock_TextBox.Text);
            JobListLogoff.Open();
            Logoff.ExecuteNonQuery();
            JobListLogoff.Close();
        }

        private void RefreshJobs()
        {
            try
            {
                SqlConnection connection = new SqlConnection(SQL_Source);
                SqlDataAdapter dataAdapter = new SqlDataAdapter(Refresh_Data, connection);
                SqlCommandBuilder commandBuilder = new SqlCommandBuilder(dataAdapter);
                ReportDataSet = new DataSet();
                dataAdapter.Fill(ReportDataSet);
                JobListGridView.DataSource = ReportDataSet.Tables[0];
            }
            catch
            {

            }
        }

        private void GroupBoxControlAddEditJob()
        {
            //Item Information GroupBox
            ItemID_TextBox.ReadOnly = false;
            JobID_TextBox.ReadOnly = false;
            Customer_TextBox.ReadOnly = false;
            CustomerItemID_TextBox.ReadOnly = false;
            Sequence_TextBox.ReadOnly = false;

            //Component Items GroupBox
            Comp1_TextBox.ReadOnly = true;
            Comp2_TextBox.ReadOnly = true;
            Comp3_TextBox.ReadOnly = true;
            Comp4_TextBox.ReadOnly = true;
            Quantity1_TextBox.ReadOnly = true;
            Quantity2_TextBox.ReadOnly = true;
            Quantity3_TextBox.ReadOnly = true;
            Quantity4_TextBox.ReadOnly = true;

            //Item Statistics GroupBox Always ReadOnly = true            

            //Folder Items GroupBox
            Fixture_TextBox.ReadOnly = false;
            FixtureLocation_TextBox.ReadOnly = false;
        }

        private void GroupBoxControlInactive()
        {
            //Item Information GroupBox
            ItemID_TextBox.ReadOnly = true;
            JobID_TextBox.ReadOnly = true;
            Customer_TextBox.ReadOnly = true;
            CustomerItemID_TextBox.ReadOnly = true;
            Sequence_TextBox.ReadOnly = true;
            Customer_ComboBox.Enabled = false;

            //Component Items GroupBox
            Comp1_TextBox.ReadOnly = true;
            Comp2_TextBox.ReadOnly = true;
            Comp3_TextBox.ReadOnly = true;
            Comp4_TextBox.ReadOnly = true;
            Quantity1_TextBox.ReadOnly = true;
            Quantity2_TextBox.ReadOnly = true;
            Quantity3_TextBox.ReadOnly = true;
            Quantity4_TextBox.ReadOnly = true;

            //Item Statistics GroupBox Always ReadOnly = true            

            //Folder Items GroupBox
            Fixture_TextBox.ReadOnly = true;
            FixtureLocation_TextBox.ReadOnly = true;
        }

        private void GroupBoxControlStart()
        {
            // Set ReadOnly
            //Item Information GroupBox
            ItemID_TextBox.ReadOnly = false;
            JobID_TextBox.ReadOnly = false;
            Customer_TextBox.ReadOnly = true;
            CustomerItemID_TextBox.ReadOnly = true;
            Sequence_TextBox.ReadOnly = true;
            AddJob_Button.Enabled = true;
            EditJob_Button.Enabled = true;
            RemoveJob_Button.Enabled = true;

            //Component Items GroupBox
            Comp1_ComboBox.Visible = false;
            Comp2_ComboBox.Visible = false;
            Comp3_ComboBox.Visible = false;
            Comp4_ComboBox.Visible = false;
            Quantity1_ComboBox.Visible = false;
            Quantity2_ComboBox.Visible = false;
            Quantity3_ComboBox.Visible = false;
            Quantity4_ComboBox.Visible = false;
            Comp1_TextBox.ReadOnly = true;
            Comp2_TextBox.ReadOnly = true;
            Comp3_TextBox.ReadOnly = true;
            Comp4_TextBox.ReadOnly = true;
            Comp1_TextBox.Visible = true;
            Comp2_TextBox.Visible = true;
            Comp3_TextBox.Visible = true;
            Comp4_TextBox.Visible = true;
            Quantity1_TextBox.ReadOnly = true;
            Quantity2_TextBox.ReadOnly = true;
            Quantity3_TextBox.ReadOnly = true;
            Quantity4_TextBox.ReadOnly = true;
            Quantity1_TextBox.Visible = true;
            Quantity2_TextBox.Visible = true;
            Quantity3_TextBox.Visible = true;
            Quantity4_TextBox.Visible = true;

            //Item Statistics GroupBox Always ReadOnly = true            

            //Folder Items GroupBox
            Fixture_TextBox.ReadOnly = true;
            FixtureLocation_TextBox.ReadOnly = true;

            // Back to ItemID
            ItemID_TextBox.Focus();
        }



        /********************************************************************************************************************
        * 
        * Methods End
        * 
        ********************************************************************************************************************/

        /********************************************************************************************************************
        * 
        * JobList End
        * 
        ********************************************************************************************************************/

        /********************************************************************************************************************
        * 
        * Methods in Testing Start
        * 
        ********************************************************************************************************************/

        private void Fixture_TextBox_Enter(object sender, EventArgs e)
        {
            if (AddJob_ButtonWasClicked == true || EditJob_ButtonWasClicked == true || RemoveJob_ButtonWasClicked == true)
            {
            }
            else
            {
                this.ActiveControl = null;
            }
        }

        private void FixtureLocation_TextBox_Enter(object sender, EventArgs e)
        {
            if (AddJob_ButtonWasClicked == true || EditJob_ButtonWasClicked == true || RemoveJob_ButtonWasClicked == true)
            {
            }
            else
            {
                this.ActiveControl = null;
            }
        }

        private void ItemID_TextBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                Search_Button_Click(null, null);
            }
        }

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

        /********************************************************************************************************************
        * 
        * Methods in Testing End
        * 
        ********************************************************************************************************************/

        private void CreateExcelFile(object sender, EventArgs e)
        {
            string ReportName = DateTime.Now.ToShortDateString() + "_" + DateTime.Now.Millisecond + "_";
            ReportName = ReportName.Replace("/", "_");
            ReportName = ReportName.Replace(":", "_");

            // Excel Initialize
            ReportApp = new Excel.Application();
            ReportApp.Visible = false;
            ReportWB = (Excel._Workbook)(ReportApp.Workbooks.Add(""));
            ReportWS = (Excel._Worksheet)ReportWB.ActiveSheet;

            string[] ColumnNames = new string[JobListGridView.Columns.Count];
            int ExcelColumns = 1;
            foreach (DataGridViewColumn dc in JobListGridView.Columns)
            {
                ReportWS.Cells[1, ExcelColumns] = dc.Name;
                ExcelColumns++;
            }

            ReportWS.get_Range("A1", "S1").Font.Bold = true;
            ReportRange = ReportWS.get_Range("A1", "S1");
            ReportRange.EntireColumn.AutoFit();

            for (int i = 0; i < ReportDataSet.Tables[0].Rows.Count; i++)
            {
                // to do: format datetime values before printing
                for (int j = 0; j < ReportDataSet.Tables[0].Columns.Count; j++)
                {
                    ReportWS.Cells[(i + 2), (j + 1)] = ReportDataSet.Tables[0].Rows[i][j];
                    CreateExcel.ReportProgress(i);
                }
            }
            ReportRange = ReportWS.get_Range("A1", "S1");
            ReportWS.get_Range("A1", "S1").AutoFilter(1, Type.Missing, Excel.XlAutoFilterOperator.xlAnd, Type.Missing, true);
            ReportRange.EntireColumn.AutoFit();
            string ReportPDFName = ReportCell + "Spot_Weld_SQL" + ".xlsx";
            ReportRange = ReportWS.get_Range("A1", "S" + RowCount.ToString());
            //ReportRange = ReportWS.get_Range("A5", "X10");
            foreach (Microsoft.Office.Interop.Excel.Range cell in ReportRange.Cells)
            {
                cell.BorderAround2();
            }
            /*
            SaveFileDialog saveFile = new SaveFileDialog();
            saveFile.Filter = "Excel Files (*.xls)|*.xls|All files (*.*)|*.*";
            saveFile.FileName = ReportPDFName;
            if (saveFile.ShowDialog() != DialogResult.OK)
            {
                ReportWS.Delete();
                ReportWB.Close();
            }
            else
            {
                ExcelFileLocation = saveFile.FileName;
                ReportWS.SaveAs(ExcelFileLocation);
                //ExcelOpen.StartInfo.FileName = ExcelFileLocation;
                //ExcelOpen.Start();
            }
            */
            ExcelFileLocation = @"\\insidedmp.com\Corporate\OH\OH Common\Engineering\Job List\SQL Data Tables\" + ReportPDFName;
            if (File.Exists(ExcelFileLocation))
            {
                bool tryAgain = true;
                try
                {
                    File.Delete(ExcelFileLocation);
                }
                catch (IOException)
                {
                    tryAgain = false;
                }
                if (tryAgain == false)
                {
                    ExcelFileLocation = @"\\insidedmp.com\Corporate\OH\OH Common\Engineering\Job List\SQL Data Tables\" + ReportName + "_" + ReportPDFName;
                }
            }

            ReportWS.SaveAs(ExcelFileLocation, Excel.XlFileFormat.xlOpenXMLWorkbook, null, null, false, false, false, false, false, false);
            //ReportWS.SaveAs(ExcelFileLocation);
            ReportWB.Close();
        }

        private void CreateExcelFile_RunWorkerComplete(object sender, RunWorkerCompletedEventArgs e)
        {
            // Process.Start(ExcelFileLocation);
            // MessageBox.Show("Saved");
            Saving_Label.Text = "Saved";
            Saving_ProgressBar.Value = 0;
            Saving_ProgressBar.Hide();
            Clear_Button.Enabled = true;
            ChangeCell_Button.Enabled = true;
            Exit_Button.Enabled = true;
            Refresh_Button.Enabled = true;
        }

        private void CreateExcelFile_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            Saving_ProgressBar.Value = e.ProgressPercentage;
        }

        private void CreateExcel_Start()
        {
            Clear_Button.Enabled = false;
            ChangeCell_Button.Enabled = false;
            Exit_Button.Enabled = false;
            Refresh_Button.Enabled = false;
            Saving_Label.Show();
            Saving_Label.Text = "Saving";
            Saving_ProgressBar.Show();
        }
        private void CreateExcel_End()
        {
            Saving_Label.Hide();
            Saving_ProgressBar.Hide();
        }

        private void Row_TotalCount()
        {
            if(CompanyCell_ComboBox.Text == "Navistar")
            {
                RowCountString = "SELECT COUNT(*) FROM [dbo].[Navistar_Item_Data]";
            }
            if (CompanyCell_ComboBox.Text == "Paccar")
            {
                RowCountString = "SELECT COUNT(*) FROM [dbo].[Paccar_Item_Data]";
            }
            try
            {
                string CheckCardCountString = RowCountString;
                SqlConnection CheckCountTotalConnection = new SqlConnection(SQL_Source);
                SqlCommand CheckCountTotalCommand = new SqlCommand(CheckCardCountString, CheckCountTotalConnection);
                CheckCountTotalConnection.Open();
                int CheckCardCountOperationTotal = (int)CheckCountTotalCommand.ExecuteScalar();
                CheckCountTotalConnection.Close();
                RowCount = CheckCardCountOperationTotal + 1;
            }
            catch (Exception e)
            {
                MessageBox.Show(e.ToString());
            }
            Saving_ProgressBar.Maximum = RowCount + 1;
        }
        
        private void JobListGridView_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
            {
                DataGridViewRow Row = JobListGridView.Rows[e.RowIndex];
                ItemID_TextBox.Text = Row.Cells[0].Value.ToString();
                Customer_TextBox.Text = Row.Cells[1].Value.ToString();
                Customer_ComboBox.Text = Row.Cells[1].Value.ToString();
                CustomerItemID_TextBox.Text = Row.Cells[2].Value.ToString();
                JobID_TextBox.Text = Row.Cells[3].Value.ToString();
                Sequence_TextBox.Text = Row.Cells[4].Value.ToString();
                Fixture_TextBox.Text = Row.Cells[5].Value.ToString();
                FixtureLocation_TextBox.Text = Row.Cells[6].Value.ToString();
                Comp1_TextBox.Text = Row.Cells[7].Value.ToString();
                Comp1_ComboBox.Text = Row.Cells[7].Value.ToString();
                Quantity1_TextBox.Text = Row.Cells[8].Value.ToString();
                Quantity1_ComboBox.Text = Row.Cells[8].Value.ToString();
                Comp2_TextBox.Text = Row.Cells[9].Value.ToString();
                Comp2_ComboBox.Text = Row.Cells[9].Value.ToString();
                Quantity2_TextBox.Text = Row.Cells[10].Value.ToString();
                Quantity2_ComboBox.Text = Row.Cells[10].Value.ToString();
                Comp3_TextBox.Text = Row.Cells[11].Value.ToString();
                Comp3_ComboBox.Text = Row.Cells[11].Value.ToString();
                Quantity3_TextBox.Text = Row.Cells[12].Value.ToString();
                Quantity3_ComboBox.Text = Row.Cells[12].Value.ToString();
                Comp4_TextBox.Text = Row.Cells[13].Value.ToString();
                Comp4_ComboBox.Text = Row.Cells[13].Value.ToString();
                Quantity4_TextBox.Text = Row.Cells[14].Value.ToString();
                Quantity4_ComboBox.Text = Row.Cells[14].Value.ToString();
                TotalRuns_TextBox.Text = Row.Cells[15].Value.ToString();
                PartsManufactured_TextBox.Text = Row.Cells[16].Value.ToString();
                PPM_TextBox.Text = Row.Cells[17].Value.ToString();
                SetupTime_TextBox.Text = Row.Cells[18].Value.ToString();
                DateParse();
                //FindItemImage();
                if (PartImage.IsBusy != true)
                {
                    PartImage.RunWorkerAsync();
                }
                JobListGridView.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
                GroupBoxControlStart();
            }
        }       

        private void JobListGridView_KeyUp(object sender, KeyEventArgs e)
        {
            DataGridViewRow Row = JobListGridView.CurrentRow;
            if (e.KeyCode == Keys.Up || e.KeyCode == Keys.Down)
            {
                ItemID_TextBox.Text = Row.Cells[0].Value.ToString();
                Customer_TextBox.Text = Row.Cells[1].Value.ToString();
                Customer_ComboBox.Text = Row.Cells[1].Value.ToString();
                CustomerItemID_TextBox.Text = Row.Cells[2].Value.ToString();
                JobID_TextBox.Text = Row.Cells[3].Value.ToString();
                Sequence_TextBox.Text = Row.Cells[4].Value.ToString();
                Fixture_TextBox.Text = Row.Cells[5].Value.ToString();
                FixtureLocation_TextBox.Text = Row.Cells[6].Value.ToString();
                Comp1_TextBox.Text = Row.Cells[7].Value.ToString();
                Comp1_ComboBox.Text = Row.Cells[7].Value.ToString();
                Quantity1_TextBox.Text = Row.Cells[8].Value.ToString();
                Quantity1_ComboBox.Text = Row.Cells[8].Value.ToString();
                Comp2_TextBox.Text = Row.Cells[9].Value.ToString();
                Comp2_ComboBox.Text = Row.Cells[9].Value.ToString();
                Quantity2_TextBox.Text = Row.Cells[10].Value.ToString();
                Quantity2_ComboBox.Text = Row.Cells[10].Value.ToString();
                Comp3_TextBox.Text = Row.Cells[11].Value.ToString();
                Comp3_ComboBox.Text = Row.Cells[11].Value.ToString();
                Quantity3_TextBox.Text = Row.Cells[12].Value.ToString();
                Quantity3_ComboBox.Text = Row.Cells[12].Value.ToString();
                Comp4_TextBox.Text = Row.Cells[13].Value.ToString();
                Comp4_ComboBox.Text = Row.Cells[13].Value.ToString();
                Quantity4_TextBox.Text = Row.Cells[14].Value.ToString();
                Quantity4_ComboBox.Text = Row.Cells[14].Value.ToString();
                TotalRuns_TextBox.Text = Row.Cells[15].Value.ToString();
                PartsManufactured_TextBox.Text = Row.Cells[16].Value.ToString();
                PPM_TextBox.Text = Row.Cells[17].Value.ToString();
                SetupTime_TextBox.Text = Row.Cells[18].Value.ToString();
                DateParse();
                //FindItemImage();
                if (PartImage.IsBusy != true)
                {
                    PartImage.RunWorkerAsync();
                }
                JobListGridView.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
                GroupBoxControlStart();
            }
        }

        //

        private void ImportData_Button_Click(object sender, EventArgs e)
        {
            DataSet ImportResult;
            IExcelDataReader ExcelReader;
            OpenFileDialog OpenExcelFile = new OpenFileDialog();
            try
            {
                OpenExcelFile.InitialDirectory = @"\\insidedmp.com\Corporate\OH\OH Common\Engineering\Job List\SQL Data Tables\";
                OpenExcelFile.FileName = CompanyCell_ComboBox.Text + "*";
                if (OpenExcelFile.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    FileStream ExcelFileStream = File.Open(OpenExcelFile.FileName, FileMode.Open, FileAccess.Read);
                    ExcelReader = ExcelReaderFactory.CreateOpenXmlReader(ExcelFileStream);
                    ImportResult = ExcelReader.AsDataSet(new ExcelDataSetConfiguration()
                    {
                        ConfigureDataTable = (_) => new ExcelDataTableConfiguration()
                        {
                            UseHeaderRow = true,
                            ReadHeaderRow = (rowReader) => {
                            }
                        }
                    });
                    ExcelReader.Close();
                    JobListGridView.DataSource = ImportResult.Tables["Sheet1"];
                    //SaveImport_Button.Show();
                    //CancelImport_Button.Show();
                    //DatabaseImport_Button.Hide();
                    DatabaseImportStart_Controls();
                }
            }
            catch (IOException)
            {
                MessageBox.Show("File: " + OpenExcelFile.FileName.ToString() + " is \ncurrently opened by another user and must be closed before import");
            }
            SecondLoadTest();
        }

        private void ImportSave_Button_Click(object sender, EventArgs e)
        {
            ExcelFileImportCreate.RunWorkerAsync();
        }

        private void CancelImport_Button_Click(object sender, EventArgs e)
        {

        }

        private void ImportUpdateSQL(object sender, EventArgs e)
        {
            DeleteTest();
            WriteTest();
            SecondLoadTest();
        }

        private void ImportUpdateSQL_Complete(object sender, RunWorkerCompletedEventArgs e)
        {
            Saving_Label.Text = "Import Complete";
            RefreshJobs();

            CreateExcel_Start();
            ExcelFileImportSave.RunWorkerAsync();
        }

        private void ImportUpdateExcel(object sender, EventArgs e)
        {
            string ReportName = DateTime.Now.ToShortDateString() + "_" + DateTime.Now.Millisecond + "_";
            ReportName = ReportName.Replace("/", "_");
            ReportName = ReportName.Replace(":", "_");

            // Excel Initialize
            ReportApp = new Excel.Application();
            ReportApp.Visible = false;
            ReportWB = (Excel._Workbook)(ReportApp.Workbooks.Add(""));
            ReportWS = (Excel._Worksheet)ReportWB.ActiveSheet;

            string[] ColumnNames = new string[JobListGridView.Columns.Count];
            int ExcelColumns = 1;
            foreach (DataGridViewColumn dc in JobListGridView.Columns)
            {
                ReportWS.Cells[1, ExcelColumns] = dc.Name;
                ExcelColumns++;
            }

            ReportWS.get_Range("A1", "S1").Font.Bold = true;
            ReportRange = ReportWS.get_Range("A1", "S1");
            ReportRange.EntireColumn.AutoFit();

            for (int i = 0; i < ReportDataSet.Tables[0].Rows.Count; i++)
            {
                // to do: format datetime values before printing
                for (int j = 0; j < ReportDataSet.Tables[0].Columns.Count; j++)
                {
                    ReportWS.Cells[(i + 2), (j + 1)] = ReportDataSet.Tables[0].Rows[i][j];
                    ExcelFileImportSave.ReportProgress(i);
                }
            }
            ReportRange = ReportWS.get_Range("A1", "S1");
            ReportWS.get_Range("A1", "S1").AutoFilter(1, Type.Missing, Excel.XlAutoFilterOperator.xlAnd, Type.Missing, true);
            ReportRange.EntireColumn.AutoFit();
            string ReportPDFName = ReportCell + "Spot_Weld_SQL" + ".xlsx";
            ReportRange = ReportWS.get_Range("A1", "S" + RowCount.ToString());
            //ReportRange = ReportWS.get_Range("A5", "X10");
            foreach (Microsoft.Office.Interop.Excel.Range cell in ReportRange.Cells)
            {
                cell.BorderAround2();
            }
            /*
            SaveFileDialog saveFile = new SaveFileDialog();
            saveFile.Filter = "Excel Files (*.xls)|*.xls|All files (*.*)|*.*";
            saveFile.FileName = ReportPDFName;
            if (saveFile.ShowDialog() != DialogResult.OK)
            {
                ReportWS.Delete();
                ReportWB.Close();
            }
            else
            {
                ExcelFileLocation = saveFile.FileName;
                ReportWS.SaveAs(ExcelFileLocation);
                //ExcelOpen.StartInfo.FileName = ExcelFileLocation;
                //ExcelOpen.Start();
            }
            */
            ExcelFileLocation = @"\\insidedmp.com\Corporate\OH\OH Common\Engineering\Job List\SQL Data Tables\" + ReportPDFName;
            if (File.Exists(ExcelFileLocation))
            {
                bool tryAgain = true;
                try
                {
                    File.Delete(ExcelFileLocation);
                }
                catch (IOException)
                {
                    tryAgain = false;
                }
                if (tryAgain == false)
                {
                    ExcelFileLocation = @"\\insidedmp.com\Corporate\OH\OH Common\Engineering\Job List\SQL Data Tables\" + ReportName + "_" + ReportPDFName;
                }
            }

            ReportWS.SaveAs(ExcelFileLocation, Excel.XlFileFormat.xlOpenXMLWorkbook, null, null, false, false, false, false, false, false);
            //ReportWS.SaveAs(ExcelFileLocation);
            ReportWB.Close();
        }

        private void ImportUpdateExcel_Complete(object sender, RunWorkerCompletedEventArgs e)
        {
            Saving_ProgressBar.Value = 0;
            Saving_ProgressBar.Hide();
            Saving_Label.Hide();
            ChangeCell_Button.Enabled = true;
            Search_Button.Enabled = true;
            Clear_Button.Enabled = true;
            Refresh_Button.Enabled = true;
            AddJob_Button.Enabled = true;
            EditJob_Button.Enabled = true;
            RemoveJob_Button.Enabled = true;
            ImportData_Button.Enabled = true;
            ImportSave_Button.Enabled = true;
            CancelImport_Button.Enabled = true;
            ImportSave_Button.Hide();
            CancelImport_Button.Hide();
            Exit_Button.Enabled = true;
            //MessageBox.Show("Complete");
        }

        private void ImportUpdateExcel_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            Saving_ProgressBar.Value = e.ProgressPercentage;
        }

        private void DeleteTest()
        {
            try
            {
                for (int i = 0; i < ImportCurrentRows; i++)
                {
                    SqlConnection Job_Connection = new SqlConnection(SQL_Source);
                    SqlCommand Delete_Job = new SqlCommand();
                    Delete_Job.CommandType = System.Data.CommandType.Text;
                    Delete_Job.CommandText = SQLRemoveCommand;
                    Delete_Job.Connection = Job_Connection;
                    Delete_Job.Parameters.AddWithValue("@ItemID", ItemID_CurrentArray[i].ToString());
                    Delete_Job.Parameters.AddWithValue("@JobID", JobID_CurrentArray[i].ToString());
                    Delete_Job.Parameters.AddWithValue("@Sequence", Sequence_CurrentArray[i].ToString());
                    Job_Connection.Open();
                    Delete_Job.ExecuteNonQuery();
                    Job_Connection.Close();
                }
            }
            catch (SqlException ExceptionValue)
            {
                int ErrorNumber = ExceptionValue.Number;
                MessageBox.Show("Error Importing Job" + "\n" + "Error Code: " + ErrorNumber.ToString());
            }
            FirstLoadTest();
        }

        private void WriteTest()
        {
            try
            {
                for (int i = 0; i < ImportUpdatedRows; i++)
                {
                    SqlConnection Job_Connection = new SqlConnection(SQL_Source);
                    SqlCommand Add_Job = new SqlCommand();
                    Add_Job.CommandType = System.Data.CommandType.Text;
                    Add_Job.CommandText = SQLAddCommand;
                    Add_Job.Connection = Job_Connection;
                    
                    Add_Job.Parameters.AddWithValue("@ItemID", ItemID_UpdatedArray[i].ToString());
                    Add_Job.Parameters.AddWithValue("@Customer", Customer_UpdatedArray[i].ToString());
                    Add_Job.Parameters.AddWithValue("@CustomerItemID", CustomerItemID_UpdatedArray[i].ToString());
                    Add_Job.Parameters.AddWithValue("@JobID", JobID_UpdatedArray[i].ToString());
                    Add_Job.Parameters.AddWithValue("@Sequence", Sequence_UpdatedArray[i].ToString());
                    Add_Job.Parameters.AddWithValue("@Fixture", Fixture_UpdatedArray[i].ToString());
                    Add_Job.Parameters.AddWithValue("@FixtureLocation", FixtureLocation_UpdatedArray[i].ToString());
                    Add_Job.Parameters.AddWithValue("@Component1", Component1_UpdatedArray[i].ToString());
                    Add_Job.Parameters.AddWithValue("@Quantity1", Quantity1_UpdatedArray[i].ToString());
                    Add_Job.Parameters.AddWithValue("@Component2", Component2_UpdatedArray[i].ToString());
                    Add_Job.Parameters.AddWithValue("@Quantity2", Quantity2_UpdatedArray[i].ToString());
                    Add_Job.Parameters.AddWithValue("@Component3", Component3_UpdatedArray[i].ToString());
                    Add_Job.Parameters.AddWithValue("@Quantity3", Quantity3_UpdatedArray[i].ToString());
                    Add_Job.Parameters.AddWithValue("@Component4", Component4_UpdatedArray[i].ToString());
                    Add_Job.Parameters.AddWithValue("@Quantity4", Quantity4_UpdatedArray[i].ToString());
                    Add_Job.Parameters.AddWithValue("@TotalRuns", TotalRuns_UpdatedArray[i].ToString());
                    Add_Job.Parameters.AddWithValue("@PartsManufactured", PartsManufactured_UpdatedArray[i].ToString());
                    Add_Job.Parameters.AddWithValue("@PartsPerMinute", PartsPerMinute_UpdatedArray[i].ToString());
                    Add_Job.Parameters.AddWithValue("@SetupTime", SetupTime_UpdatedArray[i].ToString());

                    Job_Connection.Open();
                    Add_Job.ExecuteNonQuery();
                    Job_Connection.Close();
                }
                //MessageBox.Show("Import Successful");
                //SaveImport_Button.Hide();
                //CancelImport_Button.Hide();
                //DatabaseImport_Button.Show();
                //DatabaseImportEnd_Controls();
            }
            catch (SqlException ExceptionValue)
            {
                int ErrorNumber = ExceptionValue.Number;
                if (ErrorNumber.Equals(2627))
                {
                    MessageBox.Show("Item ID: " + ItemID_TextBox.Text + " is Already on this List");
                }
                else if (ErrorNumber.Equals(245))
                {
                    MessageBox.Show("Item ID Can Only Contain Numbers");
                }
                else
                {
                    MessageBox.Show("Unable to Add Job. Please Try Again." + "\n" + "Error Code: " + ErrorNumber.ToString());
                }
            }
        }

        private void DatabaseImportStart_Controls()
        {
            ImportSave_Button.Show();
            CancelImport_Button.Show();
            ImportData_Button.Enabled = false;

            ChangeCell_Button.Enabled = false;
            Search_Button.Enabled = false;
            Clear_Button.Enabled = false;
            Refresh_Button.Enabled = false;
            AddJob_Button.Enabled = false;
            EditJob_Button.Enabled = false;
            RemoveJob_Button.Enabled = false;
            Exit_Button.Enabled = false;
        }

        private void DatabaseImportEnd_Controls()
        {
            ImportSave_Button.Hide();
            CancelImport_Button.Hide();
            ImportData_Button.Enabled = true;

            ChangeCell_Button.Enabled = true;
            Search_Button.Enabled = true;
            Clear_Button.Enabled = true;
            Refresh_Button.Enabled = true;
            AddJob_Button.Enabled = true;
            EditJob_Button.Enabled = true;
            RemoveJob_Button.Enabled = true;
            Exit_Button.Enabled = true;
        }

        private void FirstLoadTest()
        {
            try
            {
                string CheckCardCountString = RowCountString;
                SqlConnection CheckCountTotalConnection = new SqlConnection(SQL_Source);
                SqlCommand CheckCountTotalCommand = new SqlCommand(CheckCardCountString, CheckCountTotalConnection);
                CheckCountTotalConnection.Open();
                int CheckCardCountOperationTotal = (int)CheckCountTotalCommand.ExecuteScalar();
                CheckCountTotalConnection.Close();
                RowCount = CheckCardCountOperationTotal;
            }
            catch (Exception e)
            {
                MessageBox.Show(e.ToString());
            }

            Array.Clear(ItemID_CurrentArray, 0, RowCount);
            Array.Clear(JobID_CurrentArray, 0, RowCount);
            Array.Clear(Sequence_CurrentArray, 0, RowCount);
            ImportCurrentRows = 0;
            foreach (DataGridViewRow Row in JobListGridView.Rows)
            {
                ItemID_CurrentArray[ImportCurrentRows] = Row.Cells[0].Value.ToString();
                JobID_CurrentArray[ImportCurrentRows] = Row.Cells[3].Value.ToString();
                Sequence_CurrentArray[ImportCurrentRows] = Row.Cells[4].Value.ToString();
                ImportCurrentRows++;
            }
        }

        private void SecondLoadTest()
        {
            try
            {
                string CheckCardCountString = RowCountString;
                SqlConnection CheckCountTotalConnection = new SqlConnection(SQL_Source);
                SqlCommand CheckCountTotalCommand = new SqlCommand(CheckCardCountString, CheckCountTotalConnection);
                CheckCountTotalConnection.Open();
                int CheckCardCountOperationTotal = (int)CheckCountTotalCommand.ExecuteScalar();
                CheckCountTotalConnection.Close();
                RowCount = CheckCardCountOperationTotal;
            }
            catch (Exception e)
            {
                MessageBox.Show(e.ToString());
            }
            Array.Clear(ItemID_UpdatedArray, 0, ItemID_UpdatedArray.Length);
            Array.Clear(Customer_UpdatedArray, 0, Customer_UpdatedArray.Length);
            Array.Clear(CustomerItemID_UpdatedArray, 0, CustomerItemID_UpdatedArray.Length);
            Array.Clear(JobID_UpdatedArray, 0, JobID_UpdatedArray.Length);
            Array.Clear(Sequence_UpdatedArray, 0, Sequence_UpdatedArray.Length);
            Array.Clear(Fixture_UpdatedArray, 0, Fixture_UpdatedArray.Length);
            Array.Clear(FixtureLocation_UpdatedArray, 0, FixtureLocation_UpdatedArray.Length);
            Array.Clear(Component1_UpdatedArray, 0, Component1_UpdatedArray.Length);
            Array.Clear(Quantity1_UpdatedArray, 0, Quantity1_UpdatedArray.Length);
            Array.Clear(Component2_UpdatedArray, 0, Component2_UpdatedArray.Length);
            Array.Clear(Quantity2_UpdatedArray, 0, Quantity2_UpdatedArray.Length);
            Array.Clear(Component3_UpdatedArray, 0, Component3_UpdatedArray.Length);
            Array.Clear(Quantity3_UpdatedArray, 0, Quantity3_UpdatedArray.Length);
            Array.Clear(Component4_UpdatedArray, 0, Component4_UpdatedArray.Length);
            Array.Clear(Quantity4_UpdatedArray, 0, Quantity4_UpdatedArray.Length);
            Array.Clear(TotalRuns_UpdatedArray, 0, TotalRuns_UpdatedArray.Length);
            Array.Clear(PartsManufactured_UpdatedArray, 0, PartsManufactured_UpdatedArray.Length);
            Array.Clear(PartsPerMinute_UpdatedArray, 0, PartsPerMinute_UpdatedArray.Length);
            Array.Clear(SetupTime_UpdatedArray, 0, SetupTime_UpdatedArray.Length);

            ImportUpdatedRows = 0;
            foreach (DataGridViewRow Row in JobListGridView.Rows)
            {
                ItemID_UpdatedArray[ImportUpdatedRows] = Row.Cells[0].Value.ToString();
                Customer_UpdatedArray[ImportUpdatedRows] = Row.Cells[1].Value.ToString();
                CustomerItemID_UpdatedArray[ImportUpdatedRows] = Row.Cells[2].Value.ToString();
                JobID_UpdatedArray[ImportUpdatedRows] = Row.Cells[3].Value.ToString();
                Sequence_UpdatedArray[ImportUpdatedRows] = Row.Cells[4].Value.ToString();
                Fixture_UpdatedArray[ImportUpdatedRows] = Row.Cells[5].Value.ToString();
                FixtureLocation_UpdatedArray[ImportUpdatedRows] = Row.Cells[6].Value.ToString();
                Component1_UpdatedArray[ImportUpdatedRows] = Row.Cells[7].Value.ToString();
                Quantity1_UpdatedArray[ImportUpdatedRows] = Row.Cells[8].Value.ToString();
                Component2_UpdatedArray[ImportUpdatedRows] = Row.Cells[9].Value.ToString();
                Quantity2_UpdatedArray[ImportUpdatedRows] = Row.Cells[10].Value.ToString();
                Component3_UpdatedArray[ImportUpdatedRows] = Row.Cells[11].Value.ToString();
                Quantity3_UpdatedArray[ImportUpdatedRows] = Row.Cells[12].Value.ToString();
                Component4_UpdatedArray[ImportUpdatedRows] = Row.Cells[13].Value.ToString();
                Quantity4_UpdatedArray[ImportUpdatedRows] = Row.Cells[14].Value.ToString();
                TotalRuns_UpdatedArray[ImportUpdatedRows] = Row.Cells[15].Value.ToString();
                PartsManufactured_UpdatedArray[ImportUpdatedRows] = Row.Cells[16].Value.ToString();
                PartsPerMinute_UpdatedArray[ImportUpdatedRows] = Row.Cells[17].Value.ToString();
                SetupTime_UpdatedArray[ImportUpdatedRows] = Row.Cells[18].Value.ToString();
                ImportUpdatedRows++;               
            }
        }

        private void CustomSave_CheckBox_CheckedChanged(object sender, EventArgs e)
        {
            if(CustomSave_CheckBox.Checked == true && CustomExcelSave == false)
            {
                CustomExcelSave = true;
            }
            else if(CustomSave_CheckBox.Checked == false && CustomExcelSave == true)
            {
                CustomExcelSave = false;
            }
        }

        private void Comp1_ComboBox_TextChanged(object sender, EventArgs e)
        {
            if(Update == true)
            {
                var txt = Comp1_ComboBox.Text;
                var list = from d in ComponentList.ToArray() where d.ToUpper().StartsWith(Comp1_ComboBox.Text.ToUpper()) select d;
                if (list.Count() > 0)
                {
                    Comp1_ComboBox.DataSource = list.ToList();
                    //comboBox1.SelectedIndex = 0;
                    var sText = Comp1_ComboBox.Items[0].ToString();
                    Comp1_ComboBox.SelectionStart = txt.Length;
                    Comp1_ComboBox.SelectionLength = sText.Length - txt.Length;
                    return;
                }
                else
                {
                    Comp1_ComboBox.SelectionStart = txt.Length;
                }
            }
        }

        private void Comp2_ComboBox_TextChanged(object sender, EventArgs e)
        {
            var txt = Comp2_ComboBox.Text;
            var list = from d in ComponentList.ToArray() where d.ToUpper().StartsWith(Comp2_ComboBox.Text.ToUpper()) select d;
            if (list.Count() > 0)
            {
                Comp2_ComboBox.DataSource = list.ToList();
                //comboBox1.SelectedIndex = 0;
                var sText = Comp2_ComboBox.Items[0].ToString();
                Comp2_ComboBox.SelectionStart = txt.Length;
                Comp2_ComboBox.SelectionLength = sText.Length - txt.Length;
                return;
            }
            else
            {
                Comp2_ComboBox.SelectionStart = txt.Length;
            }
        }

        private void Comp3_ComboBox_TextChanged(object sender, EventArgs e)
        {
            var txt = Comp3_ComboBox.Text;
            var list = from d in ComponentList.ToArray() where d.ToUpper().StartsWith(Comp3_ComboBox.Text.ToUpper()) select d;
            if (list.Count() > 0)
            {
                Comp3_ComboBox.DataSource = list.ToList();
                //comboBox1.SelectedIndex = 0;
                var sText = Comp3_ComboBox.Items[0].ToString();
                Comp3_ComboBox.SelectionStart = txt.Length;
                Comp3_ComboBox.SelectionLength = sText.Length - txt.Length;
                return;
            }
            else
            {
                Comp3_ComboBox.SelectionStart = txt.Length;
            }
        }

        private void Comp1_ComboBox_KeyDown(object sender, KeyEventArgs e)
        {
            //Update = false;
            int length = Comp1_ComboBox.Text.Length;
            if (e.KeyCode == Keys.Back)
            {
                int sStart = Comp1_ComboBox.SelectionStart;
                //Comp1_ComboBox.Text = Comp1_ComboBox.Text.Remove(sStart, (length - sStart));
            }
            /*
            if (e.KeyCode == Keys.Back)
            {
                int length = Comp1_ComboBox.Text.Length;
                int sStart = Comp1_ComboBox.SelectionStart;
                if (sStart > 0)
                {
                    sStart--;

                }
                if (sStart == 0)
                {
                    Comp1_ComboBox.Text = "";
                }
                else
                {
                    //Comp1_ComboBox.Text = Comp1_ComboBox.Text.Substring(0, sStart);
                    Comp1_ComboBox.Text = Comp1_ComboBox.Text.Remove(sStart, (length - sStart));
                    TotalRuns_TextBox.Text = Comp1_ComboBox.Text.Remove(sStart, (length - sStart));
                }
                e.Handled = true;
            }
            else
            {
                Update = true;
            }
            */
        }

        private void Comp1_ComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            Update = false;
        }

        private void Comp1_ComboBox_TextUpdate(object sender, EventArgs e)
        {
            Update = true;
        }

        private void Comp1_ComboBox_KeyUp(object sender, KeyEventArgs e)
        {

        }

        private void Comp2_ComboBox_KeyDown(object sender, KeyEventArgs e)
        {
            int length = Comp2_ComboBox.Text.Length;
            if (e.KeyCode == Keys.Back)
            {
                int sStart = Comp2_ComboBox.SelectionStart;
                if (sStart > 0)
                {
                    sStart--;

                }
                if (sStart == 0)
                {
                    Comp2_ComboBox.Text = "";
                }
                else
                {
                    //Comp2_ComboBox.Text = Comp2_ComboBox.Text.Substring(0, sStart);
                    Comp2_ComboBox.Text = Comp2_ComboBox.Text.Remove(sStart, (length - sStart));
                }
                e.Handled = true;
            }
        }

        private void Comp2_ComboBox_KeyUp(object sender, KeyEventArgs e)
        {

        }

        private void Comp3_ComboBox_KeyDown(object sender, KeyEventArgs e)
        {
            int length = Comp3_ComboBox.Text.Length;
            if (e.KeyCode == Keys.Back)
            {
                int sStart = Comp3_ComboBox.SelectionStart;
                if (sStart > 0)
                {
                    sStart--;

                }
                if (sStart == 0)
                {
                    Comp3_ComboBox.Text = "";
                }
                else
                {
                    //Comp3_ComboBox.Text = Comp3_ComboBox.Text.Substring(0, sStart);
                    Comp3_ComboBox.Text = Comp3_ComboBox.Text.Remove(sStart, (length - sStart));
                }
                e.Handled = true;
            }
        }

        private void Comp3_ComboBox_KeyUp(object sender, KeyEventArgs e)
        {

        }

        private void Comp4_TextBox_TextChanged(object sender, EventArgs e)
        {

        }
    }
}
