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

namespace DMP_Spot_Weld_Application
{
    public partial class Cell_Manager : Form
    {
        public Cell_Manager()
        {
            InitializeComponent();
        }

        /********************************************************************************************************************
        * 
        * Global Variables 
        * 
        * 
        * 
        ********************************************************************************************************************/

        private string LoginTime = "";
        private string LoginForm = "Cell Control";

        private static int[] ArrayRunOrder = new int[10];
        private static string[] ArrayItemID = new string[10];
        private static int[] ArrayRowNumber = new int[10];
        private static string RowIndexClicked = "";
        private static int RowIndexClick = 0;
        private static string RunOrder_Array = "";
        private static int RowIndex = 0;
        static string RemoveItem = "";
        private string RemoveItemID;

        private string SQL_Source = @"Data Source = OHN7009,49172; Initial Catalog = Spot_Weld_Data; Integrated Security = True; Connect Timeout = 15;";

        // Clock_Timer();
        private static int ClockHour;
        private static int ClockMinute;
        private static int ClockSecond;

        // Search_Button_Click();
        private static string SearchValue;
        private static int SearchColumn;

        private string[] CustomerCell = { "CAT", "John Deere", "Navistar", "Paccar"};
        private string[] CAT_SpotWeldList = { "1088" };
        private string[] JohnDeere_SpotWeldList = { "150R" };
        private string[] Navistar_SpotWeldList = { "104R", "121R", "154R" };
        private string[] Paccar_SpotWeldList = { "153R", "155R" };

        static int RunOrder = 0;
        static string RunOrder_String = "";

        private static int rows;

        static string SQLRemoveCommand = "";
        static string SQLAddCommand = "";
        static string SQLUpdateCommand = "";
        static string SQLSelectCount = "";
        static string Schedule_Count = "";
        static string Refresh_Data = "";

        private void Cell_Manager_Load(object sender, EventArgs e)
        {
            SqlConnection UserLogin = new SqlConnection(SQL_Source);
            SqlCommand Login = new SqlCommand();
            Login.CommandType = System.Data.CommandType.Text;
            Login.CommandText = "INSERT INTO [dbo].[LoginData] (EmployeeName,DMPID,LoginDateTime,LoginForm) VALUES (@EmployeeName,@DMPID,@LoginDateTime,@LoginForm)";
            Login.Connection = UserLogin;
            Login.Parameters.AddWithValue("@LoginDateTime", Clock_TextBox.Text);
            Login.Parameters.AddWithValue("@EmployeeName", User_TextBox.Text);
            Login.Parameters.AddWithValue("@DMPID", DMPID_TextBox.Text);
            Login.Parameters.AddWithValue("@LoginForm", LoginForm.ToString());
            UserLogin.Open();
            Login.ExecuteNonQuery();
            UserLogin.Close();

            Clock.Enabled = true;
            LoginTime = Clock_TextBox.Text;
            CustomerCell_ComboBox.Items.AddRange(CustomerCell);
            SearchItemID_CheckBox.Checked = true;
        }

        /********************************************************************************************************************
        * 
        * Buttons Region Start
        * 
        * -- CellManager Form Button
        * - ChangeCell Button Click
        * - LogOff Click
        * 
        * -- ItemInformation GroupBox Buttons
        * - Search Button Click
        * - Clear Button Click
        * 
        * -- OrderData GroupBox Buttons
        * - AddToQueue Button Click
        * - Calculate Button Click
        * 
        * -- RunOrderGroupBox Buttons
        * - QueueUp Button Click
        * - QueueDown Button Click
        * - Remove 1 Button Click
        * - Remove 2 Button Click
        * - Remove 3 Button Click
        * - Remove 4 Button Click
        * - Remove 5 Button Click
        * - Remove 6 Button Click
        * - Remove 7 Button Click
        * - Remove 8 Button Click
        * - Remove 9 Button Click
        * - Remove 10 Button Click
        * 
        ********************************************************************************************************************/
        #region

        private void ChangeCell_Button_Click(object sender, EventArgs e)
        {
            CustomerCell_ComboBox.Enabled = true;
            CustomerCell_ComboBox.Visible = true;
            CustomerCell_TextBox.Visible = false;
            Clear();
            //GroupBoxControlStart();
            ChangeCell_Button.Hide();
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

        private void Search_Button_Click(object sender, EventArgs e)
        {
            if (SearchItemID_CheckBox.Checked == true)
            {
                SearchCustomerItemID_CheckBox.Checked = false;
                SearchValue = ItemID_TextBox.Text;
                SearchColumn = 0;
            }
            else if (SearchCustomerItemID_CheckBox.Checked == true)
            {
                SearchItemID_CheckBox.Checked = false;
                SearchValue = CustomerItemID_TextBox.Text;
                SearchColumn = 4;
            }
            CellManager_GridView.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            try
            {
                foreach (DataGridViewRow Row in CellManager_GridView.Rows)
                {
                    Row.Selected = false;
                    if (Row.Cells[SearchColumn].Value.ToString().Equals(SearchValue))
                    {
                        Row.Selected = true;
                        ItemID_TextBox.Text = Row.Cells[0].Value.ToString();
                        Customer_TextBox.Text = Row.Cells[1].Value.ToString();
                        CustomerItemID_TextBox.Text = Row.Cells[2].Value.ToString();
                        JobID_TextBox.Text = Row.Cells[3].Value.ToString();
                        Fixture_TextBox.Text = Row.Cells[5].Value.ToString();
                        FixtureLocation_TextBox.Text = Row.Cells[6].Value.ToString();
                        TotalRuns_TextBox.Text = Row.Cells[13].Value.ToString();
                        PartsManufactured_TextBox.Text = Row.Cells[14].Value.ToString();
                        PPM_TextBox.Text = Row.Cells[15].Value.ToString();
                        SetupTime_TextBox.Text = Row.Cells[16].Value.ToString();
                        CellManager_GridView.FirstDisplayedScrollingRowIndex = CellManager_GridView.SelectedRows[0].Index;
                        break;
                    }
                }
            }
            catch (Exception)
            {
                MessageBox.Show("Error Finding Job");
            }

        }
        
        private void Clear_Button_Click(object sender, EventArgs e)
        {

        }

        private void AddToQueue_Button_Click(object sender, EventArgs e)
        {
            GetJobQueueRunOrder();
            AddJobToQueue();
            RefreshJobs();
        }

        private void Calculate_Button_Click(object sender, EventArgs e)
        {
            double HoursRemaining = 0;
            double MinutesRemaining = 0;
            string RemainingTime = "";
            string Parts = PartsOnOrder_TextBox.Text;
            string PPM = PPM_TextBox.Text;
            double PartsOrdered = double.Parse(Parts);
            double AveragePPM = double.Parse(PPM);
            double TimeRemaining = (PartsOrdered / AveragePPM);

            if (TimeRemaining < 60)
            {
                MinutesRemaining = TimeRemaining;
                MinutesRemaining = Math.Round(MinutesRemaining, MidpointRounding.AwayFromZero);
                if (MinutesRemaining == 1)
                {
                    RemainingTime = MinutesRemaining + " Minute";
                }
                else
                {
                    RemainingTime = MinutesRemaining + " Minutes";
                }
            }
            else if (120 > TimeRemaining && TimeRemaining >= 60)
            {
                TimeRemaining = TimeRemaining - 60;
                HoursRemaining = 1;
                MinutesRemaining = TimeRemaining;
                MinutesRemaining = Math.Round(MinutesRemaining, MidpointRounding.AwayFromZero);
                if (MinutesRemaining == 0)
                {
                    RemainingTime = HoursRemaining + " Hour ";
                }
                else
                {
                    RemainingTime = HoursRemaining + " Hour " + MinutesRemaining + " Minutes";
                }
            }
            else if (180 > TimeRemaining && TimeRemaining >= 120)
            {
                TimeRemaining = TimeRemaining - 120;
                HoursRemaining = 2;
                MinutesRemaining = TimeRemaining;
                MinutesRemaining = Math.Round(MinutesRemaining, MidpointRounding.AwayFromZero);
                if (MinutesRemaining == 0)
                {
                    RemainingTime = HoursRemaining + " Hours";
                }
                else
                {
                    RemainingTime = HoursRemaining + " Hours " + MinutesRemaining + " Minutes";
                }
            }
            else if (240 > TimeRemaining && TimeRemaining >= 180)
            {
                TimeRemaining = TimeRemaining - 180;
                HoursRemaining = 3;
                MinutesRemaining = TimeRemaining;
                MinutesRemaining = Math.Round(MinutesRemaining, MidpointRounding.AwayFromZero);
                if (MinutesRemaining == 0)
                {
                    RemainingTime = HoursRemaining + " Hours";
                }
                else
                {
                    RemainingTime = HoursRemaining + " Hours " + MinutesRemaining + " Minutes";
                }
            }
            else if (300 > TimeRemaining && TimeRemaining >= 240)
            {
                TimeRemaining = TimeRemaining - 240;
                HoursRemaining = 4;
                MinutesRemaining = TimeRemaining;
                MinutesRemaining = Math.Round(MinutesRemaining, MidpointRounding.AwayFromZero);
                if (MinutesRemaining == 0)
                {
                    RemainingTime = HoursRemaining + " Hours";
                }
                else
                {
                    RemainingTime = HoursRemaining + " Hours " + MinutesRemaining + " Minutes";
                }
            }
            else if (360 > TimeRemaining && TimeRemaining >= 300)
            {
                TimeRemaining = TimeRemaining - 300;
                HoursRemaining = 5;
                MinutesRemaining = TimeRemaining;
                MinutesRemaining = Math.Round(MinutesRemaining, MidpointRounding.AwayFromZero);
                if (MinutesRemaining == 0)
                {
                    RemainingTime = HoursRemaining + " Hours";
                }
                else
                {
                    RemainingTime = HoursRemaining + " Hours " + MinutesRemaining + " Minutes";
                }
            }
            else if (420 > TimeRemaining && TimeRemaining >= 360)
            {
                TimeRemaining = TimeRemaining - 360;
                HoursRemaining = 6;
                MinutesRemaining = TimeRemaining;
                MinutesRemaining = Math.Round(MinutesRemaining, MidpointRounding.AwayFromZero);
                if (MinutesRemaining == 0)
                {
                    RemainingTime = HoursRemaining + " Hours";
                }
                else
                {
                    RemainingTime = HoursRemaining + " Hours " + MinutesRemaining + " Minutes";
                }
            }
            else if (480 > TimeRemaining && TimeRemaining >= 420)
            {
                TimeRemaining = TimeRemaining - 420;
                HoursRemaining = 7;
                MinutesRemaining = TimeRemaining;
                MinutesRemaining = Math.Round(MinutesRemaining, MidpointRounding.AwayFromZero);
                if (MinutesRemaining == 0)
                {
                    RemainingTime = HoursRemaining + " Hours";
                }
                else
                {
                    RemainingTime = HoursRemaining + " Hours " + MinutesRemaining + " Minutes";
                }
            }
            else if (540 > TimeRemaining && TimeRemaining >= 480)
            {
                TimeRemaining = TimeRemaining - 480;
                HoursRemaining = 8;
                MinutesRemaining = TimeRemaining;
                MinutesRemaining = Math.Round(MinutesRemaining, MidpointRounding.AwayFromZero);
                if (MinutesRemaining == 0)
                {
                    RemainingTime = HoursRemaining + " Hours";
                }
                else
                {
                    RemainingTime = HoursRemaining + " Hours " + MinutesRemaining + " Minutes";
                }
            }
            else if (600 > TimeRemaining && TimeRemaining >= 540)
            {
                TimeRemaining = TimeRemaining - 540;
                HoursRemaining = 9;
                MinutesRemaining = TimeRemaining;
                MinutesRemaining = Math.Round(MinutesRemaining, MidpointRounding.AwayFromZero);
                if (MinutesRemaining == 0)
                {
                    RemainingTime = HoursRemaining + " Hours";
                }
                else
                {
                    RemainingTime = HoursRemaining + " Hours " + MinutesRemaining + " Minutes";
                }
            }
            else if (660 > TimeRemaining && TimeRemaining >= 600)
            {
                TimeRemaining = TimeRemaining - 600;
                HoursRemaining = 9;
                MinutesRemaining = TimeRemaining;
                MinutesRemaining = Math.Round(MinutesRemaining, MidpointRounding.AwayFromZero);
                if (MinutesRemaining == 0)
                {
                    RemainingTime = HoursRemaining + " Hours";
                }
                else
                {
                    RemainingTime = HoursRemaining + " Hours " + MinutesRemaining + " Minutes";
                }
            }

            EstimatedRunTime_TextBox.Text = RemainingTime;
        }

        private void QueueUp_Button_Click(object sender, EventArgs e)
        {
            if (RowIndexClick > 0)
            {
                for (int i = 0; i < RowIndex; i++)
                {
                    if (ArrayRowNumber[i] == (RowIndexClick - 1))
                    {
                        ArrayRunOrder[i] += 1;
                        ArrayRowNumber[i] += 1;
                    }
                    else if (ArrayRowNumber[i] == RowIndexClick)
                    {
                        ArrayRunOrder[i] -= 1;
                        ArrayRowNumber[i] -= 1;
                    }
                }
                for (int m = 0; m < RowIndex; m++)
                {
                    Console.WriteLine(ArrayRunOrder[m] + " " + ArrayItemID[m] + " " + ArrayRowNumber[m]);
                }
                try
                {
                    for (int x = 0; x < rows; x++)
                    {
                        SqlConnection Job_Connection = new SqlConnection(SQL_Source);
                        SqlCommand Edit_Job = new SqlCommand();
                        Edit_Job.CommandType = System.Data.CommandType.Text;
                        Edit_Job.CommandText = SQLUpdateCommand;
                        Edit_Job.Connection = Job_Connection;
                        Edit_Job.Parameters.AddWithValue("@RunOrder", ArrayRunOrder[x].ToString());
                        Edit_Job.Parameters.AddWithValue("@ItemID", ArrayItemID[x].ToString());
                        Job_Connection.Open();
                        Edit_Job.ExecuteNonQuery();
                        Job_Connection.Close();
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString());
                }

                RefreshJobs();
            }
            else
            {
                MessageBox.Show("The Item Selected is Already at the Top of the Queue");
            }
        }

        private void QueueDown_Button_Click(object sender, EventArgs e)
        {
            if (RowIndexClick < (RowIndex - 1))
            {
                for (int i = 0; i < RowIndex; i++)
                {
                    if (ArrayRowNumber[i] == (RowIndexClick + 1))
                    {
                        ArrayRunOrder[i] -= 1;
                        ArrayRowNumber[i] -= 1;
                    }
                    else if (ArrayRowNumber[i] == RowIndexClick)
                    {
                        ArrayRunOrder[i] += 1;
                        ArrayRowNumber[i] += 1;
                    }
                }
                for (int m = 0; m < RowIndex; m++)
                {
                    Console.WriteLine(ArrayRunOrder[m] + " " + ArrayItemID[m] + " " + ArrayRowNumber[m]);
                }
                try
                {
                    for (int x = 0; x < rows; x++)
                    {
                        SqlConnection Job_Connection = new SqlConnection(SQL_Source);
                        SqlCommand Edit_Job = new SqlCommand();
                        Edit_Job.CommandType = System.Data.CommandType.Text;
                        Edit_Job.CommandText = SQLUpdateCommand;
                        Edit_Job.Connection = Job_Connection;
                        Edit_Job.Parameters.AddWithValue("@RunOrder", ArrayRunOrder[x].ToString());
                        Edit_Job.Parameters.AddWithValue("@ItemID", ArrayItemID[x].ToString());
                        Job_Connection.Open();
                        Edit_Job.ExecuteNonQuery();
                        Job_Connection.Close();
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString());
                }

                RefreshJobs();
            }
            else
            {
                MessageBox.Show("The Item Selected is Already at the Bottom of the Queue");
            }
        }

        // Remove Buttons Region
        #region

        private void Remove_1_Button_Click(object sender, EventArgs e)
        {
            RemoveItemID = RunOrder_1_TextBox.Text;
            RowIndexClick = 0;
            DeleteCombo();
            RefreshJobs();
        }

        private void Remove_2_Button_Click(object sender, EventArgs e)
        {
            RemoveItemID = RunOrder_2_TextBox.Text;
            RowIndexClick = 1;
            DeleteCombo();
            RefreshJobs();
        }

        private void Remove_3_Button_Click(object sender, EventArgs e)
        {
            RemoveItemID = RunOrder_3_TextBox.Text;
            RowIndexClick = 2;
            DeleteCombo();
            RefreshJobs();
        }

        private void Remove_4_Button_Click(object sender, EventArgs e)
        {
            RemoveItemID = RunOrder_4_TextBox.Text;
            RowIndexClick = 3;
            DeleteCombo();
            RefreshJobs();
        }

        private void Remove_5_Button_Click(object sender, EventArgs e)
        {
            RemoveItemID = RunOrder_5_TextBox.Text;
            RowIndexClick = 4;
            DeleteCombo();
            RefreshJobs();
        }

        private void Remove_6_Button_Click(object sender, EventArgs e)
        {
            RemoveItemID = RunOrder_6_TextBox.Text;
            RowIndexClick = 5;
            DeleteCombo();
            RefreshJobs();
        }

        private void Remove_7_Button_Click(object sender, EventArgs e)
        {
            RemoveItemID = RunOrder_7_TextBox.Text;
            RowIndexClick = 6;
            DeleteCombo();
            RefreshJobs();
        }

        private void Remove_8_Button_Click(object sender, EventArgs e)
        {
            RemoveItemID = RunOrder_8_TextBox.Text;
            RowIndexClick = 7;
            DeleteCombo();
            RefreshJobs();
        }

        private void Remove_9_Button_Click(object sender, EventArgs e)
        {
            RemoveItemID = RunOrder_9_TextBox.Text;
            RowIndexClick = 8;
            DeleteCombo();
            RefreshJobs();
        }

        private void Remove_10_Button_Click(object sender, EventArgs e)
        {
            RemoveItemID = RunOrder_10_TextBox.Text;
            RowIndexClick = 9;
            DeleteCombo();
            RefreshJobs();
        }
        #endregion

        /********************************************************************************************************************
        * 
        * Buttons End
        * 
        ********************************************************************************************************************/
        #endregion

        /********************************************************************************************************************
        * 
        * CheckBox Region Start
        * 
        * - SearchItemID CheckBox CheckedChanged
        * - SearchJobID CheckBox CheckedChanged
        * 
        ********************************************************************************************************************/
        #region

        private void SearchItemID_CheckBox_CheckedChanged(object sender, EventArgs e)
        {
            if (SearchItemID_CheckBox.Checked == true)
            {
                SearchCustomerItemID_CheckBox.Checked = false;
            }
        }

        private void SearchCustomerItemID_CheckBox_CheckedChanged(object sender, EventArgs e)
        {
            if (SearchCustomerItemID_CheckBox.Checked == true)
            {
                SearchItemID_CheckBox.Checked = false;
            }
        }

        private void RunOrder_1_CheckBox_CheckedChanged(object sender, EventArgs e)
        {
            if (RunOrder_1_CheckBox.Checked == true)
            {
                Array.Clear(ArrayRunOrder, 0, ArrayRunOrder.Length);
                Array.Clear(ArrayItemID, 0, ArrayItemID.Length);
                RowIndex = 0;
                foreach (DataGridViewRow Row in SpotWeld_GridView.Rows)
                {
                    RunOrder_Array = Row.Cells[0].Value.ToString();
                    ArrayRunOrder[RowIndex] = Int32.Parse(RunOrder_Array);
                    ArrayItemID[RowIndex] = Row.Cells[1].Value.ToString();
                    ArrayRowNumber[RowIndex] = SpotWeld_GridView.Rows.IndexOf(Row);
                    RowIndex++;
                }
                RowIndexClick = 0;
                RunOrder_1_TextBox.BackColor = Color.Chartreuse;
                ItemID_1_TextBox.BackColor = Color.Chartreuse;
                RunOrder_2_TextBox.BackColor = System.Drawing.SystemColors.Window;
                ItemID_2_TextBox.BackColor = System.Drawing.SystemColors.Window;
                RunOrder_3_TextBox.BackColor = System.Drawing.SystemColors.Window;
                ItemID_3_TextBox.BackColor = System.Drawing.SystemColors.Window;
                RunOrder_4_TextBox.BackColor = System.Drawing.SystemColors.Window;
                ItemID_4_TextBox.BackColor = System.Drawing.SystemColors.Window;
                RunOrder_5_TextBox.BackColor = System.Drawing.SystemColors.Window;
                ItemID_5_TextBox.BackColor = System.Drawing.SystemColors.Window;
                RunOrder_6_TextBox.BackColor = System.Drawing.SystemColors.Window;
                ItemID_6_TextBox.BackColor = System.Drawing.SystemColors.Window;
                RunOrder_7_TextBox.BackColor = System.Drawing.SystemColors.Window;
                ItemID_7_TextBox.BackColor = System.Drawing.SystemColors.Window;
                RunOrder_8_TextBox.BackColor = System.Drawing.SystemColors.Window;
                ItemID_8_TextBox.BackColor = System.Drawing.SystemColors.Window;
                RunOrder_9_TextBox.BackColor = System.Drawing.SystemColors.Window;
                ItemID_9_TextBox.BackColor = System.Drawing.SystemColors.Window;
                RunOrder_10_TextBox.BackColor = System.Drawing.SystemColors.Window;
                ItemID_10_TextBox.BackColor = System.Drawing.SystemColors.Window;
                RunOrder_2_CheckBox.Checked = false;
                RunOrder_3_CheckBox.Checked = false;
                RunOrder_4_CheckBox.Checked = false;
                RunOrder_5_CheckBox.Checked = false;
                RunOrder_6_CheckBox.Checked = false;
                RunOrder_7_CheckBox.Checked = false;
                RunOrder_8_CheckBox.Checked = false;
                RunOrder_9_CheckBox.Checked = false;
                RunOrder_10_CheckBox.Checked = false;
            }
        }

        private void RunOrder_2_CheckBox_CheckedChanged(object sender, EventArgs e)
        {
            if (RunOrder_2_CheckBox.Checked == true)
            {
                Array.Clear(ArrayRunOrder, 0, ArrayRunOrder.Length);
                Array.Clear(ArrayItemID, 0, ArrayItemID.Length);
                RowIndex = 0;
                foreach (DataGridViewRow Row in SpotWeld_GridView.Rows)
                {
                    RunOrder_Array = Row.Cells[0].Value.ToString();
                    ArrayRunOrder[RowIndex] = Int32.Parse(RunOrder_Array);
                    ArrayItemID[RowIndex] = Row.Cells[1].Value.ToString();
                    ArrayRowNumber[RowIndex] = SpotWeld_GridView.Rows.IndexOf(Row);
                    RowIndex++;
                }
                RowIndexClick = 1;
                RunOrder_2_TextBox.BackColor = Color.Chartreuse;
                ItemID_2_TextBox.BackColor = Color.Chartreuse;
                RunOrder_1_TextBox.BackColor = System.Drawing.SystemColors.Window;
                ItemID_1_TextBox.BackColor = System.Drawing.SystemColors.Window;
                RunOrder_3_TextBox.BackColor = System.Drawing.SystemColors.Window;
                ItemID_3_TextBox.BackColor = System.Drawing.SystemColors.Window;
                RunOrder_4_TextBox.BackColor = System.Drawing.SystemColors.Window;
                ItemID_4_TextBox.BackColor = System.Drawing.SystemColors.Window;
                RunOrder_5_TextBox.BackColor = System.Drawing.SystemColors.Window;
                ItemID_5_TextBox.BackColor = System.Drawing.SystemColors.Window;
                RunOrder_6_TextBox.BackColor = System.Drawing.SystemColors.Window;
                ItemID_6_TextBox.BackColor = System.Drawing.SystemColors.Window;
                RunOrder_7_TextBox.BackColor = System.Drawing.SystemColors.Window;
                ItemID_7_TextBox.BackColor = System.Drawing.SystemColors.Window;
                RunOrder_8_TextBox.BackColor = System.Drawing.SystemColors.Window;
                ItemID_8_TextBox.BackColor = System.Drawing.SystemColors.Window;
                RunOrder_9_TextBox.BackColor = System.Drawing.SystemColors.Window;
                ItemID_9_TextBox.BackColor = System.Drawing.SystemColors.Window;
                RunOrder_10_TextBox.BackColor = System.Drawing.SystemColors.Window;
                ItemID_10_TextBox.BackColor = System.Drawing.SystemColors.Window;
                RunOrder_1_CheckBox.Checked = false;
                RunOrder_3_CheckBox.Checked = false;
                RunOrder_4_CheckBox.Checked = false;
                RunOrder_5_CheckBox.Checked = false;
                RunOrder_6_CheckBox.Checked = false;
                RunOrder_7_CheckBox.Checked = false;
                RunOrder_8_CheckBox.Checked = false;
                RunOrder_9_CheckBox.Checked = false;
                RunOrder_10_CheckBox.Checked = false;
            }
            }

        private void RunOrder_3_CheckBox_CheckedChanged(object sender, EventArgs e)
        {
            if (RunOrder_3_CheckBox.Checked == true)
            {
                Array.Clear(ArrayRunOrder, 0, ArrayRunOrder.Length);
                Array.Clear(ArrayItemID, 0, ArrayItemID.Length);
                RowIndex = 0;
                foreach (DataGridViewRow Row in SpotWeld_GridView.Rows)
                {
                    RunOrder_Array = Row.Cells[0].Value.ToString();
                    ArrayRunOrder[RowIndex] = Int32.Parse(RunOrder_Array);
                    ArrayItemID[RowIndex] = Row.Cells[1].Value.ToString();
                    ArrayRowNumber[RowIndex] = SpotWeld_GridView.Rows.IndexOf(Row);
                    RowIndex++;
                }
                RowIndexClick = 2;
                RunOrder_3_TextBox.BackColor = Color.Chartreuse;
                ItemID_3_TextBox.BackColor = Color.Chartreuse;
                RunOrder_1_TextBox.BackColor = System.Drawing.SystemColors.Window;
                ItemID_1_TextBox.BackColor = System.Drawing.SystemColors.Window;
                RunOrder_2_TextBox.BackColor = System.Drawing.SystemColors.Window;
                ItemID_2_TextBox.BackColor = System.Drawing.SystemColors.Window;
                RunOrder_4_TextBox.BackColor = System.Drawing.SystemColors.Window;
                ItemID_4_TextBox.BackColor = System.Drawing.SystemColors.Window;
                RunOrder_5_TextBox.BackColor = System.Drawing.SystemColors.Window;
                ItemID_5_TextBox.BackColor = System.Drawing.SystemColors.Window;
                RunOrder_6_TextBox.BackColor = System.Drawing.SystemColors.Window;
                ItemID_6_TextBox.BackColor = System.Drawing.SystemColors.Window;
                RunOrder_7_TextBox.BackColor = System.Drawing.SystemColors.Window;
                ItemID_7_TextBox.BackColor = System.Drawing.SystemColors.Window;
                RunOrder_8_TextBox.BackColor = System.Drawing.SystemColors.Window;
                ItemID_8_TextBox.BackColor = System.Drawing.SystemColors.Window;
                RunOrder_9_TextBox.BackColor = System.Drawing.SystemColors.Window;
                ItemID_9_TextBox.BackColor = System.Drawing.SystemColors.Window;
                RunOrder_10_TextBox.BackColor = System.Drawing.SystemColors.Window;
                ItemID_10_TextBox.BackColor = System.Drawing.SystemColors.Window;
                RunOrder_1_CheckBox.Checked = false;
                RunOrder_2_CheckBox.Checked = false;
                RunOrder_4_CheckBox.Checked = false;
                RunOrder_5_CheckBox.Checked = false;
                RunOrder_6_CheckBox.Checked = false;
                RunOrder_7_CheckBox.Checked = false;
                RunOrder_8_CheckBox.Checked = false;
                RunOrder_9_CheckBox.Checked = false;
                RunOrder_10_CheckBox.Checked = false;
            }
        }

        private void RunOrder_4_CheckBox_CheckedChanged(object sender, EventArgs e)
        {
            if (RunOrder_4_CheckBox.Checked == true)
            {
                Array.Clear(ArrayRunOrder, 0, ArrayRunOrder.Length);
                Array.Clear(ArrayItemID, 0, ArrayItemID.Length);
                RowIndex = 0;
                foreach (DataGridViewRow Row in SpotWeld_GridView.Rows)
                {
                    RunOrder_Array = Row.Cells[0].Value.ToString();
                    ArrayRunOrder[RowIndex] = Int32.Parse(RunOrder_Array);
                    ArrayItemID[RowIndex] = Row.Cells[1].Value.ToString();
                    ArrayRowNumber[RowIndex] = SpotWeld_GridView.Rows.IndexOf(Row);
                    RowIndex++;
                }
                RowIndexClick = 3;
                RunOrder_4_TextBox.BackColor = Color.Chartreuse;
                ItemID_4_TextBox.BackColor = Color.Chartreuse;
                RunOrder_1_TextBox.BackColor = System.Drawing.SystemColors.Window;
                ItemID_1_TextBox.BackColor = System.Drawing.SystemColors.Window;
                RunOrder_2_TextBox.BackColor = System.Drawing.SystemColors.Window;
                ItemID_2_TextBox.BackColor = System.Drawing.SystemColors.Window;
                RunOrder_3_TextBox.BackColor = System.Drawing.SystemColors.Window;
                ItemID_3_TextBox.BackColor = System.Drawing.SystemColors.Window;
                RunOrder_5_TextBox.BackColor = System.Drawing.SystemColors.Window;
                ItemID_5_TextBox.BackColor = System.Drawing.SystemColors.Window;
                RunOrder_6_TextBox.BackColor = System.Drawing.SystemColors.Window;
                ItemID_6_TextBox.BackColor = System.Drawing.SystemColors.Window;
                RunOrder_7_TextBox.BackColor = System.Drawing.SystemColors.Window;
                ItemID_7_TextBox.BackColor = System.Drawing.SystemColors.Window;
                RunOrder_8_TextBox.BackColor = System.Drawing.SystemColors.Window;
                ItemID_8_TextBox.BackColor = System.Drawing.SystemColors.Window;
                RunOrder_9_TextBox.BackColor = System.Drawing.SystemColors.Window;
                ItemID_9_TextBox.BackColor = System.Drawing.SystemColors.Window;
                RunOrder_10_TextBox.BackColor = System.Drawing.SystemColors.Window;
                ItemID_10_TextBox.BackColor = System.Drawing.SystemColors.Window;
                RunOrder_1_CheckBox.Checked = false;
                RunOrder_2_CheckBox.Checked = false;
                RunOrder_3_CheckBox.Checked = false;
                RunOrder_5_CheckBox.Checked = false;
                RunOrder_6_CheckBox.Checked = false;
                RunOrder_7_CheckBox.Checked = false;
                RunOrder_8_CheckBox.Checked = false;
                RunOrder_9_CheckBox.Checked = false;
                RunOrder_10_CheckBox.Checked = false;
            }
        }

        private void RunOrder_5_CheckBox_CheckedChanged(object sender, EventArgs e)
        {
            if (RunOrder_5_CheckBox.Checked == true)
            {
                Array.Clear(ArrayRunOrder, 0, ArrayRunOrder.Length);
                Array.Clear(ArrayItemID, 0, ArrayItemID.Length);
                RowIndex = 0;
                foreach (DataGridViewRow Row in SpotWeld_GridView.Rows)
                {
                    RunOrder_Array = Row.Cells[0].Value.ToString();
                    ArrayRunOrder[RowIndex] = Int32.Parse(RunOrder_Array);
                    ArrayItemID[RowIndex] = Row.Cells[1].Value.ToString();
                    ArrayRowNumber[RowIndex] = SpotWeld_GridView.Rows.IndexOf(Row);
                    RowIndex++;
                }
                RowIndexClick = 4;
                RunOrder_5_TextBox.BackColor = Color.Chartreuse;
                ItemID_5_TextBox.BackColor = Color.Chartreuse;
                RunOrder_1_TextBox.BackColor = System.Drawing.SystemColors.Window;
                ItemID_1_TextBox.BackColor = System.Drawing.SystemColors.Window;
                RunOrder_2_TextBox.BackColor = System.Drawing.SystemColors.Window;
                ItemID_2_TextBox.BackColor = System.Drawing.SystemColors.Window;
                RunOrder_3_TextBox.BackColor = System.Drawing.SystemColors.Window;
                ItemID_3_TextBox.BackColor = System.Drawing.SystemColors.Window;
                RunOrder_4_TextBox.BackColor = System.Drawing.SystemColors.Window;
                ItemID_4_TextBox.BackColor = System.Drawing.SystemColors.Window;
                RunOrder_6_TextBox.BackColor = System.Drawing.SystemColors.Window;
                ItemID_6_TextBox.BackColor = System.Drawing.SystemColors.Window;
                RunOrder_7_TextBox.BackColor = System.Drawing.SystemColors.Window;
                ItemID_7_TextBox.BackColor = System.Drawing.SystemColors.Window;
                RunOrder_8_TextBox.BackColor = System.Drawing.SystemColors.Window;
                ItemID_8_TextBox.BackColor = System.Drawing.SystemColors.Window;
                RunOrder_9_TextBox.BackColor = System.Drawing.SystemColors.Window;
                ItemID_9_TextBox.BackColor = System.Drawing.SystemColors.Window;
                RunOrder_10_TextBox.BackColor = System.Drawing.SystemColors.Window;
                ItemID_10_TextBox.BackColor = System.Drawing.SystemColors.Window;
                RunOrder_1_CheckBox.Checked = false;
                RunOrder_2_CheckBox.Checked = false;
                RunOrder_3_CheckBox.Checked = false;
                RunOrder_4_CheckBox.Checked = false;
                RunOrder_6_CheckBox.Checked = false;
                RunOrder_7_CheckBox.Checked = false;
                RunOrder_8_CheckBox.Checked = false;
                RunOrder_9_CheckBox.Checked = false;
                RunOrder_10_CheckBox.Checked = false;
            }
        }

        private void RunOrder_6_CheckBox_CheckedChanged(object sender, EventArgs e)
        {
            if (RunOrder_6_CheckBox.Checked == true)
            {
                Array.Clear(ArrayRunOrder, 0, ArrayRunOrder.Length);
                Array.Clear(ArrayItemID, 0, ArrayItemID.Length);
                RowIndex = 0;
                foreach (DataGridViewRow Row in SpotWeld_GridView.Rows)
                {
                    RunOrder_Array = Row.Cells[0].Value.ToString();
                    ArrayRunOrder[RowIndex] = Int32.Parse(RunOrder_Array);
                    ArrayItemID[RowIndex] = Row.Cells[1].Value.ToString();
                    ArrayRowNumber[RowIndex] = SpotWeld_GridView.Rows.IndexOf(Row);
                    RowIndex++;
                }
                RowIndexClick = 5;
                RunOrder_6_TextBox.BackColor = Color.Chartreuse;
                ItemID_6_TextBox.BackColor = Color.Chartreuse;
                RunOrder_1_TextBox.BackColor = System.Drawing.SystemColors.Window;
                ItemID_1_TextBox.BackColor = System.Drawing.SystemColors.Window;
                RunOrder_2_TextBox.BackColor = System.Drawing.SystemColors.Window;
                ItemID_2_TextBox.BackColor = System.Drawing.SystemColors.Window;
                RunOrder_3_TextBox.BackColor = System.Drawing.SystemColors.Window;
                ItemID_3_TextBox.BackColor = System.Drawing.SystemColors.Window;
                RunOrder_4_TextBox.BackColor = System.Drawing.SystemColors.Window;
                ItemID_4_TextBox.BackColor = System.Drawing.SystemColors.Window;
                RunOrder_5_TextBox.BackColor = System.Drawing.SystemColors.Window;
                ItemID_5_TextBox.BackColor = System.Drawing.SystemColors.Window;
                RunOrder_7_TextBox.BackColor = System.Drawing.SystemColors.Window;
                ItemID_7_TextBox.BackColor = System.Drawing.SystemColors.Window;
                RunOrder_8_TextBox.BackColor = System.Drawing.SystemColors.Window;
                ItemID_8_TextBox.BackColor = System.Drawing.SystemColors.Window;
                RunOrder_9_TextBox.BackColor = System.Drawing.SystemColors.Window;
                ItemID_9_TextBox.BackColor = System.Drawing.SystemColors.Window;
                RunOrder_10_TextBox.BackColor = System.Drawing.SystemColors.Window;
                ItemID_10_TextBox.BackColor = System.Drawing.SystemColors.Window;
                RunOrder_1_CheckBox.Checked = false;
                RunOrder_2_CheckBox.Checked = false;
                RunOrder_3_CheckBox.Checked = false;
                RunOrder_4_CheckBox.Checked = false;
                RunOrder_5_CheckBox.Checked = false;
                RunOrder_7_CheckBox.Checked = false;
                RunOrder_8_CheckBox.Checked = false;
                RunOrder_9_CheckBox.Checked = false;
                RunOrder_10_CheckBox.Checked = false;
            }
        }

        private void RunOrder_7_CheckBox_CheckedChanged(object sender, EventArgs e)
        {
            if (RunOrder_7_CheckBox.Checked == true)
            {
                Array.Clear(ArrayRunOrder, 0, ArrayRunOrder.Length);
                Array.Clear(ArrayItemID, 0, ArrayItemID.Length);
                RowIndex = 0;
                foreach (DataGridViewRow Row in SpotWeld_GridView.Rows)
                {
                    RunOrder_Array = Row.Cells[0].Value.ToString();
                    ArrayRunOrder[RowIndex] = Int32.Parse(RunOrder_Array);
                    ArrayItemID[RowIndex] = Row.Cells[1].Value.ToString();
                    ArrayRowNumber[RowIndex] = SpotWeld_GridView.Rows.IndexOf(Row);
                    RowIndex++;
                }
                RowIndexClick = 6;
                RunOrder_7_TextBox.BackColor = Color.Chartreuse;
                ItemID_7_TextBox.BackColor = Color.Chartreuse;
                RunOrder_1_TextBox.BackColor = System.Drawing.SystemColors.Window;
                ItemID_1_TextBox.BackColor = System.Drawing.SystemColors.Window;
                RunOrder_2_TextBox.BackColor = System.Drawing.SystemColors.Window;
                ItemID_2_TextBox.BackColor = System.Drawing.SystemColors.Window;
                RunOrder_3_TextBox.BackColor = System.Drawing.SystemColors.Window;
                ItemID_3_TextBox.BackColor = System.Drawing.SystemColors.Window;
                RunOrder_4_TextBox.BackColor = System.Drawing.SystemColors.Window;
                ItemID_4_TextBox.BackColor = System.Drawing.SystemColors.Window;
                RunOrder_5_TextBox.BackColor = System.Drawing.SystemColors.Window;
                ItemID_5_TextBox.BackColor = System.Drawing.SystemColors.Window;
                RunOrder_6_TextBox.BackColor = System.Drawing.SystemColors.Window;
                ItemID_6_TextBox.BackColor = System.Drawing.SystemColors.Window;
                RunOrder_8_TextBox.BackColor = System.Drawing.SystemColors.Window;
                ItemID_8_TextBox.BackColor = System.Drawing.SystemColors.Window;
                RunOrder_9_TextBox.BackColor = System.Drawing.SystemColors.Window;
                ItemID_9_TextBox.BackColor = System.Drawing.SystemColors.Window;
                RunOrder_10_TextBox.BackColor = System.Drawing.SystemColors.Window;
                ItemID_10_TextBox.BackColor = System.Drawing.SystemColors.Window;
                RunOrder_1_CheckBox.Checked = false;
                RunOrder_2_CheckBox.Checked = false;
                RunOrder_3_CheckBox.Checked = false;
                RunOrder_4_CheckBox.Checked = false;
                RunOrder_5_CheckBox.Checked = false;
                RunOrder_6_CheckBox.Checked = false;
                RunOrder_8_CheckBox.Checked = false;
                RunOrder_9_CheckBox.Checked = false;
                RunOrder_10_CheckBox.Checked = false;
            }
        }

        private void RunOrder_8_CheckBox_CheckedChanged(object sender, EventArgs e)
        {
            if (RunOrder_8_CheckBox.Checked == true)
            {
                Array.Clear(ArrayRunOrder, 0, ArrayRunOrder.Length);
                Array.Clear(ArrayItemID, 0, ArrayItemID.Length);
                RowIndex = 0;
                foreach (DataGridViewRow Row in SpotWeld_GridView.Rows)
                {
                    RunOrder_Array = Row.Cells[0].Value.ToString();
                    ArrayRunOrder[RowIndex] = Int32.Parse(RunOrder_Array);
                    ArrayItemID[RowIndex] = Row.Cells[1].Value.ToString();
                    ArrayRowNumber[RowIndex] = SpotWeld_GridView.Rows.IndexOf(Row);
                    RowIndex++;
                }
                RowIndexClick = 7;
                RunOrder_8_TextBox.BackColor = Color.Chartreuse;
                ItemID_8_TextBox.BackColor = Color.Chartreuse;
                RunOrder_1_TextBox.BackColor = System.Drawing.SystemColors.Window;
                ItemID_1_TextBox.BackColor = System.Drawing.SystemColors.Window;
                RunOrder_2_TextBox.BackColor = System.Drawing.SystemColors.Window;
                ItemID_2_TextBox.BackColor = System.Drawing.SystemColors.Window;
                RunOrder_3_TextBox.BackColor = System.Drawing.SystemColors.Window;
                ItemID_3_TextBox.BackColor = System.Drawing.SystemColors.Window;
                RunOrder_4_TextBox.BackColor = System.Drawing.SystemColors.Window;
                ItemID_4_TextBox.BackColor = System.Drawing.SystemColors.Window;
                RunOrder_5_TextBox.BackColor = System.Drawing.SystemColors.Window;
                ItemID_5_TextBox.BackColor = System.Drawing.SystemColors.Window;
                RunOrder_6_TextBox.BackColor = System.Drawing.SystemColors.Window;
                ItemID_6_TextBox.BackColor = System.Drawing.SystemColors.Window;
                RunOrder_7_TextBox.BackColor = System.Drawing.SystemColors.Window;
                ItemID_7_TextBox.BackColor = System.Drawing.SystemColors.Window;
                RunOrder_9_TextBox.BackColor = System.Drawing.SystemColors.Window;
                ItemID_9_TextBox.BackColor = System.Drawing.SystemColors.Window;
                RunOrder_10_TextBox.BackColor = System.Drawing.SystemColors.Window;
                ItemID_10_TextBox.BackColor = System.Drawing.SystemColors.Window;
                RunOrder_1_CheckBox.Checked = false;
                RunOrder_2_CheckBox.Checked = false;
                RunOrder_3_CheckBox.Checked = false;
                RunOrder_4_CheckBox.Checked = false;
                RunOrder_5_CheckBox.Checked = false;
                RunOrder_6_CheckBox.Checked = false;
                RunOrder_7_CheckBox.Checked = false;
                RunOrder_9_CheckBox.Checked = false;
                RunOrder_10_CheckBox.Checked = false;
            }
        }

        private void RunOrder_9_CheckBox_CheckedChanged(object sender, EventArgs e)
        {
            if (RunOrder_9_CheckBox.Checked == true)
            {
                Array.Clear(ArrayRunOrder, 0, ArrayRunOrder.Length);
                Array.Clear(ArrayItemID, 0, ArrayItemID.Length);
                RowIndex = 0;
                foreach (DataGridViewRow Row in SpotWeld_GridView.Rows)
                {
                    RunOrder_Array = Row.Cells[0].Value.ToString();
                    ArrayRunOrder[RowIndex] = Int32.Parse(RunOrder_Array);
                    ArrayItemID[RowIndex] = Row.Cells[1].Value.ToString();
                    ArrayRowNumber[RowIndex] = SpotWeld_GridView.Rows.IndexOf(Row);
                    RowIndex++;
                }
                RowIndexClick = 8;
                RunOrder_9_TextBox.BackColor = Color.Chartreuse;
                ItemID_9_TextBox.BackColor = Color.Chartreuse;
                RunOrder_1_TextBox.BackColor = System.Drawing.SystemColors.Window;
                ItemID_1_TextBox.BackColor = System.Drawing.SystemColors.Window;
                RunOrder_2_TextBox.BackColor = System.Drawing.SystemColors.Window;
                ItemID_2_TextBox.BackColor = System.Drawing.SystemColors.Window;
                RunOrder_3_TextBox.BackColor = System.Drawing.SystemColors.Window;
                ItemID_3_TextBox.BackColor = System.Drawing.SystemColors.Window;
                RunOrder_4_TextBox.BackColor = System.Drawing.SystemColors.Window;
                ItemID_4_TextBox.BackColor = System.Drawing.SystemColors.Window;
                RunOrder_5_TextBox.BackColor = System.Drawing.SystemColors.Window;
                ItemID_5_TextBox.BackColor = System.Drawing.SystemColors.Window;
                RunOrder_6_TextBox.BackColor = System.Drawing.SystemColors.Window;
                ItemID_6_TextBox.BackColor = System.Drawing.SystemColors.Window;
                RunOrder_7_TextBox.BackColor = System.Drawing.SystemColors.Window;
                ItemID_7_TextBox.BackColor = System.Drawing.SystemColors.Window;
                RunOrder_8_TextBox.BackColor = System.Drawing.SystemColors.Window;
                ItemID_8_TextBox.BackColor = System.Drawing.SystemColors.Window;
                RunOrder_10_TextBox.BackColor = System.Drawing.SystemColors.Window;
                ItemID_10_TextBox.BackColor = System.Drawing.SystemColors.Window;
                RunOrder_1_CheckBox.Checked = false;
                RunOrder_2_CheckBox.Checked = false;
                RunOrder_3_CheckBox.Checked = false;
                RunOrder_4_CheckBox.Checked = false;
                RunOrder_5_CheckBox.Checked = false;
                RunOrder_6_CheckBox.Checked = false;
                RunOrder_7_CheckBox.Checked = false;
                RunOrder_8_CheckBox.Checked = false;
                RunOrder_10_CheckBox.Checked = false;
            }
        }

        private void RunOrder_10_CheckBox_CheckedChanged(object sender, EventArgs e)
        {
            if (RunOrder_10_CheckBox.Checked == true)
            {
                Array.Clear(ArrayRunOrder, 0, ArrayRunOrder.Length);
                Array.Clear(ArrayItemID, 0, ArrayItemID.Length);
                RowIndex = 0;
                foreach (DataGridViewRow Row in SpotWeld_GridView.Rows)
                {
                    RunOrder_Array = Row.Cells[0].Value.ToString();
                    ArrayRunOrder[RowIndex] = Int32.Parse(RunOrder_Array);
                    ArrayItemID[RowIndex] = Row.Cells[1].Value.ToString();
                    ArrayRowNumber[RowIndex] = SpotWeld_GridView.Rows.IndexOf(Row);
                    RowIndex++;
                }
                RowIndexClick = 9;
                RunOrder_10_TextBox.BackColor = Color.Chartreuse;
                ItemID_10_TextBox.BackColor = Color.Chartreuse;
                RunOrder_1_TextBox.BackColor = System.Drawing.SystemColors.Window;
                ItemID_1_TextBox.BackColor = System.Drawing.SystemColors.Window;
                RunOrder_2_TextBox.BackColor = System.Drawing.SystemColors.Window;
                ItemID_2_TextBox.BackColor = System.Drawing.SystemColors.Window;
                RunOrder_3_TextBox.BackColor = System.Drawing.SystemColors.Window;
                ItemID_3_TextBox.BackColor = System.Drawing.SystemColors.Window;
                RunOrder_4_TextBox.BackColor = System.Drawing.SystemColors.Window;
                ItemID_4_TextBox.BackColor = System.Drawing.SystemColors.Window;
                RunOrder_5_TextBox.BackColor = System.Drawing.SystemColors.Window;
                ItemID_5_TextBox.BackColor = System.Drawing.SystemColors.Window;
                RunOrder_6_TextBox.BackColor = System.Drawing.SystemColors.Window;
                ItemID_6_TextBox.BackColor = System.Drawing.SystemColors.Window;
                RunOrder_7_TextBox.BackColor = System.Drawing.SystemColors.Window;
                ItemID_7_TextBox.BackColor = System.Drawing.SystemColors.Window;
                RunOrder_8_TextBox.BackColor = System.Drawing.SystemColors.Window;
                ItemID_8_TextBox.BackColor = System.Drawing.SystemColors.Window;
                RunOrder_9_TextBox.BackColor = System.Drawing.SystemColors.Window;
                ItemID_9_TextBox.BackColor = System.Drawing.SystemColors.Window;
                RunOrder_1_CheckBox.Checked = false;
                RunOrder_2_CheckBox.Checked = false;
                RunOrder_3_CheckBox.Checked = false;
                RunOrder_4_CheckBox.Checked = false;
                RunOrder_5_CheckBox.Checked = false;
                RunOrder_6_CheckBox.Checked = false;
                RunOrder_7_CheckBox.Checked = false;
                RunOrder_8_CheckBox.Checked = false;
                RunOrder_9_CheckBox.Checked = false;
            }
        }

        /********************************************************************************************************************
        * 
        * CheckBox Region End
        * 
        *********************************************************************************************************************/
        #endregion


        /********************************************************************************************************************
        * 
        * ComboBox Region Start
        * 
        * - CustomerCell ComboBox SelectedIndexChanged
        * - BrakePress ComboBox SelectedIndexChanged
        * 
        *********************************************************************************************************************/
        #region

        private void CustomerCell_ComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            SpotWeld_ComboBox.Items.Clear();
            CustomerCell_ComboBox.Enabled = false;
            CustomerCell_ComboBox.Visible = false;
            CustomerCell_TextBox.Visible = true;
            CustomerCell_TextBox.Text = CustomerCell_ComboBox.Text;
            ChangeCell_Button.Show();

            if (CustomerCell_ComboBox.Text == "CAT")
            {
                /*
                this.CATBrakePress_GroupBox.Location = new System.Drawing.Point(12, 798);
                this.CATBrakePress_GroupBox.Size = new System.Drawing.Size(428, 231);
                CATBrakePress_GroupBox.Visible = true;
                JohnDeereBrakePress_GroupBox.Visible = false;
                NavistarBrakePress_GroupBox.Visible = false;
                PaccarBrakePress_GroupBox.Visible = false;
                */
                SpotWeld_ComboBox.Items.AddRange(CAT_SpotWeldList);

                SqlConnection CATConnection = new SqlConnection(SQL_Source);
                string CATString = "SELECT * FROM [dbo].[CAT_Item_Data]";
                SqlDataAdapter CATDataAdapter = new SqlDataAdapter(CATString, CATConnection);
                SqlCommandBuilder CATCommandBuilder = new SqlCommandBuilder(CATDataAdapter);
                DataSet CATData = new DataSet();
                CATDataAdapter.Fill(CATData);
                CellManager_GridView.DataSource = CATData.Tables[0];
            }
            else if (CustomerCell_ComboBox.Text == "John Deere")
            {
                /*
                this.JohnDeereBrakePress_GroupBox.Location = new System.Drawing.Point(12, 798);
                this.JohnDeereBrakePress_GroupBox.Size = new System.Drawing.Size(428, 231);
                CATBrakePress_GroupBox.Visible = false;
                JohnDeereBrakePress_GroupBox.Visible = true;
                NavistarBrakePress_GroupBox.Visible = false;
                PaccarBrakePress_GroupBox.Visible = false;
                */
                SpotWeld_ComboBox.Items.AddRange(JohnDeere_SpotWeldList);

                SqlConnection JohnDeereConnection = new SqlConnection(SQL_Source);
                string JohnDeereString = "SELECT * FROM [dbo].[JohnDeere_Item_Data]";
                SqlDataAdapter JohnDeereDataAdapter = new SqlDataAdapter(JohnDeereString, JohnDeereConnection);
                SqlCommandBuilder JohnDeereCommandBuilder = new SqlCommandBuilder(JohnDeereDataAdapter);
                DataSet JohnDeereData = new DataSet();
                JohnDeereDataAdapter.Fill(JohnDeereData);
                CellManager_GridView.DataSource = JohnDeereData.Tables[0];
            }
            else if (CustomerCell_ComboBox.Text == "Navistar")
            {
                /*
                this.NavistarBrakePress_GroupBox.Location = new System.Drawing.Point(12, 798);
                this.NavistarBrakePress_GroupBox.Size = new System.Drawing.Size(602, 231);
                CATBrakePress_GroupBox.Visible = false;
                JohnDeereBrakePress_GroupBox.Visible = false;
                NavistarBrakePress_GroupBox.Visible = true;
                PaccarBrakePress_GroupBox.Visible = false;
                */
                SpotWeld_ComboBox.Items.AddRange(Navistar_SpotWeldList);

                SqlConnection NavistarConnection = new SqlConnection(SQL_Source);
                string NavistarString = "SELECT * FROM [dbo].[Navistar_Item_Data]";
                SqlDataAdapter NavistarDataAdapter = new SqlDataAdapter(NavistarString, NavistarConnection);
                SqlCommandBuilder NavistarCommandBuilder = new SqlCommandBuilder(NavistarDataAdapter);
                DataSet NavistarData = new DataSet();
                NavistarDataAdapter.Fill(NavistarData);
                CellManager_GridView.DataSource = NavistarData.Tables[0];
            }
            else if (CustomerCell_ComboBox.Text == "Paccar")
            {
                /*
                this.PaccarBrakePress_GroupBox.Location = new System.Drawing.Point(12, 798);
                this.PaccarBrakePress_GroupBox.Size = new System.Drawing.Size(602, 231);
                CATBrakePress_GroupBox.Visible = false;
                JohnDeereBrakePress_GroupBox.Visible = false;
                NavistarBrakePress_GroupBox.Visible = false;
                PaccarBrakePress_GroupBox.Visible = true;
                */
                SpotWeld_ComboBox.Items.AddRange(Paccar_SpotWeldList);

                SqlConnection PaccarConnection = new SqlConnection(SQL_Source);
                string PaccarString = "SELECT * FROM [dbo].[Paccar_Item_Data]";
                //string PaccarString = "SELECT ItemID, Customer, CustomerItemID, TotalRuns, PartsManufactured, PartsPerMinute, SetupTime, BP1083, BP1155, BP1158, BP1175, BP1176 FROM [dbo].[Paccar_Item_Data]";
                SqlDataAdapter PaccarDataAdapter = new SqlDataAdapter(PaccarString, PaccarConnection);
                SqlCommandBuilder PaccarCommandBuilder = new SqlCommandBuilder(PaccarDataAdapter);
                DataSet PaccarData = new DataSet();
                PaccarDataAdapter.Fill(PaccarData);
                CellManager_GridView.DataSource = PaccarData.Tables[0];
            }
        }

        private void SpotWeld_ComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            // CAT Spot Weld
            if (SpotWeld_ComboBox.Text == "1088")
            {
                Refresh_Data = "SELECT * FROM [dbo].[SW_1088_Schedule] ORDER BY RunOrder ASC";
                Schedule_Count = "SELECT COUNT(*) FROM[dbo].[SW_1088_Schedule]";
                SQLRemoveCommand = "DELETE FROM [dbo].[SW_1088_Schedule] WHERE RunOrder=@RunOrder";
                SQLAddCommand = "INSERT INTO [dbo].[SW_1088_Schedule] (RunOrder,ItemID,JobID,Customer,CustomerItemID,Tooling,ToolingLocation,FixtureLocation,PartsOrdered,EstimatedRunTime) VALUES (@RunOrder,@ItemID,@JobID,@Customer,@CustomerItemID,@Tooling,@ToolingLocation,@FixtureLocation,@PartsOrdered,@EstimatedRunTime)";
                SQLUpdateCommand = "UPDATE [dbo].[SW_1088_Schedule] SET RunOrder=@RunOrder WHERE ItemID=@ItemID";

                SqlConnection Connection_1088 = new SqlConnection(SQL_Source);
                string SW1088 = "SELECT * FROM [dbo].[SW_1088_Schedule]";
                SqlDataAdapter DataAdapter_1088 = new SqlDataAdapter(SW1088, Connection_1088);
                SqlCommandBuilder CommandBuilder_1088 = new SqlCommandBuilder(DataAdapter_1088);
                DataSet Data_1088 = new DataSet();
                DataAdapter_1088.Fill(Data_1088);
                SpotWeld_GridView.DataSource = Data_1088.Tables[0];

            }
            else if (SpotWeld_ComboBox.Text == "Second CAT")
            {
                Refresh_Data = "SELECT * FROM [dbo].[SW_1088_Schedule] ORDER BY RunOrder ASC";
                Schedule_Count = "SELECT COUNT(*) FROM[dbo].[SW_1088_Schedule]";
                SQLRemoveCommand = "DELETE FROM [dbo].[SW_1088_Schedule] WHERE RunOrder=@RunOrder";
                SQLAddCommand = "INSERT INTO [dbo].[SW_1088_Schedule] (RunOrder,ItemID,JobID,Customer,CustomerItemID,Tooling,ToolingLocation,FixtureLocation,PartsOrdered,EstimatedRunTime) VALUES (@RunOrder,@ItemID,@JobID,@Customer,@CustomerItemID,@Tooling,@ToolingLocation,@FixtureLocation,@PartsOrdered,@EstimatedRunTime)";
                SQLUpdateCommand = "UPDATE [dbo].[SW_1088_Schedule] SET RunOrder=@RunOrder WHERE ItemID=@ItemID";

                SqlConnection Connection_1088 = new SqlConnection(SQL_Source);
                string SW1088 = "SELECT * FROM [dbo].[SW_1088_Schedule]";
                SqlDataAdapter DataAdapter_1088 = new SqlDataAdapter(SW1088, Connection_1088);
                SqlCommandBuilder CommandBuilder_1088 = new SqlCommandBuilder(DataAdapter_1088);
                DataSet Data_1088 = new DataSet();
                DataAdapter_1088.Fill(Data_1088);
                SpotWeld_GridView.DataSource = Data_1088.Tables[0];
            }
            // John Deere Spot Weld
            else if (SpotWeld_ComboBox.Text == "150R")
            {
                Refresh_Data = "SELECT * FROM [dbo].[SW_150R_Schedule] ORDER BY RunOrder ASC";
                Schedule_Count = "SELECT COUNT(*) FROM[dbo].[SW_150R_Schedule]";
                SQLRemoveCommand = "DELETE FROM [dbo].[SW_150R_Schedule] WHERE RunOrder=@RunOrder";
                SQLAddCommand = "INSERT INTO [dbo].[SW_150R_Schedule] (RunOrder,ItemID,JobID,Customer,CustomerItemID,Tooling,ToolingLocation,FixtureLocation,PartsOrdered,EstimatedRunTime) VALUES (@RunOrder,@ItemID,@JobID,@Customer,@CustomerItemID,@Tooling,@ToolingLocation,@FixtureLocation,@PartsOrdered,@EstimatedRunTime)";
                SQLUpdateCommand = "UPDATE [dbo].[SW_150R_Schedule] SET RunOrder=@RunOrder WHERE ItemID=@ItemID";

                SqlConnection Connection_150R = new SqlConnection(SQL_Source);
                string SW150R = "SELECT * FROM [dbo].[SW_150R_Schedule]";
                SqlDataAdapter DataAdapter_150R = new SqlDataAdapter(SW150R, Connection_150R);
                SqlCommandBuilder CommandBuilder_150R = new SqlCommandBuilder(DataAdapter_150R);
                DataSet Data_150R = new DataSet();
                DataAdapter_150R.Fill(Data_150R);
                SpotWeld_GridView.DataSource = Data_150R.Tables[0];
            }
            else if (SpotWeld_ComboBox.Text == "Second John Deere")
            {
                Refresh_Data = "SELECT * FROM [dbo].[SW_150R_Schedule] ORDER BY RunOrder ASC";
                Schedule_Count = "SELECT COUNT(*) FROM[dbo].[SW_150R_Schedule]";
                SQLRemoveCommand = "DELETE FROM [dbo].[SW_150R_Schedule] WHERE RunOrder=@RunOrder";
                SQLAddCommand = "INSERT INTO [dbo].[SW_150R_Schedule] (RunOrder,ItemID,JobID,Customer,CustomerItemID,Tooling,ToolingLocation,FixtureLocation,PartsOrdered,EstimatedRunTime) VALUES (@RunOrder,@ItemID,@JobID,@Customer,@CustomerItemID,@Tooling,@ToolingLocation,@FixtureLocation,@PartsOrdered,@EstimatedRunTime)";
                SQLUpdateCommand = "UPDATE [dbo].[SW_150R_Schedule] SET RunOrder=@RunOrder WHERE ItemID=@ItemID";

                SqlConnection Connection_150R = new SqlConnection(SQL_Source);
                string SW150R = "SELECT * FROM [dbo].[SW_150R_Schedule]";
                SqlDataAdapter DataAdapter_150R = new SqlDataAdapter(SW150R, Connection_150R);
                SqlCommandBuilder CommandBuilder_150R = new SqlCommandBuilder(DataAdapter_150R);
                DataSet Data_150R = new DataSet();
                DataAdapter_150R.Fill(Data_150R);
                SpotWeld_GridView.DataSource = Data_150R.Tables[0];
            }
            // Navistar Spot Weld
            else if (SpotWeld_ComboBox.Text == "First Spot Weld")
            {
                Refresh_Data = "SELECT * FROM [dbo].[SW_154R_Schedule] ORDER BY RunOrder ASC";
                Schedule_Count = "SELECT COUNT(*) FROM[dbo].[SW_154R_Schedule]";
                SQLRemoveCommand = "DELETE FROM [dbo].[SW_154R_Schedule] WHERE RunOrder=@RunOrder";
                SQLAddCommand = "INSERT INTO [dbo].[SW_154R_Schedule] (RunOrder,ItemID,JobID,Customer,CustomerItemID,Tooling,ToolingLocation,FixtureLocation,PartsOrdered,EstimatedRunTime) VALUES (@RunOrder,@ItemID,@JobID,@Customer,@CustomerItemID,@Tooling,@ToolingLocation,@FixtureLocation,@PartsOrdered,@EstimatedRunTime)";
                SQLUpdateCommand = "UPDATE [dbo].[SW_154R_Schedule] SET RunOrder=@RunOrder WHERE ItemID=@ItemID";

                SqlConnection Connection_154R = new SqlConnection(SQL_Source);
                string SW154R = "SELECT * FROM [dbo].[SW_154R_Schedule]";
                SqlDataAdapter DataAdapter_154R = new SqlDataAdapter(SW154R, Connection_154R);
                SqlCommandBuilder CommandBuilder_154R = new SqlCommandBuilder(DataAdapter_154R);
                DataSet Data_154R = new DataSet();
                DataAdapter_154R.Fill(Data_154R);
                SpotWeld_GridView.DataSource = Data_154R.Tables[0];
            }
            else if (SpotWeld_ComboBox.Text == "121R")
            {
                Refresh_Data = "SELECT * FROM [dbo].[SW_121R_Schedule] ORDER BY RunOrder ASC";
                Schedule_Count = "SELECT COUNT(*) FROM[dbo].[SW_121R_Schedule]";
                SQLRemoveCommand = "DELETE FROM [dbo].[SW_121R_Schedule] WHERE RunOrder=@RunOrder";
                SQLAddCommand = "INSERT INTO [dbo].[SW_121R_Schedule] (RunOrder,ItemID,JobID,Customer,CustomerItemID,Tooling,ToolingLocation,FixtureLocation,PartsOrdered,EstimatedRunTime) VALUES (@RunOrder,@ItemID,@JobID,@Customer,@CustomerItemID,@Tooling,@ToolingLocation,@FixtureLocation,@PartsOrdered,@EstimatedRunTime)";
                SQLUpdateCommand = "UPDATE [dbo].[SW_121R_Schedule] SET RunOrder=@RunOrder WHERE ItemID=@ItemID";

                SqlConnection Connection_121R = new SqlConnection(SQL_Source);
                string SW121R = "SELECT * FROM [dbo].[SW_121R_Schedule]";
                SqlDataAdapter DataAdapter_121R = new SqlDataAdapter(SW121R, Connection_121R);
                SqlCommandBuilder CommandBuilder_121R = new SqlCommandBuilder(DataAdapter_121R);
                DataSet Data_121R = new DataSet();
                DataAdapter_121R.Fill(Data_121R);
                SpotWeld_GridView.DataSource = Data_121R.Tables[0];
            }
            else if (SpotWeld_ComboBox.Text == "154R")
            {
                Refresh_Data = "SELECT * FROM [dbo].[SW_154R_Schedule] ORDER BY RunOrder ASC";
                Schedule_Count = "SELECT COUNT(*) FROM[dbo].[SW_154R_Schedule]";
                SQLRemoveCommand = "DELETE FROM [dbo].[SW_154R_Schedule] WHERE RunOrder=@RunOrder";
                SQLAddCommand = "INSERT INTO [dbo].[SW_154R_Schedule] (RunOrder,ItemID,JobID,Customer,CustomerItemID,Tooling,ToolingLocation,FixtureLocation,PartsOrdered,EstimatedRunTime) VALUES (@RunOrder,@ItemID,@JobID,@Customer,@CustomerItemID,@Tooling,@ToolingLocation,@FixtureLocation,@PartsOrdered,@EstimatedRunTime)";
                SQLUpdateCommand = "UPDATE [dbo].[SW_154R_Schedule] SET RunOrder=@RunOrder WHERE ItemID=@ItemID";

                SqlConnection Connection_154R = new SqlConnection(SQL_Source);
                string SW154R = "SELECT * FROM [dbo].[SW_154R_Schedule]";
                SqlDataAdapter DataAdapter_154R = new SqlDataAdapter(SW154R, Connection_154R);
                SqlCommandBuilder CommandBuilder_154R = new SqlCommandBuilder(DataAdapter_154R);
                DataSet Data_154R = new DataSet();
                DataAdapter_154R.Fill(Data_154R);
                SpotWeld_GridView.DataSource = Data_154R.Tables[0];
            }
            // Paccar Spot Weld
            else if (SpotWeld_ComboBox.Text == "153R")
            {
                Refresh_Data = "SELECT * FROM [dbo].[SW_153R_Schedule] ORDER BY RunOrder ASC";
                Schedule_Count = "SELECT COUNT(*) FROM[dbo].[SW_153R_Schedule]";
                SQLRemoveCommand = "DELETE FROM [dbo].[SW_153R_Schedule] WHERE RunOrder=@RunOrder";
                SQLAddCommand = "INSERT INTO [dbo].[SW_153R_Schedule] (RunOrder,ItemID,JobID,Customer,CustomerItemID,Tooling,ToolingLocation,FixtureLocation,PartsOrdered,EstimatedRunTime) VALUES (@RunOrder,@ItemID,@JobID,@Customer,@CustomerItemID,@Tooling,@ToolingLocation,@FixtureLocation,@PartsOrdered,@EstimatedRunTime)";
                SQLUpdateCommand = "UPDATE [dbo].[SW_153R_Schedule] SET RunOrder=@RunOrder WHERE ItemID=@ItemID";

                SqlConnection Connection_153R = new SqlConnection(SQL_Source);
                string SW153R = "SELECT * FROM [dbo].[SW_153R_Schedule]";
                SqlDataAdapter DataAdapter_153R = new SqlDataAdapter(SW153R, Connection_153R);
                SqlCommandBuilder CommandBuilder_153R = new SqlCommandBuilder(DataAdapter_153R);
                DataSet Data_153R = new DataSet();
                DataAdapter_153R.Fill(Data_153R);
                SpotWeld_GridView.DataSource = Data_153R.Tables[0];
            }
            else if (SpotWeld_ComboBox.Text == "155R")
            {
                Refresh_Data = "SELECT * FROM [dbo].[SW_155R_Schedule] ORDER BY RunOrder ASC";
                Schedule_Count = "SELECT COUNT(*) FROM[dbo].[SW_155R_Schedule]";
                SQLRemoveCommand = "DELETE FROM [dbo].[SW_155R_Schedule] WHERE RunOrder=@RunOrder";
                SQLAddCommand = "INSERT INTO [dbo].[SW_155R_Schedule] (RunOrder,ItemID,JobID,Customer,CustomerItemID,Tooling,ToolingLocation,FixtureLocation,PartsOrdered,EstimatedRunTime) VALUES (@RunOrder,@ItemID,@JobID,@Customer,@CustomerItemID,@Tooling,@ToolingLocation,@FixtureLocation,@PartsOrdered,@EstimatedRunTime)";
                SQLUpdateCommand = "UPDATE [dbo].[SW_155R_Schedule] SET RunOrder=@RunOrder WHERE ItemID=@ItemID";

                SqlConnection Connection_155R = new SqlConnection(SQL_Source);
                string SW155R = "SELECT * FROM [dbo].[SW_155R_Schedule]";
                SqlDataAdapter DataAdapter_155R = new SqlDataAdapter(SW155R, Connection_155R);
                SqlCommandBuilder CommandBuilder_155R = new SqlCommandBuilder(DataAdapter_155R);
                DataSet Data_155R = new DataSet();
                DataAdapter_155R.Fill(Data_155R);
                SpotWeld_GridView.DataSource = Data_155R.Tables[0];
            }
        }

        /********************************************************************************************************************
        * 
        * ComboBox Region End
        * 
        ********************************************************************************************************************/
        #endregion

        /********************************************************************************************************************
        * 
        * GridView Region Start 
        * 
        * - CellManager GridView CellClick
        * - SpotWeld GridView CellClick
        * 
        *********************************************************************************************************************/
        #region

        private void CellManager_GridView_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
            {
                if (CustomerCell_ComboBox.Text == "CAT")
                {
                    DataGridViewRow Row = CellManager_GridView.Rows[e.RowIndex];
                    ItemID_TextBox.Text = Row.Cells[0].Value.ToString();
                    Customer_TextBox.Text = Row.Cells[1].Value.ToString();
                    CustomerItemID_TextBox.Text = Row.Cells[2].Value.ToString();
                    JobID_TextBox.Text = Row.Cells[3].Value.ToString();
                    Fixture_TextBox.Text = Row.Cells[5].Value.ToString();
                    FixtureLocation_TextBox.Text = Row.Cells[6].Value.ToString();
                    TotalRuns_TextBox.Text = Row.Cells[13].Value.ToString();
                    PartsManufactured_TextBox.Text = Row.Cells[14].Value.ToString();
                    PPM_TextBox.Text = Row.Cells[15].Value.ToString();
                    SetupTime_TextBox.Text = Row.Cells[16].Value.ToString();
                }
                else if (CustomerCell_ComboBox.Text == "John Deere")
                {
                    DataGridViewRow Row = CellManager_GridView.Rows[e.RowIndex];
                    ItemID_TextBox.Text = Row.Cells[0].Value.ToString();
                    Customer_TextBox.Text = Row.Cells[1].Value.ToString();
                    CustomerItemID_TextBox.Text = Row.Cells[2].Value.ToString();
                    JobID_TextBox.Text = Row.Cells[3].Value.ToString();
                    Fixture_TextBox.Text = Row.Cells[5].Value.ToString();
                    FixtureLocation_TextBox.Text = Row.Cells[6].Value.ToString();
                    TotalRuns_TextBox.Text = Row.Cells[13].Value.ToString();
                    PartsManufactured_TextBox.Text = Row.Cells[14].Value.ToString();
                    PPM_TextBox.Text = Row.Cells[15].Value.ToString();
                    SetupTime_TextBox.Text = Row.Cells[16].Value.ToString();
                }
                else if (CustomerCell_ComboBox.Text == "Navistar")
                {
                    DataGridViewRow Row = CellManager_GridView.Rows[e.RowIndex];
                    ItemID_TextBox.Text = Row.Cells[0].Value.ToString();
                    Customer_TextBox.Text = Row.Cells[1].Value.ToString();
                    CustomerItemID_TextBox.Text = Row.Cells[2].Value.ToString();
                    JobID_TextBox.Text = Row.Cells[3].Value.ToString();
                    Fixture_TextBox.Text = Row.Cells[5].Value.ToString();
                    FixtureLocation_TextBox.Text = Row.Cells[6].Value.ToString();
                    TotalRuns_TextBox.Text = Row.Cells[13].Value.ToString();
                    PartsManufactured_TextBox.Text = Row.Cells[14].Value.ToString();
                    PPM_TextBox.Text = Row.Cells[15].Value.ToString();
                    SetupTime_TextBox.Text = Row.Cells[16].Value.ToString();
                }
                else if (CustomerCell_ComboBox.Text == "Paccar")
                {
                    DataGridViewRow Row = CellManager_GridView.Rows[e.RowIndex];
                    ItemID_TextBox.Text = Row.Cells[0].Value.ToString();
                    Customer_TextBox.Text = Row.Cells[1].Value.ToString();
                    CustomerItemID_TextBox.Text = Row.Cells[2].Value.ToString();
                    JobID_TextBox.Text = Row.Cells[3].Value.ToString();
                    Fixture_TextBox.Text = Row.Cells[5].Value.ToString();
                    FixtureLocation_TextBox.Text = Row.Cells[6].Value.ToString();
                    TotalRuns_TextBox.Text = Row.Cells[13].Value.ToString();
                    PartsManufactured_TextBox.Text = Row.Cells[14].Value.ToString();
                    PPM_TextBox.Text = Row.Cells[15].Value.ToString();
                    SetupTime_TextBox.Text = Row.Cells[16].Value.ToString();
                }
            }
        }

        private void SpotWeld_GridView_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            CheckBox[] RunOrderCheckBoxArray = { RunOrder_1_CheckBox, RunOrder_2_CheckBox, RunOrder_3_CheckBox, RunOrder_4_CheckBox, RunOrder_5_CheckBox, RunOrder_6_CheckBox, RunOrder_7_CheckBox, RunOrder_8_CheckBox, RunOrder_9_CheckBox, RunOrder_10_CheckBox };
            TextBox[] RunOrderTextBoxArray = { RunOrder_1_TextBox, RunOrder_2_TextBox, RunOrder_3_TextBox, RunOrder_4_TextBox, RunOrder_5_TextBox, RunOrder_6_TextBox, RunOrder_7_TextBox, RunOrder_8_TextBox, RunOrder_9_TextBox, RunOrder_10_TextBox };
            TextBox[] ItemIDTextBoxArray = { ItemID_1_TextBox, ItemID_2_TextBox, ItemID_3_TextBox, ItemID_4_TextBox, ItemID_5_TextBox, ItemID_6_TextBox, ItemID_7_TextBox, ItemID_8_TextBox, ItemID_9_TextBox, ItemID_10_TextBox };
            Array.Clear(ArrayRunOrder, 0, ArrayRunOrder.Length);
            Array.Clear(ArrayItemID, 0, ArrayItemID.Length);
            Array.Clear(ArrayRowNumber, 0, ArrayItemID.Length);
            RowIndex = 0;
            SpotWeld_GridView.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            if (e.RowIndex >= 0)
            {
                DataGridViewRow DataRow = SpotWeld_GridView.Rows[e.RowIndex];
                RowIndexClicked = SpotWeld_GridView.Rows.IndexOf(DataRow).ToString();
                RowIndexClick = Int32.Parse(RowIndexClicked);
            }
            foreach (DataGridViewRow Row in SpotWeld_GridView.Rows)
            {
                RunOrder_Array = Row.Cells[0].Value.ToString();
                ArrayRunOrder[RowIndex] = Int32.Parse(RunOrder_Array);
                ArrayItemID[RowIndex] = Row.Cells[1].Value.ToString();
                ArrayRowNumber[RowIndex] = SpotWeld_GridView.Rows.IndexOf(Row);
                RowIndex++;
            }
            for (int r = 0; r < RowIndex; ++r)
            {
                if (ArrayRowNumber[r] == RowIndexClick)
                {
                    //RunOrderTextBoxArray[r].BackColor = Color.Chartreuse;
                    //ItemIDTextBoxArray[r].BackColor = Color.Chartreuse;
                    RunOrderCheckBoxArray[r].Checked = true;
                }
                else
                {
                    //RunOrderTextBoxArray[r].BackColor = System.Drawing.SystemColors.Window;
                    //ItemIDTextBoxArray[r].BackColor = System.Drawing.SystemColors.Window;
                    RunOrderCheckBoxArray[r].Checked = false;
                }
            }
            for (int m = 0; m < RowIndex; m++)
            {
                Console.WriteLine("Run Order: " + ArrayRunOrder[m] + " Item ID: " + ArrayItemID[m] + " Row Number: " + ArrayRowNumber[m] + " Row Clicked: " + RowIndexClick + " Total Number of Rows: " + RowIndex);
            }
            SpotWeld_GridView.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
        }

        /********************************************************************************************************************
        * 
        * GridView Region End
        * 
        ********************************************************************************************************************/
        #endregion

        /********************************************************************************************************************
        * 
        * Methods Region Start
        * 
        * - EmployeeLogOff
        * - RefreshJobs
        * - GetJobQueueRunOrder
        * - AddJobToQueue
        * - RemoveJobFromQueue
        * - CountRows
        * 
        *********************************************************************************************************************/
        #region

        private void Clear()
        {
            // ItemInformation_GroupBox
            ItemID_TextBox.Clear();
            //JobID_TextBox.Clear();
            SearchItemID_CheckBox.Checked = true;
            SearchCustomerItemID_CheckBox.Checked = false;
            Customer_TextBox.Clear();
            CustomerItemID_TextBox.Clear();

            // ItemStatistics_GroupBox
            TotalRuns_TextBox.Clear();
            PartsManufactured_TextBox.Clear();
            SetupTime_TextBox.Clear();
            PPM_TextBox.Clear();

            // OrderData_GroupBox
            PartsOnOrder_TextBox.Clear();
            EstimatedRunTime_TextBox.Clear();
            //EstimatedStartTime_TextBox.Clear();
            //QueuePosition_TextBox.Clear();

            // OrderData_GroupBox
            Fixture_TextBox.Clear();
            FixtureLocation_TextBox.Clear();
            //Scanner3D_TextBox.Clear();

            /*
            // CATBrakePress_GroupBox      
            BP1107_TextBox.Clear();
            BP1139_TextBox.Clear();
            BP1177_TextBox.Clear();

            // JohnDeereBrakePress_GroupBox      
            BP1127_TextBox.Clear();
            BP1178_TextBox.Clear();

            // NavistarBrakePress_GroupBox      
            BP1065_TextBox.Clear();
            BP1108_TextBox.Clear();
            BP1156_TextBox.Clear();
            BP1720_TextBox.Clear();

            // PaccarBrakePress_GroupBox      
            BP1083_TextBox.Clear();
            BP1155_TextBox.Clear();
            BP1158_TextBox.Clear();
            BP1175_TextBox.Clear();
            BP1176_TextBox.Clear();
            */
        }

        private void EmployeeLogOff()
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

        private void RefreshJobs()
        {
            Array.Clear(ArrayRunOrder, 0, ArrayRunOrder.Length);
            Array.Clear(ArrayItemID, 0, ArrayItemID.Length);
            RowIndex = 0;
            Button[] RemoveButtonArray = { Remove_1_Button, Remove_2_Button, Remove_3_Button, Remove_4_Button, Remove_5_Button, Remove_6_Button, Remove_7_Button, Remove_8_Button, Remove_9_Button, Remove_10_Button };
            CheckBox[] RunOrderCheckBoxArray = { RunOrder_1_CheckBox, RunOrder_2_CheckBox, RunOrder_3_CheckBox, RunOrder_4_CheckBox, RunOrder_5_CheckBox, RunOrder_6_CheckBox, RunOrder_7_CheckBox, RunOrder_8_CheckBox, RunOrder_9_CheckBox, RunOrder_10_CheckBox };
            TextBox[] RunOrderTextBoxArray = { RunOrder_1_TextBox, RunOrder_2_TextBox, RunOrder_3_TextBox, RunOrder_4_TextBox, RunOrder_5_TextBox, RunOrder_6_TextBox, RunOrder_7_TextBox, RunOrder_8_TextBox, RunOrder_9_TextBox, RunOrder_10_TextBox };
            TextBox[] ItemIDTextBoxArray = { ItemID_1_TextBox, ItemID_2_TextBox, ItemID_3_TextBox, ItemID_4_TextBox, ItemID_5_TextBox, ItemID_6_TextBox, ItemID_7_TextBox, ItemID_8_TextBox, ItemID_9_TextBox, ItemID_10_TextBox };
            Label[] RunLabelArray = { Run_1_Label, Run_2_Label, Run_3_Label, Run_4_Label, Run_5_Label, Run_6_Label, Run_7_Label, Run_8_Label, Run_9_Label, Run_10_Label };
            Label[] ItemLabelArray = { Item_1_Label, Item_2_Label, Item_3_Label, Item_4_Label, Item_5_Label, Item_6_Label, Item_7_Label, Item_8_Label, Item_9_Label, Item_10_Label };
            string RefreshDataString = Refresh_Data;
            SqlConnection RefreshConnection = new SqlConnection(SQL_Source);
            SqlDataAdapter RefreshDataAdapter = new SqlDataAdapter(RefreshDataString, RefreshConnection);
            SqlCommandBuilder RefreshCommandBuilder = new SqlCommandBuilder(RefreshDataAdapter);
            DataSet RefreshData = new DataSet();
            RefreshDataAdapter.Fill(RefreshData);
            SpotWeld_GridView.DataSource = RefreshData.Tables[0];

            ClearCombo();

            rows = 0;
            string CountString = Schedule_Count;
            SqlConnection CountConnection = new SqlConnection(SQL_Source);
            SqlCommand CountCommand = new SqlCommand(CountString, CountConnection);
            CountConnection.Open();
            rows = (int)CountCommand.ExecuteScalar();
            CountConnection.Close();

            foreach (DataGridViewRow Row in SpotWeld_GridView.Rows)
            {
                if (Row.Index < rows)
                {
                    RunOrderTextBoxArray[Row.Index].Text = Row.Cells[0].Value.ToString();
                    ItemIDTextBoxArray[Row.Index].Text = Row.Cells[1].Value.ToString();
                    RunOrderTextBoxArray[Row.Index].BackColor = System.Drawing.SystemColors.Window;
                    ItemIDTextBoxArray[Row.Index].BackColor = System.Drawing.SystemColors.Window;
                    RemoveButtonArray[Row.Index].Show();
                    RunOrderCheckBoxArray[Row.Index].Show();
                    RunOrderTextBoxArray[Row.Index].Show();
                    ItemIDTextBoxArray[Row.Index].Show();
                    RunLabelArray[Row.Index].Show();
                    ItemLabelArray[Row.Index].Show();
                }
            }
        }

        private void GetJobQueueRunOrder()
        {
            string ScheduleCount = Schedule_Count;
            SqlConnection count = new SqlConnection(SQL_Source);
            SqlCommand countData = new SqlCommand(ScheduleCount, count);
            count.Open();
            RunOrder = (int)countData.ExecuteScalar();
            count.Close();
            RunOrder = RunOrder + 1;
        }

        private void AddJobToQueue()
        {
            SqlConnection Job_Connection = new SqlConnection(SQL_Source);
            SqlCommand Add_Job = new SqlCommand();
            Add_Job.CommandType = System.Data.CommandType.Text;
            Add_Job.CommandText = SQLAddCommand;
            Add_Job.Connection = Job_Connection;
            Add_Job.Parameters.AddWithValue("@RunOrder", RunOrder.ToString());
            Add_Job.Parameters.AddWithValue("@ItemID", ItemID_TextBox.Text);
            Add_Job.Parameters.AddWithValue("@JobID", JobID_TextBox.Text);
            Add_Job.Parameters.AddWithValue("@Customer", Customer_TextBox.Text);
            Add_Job.Parameters.AddWithValue("@CustomerItemID", CustomerItemID_TextBox.Text);
            Add_Job.Parameters.AddWithValue("@Fixture", Fixture_TextBox.Text);
            Add_Job.Parameters.AddWithValue("@FixtureLocation", FixtureLocation_TextBox.Text);
            Add_Job.Parameters.AddWithValue("@PartsOrdered", PartsOnOrder_TextBox.Text);
            Add_Job.Parameters.AddWithValue("@EstimatedRunTime", EstimatedRunTime_TextBox.Text);
            Job_Connection.Open();
            Add_Job.ExecuteNonQuery();
            Job_Connection.Close();
        }

        private void RemoveJobFromQueue()
        {
            try
            {
                SqlConnection Job_Connection = new SqlConnection(SQL_Source);
                SqlCommand Delete_Job = new SqlCommand();
                Delete_Job.CommandType = System.Data.CommandType.Text;
                Delete_Job.CommandText = SQLRemoveCommand;
                Delete_Job.Connection = Job_Connection;
                Delete_Job.Parameters.AddWithValue("@RunOrder", RemoveItem);
                Job_Connection.Open();
                Delete_Job.ExecuteNonQuery();
                Job_Connection.Close();
                MessageBox.Show("Job Was Successfully Removed");
            }
            catch (SqlException)
            {
                MessageBox.Show("Error Removing Job");
            }

        }

        private void CountRows()
        {
            rows = 0;
            string CountRowsString = Schedule_Count;
            SqlConnection count = new SqlConnection(SQL_Source);
            SqlCommand countRows = new SqlCommand(CountRowsString, count);
            count.Open();
            rows = (int)countRows.ExecuteScalar();
            count.Close();
        }

        private void GetCell()
        {

        }

        private void ClearCombo()
        {
            Button[] RemoveButtonArray = { Remove_1_Button, Remove_2_Button, Remove_3_Button, Remove_4_Button, Remove_5_Button, Remove_6_Button, Remove_7_Button, Remove_8_Button, Remove_9_Button, Remove_10_Button };
            TextBox[] RunOrderTextBoxArray = { RunOrder_1_TextBox, RunOrder_2_TextBox, RunOrder_3_TextBox, RunOrder_4_TextBox, RunOrder_5_TextBox, RunOrder_6_TextBox, RunOrder_7_TextBox, RunOrder_8_TextBox, RunOrder_9_TextBox, RunOrder_10_TextBox };
            TextBox[] ItemIDTextBoxArray = { ItemID_1_TextBox, ItemID_2_TextBox, ItemID_3_TextBox, ItemID_4_TextBox, ItemID_5_TextBox, ItemID_6_TextBox, ItemID_7_TextBox, ItemID_8_TextBox, ItemID_9_TextBox, ItemID_10_TextBox };
            CheckBox[] RunOrderCheckBoxArray = { RunOrder_1_CheckBox, RunOrder_2_CheckBox, RunOrder_3_CheckBox, RunOrder_4_CheckBox, RunOrder_5_CheckBox, RunOrder_6_CheckBox, RunOrder_7_CheckBox, RunOrder_8_CheckBox, RunOrder_9_CheckBox, RunOrder_10_CheckBox };
            Label[] RunLabelArray = { Run_1_Label, Run_2_Label, Run_3_Label, Run_4_Label, Run_5_Label, Run_6_Label, Run_7_Label, Run_8_Label, Run_9_Label, Run_10_Label };
            Label[] ItemLabelArray = { Item_1_Label, Item_2_Label, Item_3_Label, Item_4_Label, Item_5_Label, Item_6_Label, Item_7_Label, Item_8_Label, Item_9_Label, Item_10_Label };
            for (int i = 0; i < 10; i++)
            {
                RunLabelArray[i].Hide();
                RunOrderTextBoxArray[i].Hide();
                ItemLabelArray[i].Hide();
                RunOrderCheckBoxArray[i].Checked = false;
                RunOrderCheckBoxArray[i].Hide();
                ItemIDTextBoxArray[i].Clear();
                ItemIDTextBoxArray[i].Hide();
                RemoveButtonArray[i].Hide();
                RunOrderTextBoxArray[i].BackColor = System.Drawing.SystemColors.Window;
                ItemIDTextBoxArray[i].BackColor = System.Drawing.SystemColors.Window;
            }
        }

        private void DeleteCombo()
        {
            Array.Clear(ArrayRunOrder, 0, ArrayRunOrder.Length);
            Array.Clear(ArrayItemID, 0, ArrayItemID.Length);
            RowIndex = 0;
            foreach (DataGridViewRow Row in SpotWeld_GridView.Rows)
            {
                RunOrder_Array = Row.Cells[0].Value.ToString();
                ArrayRunOrder[RowIndex] = Int32.Parse(RunOrder_Array);
                ArrayItemID[RowIndex] = Row.Cells[1].Value.ToString();
                ArrayRowNumber[RowIndex] = SpotWeld_GridView.Rows.IndexOf(Row);
                RowIndex++;
            }
            if (RowIndexClick > 0)
            {
                for (int i = 0; i < RowIndex; i++)
                {
                    if (ArrayRowNumber[i] > RowIndexClick)
                    {
                        ArrayRunOrder[i] -= 1;
                        ArrayRowNumber[i] -= 1;
                    }
                }
            }
            try
            {
                SqlConnection Job_Connection = new SqlConnection(SQL_Source);
                SqlCommand Delete_Job = new SqlCommand();
                Delete_Job.CommandType = System.Data.CommandType.Text;
                Delete_Job.CommandText = SQLRemoveCommand;
                Delete_Job.Connection = Job_Connection;
                Delete_Job.Parameters.AddWithValue("@RunOrder", RemoveItemID);
                Job_Connection.Open();
                Delete_Job.ExecuteNonQuery();
                Job_Connection.Close();
                MessageBox.Show("Job Was Successfully Removed");
            }
            catch (SqlException)
            {
                MessageBox.Show("Error Removing Job");
            }
            try
            {
                for (int x = 0; x < rows; x++)
                {
                    SqlConnection Job_Connection = new SqlConnection(SQL_Source);
                    SqlCommand Edit_Job = new SqlCommand();
                    Edit_Job.CommandType = System.Data.CommandType.Text;
                    Edit_Job.CommandText = SQLUpdateCommand;
                    Edit_Job.Connection = Job_Connection;
                    Edit_Job.Parameters.AddWithValue("@RunOrder", ArrayRunOrder[x].ToString());
                    Edit_Job.Parameters.AddWithValue("@ItemID", ArrayItemID[x].ToString());
                    Job_Connection.Open();
                    Edit_Job.ExecuteNonQuery();
                    Job_Connection.Close();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            RefreshJobs();
        }



        /*********************************************************************************************************************
        * 
        * Methods Region End
        * 
        **********************************************************************************************************************/
        #endregion

        /********************************************************************************************************************
        * 
        * Events Region Start
        * 
        * - Clock Tick
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
        * Events Region End
        * 
        *********************************************************************************************************************/
        #endregion


    }
}
