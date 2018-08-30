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
    public partial class Admin_Control : Form
    {
        public Admin_Control()
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

        /********************************************************************************************************************
        * 
        * Form Load Variables 
        * 
        ********************************************************************************************************************/

        private string LoginTime = "";
        private string LoginForm = "Admin Access";
        private bool AddUser_ButtonWasClicked = false;
        private bool EditUser_ButtonWasClicked = false;
        private bool RemoveUser_ButtonWasClicked = false;
        private static string SQL_Source = @"Data Source=OHN7009,49172;Initial Catalog=Spot_Weld_Data;Integrated Security=True;Connect Timeout=15;";

        // Clock_Tick();
        private static int ClockHour;
        private static int ClockMinute;
        private static int ClockSecond;

        // EditUserData();
        private static string DMPIDNumber = "";

        // Search_Button_Click();
        private static int SearchColumn;
        private static string SearchValue;

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
        * Admin Control Start
        * 
        ********************************************************************************************************************/

        private void Admin_Control_Load(object sender, EventArgs e)
        {
            string SQL_Admin = "SELECT * FROM [dbo].[Employee]";
            SqlConnection AdminConnect = new SqlConnection(SQL_Source);
            SqlDataAdapter AdminAdapter = new SqlDataAdapter(SQL_Admin, AdminConnect);
            SqlCommandBuilder commandBuilder = new SqlCommandBuilder(AdminAdapter);
            DataSet AdminData = new DataSet();
            AdminAdapter.Fill(AdminData);
            AdminGridView.DataSource = AdminData.Tables[0];

            SqlConnection User_Login = new SqlConnection(SQL_Source);
            SqlCommand Login = new SqlCommand();
            Login.CommandType = System.Data.CommandType.Text;
            Login.CommandText = "INSERT INTO [dbo].[LoginData] (EmployeeName,DMPID,LoginDateTime,LoginForm) VALUES (@EmployeeName,@DMPID,@LoginDateTime,@LoginForm)";
            Login.Connection = User_Login;
            Login.Parameters.AddWithValue("@LoginDateTime", Clock_TextBox.Text);
            Login.Parameters.AddWithValue("@EmployeeName", User_TextBox.Text);
            Login.Parameters.AddWithValue("@DMPID", UserNumber_TextBox.Text);
            Login.Parameters.AddWithValue("@LoginForm", LoginForm.ToString());
            User_Login.Open();
            Login.ExecuteNonQuery();
            User_Login.Close();
            LoginTime = Clock_TextBox.Text;
            Clock.Enabled = true;
            SearchName_CheckBox.Checked = true;
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

        private void Clear_Button_Click(object sender, EventArgs e)
        {
            Clear();
        }

        private void Search_Button_Click(object sender, EventArgs e)
        {
            if (SearchName_CheckBox.Checked == true)
            {
                SearchDMPID_CheckBox.Checked = false;
                SearchValue = EmployeeName_TextBox.Text;
                SearchColumn = 1;
            }
            else if (SearchDMPID_CheckBox.Checked == true)
            {
                SearchName_CheckBox.Checked = false;
                SearchValue = DMPID_TextBox.Text;
                SearchColumn = 0;
            }
            AdminGridView.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            try
            {
                foreach (DataGridViewRow Row in AdminGridView.Rows)
                {
                    Row.Selected = false;
                    if (Row.Cells[SearchColumn].Value.ToString().Equals(SearchValue))
                    {
                        Row.Selected = true;
                        DMPID_TextBox.Text = Row.Cells[0].Value.ToString();
                        EmployeeName_TextBox.Text = Row.Cells[1].Value.ToString();
                        EmployeePassword_TextBox.Text = Row.Cells[2].Value.ToString();
                        AdminGridView.FirstDisplayedScrollingRowIndex = AdminGridView.SelectedRows[0].Index;
                        break;
                    }
                }

                SqlConnection connection = new SqlConnection(SQL_Source);
                string BP1176 = "SELECT * FROM [dbo].[LoginData] WHERE EmployeeName='" + EmployeeName_TextBox.Text + "'";
                SqlDataAdapter dataAdapter = new SqlDataAdapter(BP1176, connection);
                SqlCommandBuilder commandBuilder = new SqlCommandBuilder(dataAdapter);
                DataSet Data = new DataSet();
                dataAdapter.Fill(Data);
                LogInDataGridView.DataSource = Data.Tables[0];
            }
            catch (Exception)
            {
                if (SearchColumn == 1)
                {
                    MessageBox.Show("Unable to Find Employee: " + SearchValue + "\n" + "Please Try Again.");
                }
                else if (SearchColumn == 0)
                {
                    MessageBox.Show("Unable to Find User With DMP ID: " + SearchValue + "\n" + "Please Try Again.");
                }
            }
        }

        private void AddUser_Button_Click(object sender, EventArgs e)
        {
            Clear();
            AdminGridView.Enabled = false;
            Confirm_Button.Visible = true;
            Cancel_Button.Visible = true;
            AddUser_Button.Enabled = false;
            EditUser_Button.Enabled = false;
            RemoveUser_Button.Enabled = false;
            Search_Button.Enabled = false;
            AddUser_ButtonWasClicked = true;
            EmployeeName_TextBox.ReadOnly = false;
            EmployeePassword_TextBox.ReadOnly = false;
            DMPID_TextBox.ReadOnly = false;
            this.AddUser_Button.BackColor = Color.Silver;            
        }

        private void EditUser_Button_Click(object sender, EventArgs e)
        {
            if (EmployeeName_TextBox.Text == "" || EmployeePassword_TextBox.Text == "" || DMPID_TextBox.Text == "")
            {
                MessageBox.Show("Please Select a User to Edit");
            }
            else
            {
                AdminGridView.Enabled = false;
                Confirm_Button.Visible = true;
                Cancel_Button.Visible = true;
                AddUser_Button.Enabled = false;
                EditUser_Button.Enabled = false;
                RemoveUser_Button.Enabled = false;
                Search_Button.Enabled = false;
                EditUser_ButtonWasClicked = true;
                EmployeeName_TextBox.ReadOnly = false;
                EmployeePassword_TextBox.ReadOnly = false;
                DMPID_TextBox.ReadOnly = false;
                this.EditUser_Button.BackColor = Color.Silver;
            }
        }

        private void RemoveUser_Button_Click(object sender, EventArgs e)
        {
            if (EmployeeName_TextBox.Text == "" || EmployeePassword_TextBox.Text == "" || DMPID_TextBox.Text == "")
            {
                MessageBox.Show("Please Select a User to Remove");
            }
            else
            {
                AdminGridView.Enabled = false;
                Confirm_Button.Visible = true;
                Cancel_Button.Visible = true;
                AddUser_Button.Enabled = false;
                EditUser_Button.Enabled = false;
                RemoveUser_Button.Enabled = false;
                Search_Button.Enabled = false;
                RemoveUser_ButtonWasClicked = true;
                this.RemoveUser_Button.BackColor = Color.Silver;
            }
        }

        private void Confirm_Button_Click(object sender, EventArgs e)
        {
            if (AddUser_ButtonWasClicked == true)
            {
                AddUserConfirm();
            }
            else if (EditUser_ButtonWasClicked == true)
            {
                EditUserConfirm();
            }
            else if (RemoveUser_ButtonWasClicked == true)
            {
                RemoveUserConfirm();
            }
        }

        private void Cancel_Button_Click(object sender, EventArgs e)
        {
            Clear();
            Confirm_Button.Visible = false;
            Cancel_Button.Visible = false;
            AddUser_Button.Enabled = true;
            EditUser_Button.Enabled = true;
            RemoveUser_Button.Enabled = true;
            EmployeeName_TextBox.ReadOnly = false;
            EmployeePassword_TextBox.ReadOnly = false;
            DMPID_TextBox.ReadOnly = false;
            AddUser_ButtonWasClicked = false;
            EditUser_ButtonWasClicked = false;
            RemoveUser_ButtonWasClicked = false;
            AdminGridView.Enabled = true;
            Search_Button.Enabled = true;
            this.AddUser_Button.BackColor = SystemColors.Control;
            this.EditUser_Button.BackColor = SystemColors.Control;
            this.RemoveUser_Button.BackColor = SystemColors.Control;
        }

        private void SearchName_CheckBox_CheckedChanged(object sender, EventArgs e)
        {
            if (SearchName_CheckBox.Checked == true)
            {
                SearchDMPID_CheckBox.Checked = false;
            }
        }

        private void SearchDMPID_CheckBox_CheckedChanged(object sender, EventArgs e)
        {
            if (SearchDMPID_CheckBox.Checked == true)
            {
                SearchName_CheckBox.Checked = false;
            }
        }

        private void AdminGridView_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
            {
                DataGridViewRow Row = AdminGridView.Rows[e.RowIndex];
                DMPID_TextBox.Text = Row.Cells[0].Value.ToString();
                DMPIDNumber = Row.Cells[0].Value.ToString();
                EmployeeName_TextBox.Text = Row.Cells[1].Value.ToString();
                EmployeePassword_TextBox.Text = Row.Cells[2].Value.ToString();
            }

            EmployeeName_TextBox.ReadOnly = true;
            EmployeePassword_TextBox.ReadOnly = true;
            DMPID_TextBox.ReadOnly = true;
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

        /*********************************************************************************************************************
        * 
        * Buttons End
        * 
        *********************************************************************************************************************/
        /********************************************************************************************************************
        * [CheckBoxes]
        * 
        * ------------------------------------------------------[SearchName]-------------------------------------------------
        * 
        * Sets SearchName to True and SearchDMPID to False When Selected
        * 
        * ------------------------------------------------------[SearchDMPID]------------------------------------------------
        * 
        * Sets SearchDMPID to True and SearchName to False When Selected
        * 
        ********************************************************************************************************************/

        private void AddUserConfirm()
        {
            if (EmployeeName_TextBox.Text == "" && EmployeePassword_TextBox.Text == "")
            {
                MessageBox.Show("Please Enter Employee Name and Password");

            }
            else if (EmployeeName_TextBox.Text == "" && DMPID_TextBox.Text == "")
            {
                MessageBox.Show("Please Enter Employee Name and DMP ID Number");

            }
            else if (EmployeeName_TextBox.Text == "")
            {
                MessageBox.Show("Please Enter Employee Name");

            }
            else if (EmployeePassword_TextBox.Text == "" && DMPID_TextBox.Text == "")
            {
                MessageBox.Show("Please Enter Employee Password and DMP ID Number");

            }
            else if (EmployeePassword_TextBox.Text == "")
            {
                MessageBox.Show("Please Enter Employee Password");

            }
            else if (DMPID_TextBox.Text == "")
            {
                MessageBox.Show("Please Enter Employee DMP ID Number");

            }
            else if (EmployeeName_TextBox.Text != null && EmployeePassword_TextBox.Text != null && DMPID_TextBox.Text != null)
            {
                AddUserData();
            }
        }

        private void AddUserData()
        {
            try
            {
                SqlConnection AddUser = new SqlConnection(SQL_Source);
                SqlCommand SQL_AddUser = new SqlCommand();
                SQL_AddUser.CommandType = System.Data.CommandType.Text;
                SQL_AddUser.CommandText = "INSERT INTO [dbo].[Employee] (DMPID, EmployeeName, EmployeePassword) VALUES (@DMPID,@EmployeeName,@EmployeePassword)";
                SQL_AddUser.Connection = AddUser;
                SQL_AddUser.Parameters.AddWithValue("@DMPID", DMPID_TextBox.Text);
                SQL_AddUser.Parameters.AddWithValue("@EmployeeName", EmployeeName_TextBox.Text);
                SQL_AddUser.Parameters.AddWithValue("@EmployeePassword", EmployeePassword_TextBox.Text);
                AddUser.Open();
                SQL_AddUser.ExecuteNonQuery();
                AddUser.Close();
                DataEntryCompleted();

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

            RefreshUserData();
        }

        private void EditUserConfirm()
        {
            if (EmployeeName_TextBox.Text == "" && EmployeePassword_TextBox.Text == "")
            {
                MessageBox.Show("Employee Name and Password Cannot Be Empty");

            }
            else if (EmployeeName_TextBox.Text == "" && DMPID_TextBox.Text == "")
            {
                MessageBox.Show("Employee Name and DMP ID Number Cannot Be Empty");

            }
            else if (EmployeeName_TextBox.Text == "")
            {
                MessageBox.Show("Employee Name Cannot Be Empty");

            }
            else if (EmployeePassword_TextBox.Text == "" && DMPID_TextBox.Text == "")
            {
                MessageBox.Show("Employee Password and DMP ID Number Cannot Be Empty");

            }
            else if (EmployeePassword_TextBox.Text == "")
            {
                MessageBox.Show("Employee Password Cannot Be Empty");

            }
            else if (DMPID_TextBox.Text == "")
            {
                MessageBox.Show("Employee DMP ID Number Cannot Be Empty");

            }
            else if (EmployeeName_TextBox.Text != null && EmployeePassword_TextBox.Text != null && DMPID_TextBox.Text != null)
            {
                EditUserData();
            }
        }

        private void EditUserData()
        {
            try
            {
                SqlConnection EditUser = new SqlConnection(SQL_Source);
                SqlCommand SQL_EditUser = new SqlCommand();
                SQL_EditUser.CommandType = System.Data.CommandType.Text;
                SQL_EditUser.CommandText = "UPDATE [dbo].[Employee] SET EmployeeName=@EmployeeName, EmployeePassword=@EmployeePassword, DMPID=@DMPID WHERE DMPID='" + DMPIDNumber + "'";
                SQL_EditUser.Connection = EditUser;
                SQL_EditUser.Parameters.AddWithValue("@DMPID", DMPID_TextBox.Text);
                SQL_EditUser.Parameters.AddWithValue("@EmployeeName", EmployeeName_TextBox.Text);
                SQL_EditUser.Parameters.AddWithValue("@EmployeePassword", EmployeePassword_TextBox.Text);
                EditUser.Open();
                SQL_EditUser.ExecuteNonQuery();
                EditUser.Close();
                DataEntryCompleted();
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
                    MessageBox.Show("Unable to Edit User. Please Try Again." + "\n" + "Error Code: " + ErrorNumber.ToString());
                }
            }
            RefreshUserData();
        }

        private void RemoveUserConfirm()
        {
            RemoveUserData();
        }

        private void RemoveUserData()
        {
            try
            {
                SqlConnection RemoveUser = new SqlConnection(SQL_Source);
                SqlCommand SQL_RemoveUser = new SqlCommand();
                SQL_RemoveUser.CommandType = System.Data.CommandType.Text;
                SQL_RemoveUser.CommandText = "DELETE FROM [dbo].[Employee] WHERE DMPID=@DMPID";
                SQL_RemoveUser.Connection = RemoveUser;
                SQL_RemoveUser.Parameters.AddWithValue("@DMPID", DMPID_TextBox.Text);
                RemoveUser.Open();
                SQL_RemoveUser.ExecuteNonQuery();
                RemoveUser.Close();
                DataEntryCompleted();
            }
            catch (SqlException ExceptionValue)
            {
                int ErrorNumber = ExceptionValue.Number;
                MessageBox.Show("Unable to Remove User. Please Try Again." + "\n" + "Error Code: " + ErrorNumber.ToString());
            }
            RefreshUserData();
        }

        private void DataEntryCompleted()
        {
            Clear();
            Confirm_Button.Visible = false;
            Cancel_Button.Visible = false;
            AddUser_Button.Enabled = true;
            EditUser_Button.Enabled = true;
            RemoveUser_Button.Enabled = true;
            EmployeeName_TextBox.ReadOnly = false;
            EmployeePassword_TextBox.ReadOnly = false;
            DMPID_TextBox.ReadOnly = false;
            AddUser_ButtonWasClicked = false;
            EditUser_ButtonWasClicked = false;
            RemoveUser_ButtonWasClicked = false;
            AdminGridView.Enabled = true;
            Search_Button.Enabled = true;
            this.AddUser_Button.BackColor = SystemColors.Control;
            this.EditUser_Button.BackColor = SystemColors.Control;
            this.RemoveUser_Button.BackColor = SystemColors.Control;
        }

        private void RefreshUserData()
        {
            string SQL_Refresh = "SELECT * FROM [dbo].[Employee]";
            SqlConnection Refresh = new SqlConnection(SQL_Source);
            SqlDataAdapter dataAdapter = new SqlDataAdapter(SQL_Refresh, Refresh);
            SqlCommandBuilder commandBuilder = new SqlCommandBuilder(dataAdapter);
            DataSet Data = new DataSet();
            dataAdapter.Fill(Data);
            AdminGridView.DataSource = Data.Tables[0];
        }

        private void EmployeeLogOff()
        {
            SqlConnection User_Login = new SqlConnection(SQL_Source);
            SqlCommand Login = new SqlCommand();
            Login.CommandType = System.Data.CommandType.Text;
            Login.CommandText = "UPDATE [dbo].[LoginData] SET LogoutDateTime=@LogoutDateTime WHERE LoginDateTime=@LoginDateTime";
            Login.Connection = User_Login;
            Login.Parameters.AddWithValue("@LoginDateTime", LoginTime.ToString());
            Login.Parameters.AddWithValue("@LogoutDateTime", Clock_TextBox.Text);
            User_Login.Open();
            Login.ExecuteNonQuery();
            User_Login.Close();
        }

        private void Clear()
        {
            EmployeeName_TextBox.Clear();
            EmployeePassword_TextBox.Clear();
            DMPID_TextBox.Clear();
            EmployeeName_TextBox.ReadOnly = false;
            EmployeePassword_TextBox.ReadOnly = false;
            DMPID_TextBox.ReadOnly = false;
            
        }




        /********************************************************************************************************************
        * 
        * Methods End
        * 
        ********************************************************************************************************************/

        /********************************************************************************************************************
        * 
        * AdminAccess End
        * 
        ********************************************************************************************************************/

        /********************************************************************************************************************
        * 
        * Methods in Testing Start
        * 
        ********************************************************************************************************************/



        /********************************************************************************************************************
        * 
        * Methods in Testing End
        * 
        ********************************************************************************************************************/



    }
}
