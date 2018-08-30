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
    public partial class User_Program_Check_Card_Data_Enter : Form
    {
        public User_Program_Check_Card_Data_Enter()
        {
            this.ShowInTaskbar = false;
            InitializeComponent();
        }

        private string SQL_Source = @"Data Source=OHN7009,49172;Initial Catalog=Spot_Weld_Data;Integrated Security=True;Connect Timeout=15;Encrypt=False;TrustServerCertificate=True;";

        // Clock_Timer();
        private static int ClockHour;
        private static int ClockMinute;
        private static int ClockSecond;

        private static int ItemCheckCountTotal;
        private static int CheckCountTotal;
        private static string OperatorName = "";
        private bool OperatorNameFound = false;
        private bool BuddyCheckNameFound = false;
        private static string BuddyCheckName = "";
        private static string Date;
        private static string Time;

        string[] CodeNumbers = { "1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11" };
        string[] BoxAValues = { "Pass", "Fail" };

        private void User_Program_Check_Card_Data_Enter_Load(object sender, EventArgs e)
        {
            Code_ComboBox.Items.AddRange(CodeNumbers);

            SqlConnection EmployeeConnection = new SqlConnection(SQL_Source);
            string EmployeeString = "SELECT * FROM [dbo].[Employee] ORDER BY EmployeeName ASC";
            SqlDataAdapter EmployeeAdapter = new SqlDataAdapter(EmployeeString, EmployeeConnection);
            SqlCommandBuilder EmployeeCommandBuilder = new SqlCommandBuilder(EmployeeAdapter);
            DataSet EmployeeData = new DataSet();
            EmployeeAdapter.Fill(EmployeeData);
            Employee_DataGridView.DataSource = EmployeeData.Tables[0];
        }

        // Buttons Region
        #region

        private void Confirm_Button_Click(object sender, EventArgs e)
        {
            if (GageNumber_TextBox.Text == "")
            {
                MessageBox.Show("Please Enter the Gage #");
            }
            if (BaleNumber_TextBox.Text == "")
            {
                MessageBox.Show("Please Enter the Bale #");
            }
            if (LotNumber_TextBox.Text == "")
            {
                MessageBox.Show("Please Enter the Lot #");
            }
            if (OperatorID_TextBox.Text == "")
            {
                MessageBox.Show("Please Enter Your Operator ID #");
            }
            if (Code_ComboBox.Text == "")
            {
                MessageBox.Show("Please Select a Code #");
            }
            if (BuddyCheckID_TextBox.Text == "")
            {
                MessageBox.Show("Please Enter the Buddy Check ID #");
            }
            if (A_TextBox.Text != "Pass" && A_TextBox.Text != "Fail")
            {
                MessageBox.Show("Please Select a Value For Box A");
            }
            if (GageNumber_TextBox.Text != "" && BaleNumber_TextBox.Text != "" && LotNumber_TextBox.Text != "" && OperatorID_TextBox.Text != "" && BuddyCheckID_TextBox.Text != "" && A_TextBox.Text !=  "")
            {
                EmployeeIDSearch();
                BuddyCheckIDSearch();
                if (OperatorNameFound == true && BuddyCheckNameFound == true)
                {
                    string[] SplitString = DateTime_TextBox.Text.Split('M');
                    Date = SplitString[1];
                    Time = DateTime_TextBox.Text.Substring(0, 11);
                    CheckCard_ItemSearch();
                    CheckCard_TotalCount();
                    CheckCard_DataEntry();
                    CheckCard_Completed();
                    User_Program.UserProgram.Enabled = true;
                    this.Close();
                }
            }
        }

        private void Cancel_Button_Click(object sender, EventArgs e)
        {
            User_Program.UserProgram.Enabled = true;
            this.Close();
        }

        #endregion

        // Methods Region
        #region

        private void EmployeeIDSearch()
        {
            string SearchValue = OperatorID_TextBox.Text;
            Employee_DataGridView.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            try
            {
                foreach (DataGridViewRow Row in Employee_DataGridView.Rows)
                {
                    Row.Selected = false;
                    if (Row.Cells[0].Value.ToString().Equals(SearchValue))
                    {
                        Row.Selected = true;
                        OperatorName = Row.Cells[1].Value.ToString();
                        OperatorName_TextBox.Text = Row.Cells[1].Value.ToString();
                        OperatorName_Label.Visible = true;
                        OperatorName_TextBox.Visible = true;
                        OperatorNameFound = true;
                        break;
                    }
                }
            }
            catch (Exception e)
            {
                MessageBox.Show("Please Enter a Valid Operator ID #");
            }
        }

        private void BuddyCheckIDSearch()
        {
            string SearchValue = BuddyCheckID_TextBox.Text;
            Employee_DataGridView.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            try
            {
                foreach (DataGridViewRow Row in Employee_DataGridView.Rows)
                {
                    Row.Selected = false;
                    if (Row.Cells[0].Value.ToString().Equals(SearchValue))
                    {
                        Row.Selected = true;
                        BuddyCheckName = Row.Cells[1].Value.ToString();
                        BuddyCheckName_TextBox.Text = Row.Cells[1].Value.ToString();
                        BuddyCheckName_Label.Visible = true;
                        BuddyCheckName_TextBox.Visible = true;
                        BuddyCheckNameFound = true;
                        break;
                    }
                }
            }
            catch (Exception e)
            {
                MessageBox.Show("Please Enter a Valid Buddy Check ID #");
            }
        }

        private void CheckCard_Completed()
        {
            SqlConnection CheckCardCompletedConnection = new SqlConnection(SQL_Source);
            SqlCommand CheckCardComletedCommand = new SqlCommand();
            CheckCardComletedCommand.CommandType = System.Data.CommandType.Text;
            CheckCardComletedCommand.CommandText = "UPDATE [dbo].[ItemOperationData] SET CheckCardCompleted=@CheckCardCompleted WHERE OperationID=@OperationID";
            CheckCardComletedCommand.Connection = CheckCardCompletedConnection;
            CheckCardComletedCommand.Parameters.AddWithValue("@OperationID", OperationID_TextBox.Text);
            CheckCardComletedCommand.Parameters.AddWithValue("@CheckCardCompleted", "Yes");
            CheckCardCompletedConnection.Open();
            CheckCardComletedCommand.ExecuteNonQuery();
            CheckCardCompletedConnection.Close();
        }

        private void CheckCard_ItemSearch()
        {
            try
            {
                string ItemIDCheckCount = "SELECT COUNT(ItemID) FROM [dbo].[CheckCardData_SpotWeld] WHERE ItemID='" + ItemID_TextBox.Text + "' AND Sequence='" + Sequence_TextBox.Text + "'";
                SqlConnection CheckCountItemID_Connection = new SqlConnection(SQL_Source);
                SqlCommand CheckCountItemID_Command = new SqlCommand(ItemIDCheckCount, CheckCountItemID_Connection);
                CheckCountItemID_Connection.Open();
                int CheckCountItemID_Value = (int)CheckCountItemID_Command.ExecuteScalar();
                CheckCountItemID_Connection.Close();
                ItemCheckCountTotal = CheckCountItemID_Value + 1;
            }
            catch (Exception e)
            {
                MessageBox.Show(e.ToString());
            }
        }

        private void CheckCard_TotalCount()
        {
            try
            {
                string CheckCardCountString = "SELECT COUNT(*) FROM [dbo].[CheckCardData_SpotWeld]";
                SqlConnection CheckCountTotalConnection = new SqlConnection(SQL_Source);
                SqlCommand CheckCountTotalCommand = new SqlCommand(CheckCardCountString, CheckCountTotalConnection);
                CheckCountTotalConnection.Open();
                int CheckCardCountOperationTotal = (int)CheckCountTotalCommand.ExecuteScalar();
                CheckCountTotalConnection.Close();
                CheckCountTotal = CheckCardCountOperationTotal + 1;
            }
            catch (Exception e)
            {
                MessageBox.Show(e.ToString());
            }
        }

        private void CheckCard_DataEntry()
        {
            try
            {
                SqlConnection CheckCardConnection = new SqlConnection(SQL_Source);
                SqlCommand CheckCardCommand = new SqlCommand();
                CheckCardCommand.CommandType = System.Data.CommandType.Text;
                CheckCardCommand.CommandText = "INSERT INTO [dbo].[CheckCardData_SpotWeld] (Date,Time,ItemID,Sequence,Customer,CustomerPartID,GageNumber,BaleNumber,LotNumber,EmployeeName,EmployeeDMPID,CodeNumber,BuddyCheckName,BuddyCheckDMPID,Check_A,Check_B,Check_C,Check_D,Check_E,Check_F,Check_G,Check_H,Check_I,ItemCheckCount,CheckCountTotal) VALUES (@Date,@Time,@ItemID,@Sequence,@Customer,@CustomerPartID,@GageNumber,@BaleNumber,@LotNumber,@EmployeeName,@EmployeeDMPID,@CodeNumber,@BuddyCheckName,@BuddyCheckDMPID,@Check_A,@Check_B,@Check_C,@Check_D,@Check_E,@Check_F,@Check_G,@Check_H,@Check_I,@ItemCheckCount,@CheckCountTotal)";
                CheckCardCommand.Connection = CheckCardConnection;
                CheckCardCommand.Parameters.AddWithValue("@Date", Date);
                CheckCardCommand.Parameters.AddWithValue("@Time", Time);
                CheckCardCommand.Parameters.AddWithValue("@ItemID", ItemID_TextBox.Text);
                CheckCardCommand.Parameters.AddWithValue("@Sequence", Sequence_TextBox.Text);
                CheckCardCommand.Parameters.AddWithValue("@Customer", Customer_TextBox.Text);
                CheckCardCommand.Parameters.AddWithValue("@CustomerPartID", CustomerPartNumber_TextBox.Text);
                CheckCardCommand.Parameters.AddWithValue("@GageNumber", GageNumber_TextBox.Text);
                CheckCardCommand.Parameters.AddWithValue("@BaleNumber", BaleNumber_TextBox.Text);
                CheckCardCommand.Parameters.AddWithValue("@LotNumber", LotNumber_TextBox.Text);
                CheckCardCommand.Parameters.AddWithValue("@EmployeeName", OperatorName);
                CheckCardCommand.Parameters.AddWithValue("@EmployeeDMPID", OperatorID_TextBox.Text);
                CheckCardCommand.Parameters.AddWithValue("@CodeNumber", Code_ComboBox.Text);
                CheckCardCommand.Parameters.AddWithValue("@BuddyCheckName", BuddyCheckName);
                CheckCardCommand.Parameters.AddWithValue("@BuddyCheckDMPID", BuddyCheckID_TextBox.Text);
                CheckCardCommand.Parameters.AddWithValue("@Check_A", A_TextBox.Text);
                CheckCardCommand.Parameters.AddWithValue("@Check_B", B_TextBox.Text);
                CheckCardCommand.Parameters.AddWithValue("@Check_C", C_TextBox.Text);
                CheckCardCommand.Parameters.AddWithValue("@Check_D", D_TextBox.Text);
                CheckCardCommand.Parameters.AddWithValue("@Check_E", E_TextBox.Text);
                CheckCardCommand.Parameters.AddWithValue("@Check_F", F_TextBox.Text);
                CheckCardCommand.Parameters.AddWithValue("@Check_G", G_TextBox.Text);
                CheckCardCommand.Parameters.AddWithValue("@Check_H", H_TextBox.Text);
                CheckCardCommand.Parameters.AddWithValue("@Check_I", I_TextBox.Text);
                CheckCardCommand.Parameters.AddWithValue("@ItemCheckCount", ItemCheckCountTotal.ToString());
                CheckCardCommand.Parameters.AddWithValue("@CheckCountTotal", CheckCountTotal.ToString());
                CheckCardConnection.Open();
                CheckCardCommand.ExecuteNonQuery();
                CheckCardConnection.Close();
            }
            catch (Exception e)
            {
                MessageBox.Show(e.ToString());
            }
        }

        #endregion
    }
}
