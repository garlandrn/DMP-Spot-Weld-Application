using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using PdfSharp.Drawing;
using PdfSharp.Pdf;
using System.Data.SqlClient;
using System.IO;
using System.Diagnostics;

namespace DMP_Spot_Weld_Application
{
    public partial class Check_Card_Data_Viewer : Form
    {
        BackgroundWorker CreateExcel;
        BackgroundWorker PartImage;
        public Check_Card_Data_Viewer()
        {
            InitializeComponent();
            CreateExcel = new BackgroundWorker();
            CreateExcel.DoWork += new DoWorkEventHandler(CreateExcelFile);
            CreateExcel.RunWorkerCompleted += new RunWorkerCompletedEventHandler(CreateExcelFile_RunWorkerComplete);
            PartImage = new BackgroundWorker();
            PartImage.DoWork += new DoWorkEventHandler(FindItemImage);
            PartImage.RunWorkerCompleted += new RunWorkerCompletedEventHandler(PartImage_RunWorkerCompleted);
        }

        // Load
        private string LoginForm = "Check Card Data Viewer";
        private string LoginTime = "";

        // Clock_Tick();
        private static int ClockHour;
        private static int ClockMinute;
        private static int ClockSecond;

        // Search_Button_Click()
        private static string SearchCommand;
        private static string ReportItemID;
        private static string ReportCustomer;
        private static string ReportCustomerItemID;
        private static string ReportDate = "";

        // Search_Button_Click();
        private static int SearchColumn;

        // Excel File Creation
        private static Excel._Workbook ReportWB;
        private static Excel.Application ReportApp;
        private static Excel._Worksheet ReportWS;
        private static Excel.Range ReportRange;
        private static string ExcelFileLocation;
        private DataSet ReportDataSet;
        private static int RowCount;
        private static string RowCountString = "";

        // PDF FileLocation
        private static string PDFFileLocation;

        // FindItemImage();
        private static string ItemImagePath = "";
        private static string ItemID = "";
        private static string[] ItemIDSplit = ItemID.Split('-');
        private static double ItemID_Three;
        private static double ItemID_Five;

        string SQL_Source = @"Data Source=OHN7009,49172;Initial Catalog=Spot_Weld_Data;Integrated Security=True;Connect Timeout=15;";

        private void Check_Card_Data_Viewer_Load(object sender, EventArgs e)
        {
            if (User_TextBox.Text == "")
            {
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
                LoginTime = Clock_TextBox.Text;
            }
            Clock.Start();
            ItemIDSearch_TextBox.Focus();
            ItemID_CheckBox.Checked = true;
        }

        /********************************************************************************************************************
        * 
        * Buttons Region Start
        * 
        * - Search Button
        * - Clear Button
        * - CreatePDF Button
        * - CreateExcel Button
        * - LogOff Button
        * 
        ********************************************************************************************************************/
        #region

        private void Search_Button_Click(object sender, EventArgs e)
        {
            SearchCommand_SQL();
            SearchForItem();
        }

        private void Clear_Button_Click(object sender, EventArgs e)
        {
            Clear();
        }

        private void CreatePDF_Button_Click(object sender, EventArgs e)
        {
            if (SearchCommand == null)
            {
                MessageBox.Show("Please Create a DataTable Before Creating a PDF File");
            }
            else
            {
                CreatePDFFile();
            }
        }

        private void CreateExcel_Button_Click(object sender, EventArgs e)
        {
            if (SearchCommand == null)
            {
                MessageBox.Show("Please Create a DataTable Before Creating an Excel File");
            }
            else
            {
                CheckCard_TotalCount();
                string ReportName = DateTime.Now.ToShortDateString() + "_" + DateTime.Now.Millisecond;
                ReportName = ReportName.Replace("/", "_");
                ReportName = ReportName.Replace(":", "_");
                SaveFileDialog saveFile = new SaveFileDialog();
                saveFile.Filter = "Excel Files (*.xlsx)|*.xlsx|All files (*.*)|*.*";
                string ReportPDFName = "Check_Card_Report_" + ReportName + ".xls";
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
        * DataGridView Region Start
        * 
        ********************************************************************************************************************/
        #region


        /********************************************************************************************************************
        * 
        * DataGridView Region End
        * 
        ********************************************************************************************************************/
        #endregion

        /********************************************************************************************************************
        * 
        * Methods Region Start
        * 
        * - SearchCommand_SQL
        * - SearchForItem
        * - CheckCard_TotalCount
        * - CreateReport
        * - CreatePDFFile
        * - CreateExcelFile
        * - CreateExcelFile_RunWorkerComplete
        * - Clear
        * - EmployeeLogOff 
        * 
        ********************************************************************************************************************/
        #region

        /********************************************************************************************************************
        * 
        * Methods Region End
        * 
        ********************************************************************************************************************/
        #endregion

        private void SearchCommand_SQL()
        {
            if (ItemID_CheckBox.Checked == true && CustomerItemNumber_CheckBox.Checked == false && DateStartPicker.Checked == true)
            {
                SearchCommand = "SELECT Date, Time, GageNumber, BaleNumber, LotNumber, EmployeeDMPID, CodeNumber, BuddyCheckDMPID, Check_A, Check_B, Check_C, Check_D, Check_E, Check_F, Check_G, Check_H, Check_I, ItemID, Sequence, Customer, CustomerPartID, EmployeeName, BuddyCheckName FROM [dbo].[CheckCardData_SpotWeld] WHERE ItemID ='" + ItemIDSearch_TextBox.Text + "' AND Date BETWEEN '" + DateStartPicker.Value.Date.ToShortDateString() + "' AND '" + DateEndPicker.Value.Date.ToShortDateString() + "' ORDER BY CheckCountTotal ASC";
                //SearchCommand = "SELECT * FROM [dbo].[CheckCardData_BrakePress] WHERE ItemID ='" + ItemIDSearch_TextBox.Text + "' AND Date BETWEEN '" + DateStartPicker.Value.Date.ToShortDateString() + "' AND '" + DateEndPicker.Value.Date.ToShortDateString() + "' ORDER BY CheckCountTotal DESC";
                ReportItemID = "Item ID: " + ItemIDSearch_TextBox.Text;
                ReportCustomer = "Customer: " + Customer_TextBox.Text;
                ReportCustomerItemID = "Customer Part ID: " + CustomerItemNumber_TextBox.Text;
                ReportDate = "Date: " + DateStartPicker.Value.ToShortDateString() + " - " + DateEndPicker.Value.ToShortDateString();
                CreateReport();
            }
            if (ItemID_CheckBox.Checked == false && CustomerItemNumber_CheckBox.Checked == true && DateStartPicker.Checked == true)
            {
                SearchCommand = "SELECT Date, Time, GageNumber, BaleNumber, LotNumber, EmployeeID, CodeNumber, BuddyCheckDMPID, Check_A, Check_B, Check_C, Check_D, Check_E, Check_F, Check_G, Check_H, Check_I, ItemID, Sequence, Customer, CustomerPartID, EmployeeName, BuddyCheckName FROM [dbo].[CheckCardData_SpotWeld] WHERE CustomerPartID ='" + CustomerItemNumberSearch_TextBox.Text + "' AND Date BETWEEN '" + DateStartPicker.Value.Date.ToShortDateString() + "' AND '" + DateEndPicker.Value.Date.ToShortDateString() + "' ORDER BY CheckCountTotal ASC";
                //SearchCommand = "SELECT * FROM [dbo].[CheckCardData_BrakePress] WHERE CustomerItemNumber ='" + CustomerItemNumberSearch_TextBox.Text + "' AND Date BETWEEN '" + DateStartPicker.Value.Date.ToShortDateString() + "' AND '" + DateEndPicker.Value.Date.ToShortDateString() + "' ORDER BY CheckCountTotal DESC";
                ReportItemID = "Item ID: " + ItemIDSearch_TextBox.Text;
                ReportCustomer = "Customer: " + Customer_TextBox.Text;
                ReportCustomerItemID = "Customer Part ID: " + CustomerItemNumber_TextBox.Text;
                ReportDate = "Date: " + DateStartPicker.Value.ToShortDateString() + " - " + DateEndPicker.Value.ToShortDateString();
                CreateReport();
            }
            if (ItemID_CheckBox.Checked == true && CustomerItemNumber_CheckBox.Checked == false && DateStartPicker.Checked == false)
            {
                SearchCommand = "SELECT Date, Time, GageNumber, BaleNumber, LotNumber, EmployeeDMPID, CodeNumber, BuddyCheckDMPID, Check_A, Check_B, Check_C, Check_D, Check_E, Check_F, Check_G, Check_H, Check_I, ItemID, Sequence, Customer, CustomerPartID, EmployeeName, BuddyCheckName FROM [dbo].[CheckCardData_SpotWeld] WHERE ItemID ='" + ItemIDSearch_TextBox.Text + "'" + " ORDER BY CheckCountTotal ASC";
                //SearchCommand = "SELECT * FROM [dbo].[CheckCardData_BrakePress] WHERE ItemID ='" + ItemIDSearch_TextBox.Text + "'" + " ORDER BY CheckCountTotal DESC";
                ReportItemID = "Item ID: " + ItemIDSearch_TextBox.Text;
                ReportCustomer = "Customer: " + Customer_TextBox.Text;
                ReportCustomerItemID = "Customer Part ID: " + CustomerItemNumber_TextBox.Text;
                ReportDate = "Date: All";
                CreateReport();
            }
            if (ItemID_CheckBox.Checked == false && CustomerItemNumber_CheckBox.Checked == true && DateStartPicker.Checked == false)
            {
                SearchCommand = "SELECT Date, Time, GageNumber, BaleNumber, LotNumber, EmployeeDMPID, CodeNumber, BuddyCheckDMPID, Check_A, Check_B, Check_C, Check_D, Check_E, Check_F, Check_G, Check_H, Check_I, ItemID, Sequence, Customer, CustomerPartID, EmployeeName, BuddyCheckName FROM [dbo].[CheckCardData_SpotWeld] WHERE CustomerPartID ='" + CustomerItemNumberSearch_TextBox.Text + "'" + " ORDER BY CheckCountTotal ASC";
                //SearchCommand = "SELECT * FROM [dbo].[CheckCardData_BrakePress] WHERE CustomerItemNumber ='" + CustomerItemNumberSearch_TextBox.Text + "'" + " ORDER BY CheckCountTotal DESC";
                ReportItemID = "Item ID: " + ItemIDSearch_TextBox.Text;
                ReportCustomer = "Customer: " + Customer_TextBox.Text;
                ReportCustomerItemID = "Customer Part ID: " + CustomerItemNumber_TextBox.Text;
                ReportDate = "Date: All";
                CreateReport();
            }
            if (ItemIDSearch_TextBox.Text == "" && CustomerItemNumber_CheckBox.Checked == false && DateStartPicker.Checked == true)
            {
                SearchCommand = "SELECT Date, Time, GageNumber, BaleNumber, LotNumber, EmployeeDMPID, CodeNumber, BuddyCheckDMPID, Check_A, Check_B, Check_C, Check_D, Check_E, Check_F, Check_G, Check_H, Check_I, ItemID, Sequence, Customer, CustomerPartID, EmployeeName, BuddyCheckName FROM [dbo].[CheckCardData_SpotWeld] WHERE Date BETWEEN '" + DateStartPicker.Value.Date.ToShortDateString() + "' AND '" + DateEndPicker.Value.Date.ToShortDateString() + "' ORDER BY CheckCountTotal ASC";
                //SearchCommand = "SELECT * FROM [dbo].[CheckCardData_BrakePress] WHERE CustomerItemNumber ='" + CustomerItemNumberSearch_TextBox.Text + "'" + " ORDER BY CheckCountTotal DESC";
                ReportItemID = "Item ID: All";
                ReportCustomer = "Customer: All";
                ReportCustomerItemID = "Customer Part ID: All";
                ReportDate = "Date: All";
                CreateReport();
            }
        }

        private void SearchForItem()
        {
            string SearchValue = "";
            CheckCardData_GridView.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            if (ItemID_CheckBox.Checked == true)
            {
                CustomerItemNumber_CheckBox.Checked = false;
                SearchValue = ItemIDSearch_TextBox.Text;
                SearchColumn = 17;
            }
            else if (CustomerItemNumber_CheckBox.Checked == true)
            {
                ItemID_CheckBox.Checked = false;
                SearchValue = CustomerItemNumberSearch_TextBox.Text;
                SearchColumn = 20;
            }
            try
            {
                foreach (DataGridViewRow Row in CheckCardData_GridView.Rows)
                {
                    Row.Selected = false;
                    if (Row.Cells[SearchColumn].Value.ToString().Equals(SearchValue))
                    {
                        Row.Selected = true;
                        string Date = Row.Cells[0].Value.ToString();
                        Date = Date.Replace("12:00:00 AM", "");
                        Date_TextBox.Text = Date.Replace(" ", "");
                        Time_TextBox.Text = Row.Cells[1].Value.ToString();
                        GageNumber_TextBox.Text = Row.Cells[2].Value.ToString();
                        BaleNumber_TextBox.Text = Row.Cells[3].Value.ToString();
                        LotNumber_TextBox.Text = Row.Cells[4].Value.ToString();
                        OperatorID_TextBox.Text = Row.Cells[5].Value.ToString();
                        CodeNumber_TextBox.Text = Row.Cells[6].Value.ToString();
                        BuddyCheckID_TextBox.Text = Row.Cells[7].Value.ToString();
                        A_TextBox.Text = Row.Cells[8].Value.ToString();
                        B_TextBox.Text = Row.Cells[9].Value.ToString();
                        C_TextBox.Text = Row.Cells[10].Value.ToString();
                        D_TextBox.Text = Row.Cells[11].Value.ToString();
                        E_TextBox.Text = Row.Cells[12].Value.ToString();
                        F_TextBox.Text = Row.Cells[13].Value.ToString();
                        G_TextBox.Text = Row.Cells[14].Value.ToString();
                        H_TextBox.Text = Row.Cells[15].Value.ToString();
                        I_TextBox.Text = Row.Cells[16].Value.ToString();
                        ItemID_TextBox.Text = Row.Cells[17].Value.ToString();
                        Sequence_TextBox.Text = Row.Cells[18].Value.ToString();
                        Customer_TextBox.Text = Row.Cells[19].Value.ToString();
                        ReportCustomer = "Customer: " + Row.Cells[19].Value.ToString();
                        CustomerItemNumber_TextBox.Text = Row.Cells[20].Value.ToString();
                        ReportCustomerItemID = "Customer Part ID: " + Row.Cells[20].Value.ToString();
                        OperatorName_TextBox.Text = Row.Cells[21].Value.ToString();
                        BuddyCheckName_TextBox.Text = Row.Cells[22].Value.ToString();
                        break;
                    }
                    if (PartImage.IsBusy != true)
                    {
                        PartImage.RunWorkerAsync();
                    }
                }
            }
            catch (Exception ex)
            {
                //MessageBox.Show("The Following Item ID: " + SearchValue + " Was Not Found");
                //MessageBox.Show(ex.ToString());
            }
        }

        private void CheckCard_TotalCount()
        {
            if (ItemID_CheckBox.Checked == true && CustomerItemNumber_CheckBox.Checked == false && DateStartPicker.Checked == true)
            {
                RowCountString = "SELECT COUNT(*) FROM [dbo].[CheckCardData_SpotWeld] WHERE ItemID='" + ItemIDSearch_TextBox.Text + "' AND Date BETWEEN '" + DateStartPicker.Value.Date.ToShortDateString() + "' AND '" + DateEndPicker.Value.Date.ToShortDateString() + "'";
            }
            if (ItemID_CheckBox.Checked == false && CustomerItemNumber_CheckBox.Checked == true && DateStartPicker.Checked == true)
            {
                RowCountString = "SELECT COUNT(*) FROM [dbo].[CheckCardData_SpotWeld] WHERE CustomerPartID='" + CustomerItemNumberSearch_TextBox.Text + "' AND Date BETWEEN '" + DateStartPicker.Value.Date.ToShortDateString() + "' AND '" + DateEndPicker.Value.Date.ToShortDateString() + "'";
            }
            if (ItemID_CheckBox.Checked == true && CustomerItemNumber_CheckBox.Checked == false && DateStartPicker.Checked == false)
            {
                RowCountString = "SELECT COUNT(*) FROM [dbo].[CheckCardData_SpotWeld] WHERE ItemID='" + ItemIDSearch_TextBox.Text + "'";
            }
            if (ItemID_CheckBox.Checked == false && CustomerItemNumber_CheckBox.Checked == true && DateStartPicker.Checked == false)
            {
                RowCountString = "SELECT COUNT(*) FROM [dbo].[CheckCardData_SpotWeld] WHERE CustomerPartID='" + CustomerItemNumberSearch_TextBox.Text + "'";
            }
            try
            {
                string CheckCardCountString = RowCountString;
                SqlConnection CheckCountTotalConnection = new SqlConnection(SQL_Source);
                SqlCommand CheckCountTotalCommand = new SqlCommand(CheckCardCountString, CheckCountTotalConnection);
                CheckCountTotalConnection.Open();
                int CheckCardCountOperationTotal = (int)CheckCountTotalCommand.ExecuteScalar();
                CheckCountTotalConnection.Close();
                RowCount = CheckCardCountOperationTotal + 5;
            }
            catch (Exception e)
            {
                MessageBox.Show(e.ToString());
            }
        }

        private void CreateReport()
        {
            try
            {
                SqlConnection ReportConnect = new SqlConnection(SQL_Source);
                string SearchString = SearchCommand;
                SqlDataAdapter ReportDataAdapter = new SqlDataAdapter(SearchString, ReportConnect);
                SqlCommandBuilder ReportCommandBuilder = new SqlCommandBuilder(ReportDataAdapter);
                ReportDataSet = new DataSet();
                ReportDataAdapter.Fill(ReportDataSet);
                CheckCardData_GridView.DataSource = ReportDataSet.Tables[0];
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

        private void CreatePDFFile()
        {
            // New PDF Document
            PdfDocument CheckCardReport = new PdfDocument();
            PdfPage ReportPage = CheckCardReport.AddPage();
            ReportPage.Size = PdfSharp.PageSize.Letter;
            ReportPage.Orientation = PdfSharp.PageOrientation.Landscape;
            ReportPage.Rotate = 0;
            XGraphics ReportGraph = XGraphics.FromPdfPage(ReportPage);

            // Fonts
            XFont ReportDataHeader = new XFont("Verdana", 12, XFontStyle.Bold);
            XFont ColumnHeader = new XFont("Verdana", 7, XFontStyle.Bold | XFontStyle.Underline);
            XFont ColumnDivider = new XFont("Verdana", 25, XFontStyle.Regular);
            XFont RowFont = new XFont("Verdana", 6, XFontStyle.Regular);
            XFont PageFooterFont = new XFont("Verdana", 6, XFontStyle.Regular);

            int PointY = 0;
            int CurrentRow = 0;

            // PDF Report Name

            string ReportName = DateTime.Now.ToShortDateString() + "_" + DateTime.Now.Millisecond;
            ReportName = ReportName.Replace("/", "_");
            ReportName = ReportName.Replace(":", "_");
            string ReportFooter = " | Report Created On: " + DateTime.Now.ToShortDateString() + " | Created By: " + User_TextBox.Text;

            // PDF Header First Page Only



            ReportGraph.DrawImage(XImage.FromFile(@"\\OHN66FS01\BPprogs\Brake Press Vision\Applications\DMPLogo700.jpg"), 35, 5);
            ReportGraph.DrawString(ReportItemID, ReportDataHeader, XBrushes.Black, new XRect(400, 15, ReportPage.Width.Point, ReportPage.Height.Point), XStringFormats.TopLeft);
            ReportGraph.DrawString(ReportCustomer, ReportDataHeader, XBrushes.Black, new XRect(400, 33, ReportPage.Width.Point, ReportPage.Height.Point), XStringFormats.TopLeft);
            ReportGraph.DrawString(ReportCustomerItemID, ReportDataHeader, XBrushes.Black, new XRect(400, 51, ReportPage.Width.Point, ReportPage.Height.Point), XStringFormats.TopLeft);
            ReportGraph.DrawString(ReportDate, ReportDataHeader, XBrushes.Black, new XRect(400, 69, ReportPage.Width.Point, ReportPage.Height.Point), XStringFormats.TopLeft);
            PointY = PointY + 100;

            // Column Headers

            ReportGraph.DrawString("Date Time", ColumnHeader, XBrushes.Black, new XRect(30, PointY, ReportPage.Width.Point, ReportPage.Height.Point), XStringFormats.TopLeft);
            ReportGraph.DrawString("Gage (S)#", ColumnHeader, XBrushes.Black, new XRect(103, PointY, ReportPage.Width.Point, ReportPage.Height.Point), XStringFormats.TopLeft);
            ReportGraph.DrawString("Bale (S)#", ColumnHeader, XBrushes.Black, new XRect(153, PointY, ReportPage.Width.Point, ReportPage.Height.Point), XStringFormats.TopLeft);
            ReportGraph.DrawString("Lot (S)#", ColumnHeader, XBrushes.Black, new XRect(203, PointY, ReportPage.Width.Point, ReportPage.Height.Point), XStringFormats.TopLeft);
            ReportGraph.DrawString("Employee ID", ColumnHeader, XBrushes.Black, new XRect(248, PointY, ReportPage.Width.Point, ReportPage.Height.Point), XStringFormats.TopLeft);
            ReportGraph.DrawString("Buddy Check ID", ColumnHeader, XBrushes.Black, new XRect(308, PointY, ReportPage.Width.Point, ReportPage.Height.Point), XStringFormats.TopLeft);
            ReportGraph.DrawString("Code #", ColumnHeader, XBrushes.Black, new XRect(378, PointY, ReportPage.Width.Point, ReportPage.Height.Point), XStringFormats.TopLeft);
            ReportGraph.DrawString("A", ColumnHeader, XBrushes.Black, new XRect(430, PointY, ReportPage.Width.Point, ReportPage.Height.Point), XStringFormats.TopLeft);
            ReportGraph.DrawString("B", ColumnHeader, XBrushes.Black, new XRect(470, PointY, ReportPage.Width.Point, ReportPage.Height.Point), XStringFormats.TopLeft);
            ReportGraph.DrawString("C", ColumnHeader, XBrushes.Black, new XRect(510, PointY, ReportPage.Width.Point, ReportPage.Height.Point), XStringFormats.TopLeft);
            ReportGraph.DrawString("D", ColumnHeader, XBrushes.Black, new XRect(550, PointY, ReportPage.Width.Point, ReportPage.Height.Point), XStringFormats.TopLeft);
            ReportGraph.DrawString("E", ColumnHeader, XBrushes.Black, new XRect(590, PointY, ReportPage.Width.Point, ReportPage.Height.Point), XStringFormats.TopLeft);
            ReportGraph.DrawString("F", ColumnHeader, XBrushes.Black, new XRect(630, PointY, ReportPage.Width.Point, ReportPage.Height.Point), XStringFormats.TopLeft);
            ReportGraph.DrawString("G", ColumnHeader, XBrushes.Black, new XRect(670, PointY, ReportPage.Width.Point, ReportPage.Height.Point), XStringFormats.TopLeft);
            ReportGraph.DrawString("H", ColumnHeader, XBrushes.Black, new XRect(710, PointY, ReportPage.Width.Point, ReportPage.Height.Point), XStringFormats.TopLeft);
            ReportGraph.DrawString("I", ColumnHeader, XBrushes.Black, new XRect(750, PointY, ReportPage.Width.Point, ReportPage.Height.Point), XStringFormats.TopLeft);

            // Report Footer 
            string PageNumber = "Page: " + CheckCardReport.PageCount + ReportFooter;
            string CodelistInfo = "Code list    All codes besides 7 and 11 are to be used when a period of more than 1 hour has transpired between checks";
            string CodeList_1 = "1   Dimensional problem with part";
            string CodeList_2 = "2   Machine downtime";
            string CodeList_3 = "3   Operator scheduled break / meeting";
            string CodeList_4 = "4   Sort / Rework";
            string CodeList_5 = "5   Operator changed production job";
            string CodeList_6 = "6   Shift change";
            string CodeList_7 = "7   Initial Green Tag/Setup (Verify if Buddy Check is required)";
            string CodeList_8 = "8   Material defects";
            string CodeList_9 = "9   Operator waiting on material/hardware/components";
            string CodeList_10 = "10   Inspection to verify tool room work (capabilities etc.)";
            string CodeList_11 = "11   Last Piece Inspection";
            ReportGraph.DrawString(PageNumber, PageFooterFont, XBrushes.Black, new XRect(5, 585, ReportPage.Width.Point, ReportPage.Height.Point), XStringFormats.TopLeft);
            ReportGraph.DrawString(CodelistInfo, PageFooterFont, XBrushes.Black, new XRect(400, 540, ReportPage.Width.Point, ReportPage.Height.Point), XStringFormats.TopLeft);
            ReportGraph.DrawString(CodeList_1, PageFooterFont, XBrushes.Black, new XRect(440, 550, ReportPage.Width.Point, ReportPage.Height.Point), XStringFormats.TopLeft);
            ReportGraph.DrawString(CodeList_2, PageFooterFont, XBrushes.Black, new XRect(440, 560, ReportPage.Width.Point, ReportPage.Height.Point), XStringFormats.TopLeft);
            ReportGraph.DrawString(CodeList_3, PageFooterFont, XBrushes.Black, new XRect(440, 570, ReportPage.Width.Point, ReportPage.Height.Point), XStringFormats.TopLeft);
            ReportGraph.DrawString(CodeList_4, PageFooterFont, XBrushes.Black, new XRect(440, 580, ReportPage.Width.Point, ReportPage.Height.Point), XStringFormats.TopLeft);
            ReportGraph.DrawString(CodeList_5, PageFooterFont, XBrushes.Black, new XRect(440, 590, ReportPage.Width.Point, ReportPage.Height.Point), XStringFormats.TopLeft);
            ReportGraph.DrawString(CodeList_6, PageFooterFont, XBrushes.Black, new XRect(440, 600, ReportPage.Width.Point, ReportPage.Height.Point), XStringFormats.TopLeft);
            ReportGraph.DrawString(CodeList_7, PageFooterFont, XBrushes.Black, new XRect(570, 550, ReportPage.Width.Point, ReportPage.Height.Point), XStringFormats.TopLeft);
            ReportGraph.DrawString(CodeList_8, PageFooterFont, XBrushes.Black, new XRect(570, 560, ReportPage.Width.Point, ReportPage.Height.Point), XStringFormats.TopLeft);
            ReportGraph.DrawString(CodeList_9, PageFooterFont, XBrushes.Black, new XRect(570, 570, ReportPage.Width.Point, ReportPage.Height.Point), XStringFormats.TopLeft);
            ReportGraph.DrawString(CodeList_10, PageFooterFont, XBrushes.Black, new XRect(570, 580, ReportPage.Width.Point, ReportPage.Height.Point), XStringFormats.TopLeft);
            ReportGraph.DrawString(CodeList_11, PageFooterFont, XBrushes.Black, new XRect(570, 590, ReportPage.Width.Point, ReportPage.Height.Point), XStringFormats.TopLeft);

            PointY = PointY + 20;
            try
            {
                SqlConnection CreatePDF = new SqlConnection(SQL_Source);
                string PDFCommand = SearchCommand;
                SqlDataAdapter PDFAdapter = new SqlDataAdapter(PDFCommand, CreatePDF);
                DataSet PDFData = new DataSet();
                PDFAdapter.Fill(PDFData);

                for (int i = 0; i <= PDFData.Tables[0].Rows.Count - 1; i++)
                {
                    string DateResults = PDFData.Tables[0].Rows[i].ItemArray[0].ToString();
                    string TimeResults = PDFData.Tables[0].Rows[i].ItemArray[1].ToString();
                    string GageNumberResults = PDFData.Tables[0].Rows[i].ItemArray[2].ToString();
                    string BaleNumberResults = PDFData.Tables[0].Rows[i].ItemArray[3].ToString();
                    string LotNumberResults = PDFData.Tables[0].Rows[i].ItemArray[4].ToString();
                    string EmployeeIDResults = PDFData.Tables[0].Rows[i].ItemArray[5].ToString();
                    string CodeNumberResults = PDFData.Tables[0].Rows[i].ItemArray[6].ToString();
                    string BuddyCheckIDResults = PDFData.Tables[0].Rows[i].ItemArray[7].ToString();
                    string A_Results = PDFData.Tables[0].Rows[i].ItemArray[8].ToString();
                    string B_Results = PDFData.Tables[0].Rows[i].ItemArray[9].ToString();
                    string C_Results = PDFData.Tables[0].Rows[i].ItemArray[10].ToString();
                    string D_Results = PDFData.Tables[0].Rows[i].ItemArray[11].ToString();
                    string E_Results = PDFData.Tables[0].Rows[i].ItemArray[12].ToString();
                    string F_Results = PDFData.Tables[0].Rows[i].ItemArray[13].ToString();
                    string G_Results = PDFData.Tables[0].Rows[i].ItemArray[14].ToString();
                    string H_Results = PDFData.Tables[0].Rows[i].ItemArray[15].ToString();
                    string I_Results = PDFData.Tables[0].Rows[i].ItemArray[16].ToString();

                    DateResults = DateResults.Replace("12:00:00 AM", "");

                    // Report Row Data

                    ReportGraph.DrawString(DateResults, RowFont, XBrushes.Black, new XRect(15, PointY, ReportPage.Width.Point, ReportPage.Height.Point), XStringFormats.TopLeft);
                    ReportGraph.DrawString(TimeResults, RowFont, XBrushes.Black, new XRect(50, PointY, ReportPage.Width.Point, ReportPage.Height.Point), XStringFormats.TopLeft);
                    //ReportGraph.DrawString("|", ColumnDivider, XBrushes.Black, new XRect(90, PointY, ReportPage.Width.Point, ReportPage.Height.Point), XStringFormats.TopLeft);
                    ReportGraph.DrawString(GageNumberResults, RowFont, XBrushes.Black, new XRect(110, PointY, ReportPage.Width.Point, ReportPage.Height.Point), XStringFormats.TopLeft);
                    ReportGraph.DrawString(BaleNumberResults, RowFont, XBrushes.Black, new XRect(160, PointY, ReportPage.Width.Point, ReportPage.Height.Point), XStringFormats.TopLeft);
                    ReportGraph.DrawString(LotNumberResults, RowFont, XBrushes.Black, new XRect(208, PointY, ReportPage.Width.Point, ReportPage.Height.Point), XStringFormats.TopLeft);
                    ReportGraph.DrawString(EmployeeIDResults, RowFont, XBrushes.Black, new XRect(260, PointY, ReportPage.Width.Point, ReportPage.Height.Point), XStringFormats.TopLeft);
                    ReportGraph.DrawString(BuddyCheckIDResults, RowFont, XBrushes.Black, new XRect(325, PointY, ReportPage.Width.Point, ReportPage.Height.Point), XStringFormats.TopLeft);
                    ReportGraph.DrawString(CodeNumberResults, RowFont, XBrushes.Black, new XRect(390, PointY, ReportPage.Width.Point, ReportPage.Height.Point), XStringFormats.TopLeft);
                    ReportGraph.DrawString(A_Results, RowFont, XBrushes.Black, new XRect(428, PointY, ReportPage.Width.Point, ReportPage.Height.Point), XStringFormats.TopLeft);
                    ReportGraph.DrawString(B_Results, RowFont, XBrushes.Black, new XRect(464, PointY, ReportPage.Width.Point, ReportPage.Height.Point), XStringFormats.TopLeft);
                    ReportGraph.DrawString(C_Results, RowFont, XBrushes.Black, new XRect(504, PointY, ReportPage.Width.Point, ReportPage.Height.Point), XStringFormats.TopLeft);
                    ReportGraph.DrawString(D_Results, RowFont, XBrushes.Black, new XRect(544, PointY, ReportPage.Width.Point, ReportPage.Height.Point), XStringFormats.TopLeft);
                    ReportGraph.DrawString(E_Results, RowFont, XBrushes.Black, new XRect(584, PointY, ReportPage.Width.Point, ReportPage.Height.Point), XStringFormats.TopLeft);
                    ReportGraph.DrawString(F_Results, RowFont, XBrushes.Black, new XRect(624, PointY, ReportPage.Width.Point, ReportPage.Height.Point), XStringFormats.TopLeft);
                    ReportGraph.DrawString(G_Results, RowFont, XBrushes.Black, new XRect(664, PointY, ReportPage.Width.Point, ReportPage.Height.Point), XStringFormats.TopLeft);
                    ReportGraph.DrawString(H_Results, RowFont, XBrushes.Black, new XRect(704, PointY, ReportPage.Width.Point, ReportPage.Height.Point), XStringFormats.TopLeft);
                    ReportGraph.DrawString(I_Results, RowFont, XBrushes.Black, new XRect(744, PointY, ReportPage.Width.Point, ReportPage.Height.Point), XStringFormats.TopLeft);
                    PointY = PointY + 20;
                    CurrentRow = CurrentRow + 1;

                    // Report Creates Adds Another Page If Data is Larger than 22 Rows
                    // Only 22 Entries on Page One Due To Report Header
                    if (CurrentRow == 21 && CheckCardReport.PageCount == 1)
                    {
                        PointY = 0;
                        ReportPage = CheckCardReport.AddPage();
                        ReportGraph = XGraphics.FromPdfPage(ReportPage);
                        ReportPage.Size = PdfSharp.PageSize.Letter;
                        ReportPage.Orientation = PdfSharp.PageOrientation.Landscape;
                        ReportPage.Rotate = 0;
                        PointY = PointY + 50;

                        // Column Headers For Second Page

                        ReportGraph.DrawString("Date Time", ColumnHeader, XBrushes.Black, new XRect(30, PointY, ReportPage.Width.Point, ReportPage.Height.Point), XStringFormats.TopLeft);
                        ReportGraph.DrawString("Gage (S)#", ColumnHeader, XBrushes.Black, new XRect(103, PointY, ReportPage.Width.Point, ReportPage.Height.Point), XStringFormats.TopLeft);
                        ReportGraph.DrawString("Bale (S)#", ColumnHeader, XBrushes.Black, new XRect(153, PointY, ReportPage.Width.Point, ReportPage.Height.Point), XStringFormats.TopLeft);
                        ReportGraph.DrawString("Lot (S)#", ColumnHeader, XBrushes.Black, new XRect(203, PointY, ReportPage.Width.Point, ReportPage.Height.Point), XStringFormats.TopLeft);
                        ReportGraph.DrawString("Employee ID", ColumnHeader, XBrushes.Black, new XRect(248, PointY, ReportPage.Width.Point, ReportPage.Height.Point), XStringFormats.TopLeft);
                        ReportGraph.DrawString("Buddy Check ID", ColumnHeader, XBrushes.Black, new XRect(308, PointY, ReportPage.Width.Point, ReportPage.Height.Point), XStringFormats.TopLeft);
                        ReportGraph.DrawString("Code #", ColumnHeader, XBrushes.Black, new XRect(378, PointY, ReportPage.Width.Point, ReportPage.Height.Point), XStringFormats.TopLeft);
                        ReportGraph.DrawString("A", ColumnHeader, XBrushes.Black, new XRect(430, PointY, ReportPage.Width.Point, ReportPage.Height.Point), XStringFormats.TopLeft);
                        ReportGraph.DrawString("B", ColumnHeader, XBrushes.Black, new XRect(470, PointY, ReportPage.Width.Point, ReportPage.Height.Point), XStringFormats.TopLeft);
                        ReportGraph.DrawString("C", ColumnHeader, XBrushes.Black, new XRect(510, PointY, ReportPage.Width.Point, ReportPage.Height.Point), XStringFormats.TopLeft);
                        ReportGraph.DrawString("D", ColumnHeader, XBrushes.Black, new XRect(550, PointY, ReportPage.Width.Point, ReportPage.Height.Point), XStringFormats.TopLeft);
                        ReportGraph.DrawString("E", ColumnHeader, XBrushes.Black, new XRect(590, PointY, ReportPage.Width.Point, ReportPage.Height.Point), XStringFormats.TopLeft);
                        ReportGraph.DrawString("F", ColumnHeader, XBrushes.Black, new XRect(630, PointY, ReportPage.Width.Point, ReportPage.Height.Point), XStringFormats.TopLeft);
                        ReportGraph.DrawString("G", ColumnHeader, XBrushes.Black, new XRect(670, PointY, ReportPage.Width.Point, ReportPage.Height.Point), XStringFormats.TopLeft);
                        ReportGraph.DrawString("H", ColumnHeader, XBrushes.Black, new XRect(710, PointY, ReportPage.Width.Point, ReportPage.Height.Point), XStringFormats.TopLeft);
                        ReportGraph.DrawString("I", ColumnHeader, XBrushes.Black, new XRect(750, PointY, ReportPage.Width.Point, ReportPage.Height.Point), XStringFormats.TopLeft);
                        PageNumber = "Page: " + CheckCardReport.PageCount + ReportFooter;
                        ReportGraph.DrawString(PageNumber, PageFooterFont, XBrushes.Black, new XRect(5, 585, ReportPage.Width.Point, ReportPage.Height.Point), XStringFormats.TopLeft); PointY = PointY + 25;
                        CurrentRow = 0;
                    }
                    else if (CurrentRow == 25 && CheckCardReport.PageCount >= 2)
                    {
                        PointY = 0;
                        ReportPage = CheckCardReport.AddPage();
                        ReportGraph = XGraphics.FromPdfPage(ReportPage);
                        ReportPage.Size = PdfSharp.PageSize.Letter;
                        ReportPage.Orientation = PdfSharp.PageOrientation.Landscape;
                        ReportPage.Rotate = 0;
                        PointY = PointY + 50;

                        // Column Headers For Any Page After Two

                        ReportGraph.DrawString("Date Time", ColumnHeader, XBrushes.Black, new XRect(30, PointY, ReportPage.Width.Point, ReportPage.Height.Point), XStringFormats.TopLeft);
                        ReportGraph.DrawString("Gage (S)#", ColumnHeader, XBrushes.Black, new XRect(103, PointY, ReportPage.Width.Point, ReportPage.Height.Point), XStringFormats.TopLeft);
                        ReportGraph.DrawString("Bale (S)#", ColumnHeader, XBrushes.Black, new XRect(153, PointY, ReportPage.Width.Point, ReportPage.Height.Point), XStringFormats.TopLeft);
                        ReportGraph.DrawString("Lot (S)#", ColumnHeader, XBrushes.Black, new XRect(203, PointY, ReportPage.Width.Point, ReportPage.Height.Point), XStringFormats.TopLeft);
                        ReportGraph.DrawString("Employee ID", ColumnHeader, XBrushes.Black, new XRect(248, PointY, ReportPage.Width.Point, ReportPage.Height.Point), XStringFormats.TopLeft);
                        ReportGraph.DrawString("Buddy Check ID", ColumnHeader, XBrushes.Black, new XRect(308, PointY, ReportPage.Width.Point, ReportPage.Height.Point), XStringFormats.TopLeft);
                        ReportGraph.DrawString("Code #", ColumnHeader, XBrushes.Black, new XRect(378, PointY, ReportPage.Width.Point, ReportPage.Height.Point), XStringFormats.TopLeft);
                        ReportGraph.DrawString("A", ColumnHeader, XBrushes.Black, new XRect(430, PointY, ReportPage.Width.Point, ReportPage.Height.Point), XStringFormats.TopLeft);
                        ReportGraph.DrawString("B", ColumnHeader, XBrushes.Black, new XRect(470, PointY, ReportPage.Width.Point, ReportPage.Height.Point), XStringFormats.TopLeft);
                        ReportGraph.DrawString("C", ColumnHeader, XBrushes.Black, new XRect(510, PointY, ReportPage.Width.Point, ReportPage.Height.Point), XStringFormats.TopLeft);
                        ReportGraph.DrawString("D", ColumnHeader, XBrushes.Black, new XRect(550, PointY, ReportPage.Width.Point, ReportPage.Height.Point), XStringFormats.TopLeft);
                        ReportGraph.DrawString("E", ColumnHeader, XBrushes.Black, new XRect(590, PointY, ReportPage.Width.Point, ReportPage.Height.Point), XStringFormats.TopLeft);
                        ReportGraph.DrawString("F", ColumnHeader, XBrushes.Black, new XRect(630, PointY, ReportPage.Width.Point, ReportPage.Height.Point), XStringFormats.TopLeft);
                        ReportGraph.DrawString("G", ColumnHeader, XBrushes.Black, new XRect(670, PointY, ReportPage.Width.Point, ReportPage.Height.Point), XStringFormats.TopLeft);
                        ReportGraph.DrawString("H", ColumnHeader, XBrushes.Black, new XRect(710, PointY, ReportPage.Width.Point, ReportPage.Height.Point), XStringFormats.TopLeft);
                        ReportGraph.DrawString("I", ColumnHeader, XBrushes.Black, new XRect(750, PointY, ReportPage.Width.Point, ReportPage.Height.Point), XStringFormats.TopLeft);
                        PageNumber = "Page: " + CheckCardReport.PageCount + ReportFooter;
                        ReportGraph.DrawString(PageNumber, PageFooterFont, XBrushes.Black, new XRect(5, 585, ReportPage.Width.Point, ReportPage.Height.Point), XStringFormats.TopLeft); PointY = PointY + 25;
                        CurrentRow = 0;
                    }
                }

                //string ReportPDFName = "Brake_Press_Report_" + ReportName + ".pdf";
                //BrakePressReport.Save(ReportPDFName);
                //BrakePressReport.Save(@"C:\Users\rgarland\Desktop\"+ReportPDFName);
                //Process.Start(ReportPDFName);

                string ReportPDFName = "Check_Card_Report_" + ReportName + ".pdf";
                SaveFileDialog saveFile = new SaveFileDialog();
                saveFile.Filter = "PDF Files (*.pdf)|*.pdf|All files (*.*)|*.*";
                saveFile.FileName = ReportPDFName;
                if (saveFile.ShowDialog() == DialogResult.OK)
                {
                    PDFFileLocation = saveFile.FileName;
                    CheckCardReport.Save(PDFFileLocation);
                    //BrakePressReport.Save(ReportPDFName);
                    //BrakePressReport.Save(@"C:\Users\rgarland\Desktop\"+ReportPDFName);
                    Process.Start(PDFFileLocation);
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void CreateExcelFile(object sender, EventArgs e)
        {
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
            ReportWS.Shapes.AddPicture(@"\\OHN66FS01\BPprogs\Brake Press Vision\Applications\DMPLogo700.jpg", Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, 0, 0, 300, 75);
            //ReportWS.Range["F1", "H1"].Merge();
            string Name = User_TextBox.Text;
            ReportWS.Cells[1, 8] = ReportItemID;
            ReportWS.Cells[2, 8] = ReportCustomer;
            ReportWS.Cells[3, 8] = ReportCustomerItemID;
            ReportWS.Cells[4, 8] = ReportDate;
            ReportWS.get_Range("H1", "P4").Font.Bold = true;
            ReportWS.get_Range("H1", "P4").Font.Size = 14;
            ReportRange.get_Range("H1", "P1").Merge();
            ReportRange.get_Range("H2", "P2").Merge();
            ReportRange.get_Range("H3", "P3").Merge();
            ReportRange.get_Range("H4", "P4").Merge();
            ReportRange.EntireColumn.AutoFit();

            string[] ColumnNames = new string[CheckCardData_GridView.Columns.Count];
            int ExcelColumns = 1;
            /*
            foreach (DataGridViewColumn dc in CheckCardData_GridView.Columns)
            {
                ReportWS.Cells[5, ExcelColumns] = dc.Name;
                ExcelColumns++;
            }
            */

            //ReportWS.get_Range("A5", "Y5").Font.Bold = true;
            //ReportWS.get_Range("A5", "Y5").AutoFilter(1, Type.Missing, Excel.XlAutoFilterOperator.xlAnd, Type.Missing, true);
            //ReportRange = ReportWS.get_Range("A5", "Y5");
            ReportWS.get_Range("A5", "Q5").Font.Bold = true;
            ReportWS.get_Range("A5", "Q5").AutoFilter(1, Type.Missing, Excel.XlAutoFilterOperator.xlAnd, Type.Missing, true);
            ReportRange = ReportWS.get_Range("A5", "Q5");
            //ReportRange.EntireColumn.AutoFit();

            for (int i = 0; i < ReportDataSet.Tables[0].Rows.Count; i++)
            {
                // to do: format datetime values before printing
                for (int j = 0; j < 17; j++)
                {
                    ReportWS.Cells[(i + 6), (j + 1)] = ReportDataSet.Tables[0].Rows[i][j];
                    //ReportWS.Cells.BorderAround2();
                }
            }

            ReportRange = ReportWS.get_Range("A5", "Q" + RowCount.ToString());
            ReportWS.get_Range("A5", "Q" + RowCount.ToString()).Font.Size = 8;
            ReportWS.get_Range("A5", "Q" + RowCount.ToString()).Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;
            //ReportRange = ReportWS.get_Range("A5", "X10");
            foreach (Microsoft.Office.Interop.Excel.Range cell in ReportRange.Cells)
            {
                cell.BorderAround2();
            }

            /*
            ReportWS.Cells[5, 7] = "Gage #";
            ReportWS.Cells[5, 8] = "Bale #";
            ReportWS.Cells[5, 9] = "Lot #";
            ReportWS.Cells[5, 11] = "Employee ID";
            ReportWS.Cells[5, 12] = "Code #";
            ReportWS.Cells[5, 14] = "Buddy Check ID";
            ReportWS.Cells[5, 15] = "A";
            ReportWS.Cells[5, 16] = "B";
            ReportWS.Cells[5, 17] = "C";
            ReportWS.Cells[5, 18] = "D";
            ReportWS.Cells[5, 19] = "E";
            ReportWS.Cells[5, 20] = "F";
            ReportWS.Cells[5, 21] = "G";
            ReportWS.Cells[5, 22] = "H";
            ReportWS.Cells[5, 23] = "I";          */

            ReportWS.Cells[5, 1] = "Date";
            ReportWS.Cells[5, 2] = "Time";
            ReportWS.Cells[5, 3] = "Gage #";
            ReportWS.Cells[5, 3] = "Gage #";
            ReportWS.Cells[5, 4] = "Bale #";
            ReportWS.Cells[5, 5] = "Lot #";
            ReportWS.Cells[5, 6] = "Employee ID";
            ReportWS.Cells[5, 7] = "Code #";
            ReportWS.Cells[5, 8] = "Buddy Check ID";
            ReportWS.Cells[5, 9] = "A";
            ReportWS.Cells[5, 10] = "B";
            ReportWS.Cells[5, 11] = "C";
            ReportWS.Cells[5, 12] = "D";
            ReportWS.Cells[5, 13] = "E";
            ReportWS.Cells[5, 14] = "F";
            ReportWS.Cells[5, 15] = "G";
            ReportWS.Cells[5, 16] = "H";
            ReportWS.Cells[5, 17] = "I";

            ReportRange = ReportWS.get_Range("A6", "Y6");
            //ReportWS.Cells.BorderAround2();
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

        private void CreateExcelFile_RunWorkerComplete(object sender, RunWorkerCompletedEventArgs e)
        {
            Process.Start(ExcelFileLocation);
        }

        private void Clear()
        {
            // Search_GroupBox Clear
            ItemIDSearch_TextBox.Clear();
            CustomerItemNumberSearch_TextBox.Clear();
            OperatorIDSearch_TextBox.Clear();
            BrakePressSearch_ComboBox.Text = "";
            ItemID_CheckBox.Checked = true;
            CustomerItemNumber_CheckBox.Checked = false;
            DateStartPicker.Checked = false;
            DateStartPicker.Size = new System.Drawing.Size(323, 30);
            DateStartPicker.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            DateStartPicker.ResetText();
            DateEndPicker.Size = new System.Drawing.Size(323, 30);
            DateEndPicker.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            DateEndPicker.ResetText();

            // GridView Clear
            CheckCardData_GridView.DataSource = null;

            // Results_GroupBox Clear
            ItemID_TextBox.Clear();
            Sequence_TextBox.Clear();
            Customer_TextBox.Clear();
            CustomerItemNumber_TextBox.Clear();
            Date_TextBox.Clear();
            Time_TextBox.Clear();
            GageNumber_TextBox.Clear();
            BaleNumber_TextBox.Clear();
            LotNumber_TextBox.Clear();
            OperatorID_TextBox.Clear();
            OperatorName_TextBox.Clear();
            BuddyCheckID_TextBox.Clear();
            BuddyCheckName_TextBox.Clear();
            CodeNumber_TextBox.Clear();
            A_TextBox.Clear();
            B_TextBox.Clear();
            C_TextBox.Clear();
            D_TextBox.Clear();
            E_TextBox.Clear();
            F_TextBox.Clear();
            G_TextBox.Clear();
            H_TextBox.Clear();
            I_TextBox.Clear();

            // CreatePDF File and Create Excel File
            SearchCommand = null;

            ItemIDSearch_TextBox.Focus();
        }

        private void EmployeeLogOff()
        {
            if (LoginTime != "")
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

        void PartImage_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {

        }

        /********************************************************************************************************************
        * 
        * Events Region Start
        * 
        * - SearchItemID CheckBox CheckedChanged
        * - CustomerItemNumber CheckBox CheckedChanged
        * - DateStartPicker DropDown
        * - DateEndPicker DropDown
        * - Clock Tick
        * 
        ********************************************************************************************************************/
        #region

        private void ItemID_CheckBox_CheckedChanged(object sender, EventArgs e)
        {
            if (ItemID_CheckBox.Checked == true)
            {
                CustomerItemNumber_CheckBox.Checked = false;
            }
        }

        private void CustomerItemNumber_CheckBox_CheckedChanged(object sender, EventArgs e)
        {
            if (CustomerItemNumber_CheckBox.Checked == true)
            {
                ItemID_CheckBox.Checked = false;
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
            else if (ClockHour < 10)
            {
                Time += "0" + ClockHour;
                AMPM = "AM";
            }
            else if (ClockHour >= 10 && ClockHour <= 12)
            {
                Time += ClockHour;
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
        * Events Region End
        * 
        ********************************************************************************************************************/
        #endregion

    }
}
