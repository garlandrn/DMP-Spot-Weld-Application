using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Opc.Da;

/*
 * Program: DMP Spot Weld Application
 * Form: User Program Job Data
 * Created By: Ryan Garland
 * Last Updated on 8/28/18
 * 
 */

namespace DMP_Spot_Weld_Application
{
    public partial class User_Program_Job_Data : Form
    {
        User_Program owner = null;
        BackgroundWorker ConnectToOPC;
        public User_Program_Job_Data(User_Program owner)
        {
            InitializeComponent();
            this.owner = owner;
            ConnectToOPC = new BackgroundWorker();
            ConnectToOPC.DoWork += new DoWorkEventHandler(JobData_ConnectToOPCServer);
            ConnectToOPC.RunWorkerCompleted += new RunWorkerCompletedEventHandler(JobData_ConnectToServer_OPC_RunWorkerCompleted);
            Run_Button.DialogResult = DialogResult.Yes;
            Cancel_Button.DialogResult = DialogResult.No;
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
        * Barcode Scanner Variables 
        * 
        ********************************************************************************************************************/

        /********************************************************************************************************************
        * 
        * OPC Tag Variables 
        * 
        ********************************************************************************************************************/

        // OPC Server
        private Opc.Da.Server OPCServer;
        private OpcCom.Factory OPCFactory = new OpcCom.Factory();

        // Subsciptions
        private Opc.Da.Subscription Component_Read;
        private Opc.Da.SubscriptionState Component_StateRead;
        private Opc.Da.Subscription ScanNewPart_Write;
        private Opc.Da.SubscriptionState ScanNewPart_StateWrite;
        private Opc.Da.Subscription Hardware_Write;
        private Opc.Da.SubscriptionState Hardware_StateWrite;

        // Toggle Tag HMI_PB_SCAN_HARDWARE_VALUE on and off
        private static int HMI_PB_SCAN_HARDWARE_VALUE;


        private static string Spotweld_Tag_Name = "";
        private static string Correct_Component_Bit = "";
        

        /********************************************************************************************************************
        * 
        * Form Load Variables 
        * 
        ********************************************************************************************************************/

        private static int QuantityOfParts;
        private static string ScannerString = "";
        private static string ReferenceNumber = "";

        // Not Needed
        private static bool Comp1Scanned = false;
        private static bool Comp2Scanned = false;
        private static bool Comp3Scanned = false;
        private static bool Comp4Scanned = false;
        private static bool Comp1Needed = false;
        private static bool Comp2Needed = false;
        private static bool Comp3Needed = false;
        private static bool Comp4Needed = false;
        private static bool FixtureScanned = false;
        private static bool FixtureNeeded = false;
        private static string AllComponentsFound = "";
        private static int NumberOfComponents;
        private static int YesNoFixture;
        private static string[] ItemsToScan = new string[] { };
        private TextBox[] JobDataTextBoxes;

        /********************************************************************************************************************
        * 
        * User Program: Job Data Start
        * 
        ********************************************************************************************************************/

        private void User_Program_Job_Data_Load(object sender, EventArgs e)
        {
            SpotWeldID();
            ReferenceNumber_TextBox.Focus();
            
            ConnectToOPC.RunWorkerAsync();
            Scan_ListBox.Items.Add("Please Scan The Job Number");
            ItemID_TextBox.ReadOnly = true;
        }

        /********************************************************************************************************************
        *  
        *  [User Interface]
        *  
        *  Buttons
        * 
        ********************************************************************************************************************
        ********************************************************************************************************************
        * [Buttons]
        * 
        * -----------------------------------------------------[Enter]------------------------------------------------------
        * -- 
        * 
        * ------------------------------------------------------[Run]-------------------------------------------------------
        * --  
        *   
        * ---------------------------------------------------- [Cancel]-----------------------------------------------------
        * --Global Variables:
        *   JobFound = false;
        * 
        ********************************************************************************************************************/
        
        private void Enter_Button_Click(object sender, EventArgs e)
        {
            try
            {
                QuantityOfParts = Int32.Parse(PartsNeeded_TextBox.Text);
                if (QuantityOfParts >= 1)
                {
                    PartsNeeded_TextBox.ReadOnly = true;
                    Enter_Button.Enabled = false;
                    Enter_Button.Visible = false;
                    ShowComponents();
                    Scan_TextBox.Focus();
                    Scan_ListBox.Items.Clear();
                    Scan_ListBox.Items.Add("Please Scan In Components Needed");
                }
                else if (QuantityOfParts == 0)
                {
                    MessageBox.Show("Please Enter a Value Greater Than 0");
                }
            }
            catch
            {
                MessageBox.Show("Please Enter a Value Greater Than 0");
            }
        }

        private void ReferenceEnter_Button_Click(object sender, EventArgs e)
        {
            try
            {
                string JobNumber = "";
                ReferenceNumber = ReferenceNumber_TextBox.Text;
                JobNumber = ReferenceNumber.Substring(0, 11);
                if (JobNumber.Length == 11 && (JobNumber.StartsWith("j") || JobNumber.StartsWith("J")))
                {
                    PartsNeeded_TextBox.Visible = true;
                    PartsNeeded_Label.Visible = true;
                    PartsNeeded_TextBox.Focus();
                    ReferenceNumber_TextBox.ReadOnly = true;
                    ReferenceEnter_Button.Visible = false;
                    Enter_Button.Visible = true;
                    Scan_ListBox.Items.Clear();
                    Scan_ListBox.Items.Add("Please Enter The Number of Parts Needed");
                    PartsNeeded_TextBox.Focus();
                }
                else if (JobNumber.Length != 11 || (JobNumber.StartsWith("j") || JobNumber.StartsWith("J")) == false)
                {
                    MessageBox.Show("Reference Number Invalid" + "\nPlease Scan Reference Number Again");
                    ReferenceNumber_TextBox.Clear();
                }
            }
            catch (Exception)
            {
                MessageBox.Show("Reference Number Invalid" + "\nPlease Scan Reference Number Again");
                ReferenceNumber_TextBox.Clear();
            }
        }

        private void Run_Button_Click(object sender, EventArgs e)
        {
            HMI_PB_SCAN_HARDWARE_VALUE = 0;
            SetHardware_OPC();
            this.owner.PassValue(PartsNeeded_TextBox.Text);
            this.owner.PassReferenceNumber(ReferenceNumber_TextBox.Text);
            User_Program.UserProgram.Enabled = true;
            OPCServer.Disconnect();
            this.Close();
        }

        private void Cancel_Button_Click(object sender, EventArgs e)
        {
            HMI_PB_SCAN_HARDWARE_VALUE = 0;
            SetHardware_OPC();
            User_Program.UserProgram.Enabled = true;
            owner.StartNewJob_Button.Focus();
            OPCServer.Disconnect();
            this.Close();
        }

        /*********************************************************************************************************************
        * 
        * Buttons End
        * 
        *********************************************************************************************************************/

        /*********************************************************************************************************************
        *  
        *  Methods
        *  
        *  Show Components():
        *  Search For Components():
        *  Search For Job():
        *  
        * 
        *********************************************************************************************************************/

        private void JobData_ConnectToOPCServer(object sender, EventArgs e)
        {
            try
            {
                OPCServer = new Opc.Da.Server(OPCFactory, null);
                OPCServer.Url = new Opc.URL("opcda://OHN66OPC/Kepware.KEPServerEX.V6");
                OPCServer.Connect();

                Component_StateRead = new Opc.Da.SubscriptionState();
                Component_StateRead.Name = "Component_Reading";
                Component_StateRead.UpdateRate = 1000;
                Component_StateRead.Active = true;
                Component_Read = (Opc.Da.Subscription)OPCServer.CreateSubscription(Component_StateRead);
                Component_Read.DataChanged += new Opc.Da.DataChangedEventHandler(Component_Read_DataChanged);

                ScanNewPart_StateWrite = new Opc.Da.SubscriptionState();
                ScanNewPart_StateWrite.Name = "ScanNewPart_WriteGroup";
                ScanNewPart_StateWrite.Active = true;
                ScanNewPart_Write = (Opc.Da.Subscription)OPCServer.CreateSubscription(ScanNewPart_StateWrite);


                Hardware_StateWrite = new Opc.Da.SubscriptionState();
                Hardware_StateWrite.Name = "Hardware_WriteGroup";
                Hardware_StateWrite.Active = true;
                Hardware_Write = (Opc.Da.Subscription)OPCServer.CreateSubscription(Hardware_StateWrite);                
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void JobData_ConnectToServer_OPC_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            HMI_PB_SCAN_HARDWARE_VALUE = 1;
            SetHardware_OPC();
            ComponentCheck_OPC();
        }
        
        // Turn off Scan New Part Input
        // Toggle Scan Hardware Input
        private void SetHardware_OPC()
        {
            Opc.Da.Item[] OPC_SetHardware = new Opc.Da.Item[2];
            OPC_SetHardware[0] = new Opc.Da.Item();
            OPC_SetHardware[0].ItemName = Spotweld_Tag_Name + "HMI_PB_SCAN_NEW_PART";
            OPC_SetHardware[1] = new Opc.Da.Item();
            OPC_SetHardware[1].ItemName = Spotweld_Tag_Name + "HMI_PB_SCAN_HARDWARE";
            OPC_SetHardware = Hardware_Write.AddItems(OPC_SetHardware);

            Opc.Da.ItemValue[] OPC_SetHardware_Value = new Opc.Da.ItemValue[2];
            OPC_SetHardware_Value[0] = new Opc.Da.ItemValue();
            OPC_SetHardware_Value[0].ServerHandle = Hardware_Write.Items[0].ServerHandle;
            OPC_SetHardware_Value[0].Value = 0;
            OPC_SetHardware_Value[1] = new Opc.Da.ItemValue();
            OPC_SetHardware_Value[1].ServerHandle = Hardware_Write.Items[1].ServerHandle;
            OPC_SetHardware_Value[1].Value = HMI_PB_SCAN_HARDWARE_VALUE;
            Opc.IRequest OPCRequest;
            Hardware_Write.Write(OPC_SetHardware_Value, 123, new Opc.Da.WriteCompleteEventHandler(WriteCompleteCallback), out OPCRequest);
        }

        // Write The Component Value From Barcode Scanner
        private void ScanComponents_OPC()
        {
            Opc.Da.Item[] OPC_ScanComponent = new Opc.Da.Item[1];
            OPC_ScanComponent[0] = new Opc.Da.Item();
            OPC_ScanComponent[0].ItemName = Spotweld_Tag_Name + "DM8050_READ_RESULTS.DATA";
            OPC_ScanComponent = ScanNewPart_Write.AddItems(OPC_ScanComponent);

            Opc.Da.ItemValue[] OPC_ScanComponent_Value = new Opc.Da.ItemValue[1];
            OPC_ScanComponent_Value[0] = new Opc.Da.ItemValue();
            OPC_ScanComponent_Value[0].ServerHandle = ScanNewPart_Write.Items[0].ServerHandle;
            OPC_ScanComponent_Value[0].Value = ScannerString;
            Opc.IRequest OPCRequest;
            ScanNewPart_Write.Write(OPC_ScanComponent_Value, 123, new Opc.Da.WriteCompleteEventHandler(WriteCompleteCallback), out OPCRequest);
        }

        // Initialize the Correct Components Bit From the PLC
        private void ComponentCheck_OPC()
        {
            List<Item> OPC_ReadList = new List<Item>();
            Opc.Da.Item[] OPC_ReadComponent = new Opc.Da.Item[1];
            OPC_ReadComponent[0] = new Opc.Da.Item();
            OPC_ReadComponent[0].ItemName = Spotweld_Tag_Name + "CORRECT_COMPONENTS_BIT";
            OPC_ReadList.Add(OPC_ReadComponent[0]);
            Component_Read.AddItems(OPC_ReadList.ToArray());

            Opc.IRequest req;
            Component_Read.Read(Component_Read.Items, 123, new Opc.Da.ReadCompleteEventHandler(ReadCompleteCallback), out req);
        }

        // When the Value of Correct Component Bit Changes Enable the OK Button
        public void Component_Read_DataChanged(object subscriptionHandle, object requestHandle, ItemValueResult[] values)
        {
            // CAT Spot Welders
            if (System.Environment.MachineName == "123R") // CAT - 123R
            {
                foreach (ItemValueResult itemValue in values)
                {
                    if (itemValue.ItemName == "OHN66OPC.Spot_Weld_123R.Global.CORRECT_COMPONENTS_BIT")
                    {
                        Correct_Component_Bit = Convert.ToString(itemValue.Value);
                    }
                }
            }
            else if (System.Environment.MachineName == "1088") // CAT - 1088
            {
                foreach (ItemValueResult itemValue in values)
                {
                    if (itemValue.ItemName == "OHN66OPC.Spot_Weld_1088.Global.CORRECT_COMPONENTS_BIT")
                    {
                        Correct_Component_Bit = Convert.ToString(itemValue.Value);
                    }
                }
            }
            // John Deere Spot Welders
            else if (System.Environment.MachineName == "108R") // John Deere - 108R
            {
                foreach (ItemValueResult itemValue in values)
                {
                    if (itemValue.ItemName == "OHN66OPC.Spot_Weld_108R.Global.CORRECT_COMPONENTS_BIT")
                    {
                        Correct_Component_Bit = Convert.ToString(itemValue.Value);
                    }
                }
            }
            else if (System.Environment.MachineName == "150R") // John Deere - 150R
            {
                foreach (ItemValueResult itemValue in values)
                {
                    if (itemValue.ItemName == "OHN66OPC.Spot_Weld_150R.Global.CORRECT_COMPONENTS_BIT")
                    {
                        Correct_Component_Bit = Convert.ToString(itemValue.Value);
                    }
                }
            }
            // Navistar Spot Welders
            else if (System.Environment.MachineName == "OHN7149") // Navistar - 121R
            {
                foreach (ItemValueResult itemValue in values)
                {
                    if (itemValue.ItemName == "OHN66OPC.Spot_Weld_121R.Global.SW121R_CORRECT_COMPONENTS_BIT")
                    {
                        Correct_Component_Bit = Convert.ToString(itemValue.Value);
                    }
                }
            }
            else if (System.Environment.MachineName == "OHN7111") // Navistar - 154R
            {
                foreach (ItemValueResult itemValue in values)
                {
                    if (itemValue.ItemName == "OHN66OPC.Spot_Weld_154R.Global.CORRECT_COMPONENTS_BIT")
                    {
                        Correct_Component_Bit = Convert.ToString(itemValue.Value);
                    }
                }
            }
            // Paccar Spot Welders
            else if (System.Environment.MachineName == "OHN7124") // Paccar 153R
            {
                foreach (ItemValueResult itemValue in values)
                {
                    if (itemValue.ItemName == "OHN66OPC.Spot_Weld_153R.Global.CORRECT_COMPONENTS_BIT")
                    {
                        Correct_Component_Bit = Convert.ToString(itemValue.Value);
                    }
                }
            }
            else if (System.Environment.MachineName == "OHN7123") // Paccar 155R
            {
                foreach (ItemValueResult itemValue in values)
                {
                    if (itemValue.ItemName == "OHN66OPC.Spot_Weld_155R.Global.CORRECT_COMPONENTS_BIT")
                    {
                        Correct_Component_Bit = Convert.ToString(itemValue.Value);
                    }
                }
            }
            // My Computer For Testing
            else if (System.Environment.MachineName == "OHN7047NL") // My PC
            {
                foreach (ItemValueResult itemValue in values)
                {
                    if (itemValue.ItemName == "OHN66OPC.Spot_Weld_121R.Global.SW121R_CORRECT_COMPONENTS_BIT")
                    {
                        Correct_Component_Bit = Convert.ToString(itemValue.Value);
                    }
                }
            }
            CorrectComponents_TextBox.Invoke(new EventHandler(delegate { CorrectComponents_TextBox.Text = Correct_Component_Bit; }));
            if (CorrectComponents_TextBox.Text == "True")
            {
                Run_Button.Invoke(new EventHandler(delegate { Run_Button.Show(); }));
                Run_Button.Invoke(new EventHandler(delegate { Run_Button.Focus(); }));
            }
            else if (CorrectComponents_TextBox.Text == "False")
            {
                Run_Button.Invoke(new EventHandler(delegate { Run_Button.Hide(); }));
            }
        }

        private void ReadCompleteCallback(object clientHandle, Opc.Da.ItemValueResult[] results)
        {
            CorrectComponents_TextBox.Invoke(new EventHandler(delegate { CorrectComponents_TextBox.Text = (results[0].Value).ToString(); }));
        }

        private void WriteCompleteCallback(object clientHandle, Opc.IdentifiedResult[] results)
        {
            foreach (Opc.IdentifiedResult writeResult in results)
            {
                Console.WriteLine("\t{0} write result: {1}", writeResult.ItemName, writeResult.ResultID);
            }
        }


        /*********************************************************************************************************************
        * 
        * Methods Region Start
        * 
        *********************************************************************************************************************/
        #region

        // Scan in Reference Number and Check to See if it is Valid
        private void ReferenceNumber_TextBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                try
                {
                    string JobNumber = "";
                    ReferenceNumber = ReferenceNumber_TextBox.Text;
                    JobNumber = ReferenceNumber.Substring(0, 11);
                    if (JobNumber.Length == 11 && (JobNumber.StartsWith("j") || JobNumber.StartsWith("J")))
                    {
                        PartsNeeded_TextBox.Visible = true;
                        PartsNeeded_Label.Visible = true;
                        PartsNeeded_TextBox.Focus();
                        ReferenceNumber_TextBox.ReadOnly = true;
                        ReferenceEnter_Button.Visible = false;
                        Enter_Button.Visible = true;
                        Scan_ListBox.Items.Clear();
                        Scan_ListBox.Items.Add("Please Enter The Number of Parts Needed");
                        PartsNeeded_TextBox.Focus();
                    }
                    else if (JobNumber.Length != 11 || (JobNumber.StartsWith("j") || JobNumber.StartsWith("J")) == false)
                    {
                        MessageBox.Show("Reference Number Invalid" + "\nPlease Scan Reference Number Again");
                        ReferenceNumber_TextBox.Clear();
                    }
                }
                catch (Exception)
                {
                    MessageBox.Show("Reference Number Invalid" + "\nPlease Scan Reference Number Again");
                    ReferenceNumber_TextBox.Clear();
                }
            }
        }

        // Occurs When The Scan TextBox is Focused on and a Key is Pressed 
        private void Scan_TextBox_KeyDown(object sender, KeyEventArgs e)
        {
            // The Barcode Scanner Acts as the Enter Button
            if (e.KeyCode == Keys.Enter)
            {
                ScannerString = Scan_TextBox.Text; // Move Barcode Scanned to ScannerString
                int StringError;
                StringError = Regex.Matches(ScannerString, @"[\D]").Count; // Check to see if barcode contains any invalid characters
                if (ScannerString.Length >= 9 && StringError == 1)
                {
                    ScanComponents_OPC(); // Write Barcode To PLC
                    ScanItems();  // Check to See if Barcode Scanned Matches any Components
                }
                else
                {
                    Scan_ListBox.Items.Add("Item " + ScannerString + " Was Not a Correct Component or Fixture");
                    Scan_TextBox.Clear();
                }
            }
        }
         
        // Check to See if Barcode Matches any Components
        private void ScanItems()
        {
            if (ScannerString == Comp1_TextBox.Text)
            {
                if (Comp1Scan_Label.Visible == false)
                {
                    Scan_ListBox.Items.Add("Component: " + ScannerString + " Scanned" + "\n");
                    Comp1Scanned = true;
                    this.Comp1_TextBox.BackColor = Color.Chartreuse;
                    Comp1Scan_Label.Show();
                    Scan_TextBox.Clear();
                }
                else if (Comp1Scan_Label.Visible == true)
                {
                    Scan_ListBox.Items.Add("Component: " + ScannerString + " Has Already Been Scanned" + "\n");
                    Scan_TextBox.Clear();
                }
            }
            else if (ScannerString == Comp2_TextBox.Text)
            {
                if (Comp2Scan_Label.Visible == false)
                {
                    Scan_ListBox.Items.Add("Component: " + ScannerString + " Scanned" + "\n");
                    Comp2Scanned = true;
                    this.Comp2_TextBox.BackColor = Color.Chartreuse;
                    Comp2Scan_Label.Show();
                    Scan_TextBox.Clear();
                }
                else if (Comp2Scan_Label.Visible == true)
                {
                    Scan_ListBox.Items.Add("Component: " + ScannerString + " Has Already Been Scanned" + "\n");
                    Scan_TextBox.Clear();
                }
            }
            else if (ScannerString == Comp3_TextBox.Text)
            {
                if (Comp3Scan_Label.Visible == false)
                {
                    Scan_ListBox.Items.Add("Component: " + ScannerString + " Scanned" + "\n");
                    Comp3Scanned = true;
                    this.Comp3_TextBox.BackColor = Color.Chartreuse;
                    Comp3Scan_Label.Show();
                    Scan_TextBox.Clear();
                }
                else if (Comp3Scan_Label.Visible == true)
                {
                    Scan_ListBox.Items.Add("Component: " + ScannerString + " Has Already Been Scanned" + "\n");
                    Scan_TextBox.Clear();
                }
            }
            else if (ScannerString == Comp4_TextBox.Text)
            {
                if (Comp4Scan_Label.Visible == false)
                {
                    Scan_ListBox.Items.Add("Component: " + ScannerString + " Scanned" + "\n");
                    Comp4Scanned = true;
                    this.Comp4_TextBox.BackColor = Color.Chartreuse;
                    Comp4Scan_Label.Show();
                    Scan_TextBox.Clear();
                }
                else if (Comp4Scan_Label.Visible == true)
                {
                    Scan_ListBox.Items.Add("Component: " + ScannerString + " Has Already Been Scanned" + "\n");
                    Scan_TextBox.Clear();
                }
            }
            else if (ScannerString == Fixture_TextBox.Text)
            {
                if (FixtureScan_Label.Visible == false)
                {
                    Scan_ListBox.Items.Add("Component: " + ScannerString + " Scanned" + "\n");
                    FixtureScanned = true;
                    this.Fixture_TextBox.BackColor = Color.Chartreuse;
                    FixtureScan_Label.Show();
                    Scan_TextBox.Clear();
                }
                else if (FixtureScan_Label.Visible == true)
                {
                    Scan_ListBox.Items.Add("Fixture: " + ScannerString + " Has Already Been Scanned" + "\n");
                    Scan_TextBox.Clear();
                }
            }
            else
            {
                Scan_ListBox.Items.Add("Item " + ScannerString + " Was Not a Correct Component or Fixture");
                Scan_TextBox.Clear();
            }
        }        

        // Set the OPC Tag Name
        private void SpotWeldID()
        {
            string SpotWeldComputerID = System.Environment.MachineName;

            // CAT Spot Weld
            if (SpotWeldComputerID == "123R") // CAT - 123R
            {
                Spotweld_Tag_Name = "OHN66OPC.Spot_Weld_123R.Global.";
            }
            if (SpotWeldComputerID == "1088") // CAT - 1088
            {
                Spotweld_Tag_Name = "OHN66OPC.Spot_Weld_1088.Global.";
            }
            // John Deere Spot Weld
            if (SpotWeldComputerID == "108R") // John Deere - 108R
            {
                Spotweld_Tag_Name = "OHN66OPC.Spot_Weld_108R.Global.";
            }
            if (SpotWeldComputerID == "150R") // John Deere - 150R
            {
                Spotweld_Tag_Name = "OHN66OPC.Spot_Weld_150R.Global.";
            }
            // Navistar
            if (SpotWeldComputerID == "104R") // Navistar - 104R
            {
                Spotweld_Tag_Name = "OHN66OPC.Spot_Weld_104R.Global.";
            }
            if (SpotWeldComputerID == "OHN7149") // Navistar - 121R
            {
                Spotweld_Tag_Name = "OHN66OPC.Spot_Weld_121R.Global.SW121R_";
            }
            if (SpotWeldComputerID == "OHN7111") // Navistar - 154R
            {
                Spotweld_Tag_Name = "OHN66OPC.Spot_Weld_154R.Global.";
            }
            // Paccar
            if (SpotWeldComputerID == "OHN7124") // Paccar - 153R
            {
                Spotweld_Tag_Name = "OHN66OPC.Spot_Weld_153R.Global.";
            }
            if (SpotWeldComputerID == "OHN7123") // Paccar - 155R
            {
                Spotweld_Tag_Name = "OHN66OPC.Spot_Weld_155R.Global.";
            }
            if (SpotWeldComputerID == "OHN7047NL") //  My Laptop
            {
                Spotweld_Tag_Name = "OHN66OPC.Spot_Weld_121R.Global.SW121R_";
            }
        }

        // Check To See What Components Need Scanned and Make Them Visible
        private void ShowComponents()
        {
            Scan_ListBox.Show();
            Fixture_Label.Show();
            Fixture_TextBox.Show();
            //FixtureLocation_TextBox.Show();
            //FixtureLocation_Label.Show();

            if (Comp1_TextBox.TextLength == 9)
            {
                Comp1_Label.Show();
                Comp1_TextBox.Show();
                //Scan_TextBox.Show();
            }
            if (Comp2_TextBox.TextLength == 9)
            {
                Comp2_Label.Show();
                Comp2_TextBox.Show();
            }
            if (Comp3_TextBox.TextLength == 9)
            {
                Comp3_Label.Show();
                Comp3_TextBox.Show();
            }
            if (Comp4_TextBox.TextLength == 9)
            {
                Comp4_Label.Show();
                Comp4_TextBox.Show();
            }
        }

        /*********************************************************************************************************************
        * 
        * Methods Region End
        * 
        *********************************************************************************************************************/
        #endregion

        /*********************************************************************************************************************
        * 
        * Events Region Start
        * 
        *********************************************************************************************************************/
        #region

        private void PartsNeeded_TextBox_Enter(object sender, EventArgs e)
        {
            if (PartsNeeded_TextBox.ReadOnly == true && Scan_TextBox.Visible == true)
            {
                Scan_TextBox.Focus();
            }
            else if (PartsNeeded_TextBox.ReadOnly == true && Run_Button.Visible == true)
            {
                Run_Button.Focus();
            }
        }

        private void PartsNeeded_TextBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter && PartsNeeded_TextBox.ReadOnly == false)
            {
                Enter_Button_Click(null, null);
            }
        }

        private void ItemID_TextBox_Enter(object sender, EventArgs e)
        {
            if (Scan_TextBox.Visible == true && Run_Button.Visible == false)
            {
                Scan_TextBox.Focus();
            }
            else if (Scan_TextBox.Visible == false && Run_Button.Visible == true)
            {
                Run_Button.Focus();
            }
            else if (Scan_TextBox.Visible == false && Run_Button.Visible == false)
            {
                Enter_Button.Focus();
            }
        }

        private void JobID_TextBox_Enter(object sender, EventArgs e)
        {
            if (Scan_TextBox.Visible == true && Run_Button.Visible == false)
            {
                Scan_TextBox.Focus();
            }
            else if (Scan_TextBox.Visible == false && Run_Button.Visible == true)
            {
                Run_Button.Focus();
            }
            else if (Scan_TextBox.Visible == false && Run_Button.Visible == false)
            {
                Enter_Button.Focus();
            }
        }

        private void Comp1_TextBox_Enter(object sender, EventArgs e)
        {
            if (Scan_TextBox.Visible == true && Run_Button.Visible == false)
            {
                Scan_TextBox.Focus();
            }
            else if (Scan_TextBox.Visible == false && Run_Button.Visible == true)
            {
                Run_Button.Focus();
            }
        }

        private void Comp2_TextBox_Enter(object sender, EventArgs e)
        {
            if (Scan_TextBox.Visible == true && Run_Button.Visible == false)
            {
                Scan_TextBox.Focus();
            }
            else if (Scan_TextBox.Visible == false && Run_Button.Visible == true)
            {
                Run_Button.Focus();
            }
        }

        private void Comp3_TextBox_Enter(object sender, EventArgs e)
        {
            if (Scan_TextBox.Visible == true && Run_Button.Visible == false)
            {
                Scan_TextBox.Focus();
            }
            else if (Scan_TextBox.Visible == false && Run_Button.Visible == true)
            {
                Run_Button.Focus();
            }
        }

        private void Comp4_TextBox_Enter(object sender, EventArgs e)
        {
            if (Scan_TextBox.Visible == true && Run_Button.Visible == false)
            {
                Scan_TextBox.Focus();
            }
            else if (Scan_TextBox.Visible == false && Run_Button.Visible == true)
            {
                Run_Button.Focus();
            }
        }

        private void ReferenceNumber_TextBox_Enter(object sender, EventArgs e)
        {
            if (ReferenceNumber_TextBox.Enabled == true)
            {
                ReferenceNumber_TextBox.Focus();
            }
            else if (ReferenceNumber_TextBox.Enabled == false)
            {
                PartsNeeded_TextBox.Focus();
            }
        }

        private void Fixture_TextBox_Enter(object sender, EventArgs e)
        {
            if (Scan_TextBox.Visible == true && Run_Button.Visible == false)
            {
                Scan_TextBox.Focus();
            }
            else if (Scan_TextBox.Visible == false && Run_Button.Visible == true)
            {
                Run_Button.Focus();
            }
        }

        private void FixtureLocation_TextBox_Enter(object sender, EventArgs e)
        {
            if (Scan_TextBox.Visible == true && Run_Button.Visible == false)
            {
                Scan_TextBox.Focus();
            }
            else if (Scan_TextBox.Visible == false && Run_Button.Visible == true)
            {
                Run_Button.Focus();
            }
        }

        private void Scan_ListBox_Enter(object sender, EventArgs e)
        {
            if (Scan_TextBox.Visible == true && Run_Button.Visible == false)
            {
                Scan_TextBox.Focus();
            }
            else if (Scan_TextBox.Visible == false && Run_Button.Visible == true)
            {
                Run_Button.Focus();
            }
            else if (Scan_TextBox.Visible == false && Run_Button.Visible == false)
            {
                ReferenceNumber_TextBox.Focus();
            }
        }

        private void ReferenceEnter_Button_Enter(object sender, EventArgs e)
        {
            if (ReferenceNumber_TextBox.Text == "")
            {
                ReferenceNumber_TextBox.Focus();
            }
        }

        /*********************************************************************************************************************
        * 
        * Events Region End
        * 
        *********************************************************************************************************************/
        #endregion


        /*********************************************************************************************************************
        * 
        * No Longer in Use Region
        * 
        *********************************************************************************************************************/
        #region

        private void CheckCompletion()
        {
            switch (NumberOfComponents)
            {
                case 1: // Single Component
                    switch (YesNoFixture)
                    {
                        case 0: // No Fixture

                            if (Comp1Scanned == true && CorrectComponents_TextBox.Text == "True")
                            {
                                Scan_TextBox.Hide();
                                Run_Button.Show();
                                //Run_Button.Focus();
                                //CorrectComponents_Timer.Stop();
                            }
                            break;

                        case 1: // Yes Fixture
                            if (Comp1Scanned == true && FixtureScanned == true)
                            {
                                Scan_TextBox.Hide();
                                Run_Button.Show();
                                Run_Button.Focus();
                            }
                            break;
                    }
                    break;
                case 2: // Two Components
                    switch (YesNoFixture)
                    {
                        case 0: // No Fixture

                            if (Comp1Scanned == true && Comp2Scanned == true)
                            {
                                Scan_TextBox.Hide();
                                Run_Button.Show();
                                Run_Button.Focus();
                            }
                            break;

                        case 1: // Yes Fixture
                            if (Comp1Scanned == true && Comp2Scanned == true && FixtureScanned == true)
                            {
                                Scan_TextBox.Hide();
                                Run_Button.Show();
                                Run_Button.Focus();
                            }
                            break;
                    }
                    break;

                case 3: // Three Components
                    switch (YesNoFixture)
                    {
                        case 0: // No Fixture

                            if (Comp1Scanned == true && Comp2Scanned == true && Comp3Scanned == true)
                            {
                                Scan_TextBox.Hide();
                                Run_Button.Show();
                                Run_Button.Focus();
                            }
                            break;

                        case 1: // Yes Fixture
                            if (Comp1Scanned == true && Comp2Scanned == true && Comp3Scanned == true && FixtureScanned == true)
                            {
                                Scan_TextBox.Hide();
                                Run_Button.Show();
                                Run_Button.Focus();
                            }
                            break;
                    }
                    break;

                case 4: // Three Components
                    switch (YesNoFixture)
                    {
                        case 0: // No Fixture

                            if (Comp1Scanned == true && Comp2Scanned == true && Comp3Scanned == true && Comp4Scanned == true)
                            {
                                Scan_TextBox.Hide();
                                Run_Button.Show();
                                Run_Button.Focus();
                            }
                            break;

                        case 1: // Yes Fixture
                            if (Comp1Scanned == true && Comp2Scanned == true && Comp3Scanned == true && Comp4Scanned == true && FixtureScanned == true)
                            {
                                Scan_TextBox.Hide();
                                Run_Button.Show();
                                Run_Button.Focus();
                            }
                            break;
                    }
                    break;
            }
        }

        private void ItemsNeeded()
        {
            bool[] NeededItems = { Comp1Needed, Comp2Needed, Comp3Needed, Comp4Needed };
            TextBox[] JobDataTextBoxes = { Comp1_TextBox, Comp2_TextBox, Comp3_TextBox, Comp4_TextBox };
            List<string> ItemID = new List<string>();
            for (int i = 0; i < JobDataTextBoxes.Length; i++)
            {
                if (JobDataTextBoxes[i].TextLength == 9)
                {
                    ItemID.Add(JobDataTextBoxes[i].Text);
                    NeededItems[i] = true;
                    NumberOfComponents = NumberOfComponents + 1;
                }
            }
            if (Fixture_TextBox.Text != "N/A")
            {
                YesNoFixture = 1;
                FixtureNeeded = true;
            }
            ItemsToScan = ItemID.ToArray();
        }

        private void ResetForm()
        {
            Comp1Scanned = false;
            Comp2Scanned = false;
            Comp3Scanned = false;
            Comp4Scanned = false;
            NumberOfComponents = 0;
            YesNoFixture = 0;
            OPCServer.Disconnect();
        }
        #endregion
    }
}
