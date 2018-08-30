using Opc.Da;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

/*
 * 
 * Program: DMP Spot Weld Application
 * Form: User Program Reset Weld Count Dialog
 * Created By: Ryan Garland
 * Last Updated on 8/30/18
 * 
 */

namespace DMP_Spot_Weld_Application
{
    public partial class User_Program_Reset_Weld_Count_Dialog : Form
    {
        //User_Program Owner = null;
        BackgroundWorker ResetWeldCount;

        public User_Program_Reset_Weld_Count_Dialog()
        {
            InitializeComponent();
            this.ShowInTaskbar = false;

            Confirm_Button.DialogResult = DialogResult.Yes;
            Cancel_Button.DialogResult = DialogResult.No;

            ResetWeldCount = new BackgroundWorker();
            ResetWeldCount.DoWork += new DoWorkEventHandler(ConfirmWeldReset_On_OPC);
            ResetWeldCount.RunWorkerCompleted += new RunWorkerCompletedEventHandler(ConfirmWeldReset_On_OPC_RunWorkerComplete);
        }

        /********************************************************************************************************************
        * 
        * OPC Tag Variables 
        * 
        ********************************************************************************************************************/

        // Connect To Server on Form Load
        private Opc.Da.Server OPCServer;
        private OpcCom.Factory OPCFactory = new OpcCom.Factory();

        // Reset 
        private Opc.Da.Subscription ResetWeldCount_Write;
        private Opc.Da.SubscriptionState ResetWeldCount_StateWrite;
        private Opc.Da.Subscription ResetWeldCount_Off_Write;
        private Opc.Da.SubscriptionState ResetWeldCount_Off_StateWrite;
        private Opc.Da.Subscription ResetWeldCount_Read;
        private Opc.Da.SubscriptionState ResetWeldCount_StateRead;

        private static string SpotWeld_Tag_Name = "";

        // Form Loads. 
        // Checks Spot Weld Computer. 
        // Connects to OPC Server
        private void User_Program_Reset_Weld_Count_Dialog_Load(object sender, EventArgs e)
        {
            SpotWeldID();
            OPCServer = new Opc.Da.Server(OPCFactory, null);
            OPCServer.Url = new Opc.URL("opcda://OHN66OPC/Kepware.KEPServerEX.V6");
            OPCServer.Connect();

            ResetWeldCount_StateRead = new Opc.Da.SubscriptionState();
            ResetWeldCount_StateRead.Name = "Reset_Weld_Count_Spotweld";
            ResetWeldCount_StateRead.UpdateRate = 1000;
            ResetWeldCount_StateRead.Active = true;

            ResetWeldCount_Read = (Opc.Da.Subscription)OPCServer.CreateSubscription(ResetWeldCount_StateRead);

            ResetWeldCount_StateWrite = new Opc.Da.SubscriptionState();
            ResetWeldCount_StateWrite.Name = "PB_Reset_Part_Weld_Count_On";
            ResetWeldCount_StateWrite.Active = true;
            ResetWeldCount_Write = (Opc.Da.Subscription)OPCServer.CreateSubscription(ResetWeldCount_StateWrite);

            ResetWeldCount_Off_StateWrite = new Opc.Da.SubscriptionState();
            ResetWeldCount_Off_StateWrite.Name = "PB_Reset_Part_Weld_Count_Off";
            ResetWeldCount_Off_StateWrite.Active = true;
            ResetWeldCount_Off_Write = (Opc.Da.Subscription)OPCServer.CreateSubscription(ResetWeldCount_Off_StateWrite);
        }
        
        private void Confirm_Button_Click(object sender, EventArgs e)
        {
            ResetWeldCount.RunWorkerAsync();
        }

        private void Cancel_Button_Click(object sender, EventArgs e)
        {
            OPCServer.Disconnect();
            User_Program.UserProgram.Enabled = true;
            this.Close();
        }

        // Turn on Reset Input
        private void ConfirmWeldReset_On_OPC(object sender, EventArgs e)
        {
            Opc.Da.Item[] OPC_Reset_On = new Opc.Da.Item[1];
            OPC_Reset_On[0] = new Opc.Da.Item();
            OPC_Reset_On[0].ItemName = SpotWeld_Tag_Name + "HMI_PB_Reset_Part_Weld_Count";
            OPC_Reset_On = ResetWeldCount_Write.AddItems(OPC_Reset_On);

            Opc.Da.ItemValue[] OPC_ResetValue_On = new Opc.Da.ItemValue[1];
            OPC_ResetValue_On[0] = new Opc.Da.ItemValue();
            OPC_ResetValue_On[0].ServerHandle = ResetWeldCount_Write.Items[0].ServerHandle;
            OPC_ResetValue_On[0].Value = 1;
            Opc.IRequest OPCRequest;
            ResetWeldCount_Write.Write(OPC_ResetValue_On, 123, new Opc.Da.WriteCompleteEventHandler(WriteCompleteCallback), out OPCRequest);
            //ResetOff_Timer.Start(); // Start a Timer to turn off Input
        }

        private void ConfirmWeldReset_On_OPC_RunWorkerComplete(object sender, RunWorkerCompletedEventArgs e)
        {
            ConfirmWeldReset_Off_OPC();
            OPCServer.Disconnect();
            User_Program.UserProgram.Enabled = true;
            this.Close();
        }

        // Turn off Reset Input
        private void ConfirmWeldReset_Off_OPC()
        {
            Opc.Da.Item[] OPC_Reset_Off = new Opc.Da.Item[1];
            OPC_Reset_Off[0] = new Opc.Da.Item();
            OPC_Reset_Off[0].ItemName = SpotWeld_Tag_Name + "HMI_PB_Reset_Part_Weld_Count";
            OPC_Reset_Off = ResetWeldCount_Off_Write.AddItems(OPC_Reset_Off);

            Opc.Da.ItemValue[] OPC_ResetValue_Off = new Opc.Da.ItemValue[1];
            OPC_ResetValue_Off[0] = new Opc.Da.ItemValue();
            OPC_ResetValue_Off[0].ServerHandle = ResetWeldCount_Write.Items[0].ServerHandle;
            OPC_ResetValue_Off[0].Value = 0;

            Opc.IRequest OPCRequest;
            ResetWeldCount_Off_Write.Write(OPC_ResetValue_Off, 123, new Opc.Da.WriteCompleteEventHandler(WriteCompleteCallback), out OPCRequest);
        }

        private void WriteCompleteCallback(object clientHandle, Opc.IdentifiedResult[] results)
        {

        }

        // Set the SpotWeld_Tag_Name
        private void SpotWeldID()
        {
            string SpotWeldComputerID = System.Environment.MachineName;

            // CAT Spot Weld
            if (SpotWeldComputerID == "123R") // CAT - 123
            {
                SpotWeld_Tag_Name = "OHN66OPC.Spot_Weld_123R.Global.";
            }
            if (SpotWeldComputerID == "1088") // CAT - 1088
            {
                SpotWeld_Tag_Name = "OHN66OPC.Spot_Weld_1088.Global.";
            }
            // John Deere Spot Weld
            if (SpotWeldComputerID == "108R") // John Deere - 108R
            {
                SpotWeld_Tag_Name = "OHN66OPC.Spot_Weld_108R.Global.";
            }
            if (SpotWeldComputerID == "150R") // John Deere - 150R
            {
                SpotWeld_Tag_Name = "OHN66OPC.Spot_Weld_150R.Global.";
            }
            // Navistar
            if (SpotWeldComputerID == "104R") // Navistar - 104R
            {
                SpotWeld_Tag_Name = "OHN66OPC.Spot_Weld_104R.Global.";
            }
            if (SpotWeldComputerID == "OHN7149") // Navistar - 121R
            {
                SpotWeld_Tag_Name = "OHN66OPC.Spot_Weld_121R.Global.";
            }
            if (SpotWeldComputerID == "OHN7111") // Navistar - 154R
            {
                SpotWeld_Tag_Name = "OHN66OPC.Spot_Weld_154R.Global.";
            }
            // Paccar
            if (SpotWeldComputerID == "OHN7124") // Paccar - 153R
            {
                SpotWeld_Tag_Name = "OHN66OPC.Spot_Weld_153R.Global.";
            }
            if (SpotWeldComputerID == "OHN7123") // Paccar - 155R
            {
                SpotWeld_Tag_Name = "OHN66OPC.Spot_Weld_155R.Global.";
            }
            // My Computer
            if (SpotWeldComputerID == "OHN7047NL") //  My Laptop
            {
                SpotWeld_Tag_Name = "OHN66OPC.Spot_Weld_121R.Global.";
            }
        }
    }
}
