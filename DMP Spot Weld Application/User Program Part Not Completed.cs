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
 * Program: DMP Spot Weld Application
 * Form: User Program Part Not Completed
 * Created By: Ryan Garland
 * Last Updated on 8/30/18
 * 
 */

namespace DMP_Spot_Weld_Application
{
    public partial class User_Program_Part_Not_Completed : Form
    {
        BackgroundWorker FaultOff;

        public User_Program_Part_Not_Completed()
        {
            InitializeComponent();
            OK_Button.DialogResult = DialogResult.Yes;
            this.ShowInTaskbar = false;
            // 
            FaultOff = new BackgroundWorker();
            FaultOff.DoWork += new DoWorkEventHandler(ConfirmFaultReset_Off_OPC);
            FaultOff.RunWorkerCompleted += new RunWorkerCompletedEventHandler(FaultOff_RunWorkerCompleted);
        }

        // Connect To Server on Form Load
        private Opc.Da.Server OPCServer;
        private OpcCom.Factory OPCFactory = new OpcCom.Factory();

        // ConfirmFaultReset_Off_OPC Method
        private Opc.Da.Subscription Fault_Off_Write;
        private Opc.Da.SubscriptionState Fault_Off_StateWrite;

        // PLC Tag Name
        private static string Spotweld_Tag_Name = "";

        // Not Used
        private Opc.Da.Subscription Fault_On_Write;
        private Opc.Da.SubscriptionState Fault_On_StateWrite;

        // Form Loads. 
        // Checks Spot Weld Computer. 
        // Connects to OPC Server
        private void User_Program_Part_Not_Completed_Load(object sender, EventArgs e)
        {
            SpotWeldID();
            OPCServer = new Opc.Da.Server(OPCFactory, null);
            OPCServer.Url = new Opc.URL("opcda://OHN66OPC/Kepware.KEPServerEX.V6");
            OPCServer.Connect();

            Fault_Off_StateWrite = new Opc.Da.SubscriptionState();
            Fault_Off_StateWrite.Name = "PB_Reset_Off_Fault";
            Fault_Off_StateWrite.Active = true;
            Fault_Off_Write = (Opc.Da.Subscription)OPCServer.CreateSubscription(Fault_Off_StateWrite);

            Fault_On_StateWrite = new Opc.Da.SubscriptionState();
            Fault_On_StateWrite.Name = "PB_Reset_On_Fault";
            Fault_On_StateWrite.Active = true;
            Fault_On_Write = (Opc.Da.Subscription)OPCServer.CreateSubscription(Fault_On_StateWrite);
        }

        // Start The FaultOff Background Worker
        private void OK_Button_Click(object sender, EventArgs e)
        {
            FaultOff.RunWorkerAsync(); // Run ConfirmFaultReset_Off_OPC Method

        }

        // Set the Fault Value in the PLC back to 0
        private void ConfirmFaultReset_Off_OPC(object sender, EventArgs e)
        {
            Opc.Da.Item[] OPC_Fault_Off = new Opc.Da.Item[1];
            OPC_Fault_Off[0] = new Opc.Da.Item();
            OPC_Fault_Off[0].ItemName = Spotweld_Tag_Name + "Fault";
            OPC_Fault_Off = Fault_Off_Write.AddItems(OPC_Fault_Off);

            Opc.Da.ItemValue[] OPC_ResetValue_Off = new Opc.Da.ItemValue[1];
            OPC_ResetValue_Off[0] = new Opc.Da.ItemValue();
            OPC_ResetValue_Off[0].ServerHandle = Fault_Off_Write.Items[0].ServerHandle;
            OPC_ResetValue_Off[0].Value = 0;

            Opc.IRequest OPCRequest;
            Fault_Off_Write.Write(OPC_ResetValue_Off, 123, new Opc.Da.WriteCompleteEventHandler(WriteCompleteCallback), out OPCRequest);
        }

        // When the Worker is Complete we disconnect from the OPC Server, enable the user program, and close this form
        private void FaultOff_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            OPCServer.Disconnect();
            User_Program.UserProgram.Enabled = true;
            this.Close();
        }

        private void WriteCompleteCallback(object clientHandle, Opc.IdentifiedResult[] results)
        {
            foreach (Opc.IdentifiedResult writeResult in results)
            {
                Console.WriteLine("\t{0} write result: {1}", writeResult.ItemName, writeResult.ResultID);
            }
        }

        // Set the SpotWeld_Tag_Name
        private void SpotWeldID()
        {
            string SpotWeldComputerID = System.Environment.MachineName;

            // CAT Spot Weld
            if (SpotWeldComputerID == "123R") // CAT - 123
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
            // My Computer
            if (SpotWeldComputerID == "OHN7047NL") //  My Laptop
            {
                Spotweld_Tag_Name = "OHN66OPC.Spot_Weld_121R.Global.SW121R_";
            }
        }
    }    
}
