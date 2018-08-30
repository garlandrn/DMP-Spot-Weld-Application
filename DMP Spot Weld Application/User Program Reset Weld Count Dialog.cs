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
 * Last Updated on 8/28/18
 * 
 */

namespace DMP_Spot_Weld_Application
{
    public partial class User_Program_Reset_Weld_Count_Dialog : Form
    {
        User_Program Owner = null;
        public User_Program_Reset_Weld_Count_Dialog(User_Program Owner)
        {
            InitializeComponent();
            this.ShowInTaskbar = false;
            this.Owner = Owner;
            Confirm_Button.DialogResult = DialogResult.Yes;
            Cancel_Button.DialogResult = DialogResult.No;
        }

        /********************************************************************************************************************
        * 
        * OPC Tag Variables 
        * 
        ********************************************************************************************************************/

        // ConnectToServer_OPC Method
        private Opc.Da.Server OPCServer;
        private OpcCom.Factory OPCFactory = new OpcCom.Factory();

        // Reset 
        private Opc.Da.Subscription ResetWeldCount_Write;
        private Opc.Da.SubscriptionState ResetWeldCount_StateWrite;
        private Opc.Da.Subscription ResetWeldCount_Off_Write;
        private Opc.Da.SubscriptionState ResetWeldCount_Off_StateWrite;
        private Opc.Da.Subscription ResetWeldCount_Read;
        private Opc.Da.SubscriptionState ResetWeldCount_StateRead;
        private static string Spotweld_Tag_Name = "";

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
            ConfirmWeldReset_On_OPC();
        }

        private void Cancel_Button_Click(object sender, EventArgs e)
        {
            OPCServer.Disconnect();
            User_Program.UserProgram.Enabled = true;
            this.Close();
        }

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
                Spotweld_Tag_Name = "OHN66OPC.Spot_Weld_121R.Global.";
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
                Spotweld_Tag_Name = "OHN66OPC.Spot_Weld_121R.Global.";
            }
        }

        private void WriteCompleteCallback(object clientHandle, Opc.IdentifiedResult[] results)
        {
            foreach (Opc.IdentifiedResult writeResult in results)
            {
                Console.WriteLine("\t{0} write result: {1}", writeResult.ItemName, writeResult.ResultID);
            }
        }

        // Turn on Reset Input
        private void ConfirmWeldReset_On_OPC()
        {
            Opc.Da.Item[] OPC_Reset_On = new Opc.Da.Item[1];
            OPC_Reset_On[0] = new Opc.Da.Item();
            OPC_Reset_On[0].ItemName = Spotweld_Tag_Name + "HMI_PB_Reset_Part_Weld_Count";
            OPC_Reset_On = ResetWeldCount_Write.AddItems(OPC_Reset_On);

            Opc.Da.ItemValue[] OPC_ResetValue_On = new Opc.Da.ItemValue[1];
            OPC_ResetValue_On[0] = new Opc.Da.ItemValue();
            OPC_ResetValue_On[0].ServerHandle = ResetWeldCount_Write.Items[0].ServerHandle;
            OPC_ResetValue_On[0].Value = 1;
            Opc.IRequest OPCRequest;
            ResetWeldCount_Write.Write(OPC_ResetValue_On, 123, new Opc.Da.WriteCompleteEventHandler(WriteCompleteCallback), out OPCRequest);
            ResetOff_Timer.Start(); // Start a Timer to turn off Input
        }

        // Turn off Reset Input
        private void ConfirmWeldReset_Off_OPC()
        {
            Opc.Da.Item[] OPC_Reset_Off = new Opc.Da.Item[1];
            OPC_Reset_Off[0] = new Opc.Da.Item();
            OPC_Reset_Off[0].ItemName = Spotweld_Tag_Name + "HMI_PB_Reset_Part_Weld_Count";
            OPC_Reset_Off = ResetWeldCount_Off_Write.AddItems(OPC_Reset_Off);

            Opc.Da.ItemValue[] OPC_ResetValue_Off = new Opc.Da.ItemValue[1];
            OPC_ResetValue_Off[0] = new Opc.Da.ItemValue();
            OPC_ResetValue_Off[0].ServerHandle = ResetWeldCount_Write.Items[0].ServerHandle;
            OPC_ResetValue_Off[0].Value = 0;

            Opc.IRequest OPCRequest;
            ResetWeldCount_Off_Write.Write(OPC_ResetValue_Off, 123, new Opc.Da.WriteCompleteEventHandler(WriteCompleteCallback), out OPCRequest);
        }
       
        private void ResetOff_Timer_Tick(object sender, EventArgs e)
        {
            ConfirmWeldReset_Off_OPC();
            OPCServer.Disconnect();
            ResetOff_Timer.Stop();
            User_Program.UserProgram.Enabled = true;
            this.Close();
        }
    }
}
