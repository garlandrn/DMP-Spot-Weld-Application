using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace DMP_Spot_Weld_Application
{
    public partial class User_Program_Part_Not_Completed : Form
    {
        public User_Program_Part_Not_Completed()
        {
            InitializeComponent();
            OK_Button.DialogResult = DialogResult.Yes;
            this.ShowInTaskbar = false;
        }

        private Opc.URL OPCUrl;
        private Opc.Da.Server OPCServer;
        private OpcCom.Factory OPCFactory = new OpcCom.Factory();
        private Opc.Da.Subscription Fault_Off_Write;
        private Opc.Da.SubscriptionState Fault_Off_StateWrite;
        private Opc.Da.Subscription Fault_On_Write;
        private Opc.Da.SubscriptionState Fault_On_StateWrite;
        private static string Spotweld_Tag_Name = "";

        private void OK_Button_Click(object sender, EventArgs e)
        {
            ConfirmFaultReset_Off_OPC();
            //OPCServer.Disconnect();
            //User_Program.UserProgram.Enabled = true;            
            //this.Close();
        }

        private void ConfirmFaultReset_Off_OPC()
        {
            Opc.Da.Item[] OPC_Fault_Off = new Opc.Da.Item[1];
            OPC_Fault_Off[0] = new Opc.Da.Item();
            OPC_Fault_Off[0].ItemName = Spotweld_Tag_Name + "HMI_PB_Alarm_Reset";
            OPC_Fault_Off = Fault_Off_Write.AddItems(OPC_Fault_Off);

            Opc.Da.ItemValue[] OPC_ResetValue_Off = new Opc.Da.ItemValue[1];
            OPC_ResetValue_Off[0] = new Opc.Da.ItemValue();
            OPC_ResetValue_Off[0].ServerHandle = Fault_Off_Write.Items[0].ServerHandle;
            OPC_ResetValue_Off[0].Value = 1;

            Opc.IRequest OPCRequest;
            Fault_Off_Write.Write(OPC_ResetValue_Off, 123, new Opc.Da.WriteCompleteEventHandler(WriteCompleteCallback), out OPCRequest);
            Reset_Timer.Start();
        }

        private void ConfirmFaultReset_On_OPC()
        {
            Opc.Da.Item[] OPC_Fault_On = new Opc.Da.Item[1];
            OPC_Fault_On[0] = new Opc.Da.Item();
            OPC_Fault_On[0].ItemName = Spotweld_Tag_Name + "HMI_PB_Alarm_Reset";
            OPC_Fault_On = Fault_On_Write.AddItems(OPC_Fault_On);

            Opc.Da.ItemValue[] OPC_ResetValue_On = new Opc.Da.ItemValue[1];
            OPC_ResetValue_On[0] = new Opc.Da.ItemValue();
            OPC_ResetValue_On[0].ServerHandle = Fault_On_Write.Items[0].ServerHandle;
            OPC_ResetValue_On[0].Value = 0;

            Opc.IRequest OPCRequest;
            Fault_On_Write.Write(OPC_ResetValue_On, 123, new Opc.Da.WriteCompleteEventHandler(WriteCompleteCallback), out OPCRequest);
        }

        private void WriteCompleteCallback(object clientHandle, Opc.IdentifiedResult[] results)
        {
            foreach (Opc.IdentifiedResult writeResult in results)
            {
                Console.WriteLine("\t{0} write result: {1}", writeResult.ItemName, writeResult.ResultID);
            }
        }

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

        private void Reset_Timer_Tick(object sender, EventArgs e)
        {
            ConfirmFaultReset_On_OPC();
            OPCServer.Disconnect();
            Reset_Timer.Stop();
            User_Program.UserProgram.Enabled = true;
            this.Close();
        }
    }    
}
