using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace DMP_Spot_Weld_Application
{
    public partial class TestForm : Form
    {
        public TestForm()
        {
            InitializeComponent();
        }

        private Opc.URL OPCUrl;
        private Opc.Da.Server OPCServer;
        private OpcCom.Factory OPCFactory = new OpcCom.Factory();
        private Opc.Da.Subscription Fault_Off_Write;
        private Opc.Da.SubscriptionState Fault_Off_StateWrite;
        private Opc.Da.Subscription Fault_On_Write;
        private Opc.Da.SubscriptionState Fault_On_StateWrite;
        private static string Spotweld_Tag_Name = "";

        TimeSpan TimeOfOperation = new TimeSpan();

        private void textBox1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                int error;
                string s = textBox1.Text;
                //Regex r = new Regex("[a-zA-Z]|$");
                //bool containsAny = r.IsMatch(s);
                //listBox1.Items.Add(containsAny.ToString());

                error = Regex.Matches(s, @"[\D]").Count;
                listBox1.Items.Add(error.ToString());
                TimeOfOperation = DateTime.Parse("10:00:00 PM").Subtract(DateTime.Parse("11:00:00 AM"));
                listBox1.Items.Add("Total Time: " + TimeOfOperation);
                listBox1.Items.Add("TimeOfOperation.TotalMinutes: " + TimeOfOperation.TotalMinutes);
                listBox1.Items.Add("");
            }
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
            timer1.Start();
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
            if (SpotWeldComputerID == "1088") // CAT - 1088
            {
                Spotweld_Tag_Name = "OHN66OPC.Spot_Weld_1088.Global.";
            }
            else if (SpotWeldComputerID == "?") // CAT - ?
            {
                Spotweld_Tag_Name = "OHN66OPC.Spot_Weld_?.Global.";
            }
            // John Deere Spot Weld
            else if (SpotWeldComputerID == "OHN?") // John Deere 1
            {
                Spotweld_Tag_Name = "OHN66OPC.Spot_Weld_?.Global.";
            }
            else if (SpotWeldComputerID == "OHN?") // John Deere 2
            {
                Spotweld_Tag_Name = "OHN66OPC.Spot_Weld_?.Global.";
            }
            // Navistar
            else if (SpotWeldComputerID == "104R") // Navistar - 104R
            {
                Spotweld_Tag_Name = "OHN66OPC.Spot_Weld_154R.Global.";
            }
            else if (SpotWeldComputerID == "121R") // Navistar - 121R
            {
                Spotweld_Tag_Name = "OHN66OPC.Spot_Weld_154R.Global.";
            }
            else if (SpotWeldComputerID == "OHN7111") // Navistar - 154R
            {
                Spotweld_Tag_Name = "OHN66OPC.Spot_Weld_154R.Global.";
            }
            // Paccar
            else if (SpotWeldComputerID == "OHN7124") // Paccar - 153R
            {
                Spotweld_Tag_Name = "OHN66OPC.Spot_Weld_153R.Global.";
            }
            else if (SpotWeldComputerID == "OHN7123") // Paccar - 155R
            {
                Spotweld_Tag_Name = "OHN66OPC.Spot_Weld_155R.Global.";
            }
            else if (SpotWeldComputerID == "OHN7047NL") //  My Laptop
            {
                Spotweld_Tag_Name = "OHN66OPC.Spot_Weld_154R.Global.";
            }
        }

        private void TestForm_Load(object sender, EventArgs e)
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

        private void OK_Button_Click(object sender, EventArgs e)
        {
            ConfirmFaultReset_Off_OPC();
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            ConfirmFaultReset_On_OPC();
            timer1.Stop();
            this.Close();
        }
    }
}
