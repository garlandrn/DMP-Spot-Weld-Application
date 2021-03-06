﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Opc.Da;

/* 
 * Program: DMP Spot Weld Application
 * Form: User Program Select Operation
 * Created By: Ryan Garland
 * Last Updated on 8/30/18 
 */

namespace DMP_Spot_Weld_Application
{
    public partial class User_Program_Select_Operation : Form
    {
        User_Program Owner = null;

        public User_Program_Select_Operation(User_Program Owner)
        {
            InitializeComponent();
            this.Owner = Owner;
            Operation_1_Button.DialogResult = DialogResult.Yes;
            Operation_2_Button.DialogResult = DialogResult.Yes;
            Operation_3_Button.DialogResult = DialogResult.Yes;
            Operation_4_Button.DialogResult = DialogResult.Yes;
        }

        public int SelectedOperation;

        // Connect to OPC Server on Form Load
        private Opc.Da.Server OPCServer;
        private OpcCom.Factory OPCFactory = new OpcCom.Factory();

        // OPC Write
        private Opc.Da.Subscription OperationSelection_Write;
        private Opc.Da.SubscriptionState OperationSelection_StateWrite;

        // OPC Read 
        private Opc.Da.Subscription OperationSelection_Read;
        private Opc.Da.SubscriptionState OperationSelection_StateRead;

        private static string SpotWeld_Tag_Name = "";

        private void User_Program_Select_Operation_Load(object sender, EventArgs e)
        {
            SpotWeldID();

            // Kepware OPC Connection
            OPCServer = new Opc.Da.Server(OPCFactory, null);
            OPCServer.Url = new Opc.URL("opcda://OHN66OPC/Kepware.KEPServerEX.V6");
            OPCServer.Connect();

            // OPC Write
            OperationSelection_StateWrite = new Opc.Da.SubscriptionState();
            OperationSelection_StateWrite.Name = "PB_OperationSelect_WriteGroup";
            OperationSelection_StateWrite.Active = false;
            OperationSelection_Write = (Opc.Da.Subscription)OPCServer.CreateSubscription(OperationSelection_StateWrite);
            
            // OPC Read
            OperationSelection_StateRead = new Opc.Da.SubscriptionState();
            OperationSelection_StateRead.Name = "153R_Spotweld";
            OperationSelection_StateRead.UpdateRate = 200;
            OperationSelection_StateRead.Active = true;
            OperationSelection_Read = (Opc.Da.Subscription)OPCServer.CreateSubscription(OperationSelection_StateRead);
        }

        // Select Operation 1
        // Send Selected Operation to User Program
        private void Operation_1_Button_Click(object sender, EventArgs e)
        {
            SelectedOperation = 1;
            OperationWriteOPC();
            this.Owner.PassOperationValue(SelectedOperation);
            User_Program.UserProgram.Enabled = true;
            this.Close();
        }

        private void Operation_2_Button_Click(object sender, EventArgs e)
        {
            SelectedOperation = 2;
            OperationWriteOPC();
            this.Owner.PassOperationValue(SelectedOperation);
            User_Program.UserProgram.Enabled = true;
            this.Close();
        }

        private void Operation_3_Button_Click(object sender, EventArgs e)
        {
            SelectedOperation = 3;
            OperationWriteOPC();
            this.Owner.PassOperationValue(SelectedOperation);
            User_Program.UserProgram.Enabled = true;
            this.Close();
        }

        private void Operation_4_Button_Click(object sender, EventArgs e)
        {
            SelectedOperation = 4;
            OperationWriteOPC();
            this.Owner.PassOperationValue(SelectedOperation);
            User_Program.UserProgram.Enabled = true;
            this.Close();
        }

        // Set SpotWeld_Tag_Name
        private void SpotWeldID()
        {
            string SpotWeldComputerID = System.Environment.MachineName;

            // CAT Spot Welders
            if (SpotWeldComputerID == "123R") // CAT - 123R
            {
                SpotWeld_Tag_Name = "OHN66OPC.Spot_Weld_123R.Global.";
            }
            else if (SpotWeldComputerID == "1088") // CAT - 1088
            {
                SpotWeld_Tag_Name = "OHN66OPC.Spot_Weld_1088.Global.";
            }
            // John Deere Spot Welders
            else if (SpotWeldComputerID == "108R") // John Deere - 108R
            {
                SpotWeld_Tag_Name = "OHN66OPC.Spot_Weld_108R.Global.";
            }
            else if (SpotWeldComputerID == "150R") // John Deere - 150R
            {
                SpotWeld_Tag_Name = "OHN66OPC.Spot_Weld_150R.Global.";
            }
            // Navistar Spot Welders
            else if (SpotWeldComputerID == "104R") // Navistar - 104R
            {
                SpotWeld_Tag_Name = "OHN66OPC.Spot_Weld_104R.Global.";
            }
            else if (SpotWeldComputerID == "OHN7149") // Navistar - 121R
            {
                SpotWeld_Tag_Name = "OHN66OPC.Spot_Weld_121R.Global.SW121R_";
            }
            else if (SpotWeldComputerID == "OHN7111") // Navistar - 154R
            {
                SpotWeld_Tag_Name = "OHN66OPC.Spot_Weld_154R.Global.";
            }
            // Paccar Spot Welders
            else if (SpotWeldComputerID == "OHN7124") // Paccar - 153R
            {
                SpotWeld_Tag_Name = "OHN66OPC.Spot_Weld_153R.Global.";
            }
            else if (SpotWeldComputerID == "OHN7123") // Paccar - 155R
            {
                SpotWeld_Tag_Name = "OHN66OPC.Spot_Weld_155R.Global.";
            }
            // My Computer 
            else if (SpotWeldComputerID == "OHN7047NL") //  My Computer
            {
                SpotWeld_Tag_Name = "OHN66OPC.Spot_Weld_121R.Global.SW121R_";
            }
            // Default Computer
            else                                       // Default Computer
            {
                SpotWeld_Tag_Name = "OHN66OPC.Spot_Weld_121R.Global.SW121R_";
            }
        }

        // Write the Selected Operation Value to the PLC
        private void OperationWriteOPC()
        {
            Opc.Da.Item[] OPC_ItemID = new Opc.Da.Item[8];
            OPC_ItemID[0] = new Opc.Da.Item();
            OPC_ItemID[0].ItemName = SpotWeld_Tag_Name + "HMI_Operation_One_PB";
            OPC_ItemID[1] = new Opc.Da.Item();
            OPC_ItemID[1].ItemName = SpotWeld_Tag_Name + "HMI_Operation_Two_PB";
            OPC_ItemID[2] = new Opc.Da.Item();
            OPC_ItemID[2].ItemName = SpotWeld_Tag_Name + "HMI_Operation_Three_PB";
            OPC_ItemID[3] = new Opc.Da.Item();
            OPC_ItemID[3].ItemName = SpotWeld_Tag_Name + "HMI_Operation_Four_PB";

            OPC_ItemID[4] = new Opc.Da.Item();
            OPC_ItemID[4].ItemName = SpotWeld_Tag_Name + "HMI_Operation_One_Selected";
            OPC_ItemID[5] = new Opc.Da.Item();
            OPC_ItemID[5].ItemName = SpotWeld_Tag_Name + "HMI_Operation_Two_Selected";
            OPC_ItemID[6] = new Opc.Da.Item();
            OPC_ItemID[6].ItemName = SpotWeld_Tag_Name + "HMI_Operation_Three_Selected";
            OPC_ItemID[7] = new Opc.Da.Item();
            OPC_ItemID[7].ItemName = SpotWeld_Tag_Name + "HMI_Operation_Four_Selected";
            OPC_ItemID = OperationSelection_Write.AddItems(OPC_ItemID);

            Opc.Da.ItemValue[] ItemID_OPCValue = new Opc.Da.ItemValue[8];
            ItemID_OPCValue[0] = new Opc.Da.ItemValue();
            ItemID_OPCValue[1] = new Opc.Da.ItemValue();
            ItemID_OPCValue[2] = new Opc.Da.ItemValue();
            ItemID_OPCValue[3] = new Opc.Da.ItemValue();
            ItemID_OPCValue[4] = new Opc.Da.ItemValue();
            ItemID_OPCValue[5] = new Opc.Da.ItemValue();
            ItemID_OPCValue[6] = new Opc.Da.ItemValue();
            ItemID_OPCValue[7] = new Opc.Da.ItemValue();

            // switch the values written based on which operation was selected
            switch (SelectedOperation)
            {
                case 1: // Operation #1
                    ItemID_OPCValue[0].ServerHandle = OperationSelection_Write.Items[0].ServerHandle;
                    ItemID_OPCValue[0].Value = 1;
                    ItemID_OPCValue[1].ServerHandle = OperationSelection_Write.Items[1].ServerHandle;
                    ItemID_OPCValue[1].Value = 0;
                    ItemID_OPCValue[2].ServerHandle = OperationSelection_Write.Items[2].ServerHandle;
                    ItemID_OPCValue[2].Value = 0;
                    ItemID_OPCValue[3].ServerHandle = OperationSelection_Write.Items[3].ServerHandle;
                    ItemID_OPCValue[3].Value = 0;

                    ItemID_OPCValue[4].ServerHandle = OperationSelection_Write.Items[4].ServerHandle;
                    ItemID_OPCValue[4].Value = 1;
                    ItemID_OPCValue[5].ServerHandle = OperationSelection_Write.Items[5].ServerHandle;
                    ItemID_OPCValue[5].Value = 0;
                    ItemID_OPCValue[6].ServerHandle = OperationSelection_Write.Items[6].ServerHandle;
                    ItemID_OPCValue[6].Value = 0;
                    ItemID_OPCValue[7].ServerHandle = OperationSelection_Write.Items[7].ServerHandle;
                    ItemID_OPCValue[7].Value = 0;
                    break;

                case 2: // Operation #2
                    ItemID_OPCValue[0].ServerHandle = OperationSelection_Write.Items[0].ServerHandle;
                    ItemID_OPCValue[0].Value = 0;
                    ItemID_OPCValue[1].ServerHandle = OperationSelection_Write.Items[1].ServerHandle;
                    ItemID_OPCValue[1].Value = 1;
                    ItemID_OPCValue[2].ServerHandle = OperationSelection_Write.Items[2].ServerHandle;
                    ItemID_OPCValue[2].Value = 0;
                    ItemID_OPCValue[3].ServerHandle = OperationSelection_Write.Items[3].ServerHandle;
                    ItemID_OPCValue[3].Value = 0;

                    ItemID_OPCValue[4].ServerHandle = OperationSelection_Write.Items[4].ServerHandle;
                    ItemID_OPCValue[4].Value = 0;
                    ItemID_OPCValue[5].ServerHandle = OperationSelection_Write.Items[5].ServerHandle;
                    ItemID_OPCValue[5].Value = 1;
                    ItemID_OPCValue[6].ServerHandle = OperationSelection_Write.Items[6].ServerHandle;
                    ItemID_OPCValue[6].Value = 0;
                    ItemID_OPCValue[7].ServerHandle = OperationSelection_Write.Items[7].ServerHandle;
                    ItemID_OPCValue[7].Value = 0;
                    break;

                case 3: // Operation #3
                    ItemID_OPCValue[0].ServerHandle = OperationSelection_Write.Items[0].ServerHandle;
                    ItemID_OPCValue[0].Value = 0;
                    ItemID_OPCValue[1].ServerHandle = OperationSelection_Write.Items[1].ServerHandle;
                    ItemID_OPCValue[1].Value = 0;
                    ItemID_OPCValue[2].ServerHandle = OperationSelection_Write.Items[2].ServerHandle;
                    ItemID_OPCValue[2].Value = 1;
                    ItemID_OPCValue[3].ServerHandle = OperationSelection_Write.Items[3].ServerHandle;
                    ItemID_OPCValue[3].Value = 0;

                    ItemID_OPCValue[4].ServerHandle = OperationSelection_Write.Items[4].ServerHandle;
                    ItemID_OPCValue[4].Value = 0;
                    ItemID_OPCValue[5].ServerHandle = OperationSelection_Write.Items[5].ServerHandle;
                    ItemID_OPCValue[5].Value = 0;
                    ItemID_OPCValue[6].ServerHandle = OperationSelection_Write.Items[6].ServerHandle;
                    ItemID_OPCValue[6].Value = 1;
                    ItemID_OPCValue[7].ServerHandle = OperationSelection_Write.Items[7].ServerHandle;
                    ItemID_OPCValue[7].Value = 0;
                    break;

                case 4: // Operation #4
                    ItemID_OPCValue[0].ServerHandle = OperationSelection_Write.Items[0].ServerHandle;
                    ItemID_OPCValue[0].Value = 0;
                    ItemID_OPCValue[1].ServerHandle = OperationSelection_Write.Items[1].ServerHandle;
                    ItemID_OPCValue[1].Value = 0;
                    ItemID_OPCValue[2].ServerHandle = OperationSelection_Write.Items[2].ServerHandle;
                    ItemID_OPCValue[2].Value = 0;
                    ItemID_OPCValue[3].ServerHandle = OperationSelection_Write.Items[3].ServerHandle;
                    ItemID_OPCValue[3].Value = 1;

                    ItemID_OPCValue[4].ServerHandle = OperationSelection_Write.Items[4].ServerHandle;
                    ItemID_OPCValue[4].Value = 0;
                    ItemID_OPCValue[5].ServerHandle = OperationSelection_Write.Items[5].ServerHandle;
                    ItemID_OPCValue[5].Value = 0;
                    ItemID_OPCValue[6].ServerHandle = OperationSelection_Write.Items[6].ServerHandle;
                    ItemID_OPCValue[6].Value = 0;
                    ItemID_OPCValue[7].ServerHandle = OperationSelection_Write.Items[7].ServerHandle;
                    ItemID_OPCValue[7].Value = 1;
                    break;
            }
            Opc.IRequest OPCRequest;
            OperationSelection_Write.Write(ItemID_OPCValue, 123, new Opc.Da.WriteCompleteEventHandler(WriteCompleteCallback), out OPCRequest);
        }

        private void WriteCompleteCallback(object clientHandle, Opc.IdentifiedResult[] results)
        {
            foreach (Opc.IdentifiedResult writeResult in results)
            {
                Console.WriteLine("\t{0} write result: {1}", writeResult.ItemName, writeResult.ResultID);
            }
        }

        // OPC Server Disconnects When the Form is Closing
        private void User_Program_Select_Operation_FormClosing(object sender, FormClosingEventArgs e)
        {
            OPCServer.Disconnect();
        }
    }
}
