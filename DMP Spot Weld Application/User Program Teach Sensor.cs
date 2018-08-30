
using System;
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
 * Form: User Program Teach Sensor
 * Created By: Ryan Garland
 * Last Updated on 8/28/18
 * 
 */

namespace DMP_Spot_Weld_Application
{
    public partial class User_Program_Teach_Senson : Form
    {
        User_Program owner = null;
        public User_Program_Teach_Senson(User_Program owner)
        {
            InitializeComponent();
            this.ShowInTaskbar = false;
            this.owner = owner;
            OK_Button.DialogResult = DialogResult.Yes;
        }
        
        // OPC Server
        private Opc.Da.Server OPCServer;
        private OpcCom.Factory OPCFactory = new OpcCom.Factory();

        // Read Sensor Subscription
        private Opc.Da.Subscription Sensor_Read;
        private Opc.Da.SubscriptionState Sensor_StateRead;
        
        // Write To Teach Input
        private Opc.Da.Subscription TeachSensor_Write;
        private Opc.Da.SubscriptionState TeachSensor_StateWrite;

        // Component Write Groups
        private Opc.Da.Subscription GroupWriteCompOne;
        private Opc.Da.SubscriptionState GroupStateWriteCompOne;
        private Opc.Da.Subscription GroupWriteCompTwo;
        private Opc.Da.SubscriptionState GroupStateWriteCompTwo;
        private Opc.Da.Subscription GroupWriteCompThree;
        private Opc.Da.SubscriptionState GroupStateWriteCompThree;

        // OPC Item Lists
        private List<Item> OPC_ItemsList = new List<Item>();
        private List<Item> OPC_WriteSensor = new List<Item>();

        // Tag String
        private static string Spotweld_Tag_Name = "";

        // Variable for Input Value
        private static int OPC_TeachSensorValue;
        
        private static string HMI_Probe_Value = "";
        private static string Target_Travel_Value = "";
        private static string Multi_Component_Target2 = "";
        private static string Multi_Component_Target3 = "";
        private static string Multi_Component_Target4 = "";
        private static string Target_Travel_Value_Set = "";
        private static string Sequence_One_Set = "";
        private static string Sequence_Two_Set = "";
        private static string Sequence_Three_Set = "";

        private void User_Program_Teach_Senson_Load(object sender, EventArgs e)
        {
            SpotWeldID(); // Get Computer ID and Set Tag Name

            // Connect to OPC Server
            OPCServer = new Opc.Da.Server(OPCFactory, null);
            OPCServer.Url = new Opc.URL("opcda://OHN66OPC/Kepware.KEPServerEX.V6");
            OPCServer.Connect();

            Sensor_StateRead = new Opc.Da.SubscriptionState();
            Sensor_StateRead.Name = "Teach_Sensor_OPC";
            Sensor_StateRead.UpdateRate = 1000;
            Sensor_StateRead.Active = true;
            Sensor_Read = (Opc.Da.Subscription)OPCServer.CreateSubscription(Sensor_StateRead);
            Sensor_Read.DataChanged += new Opc.Da.DataChangedEventHandler(Sensor_Read_DataChanged);

            TeachSensor_StateWrite = new Opc.Da.SubscriptionState();
            TeachSensor_StateWrite.Name = "OPCWriteGroup";
            TeachSensor_StateWrite.Active = false;
            TeachSensor_Write = (Opc.Da.Subscription)OPCServer.CreateSubscription(TeachSensor_StateWrite);

            GroupStateWriteCompOne = new Opc.Da.SubscriptionState();
            GroupStateWriteCompOne.Name = "OPCWriteOneGroup";
            GroupStateWriteCompOne.Active = false;
            GroupWriteCompOne = (Opc.Da.Subscription)OPCServer.CreateSubscription(GroupStateWriteCompOne);

            GroupStateWriteCompTwo = new Opc.Da.SubscriptionState();
            GroupStateWriteCompTwo.Name = "OPCWriteTwoGroup";
            GroupStateWriteCompTwo.Active = false;
            GroupWriteCompTwo = (Opc.Da.Subscription)OPCServer.CreateSubscription(GroupStateWriteCompTwo);

            GroupStateWriteCompThree = new Opc.Da.SubscriptionState();
            GroupStateWriteCompThree.Name = "OPCWriteThreeGroup";
            GroupStateWriteCompThree.Active = false;
            GroupWriteCompThree = (Opc.Da.Subscription)OPCServer.CreateSubscription(GroupStateWriteCompThree);
            
            // Get The Component ID From the User Program
            Component_1_TextBox.Text = owner.Comp1_TextBox.Text;
            Component_2_TextBox.Text = owner.Comp2_TextBox.Text;
            Component_3_TextBox.Text = owner.Comp3_TextBox.Text;

            CheckComponentValue(); // Find the Number of Components
            //TargetValueSet_TextBox.Clear();
            OPC_TeachSensorValue = 1;
            TeachSensor_OPC();
            SensorRead_OPC();
            WriteComponentOne_OPC();
        }

        /*********************************************************************************************************************
        * 
        * Buttons Region Start
        * -- Total: 5
        * 
        * - OK Button Click
        * - Cancel Button Click
        * - Component 1 Button Click
        * - Component 2 Button Click
        * - Component 3 Button Click
        * 
        *********************************************************************************************************************/
        #region

        // Turn Sensor Input off and Close
        private void OK_Button_Click(object sender, EventArgs e)
        {
            OPC_TeachSensorValue = 0;
            TeachSensor_OPC();
            User_Program.UserProgram.Enabled = true;
            this.Close();
        }

        // Turn Sensor Input off and Close
        private void Cancel_Button_Click(object sender, EventArgs e)
        {
            OPC_TeachSensorValue = 0;
            TeachSensor_OPC();
            User_Program.UserProgram.Enabled = true;
            this.Dispose();
        }

        private void Component_1_Button_Click(object sender, EventArgs e)
        {
            WriteComponentOne_OPC();
        }

        private void Component_2_Button_Click(object sender, EventArgs e)
        {
            WriteComponentTwo_OPC();
        }

        private void Component_3_Button_Click(object sender, EventArgs e)
        {
            WriteComponentThree_OPC();
        }

        /*********************************************************************************************************************
        * 
        * Buttons Region End
        * 
        *********************************************************************************************************************/
        #endregion

        /********************************************************************************************************************
        *  
        *  OPC Region Start
        *  -- Total: 
        *  
        ********************************************************************************************************************/
        #region

        // Write Both On and Off Using OPC_TeachSensorValue
        private void TeachSensor_OPC()
        {
            Opc.Da.Item[] OPC_SensorWrite = new Opc.Da.Item[1];
            OPC_SensorWrite[0] = new Opc.Da.Item();
            OPC_SensorWrite[0].ItemName = Spotweld_Tag_Name + "HMI_PB_TEACH_SENSOR";
            OPC_WriteSensor.Add(OPC_SensorWrite[0]);
            TeachSensor_Write.AddItems(OPC_WriteSensor.ToArray());

            Opc.Da.ItemValue[] WriteValue = new Opc.Da.ItemValue[1];
            WriteValue[0] = new Opc.Da.ItemValue();
            WriteValue[0].ServerHandle = TeachSensor_Write.Items[0].ServerHandle;
            WriteValue[0].Value = OPC_TeachSensorValue;

            Opc.IRequest req;
            TeachSensor_Write.Write(WriteValue, 123, new Opc.Da.WriteCompleteEventHandler(WriteCompleteCallback), out req);
        }

        private void WriteComponentOne_OPC()
        {
            Opc.Da.Item[] OPC_CompOneWrite = new Opc.Da.Item[4];
            OPC_CompOneWrite[0] = new Opc.Da.Item();
            OPC_CompOneWrite[0].ItemName = Spotweld_Tag_Name + "HMI_Operation_One_PB";
            OPC_CompOneWrite[1] = new Opc.Da.Item();
            OPC_CompOneWrite[1].ItemName = Spotweld_Tag_Name + "HMI_Operation_Two_PB";
            OPC_CompOneWrite[2] = new Opc.Da.Item();
            OPC_CompOneWrite[2].ItemName = Spotweld_Tag_Name + "HMI_Operation_Three_PB";
            OPC_CompOneWrite[3] = new Opc.Da.Item();
            OPC_CompOneWrite[3].ItemName = Spotweld_Tag_Name + "HMI_Operation_Four_PB";
            OPC_CompOneWrite = GroupWriteCompOne.AddItems(OPC_CompOneWrite);


            Opc.Da.ItemValue[] WriteCompOneValue = new Opc.Da.ItemValue[4];
            WriteCompOneValue[0] = new Opc.Da.ItemValue();
            WriteCompOneValue[0].ServerHandle = GroupWriteCompOne.Items[0].ServerHandle;
            WriteCompOneValue[0].Value = 1;
            WriteCompOneValue[1] = new Opc.Da.ItemValue();
            WriteCompOneValue[1].ServerHandle = GroupWriteCompOne.Items[1].ServerHandle;
            WriteCompOneValue[1].Value = 0;
            WriteCompOneValue[2] = new Opc.Da.ItemValue();
            WriteCompOneValue[2].ServerHandle = GroupWriteCompOne.Items[2].ServerHandle;
            WriteCompOneValue[2].Value = 0;
            WriteCompOneValue[3] = new Opc.Da.ItemValue();
            WriteCompOneValue[3].ServerHandle = GroupWriteCompOne.Items[3].ServerHandle;
            WriteCompOneValue[3].Value = 0;

            Opc.IRequest req;
            GroupWriteCompOne.Write(WriteCompOneValue, 123, new Opc.Da.WriteCompleteEventHandler(WriteCompleteCallback), out req);
        }

        private void WriteComponentTwo_OPC()
        {
            Opc.Da.Item[] OPC_CompWrite = new Opc.Da.Item[4];
            OPC_CompWrite[0] = new Opc.Da.Item();
            OPC_CompWrite[0].ItemName = Spotweld_Tag_Name + "HMI_Operation_One_PB";
            OPC_CompWrite[1] = new Opc.Da.Item();
            OPC_CompWrite[1].ItemName = Spotweld_Tag_Name + "HMI_Operation_Two_PB";
            OPC_CompWrite[2] = new Opc.Da.Item();
            OPC_CompWrite[2].ItemName = Spotweld_Tag_Name + "HMI_Operation_Three_PB";
            OPC_CompWrite[3] = new Opc.Da.Item();
            OPC_CompWrite[3].ItemName = Spotweld_Tag_Name + "HMI_Operation_Four_PB";
            OPC_CompWrite = GroupWriteCompTwo.AddItems(OPC_CompWrite);

            Opc.Da.ItemValue[] WriteCompTwoValue = new Opc.Da.ItemValue[4];
            WriteCompTwoValue[0] = new Opc.Da.ItemValue();
            WriteCompTwoValue[0].ServerHandle = GroupWriteCompTwo.Items[0].ServerHandle;
            WriteCompTwoValue[0].Value = 0;
            WriteCompTwoValue[1] = new Opc.Da.ItemValue();
            WriteCompTwoValue[1].ServerHandle = GroupWriteCompTwo.Items[1].ServerHandle;
            WriteCompTwoValue[1].Value = 1;
            WriteCompTwoValue[2] = new Opc.Da.ItemValue();
            WriteCompTwoValue[2].ServerHandle = GroupWriteCompTwo.Items[2].ServerHandle;
            WriteCompTwoValue[2].Value = 0;
            WriteCompTwoValue[3] = new Opc.Da.ItemValue();
            WriteCompTwoValue[3].ServerHandle = GroupWriteCompTwo.Items[3].ServerHandle;
            WriteCompTwoValue[3].Value = 0;

            Opc.IRequest req;
            GroupWriteCompTwo.Write(WriteCompTwoValue, 123, new Opc.Da.WriteCompleteEventHandler(WriteCompleteCallback), out req);
        }

        private void WriteComponentThree_OPC()
        {
            Opc.Da.Item[] OPC_CompThreeWrite = new Opc.Da.Item[4];
            OPC_CompThreeWrite[0] = new Opc.Da.Item();
            OPC_CompThreeWrite[0].ItemName = Spotweld_Tag_Name + "HMI_Operation_One_PB";
            OPC_CompThreeWrite[1] = new Opc.Da.Item();
            OPC_CompThreeWrite[1].ItemName = Spotweld_Tag_Name + "HMI_Operation_Two_PB";
            OPC_CompThreeWrite[2] = new Opc.Da.Item();
            OPC_CompThreeWrite[2].ItemName = Spotweld_Tag_Name + "HMI_Operation_Three_PB";
            OPC_CompThreeWrite[3] = new Opc.Da.Item();
            OPC_CompThreeWrite[3].ItemName = Spotweld_Tag_Name + "HMI_Operation_Four_PB";
            OPC_CompThreeWrite = GroupWriteCompThree.AddItems(OPC_CompThreeWrite);

            Opc.Da.ItemValue[] WriteCompThreeValue = new Opc.Da.ItemValue[4];
            WriteCompThreeValue[0] = new Opc.Da.ItemValue();
            WriteCompThreeValue[0].ServerHandle = GroupWriteCompThree.Items[0].ServerHandle;
            WriteCompThreeValue[0].Value = 0;
            WriteCompThreeValue[1] = new Opc.Da.ItemValue();
            WriteCompThreeValue[1].ServerHandle = GroupWriteCompThree.Items[1].ServerHandle;
            WriteCompThreeValue[1].Value = 0;
            WriteCompThreeValue[2] = new Opc.Da.ItemValue();
            WriteCompThreeValue[2].ServerHandle = GroupWriteCompThree.Items[2].ServerHandle;
            WriteCompThreeValue[2].Value = 1;
            WriteCompThreeValue[3] = new Opc.Da.ItemValue();
            WriteCompThreeValue[3].ServerHandle = GroupWriteCompThree.Items[3].ServerHandle;
            WriteCompThreeValue[3].Value = 0;

            Opc.IRequest req;
            GroupWriteCompThree.Write(WriteCompThreeValue, 123, new Opc.Da.WriteCompleteEventHandler(WriteCompleteCallback), out req);
        }

        // Add Items To Group Read 
        private void SensorRead_OPC()
        {
            List<Item> OPC_ItemsList = new List<Item>();
            Opc.Da.Item[] OPC_ItemID = new Opc.Da.Item[9];
            OPC_ItemID[0] = new Opc.Da.Item();
            OPC_ItemID[0].ItemName = Spotweld_Tag_Name + "HMI_PROBE_VALUE";
            OPC_ItemsList.Add(OPC_ItemID[0]);
            OPC_ItemID[1] = new Opc.Da.Item();
            OPC_ItemID[1].ItemName = Spotweld_Tag_Name + "TARGET_TRAVEL_VALUE";
            OPC_ItemsList.Add(OPC_ItemID[1]);
            OPC_ItemID[2] = new Opc.Da.Item();
            OPC_ItemID[2].ItemName = Spotweld_Tag_Name + "MULTI_COMPONENT_TARGET2";
            OPC_ItemsList.Add(OPC_ItemID[2]);
            OPC_ItemID[3] = new Opc.Da.Item();
            OPC_ItemID[3].ItemName = Spotweld_Tag_Name + "MULTI_COMPONENT_TARGET3";
            OPC_ItemsList.Add(OPC_ItemID[3]);
            OPC_ItemID[4] = new Opc.Da.Item();
            OPC_ItemID[4].ItemName = Spotweld_Tag_Name + "MULTI_COMPONENT_TARGET4";
            OPC_ItemsList.Add(OPC_ItemID[4]);
            OPC_ItemID[5] = new Opc.Da.Item();
            OPC_ItemID[5].ItemName = Spotweld_Tag_Name + "TARGET_TRAVEL_VALUE_SET";
            OPC_ItemsList.Add(OPC_ItemID[5]);
            OPC_ItemID[6] = new Opc.Da.Item();
            OPC_ItemID[6].ItemName = Spotweld_Tag_Name + "SEQUENCE_ONE_SET";
            OPC_ItemsList.Add(OPC_ItemID[6]);
            OPC_ItemID[7] = new Opc.Da.Item();
            OPC_ItemID[7].ItemName = Spotweld_Tag_Name + "SEQUENCE_TWO_SET";
            OPC_ItemsList.Add(OPC_ItemID[7]);
            OPC_ItemID[8] = new Opc.Da.Item();
            OPC_ItemID[8].ItemName = Spotweld_Tag_Name + "SEQUENCE_THREE_SET";
            OPC_ItemsList.Add(OPC_ItemID[8]);
            Sensor_Read.AddItems(OPC_ItemsList.ToArray());
        }
        
        // Read OPC Values When Data Changes
        // Tag Names are Determined by PC Name
        public void Sensor_Read_DataChanged(object subscriptionHandle, object requestHandle, ItemValueResult[] values)
        {
            // CAT Spot Welders
            if (System.Environment.MachineName == "123R") // CAT - 123R
            {
                foreach (ItemValueResult itemValue in values)
                {
                    switch (itemValue.ItemName)
                    {
                        case "OHN66OPC.Spot_Weld_123R.Global.HMI_PROBE_VALUE":
                            HMI_Probe_Value = Convert.ToString(itemValue.Value);
                            break;

                        case "OHN66OPC.Spot_Weld_123R.Global.TARGET_TRAVEL_VALUE":
                            Target_Travel_Value = Convert.ToString(itemValue.Value);
                            break;

                        case "OHN66OPC.Spot_Weld_123R.Global.MULTI_COMPONENT_TARGET2":
                            Multi_Component_Target2 = Convert.ToString(itemValue.Value);
                            break;

                        case "OHN66OPC.Spot_Weld_123R.Global.MULTI_COMPONENT_TARGET3":
                            Multi_Component_Target3 = Convert.ToString(itemValue.Value);
                            break;

                        case "OHN66OPC.Spot_Weld_123R.Global.MULTI_COMPONENT_TARGET4":
                            Multi_Component_Target4 = Convert.ToString(itemValue.Value);
                            break;

                        case "OHN66OPC.Spot_Weld_123R.Global.TARGET_TRAVEL_VALUE_SET":
                            Target_Travel_Value_Set = Convert.ToString(itemValue.Value);
                            break;

                        case "OHN66OPC.Spot_Weld_123R.Global.SEQUENCE_TWO_SET":
                            Sequence_Two_Set = Convert.ToString(itemValue.Value);
                            break;

                        case "OHN66OPC.Spot_Weld_123R.Global.SEQUENCE_THREE_SET":
                            Sequence_Three_Set = Convert.ToString(itemValue.Value);
                            break;
                    }
                }
            }
            else if (System.Environment.MachineName == "1088")  // CAT - 1088
            {
                foreach (ItemValueResult itemValue in values)
                {
                    switch (itemValue.ItemName)
                    {
                        case "OHN66OPC.Spot_Weld_1088.Global.HMI_PROBE_VALUE":
                            HMI_Probe_Value = Convert.ToString(itemValue.Value);
                            break;

                        case "OHN66OPC.Spot_Weld_1088.Global.TARGET_TRAVEL_VALUE":
                            Target_Travel_Value = Convert.ToString(itemValue.Value);
                            break;

                        case "OHN66OPC.Spot_Weld_1088.Global.MULTI_COMPONENT_TARGET2":
                            Multi_Component_Target2 = Convert.ToString(itemValue.Value);
                            break;

                        case "OHN66OPC.Spot_Weld_1088.Global.MULTI_COMPONENT_TARGET3":
                            Multi_Component_Target3 = Convert.ToString(itemValue.Value);
                            break;

                        case "OHN66OPC.Spot_Weld_1088.Global.MULTI_COMPONENT_TARGET4":
                            Multi_Component_Target4 = Convert.ToString(itemValue.Value);
                            break;

                        case "OHN66OPC.Spot_Weld_1088.Global.TARGET_TRAVEL_VALUE_SET":
                            Target_Travel_Value_Set = Convert.ToString(itemValue.Value);
                            break;

                        case "OHN66OPC.Spot_Weld_1088.Global.SEQUENCE_TWO_SET":
                            Sequence_Two_Set = Convert.ToString(itemValue.Value);
                            break;

                        case "OHN66OPC.Spot_Weld_1088.Global.SEQUENCE_THREE_SET":
                            Sequence_Three_Set = Convert.ToString(itemValue.Value);
                            break;
                    }
                }
            }
            // John Deere Spot Welders
            else if (System.Environment.MachineName == "108R")  // John Deere - 108R
            {
                foreach (ItemValueResult itemValue in values)
                {
                    switch (itemValue.ItemName)
                    {
                        case "OHN66OPC.Spot_Weld_108R.Global.HMI_PROBE_VALUE":
                            HMI_Probe_Value = Convert.ToString(itemValue.Value);
                            break;

                        case "OHN66OPC.Spot_Weld_108R.Global.TARGET_TRAVEL_VALUE":
                            Target_Travel_Value = Convert.ToString(itemValue.Value);
                            break;

                        case "OHN66OPC.Spot_Weld_108R.Global.MULTI_COMPONENT_TARGET2":
                            Multi_Component_Target2 = Convert.ToString(itemValue.Value);
                            break;

                        case "OHN66OPC.Spot_Weld_108R.Global.MULTI_COMPONENT_TARGET3":
                            Multi_Component_Target3 = Convert.ToString(itemValue.Value);
                            break;

                        case "OHN66OPC.Spot_Weld_108R.Global.MULTI_COMPONENT_TARGET4":
                            Multi_Component_Target4 = Convert.ToString(itemValue.Value);
                            break;

                        case "OHN66OPC.Spot_Weld_108R.Global.TARGET_TRAVEL_VALUE_SET":
                            Target_Travel_Value_Set = Convert.ToString(itemValue.Value);
                            break;

                        case "OHN66OPC.Spot_Weld_108R.Global.SEQUENCE_TWO_SET":
                            Sequence_Two_Set = Convert.ToString(itemValue.Value);
                            break;

                        case "OHN66OPC.Spot_Weld_108R.Global.SEQUENCE_THREE_SET":
                            Sequence_Three_Set = Convert.ToString(itemValue.Value);
                            break;
                    }
                }
            }
            else if (System.Environment.MachineName == "150R") // John Deere - 150R
            {
                foreach (ItemValueResult itemValue in values)
                {
                    switch (itemValue.ItemName)
                    {
                        case "OHN66OPC.Spot_Weld_150R.Global.HMI_PROBE_VALUE":
                            HMI_Probe_Value = Convert.ToString(itemValue.Value);
                            break;

                        case "OHN66OPC.Spot_Weld_150R.Global.TARGET_TRAVEL_VALUE":
                            Target_Travel_Value = Convert.ToString(itemValue.Value);
                            break;

                        case "OHN66OPC.Spot_Weld_150R.Global.MULTI_COMPONENT_TARGET2":
                            Multi_Component_Target2 = Convert.ToString(itemValue.Value);
                            break;

                        case "OHN66OPC.Spot_Weld_150R.Global.MULTI_COMPONENT_TARGET3":
                            Multi_Component_Target3 = Convert.ToString(itemValue.Value);
                            break;

                        case "OHN66OPC.Spot_Weld_150R.Global.MULTI_COMPONENT_TARGET4":
                            Multi_Component_Target4 = Convert.ToString(itemValue.Value);
                            break;

                        case "OHN66OPC.Spot_Weld_150R.Global.TARGET_TRAVEL_VALUE_SET":
                            Target_Travel_Value_Set = Convert.ToString(itemValue.Value);
                            break;

                        case "OHN66OPC.Spot_Weld_150R.Global.SEQUENCE_TWO_SET":
                            Sequence_Two_Set = Convert.ToString(itemValue.Value);
                            break;

                        case "OHN66OPC.Spot_Weld_150R.Global.SEQUENCE_THREE_SET":
                            Sequence_Three_Set = Convert.ToString(itemValue.Value);
                            break;
                    }
                }
            }
            // Navistar Spot Welders
            else if (System.Environment.MachineName == "OHN7149") // Navistar - 121R
            {
                foreach (ItemValueResult itemValue in values)
                {
                    switch (itemValue.ItemName)
                    {

                        case "OHN66OPC.Spot_Weld_121R.Global.SW121R_HMI_PROBE_VALUE":
                            HMI_Probe_Value = Convert.ToString(itemValue.Value);
                            break;

                        case "OHN66OPC.Spot_Weld_121R.Global.SW121R_TARGET_TRAVEL_VALUE":
                            Target_Travel_Value = Convert.ToString(itemValue.Value);
                            break;

                        case "OHN66OPC.Spot_Weld_121R.Global.SW121R_MULTI_COMPONENT_TARGET2":
                            Multi_Component_Target2 = Convert.ToString(itemValue.Value);
                            break;

                        case "OHN66OPC.Spot_Weld_121R.Global.SW121R_MULTI_COMPONENT_TARGET3":
                            Multi_Component_Target3 = Convert.ToString(itemValue.Value);
                            break;

                        case "OHN66OPC.Spot_Weld_121R.Global.SW121R_MULTI_COMPONENT_TARGET4":
                            Multi_Component_Target4 = Convert.ToString(itemValue.Value);
                            break;

                        case "OHN66OPC.Spot_Weld_121R.Global.SW121R_TARGET_TRAVEL_VALUE_SET":
                            Target_Travel_Value_Set = Convert.ToString(itemValue.Value);
                            break;

                        case "OHN66OPC.Spot_Weld_121R.Global.SW121R_SEQUENCE_TWO_SET":
                            Sequence_Two_Set = Convert.ToString(itemValue.Value);
                            break;

                        case "OHN66OPC.Spot_Weld_121R.Global.SW121R_SEQUENCE_THREE_SET":
                            Sequence_Three_Set = Convert.ToString(itemValue.Value);
                            break;
                    }
                }
            }
            else if (System.Environment.MachineName == "OHN7111") // Navistar - 154R
            {
                foreach (ItemValueResult itemValue in values)
                {
                    switch (itemValue.ItemName)
                    {
                        case "OHN66OPC.Spot_Weld_154R.Global.HMI_PROBE_VALUE":
                            HMI_Probe_Value = Convert.ToString(itemValue.Value);
                            break;

                        case "OHN66OPC.Spot_Weld_154R.Global.TARGET_TRAVEL_VALUE":
                            Target_Travel_Value = Convert.ToString(itemValue.Value);
                            break;

                        case "OHN66OPC.Spot_Weld_154R.Global.MULTI_COMPONENT_TARGET2":
                            Multi_Component_Target2 = Convert.ToString(itemValue.Value);
                            break;

                        case "OHN66OPC.Spot_Weld_154R.Global.MULTI_COMPONENT_TARGET3":
                            Multi_Component_Target3 = Convert.ToString(itemValue.Value);
                            break;

                        case "OHN66OPC.Spot_Weld_154R.Global.MULTI_COMPONENT_TARGET4":
                            Multi_Component_Target4 = Convert.ToString(itemValue.Value);
                            break;

                        case "OHN66OPC.Spot_Weld_154R.Global.TARGET_TRAVEL_VALUE_SET":
                            Target_Travel_Value_Set = Convert.ToString(itemValue.Value);
                            break;

                        case "OHN66OPC.Spot_Weld_154R.Global.SEQUENCE_TWO_SET":
                            Sequence_Two_Set = Convert.ToString(itemValue.Value);
                            break;

                        case "OHN66OPC.Spot_Weld_154R.Global.SEQUENCE_THREE_SET":
                            Sequence_Three_Set = Convert.ToString(itemValue.Value);
                            break;
                    }
                }
            }
            // Paccar Spot Welders
            else if (System.Environment.MachineName == "OHN7124")  // Paccar - 153R
            {
                foreach (ItemValueResult itemValue in values)
                {
                    switch (itemValue.ItemName)
                    {
                        case "OHN66OPC.Spot_Weld_153R.Global.HMI_PROBE_VALUE":
                            HMI_Probe_Value = Convert.ToString(itemValue.Value);
                            break;

                        case "OHN66OPC.Spot_Weld_153R.Global.TARGET_TRAVEL_VALUE":
                            Target_Travel_Value = Convert.ToString(itemValue.Value);
                            break;

                        case "OHN66OPC.Spot_Weld_153R.Global.MULTI_COMPONENT_TARGET2":
                            Multi_Component_Target2 = Convert.ToString(itemValue.Value);
                            break;

                        case "OHN66OPC.Spot_Weld_153R.Global.MULTI_COMPONENT_TARGET3":
                            Multi_Component_Target3 = Convert.ToString(itemValue.Value);
                            break;

                        case "OHN66OPC.Spot_Weld_153R.Global.MULTI_COMPONENT_TARGET4":
                            Multi_Component_Target4 = Convert.ToString(itemValue.Value);
                            break;

                        case "OHN66OPC.Spot_Weld_153R.Global.TARGET_TRAVEL_VALUE_SET":
                            Target_Travel_Value_Set = Convert.ToString(itemValue.Value);
                            break;

                        case "OHN66OPC.Spot_Weld_153R.Global.SEQUENCE_TWO_SET":
                            Sequence_Two_Set = Convert.ToString(itemValue.Value);
                            break;

                        case "OHN66OPC.Spot_Weld_153R.Global.SEQUENCE_THREE_SET":
                            Sequence_Three_Set = Convert.ToString(itemValue.Value);
                            break;
                    }
                }
            }
            else if (System.Environment.MachineName == "OHN7123") // Paccar - 155R
            {
                foreach (ItemValueResult itemValue in values)
                {
                    switch (itemValue.ItemName)
                    {
                        case "OHN66OPC.Spot_Weld_155R.Global.HMI_PROBE_VALUE":
                            HMI_Probe_Value = Convert.ToString(itemValue.Value);
                            break;

                        case "OHN66OPC.Spot_Weld_155R.Global.TARGET_TRAVEL_VALUE":
                            Target_Travel_Value = Convert.ToString(itemValue.Value);
                            break;

                        case "OHN66OPC.Spot_Weld_155R.Global.MULTI_COMPONENT_TARGET2":
                            Multi_Component_Target2 = Convert.ToString(itemValue.Value);
                            break;

                        case "OHN66OPC.Spot_Weld_155R.Global.MULTI_COMPONENT_TARGET3":
                            Multi_Component_Target3 = Convert.ToString(itemValue.Value);
                            break;

                        case "OHN66OPC.Spot_Weld_155R.Global.MULTI_COMPONENT_TARGET4":
                            Multi_Component_Target4 = Convert.ToString(itemValue.Value);
                            break;

                        case "OHN66OPC.Spot_Weld_155R.Global.TARGET_TRAVEL_VALUE_SET":
                            Target_Travel_Value_Set = Convert.ToString(itemValue.Value);
                            break;

                        case "OHN66OPC.Spot_Weld_155R.Global.SEQUENCE_TWO_SET":
                            Sequence_Two_Set = Convert.ToString(itemValue.Value);
                            break;

                        case "OHN66OPC.Spot_Weld_155R.Global.SEQUENCE_THREE_SET":
                            Sequence_Three_Set = Convert.ToString(itemValue.Value);
                            break;
                    }
                }
            }
            // My PC for Testing
            else if (System.Environment.MachineName == "OHN7047NL") // My PC
            {
                foreach (ItemValueResult itemValue in values)
                {
                    switch (itemValue.ItemName)
                    {
                        case "OHN66OPC.Spot_Weld_121R.Global.SW121R_HMI_PROBE_VALUE":
                            HMI_Probe_Value = Convert.ToString(itemValue.Value);
                            break;

                        case "OHN66OPC.Spot_Weld_121R.Global.SW121R_TARGET_TRAVEL_VALUE":
                            Target_Travel_Value = Convert.ToString(itemValue.Value);
                            break;

                        case "OHN66OPC.Spot_Weld_121R.Global.SW121R_MULTI_COMPONENT_TARGET2":
                            Multi_Component_Target2 = Convert.ToString(itemValue.Value);
                            break;

                        case "OHN66OPC.Spot_Weld_121R.Global.SW121R_MULTI_COMPONENT_TARGET3":
                            Multi_Component_Target3 = Convert.ToString(itemValue.Value);
                            break;

                        case "OHN66OPC.Spot_Weld_121R.Global.SW121R_MULTI_COMPONENT_TARGET4":
                            Multi_Component_Target4 = Convert.ToString(itemValue.Value);
                            break;

                        case "OHN66OPC.Spot_Weld_121R.Global.SW121R_TARGET_TRAVEL_VALUE_SET":
                            Target_Travel_Value_Set = Convert.ToString(itemValue.Value);
                            break;

                        case "OHN66OPC.Spot_Weld_121R.Global.SW121R_SEQUENCE_TWO_SET":
                            Sequence_Two_Set = Convert.ToString(itemValue.Value);
                            break;

                        case "OHN66OPC.Spot_Weld_121R.Global.SW121R_SEQUENCE_THREE_SET":
                            Sequence_Three_Set = Convert.ToString(itemValue.Value);
                            break;
                    }
                }
            }

            // Write The Values to Their Text Box
            ProbeValue_TextBox.Invoke(new EventHandler(delegate { ProbeValue_TextBox.Text = HMI_Probe_Value; }));
            TargetValue_TextBox.Invoke(new EventHandler(delegate { TargetValue_TextBox.Text = Target_Travel_Value; }));
            MultiComponent1_TextBox.Invoke(new EventHandler(delegate { MultiComponent1_TextBox.Text = Target_Travel_Value; }));
            MultiComponent2_TextBox.Invoke(new EventHandler(delegate { MultiComponent2_TextBox.Text = Multi_Component_Target2; }));
            MultiComponent3_TextBox.Invoke(new EventHandler(delegate { MultiComponent3_TextBox.Text = Multi_Component_Target3; }));
            TargetValueSet_TextBox.Invoke(new EventHandler(delegate { TargetValueSet_TextBox.Text = Target_Travel_Value_Set; }));
            TargetTravel_Comp2_Set_TextBox.Invoke(new EventHandler(delegate { TargetTravel_Comp2_Set_TextBox.Text = Sequence_Two_Set; }));
            TargetTravel_Comp3_Set_TextBox.Invoke(new EventHandler(delegate { TargetTravel_Comp3_Set_TextBox.Text = Sequence_Three_Set; }));
            
            // Once the Target Value(s) have been set we update the interface
            if (TargetValueSet_TextBox.Text == "True")
            {
                SensorSet_TextBox.Invoke(new EventHandler(delegate { SensorSet_TextBox.Visible = true; }));
                OK_Button.Invoke(new EventHandler(delegate { OK_Button.Show(); }));
            }
            else if(TargetValueSet_TextBox.Text == "False")
            {
                SensorSet_TextBox.Invoke(new EventHandler(delegate { SensorSet_TextBox.Visible = false; }));
                OK_Button.Invoke(new EventHandler(delegate { OK_Button.Hide(); }));
            }
            if (TargetTravel_Comp2_Set_TextBox.Text == "True")
            {
                SensorSetComponent2_TextBox.Invoke(new EventHandler(delegate { SensorSetComponent2_TextBox.Visible = true; }));
            }
            else if (TargetTravel_Comp2_Set_TextBox.Text == "False")
            {
                SensorSetComponent2_TextBox.Invoke(new EventHandler(delegate { SensorSetComponent2_TextBox.Visible = false; }));
            }

            if (TargetTravel_Comp3_Set_TextBox.Text == "True")
            {
                SensorSetComponent3_TextBox.Invoke(new EventHandler(delegate { SensorSetComponent3_TextBox.Visible = true; }));
            }
            else if (TargetTravel_Comp3_Set_TextBox.Text == "False")
            {
                SensorSetComponent3_TextBox.Invoke(new EventHandler(delegate { SensorSetComponent3_TextBox.Visible = false; }));
            }
        }
        
        private void WriteCompleteCallback(object clientHandle, Opc.IdentifiedResult[] results)
        {

        }

        /********************************************************************************************************************
        *  
        *  OPC Region End
        *  
        ********************************************************************************************************************/
        #endregion

        /*********************************************************************************************************************
        * 
        * Methods Region Start
        * -- Total: 2
        * 
        * - Check TeachSensor
        * - CheckComponentValue
        * 
        *********************************************************************************************************************/
        #region

        private void CheckComponentValue()
        {
            if (Component_2_TextBox.Text == "")
            {
                Component_2_Button.Hide();
                Component_2_TextBox.Hide();
                MultiComponent2_TextBox.Hide();
            }
            if (Component_3_TextBox.Text == "")
            {
                Component_3_Button.Hide();
                Component_3_TextBox.Hide();
                MultiComponent3_TextBox.Hide();
            }
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

        /*********************************************************************************************************************
        * 
        * Methods Region End
        * 
        *********************************************************************************************************************/
        #endregion

        /*********************************************************************************************************************
        * 
        * Events Region Start
        * -- Inner Regions: 2
        * 
        * - Timer Region
        * - TextBox Enter Region 
        * 
        *********************************************************************************************************************/
        #region

        /*********************************************************************************************************************
        * TextBox Enter Region Start
        * 
        * -- Total TextBox: 8
        * -- Set the Actice Control to null if TextBoxes are Entered
        * 
        *********************************************************************************************************************/
        #region

        private void ProbeValue_TextBox_Enter(object sender, EventArgs e)
        {
            this.ActiveControl = null;
        }

        private void TargetValue_TextBox_Enter(object sender, EventArgs e)
        {
            this.ActiveControl = null;
        }

        private void MultiComponent1_TextBox_Enter(object sender, EventArgs e)
        {
            this.ActiveControl = null;
        }

        private void MultiComponent2_TextBox_Enter(object sender, EventArgs e)
        {
            this.ActiveControl = null;
        }

        private void MultiComponent3_TextBox_Enter(object sender, EventArgs e)
        {
            this.ActiveControl = null;
        }

        private void SensorSet_TextBox_Enter(object sender, EventArgs e)
        {
            this.ActiveControl = null;
        }

        private void SensorSetComponent2_TextBox_Enter(object sender, EventArgs e)
        {
            this.ActiveControl = null;
        }

        private void SensorSetComponent3_TextBox_Enter(object sender, EventArgs e)
        {
            this.ActiveControl = null;
        }

        /*********************************************************************************************************************
        * TextBox Enter Region End
        *********************************************************************************************************************/
        #endregion

        /*********************************************************************************************************************
        * 
        * Events Region End
        * 
        *********************************************************************************************************************/
        #endregion

        /*********************************************************************************************************************
        * 
        * User Teach Sensor End
        * 
        *********************************************************************************************************************/

        private void User_Program_Teach_Senson_FormClosing(object sender, FormClosingEventArgs e)
        {
            OPCServer.Disconnect();
        }        
    }
}
