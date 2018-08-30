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
    public partial class User_Program_Sign_Off : Form
    {
        public User_Program_Sign_Off()
        {
            InitializeComponent();
        }

        private static int CountdownMinute;
        private static int CountdownSecond;
        string CountdownTime = "";

        private void User_Program_Sign_Off_Load(object sender, EventArgs e)
        {
            Timer.Start();
            CountdownMinute = 1;
            CountdownSecond = 59;
        }

        private void Timer_Tick(object sender, EventArgs e)
        {
            CountdownSecond = CountdownSecond - 1;
            if (CountdownMinute == 0 && CountdownSecond == 0)
            {
                Timer.Stop();
                LogOff_Button_Click(null, null);
            }
            if (CountdownSecond == 0)
            {
                CountdownMinute = CountdownMinute - 1;
                CountdownSecond = 59;
                CountdownTime = CountdownMinute.ToString() + ":" + CountdownSecond.ToString();
                Countdown_TextBox.Text = CountdownTime;
            }
            else if (CountdownSecond <= 9)
            {
                CountdownTime = CountdownMinute.ToString() + ":0" + CountdownSecond.ToString();
                Countdown_TextBox.Text = CountdownTime;
            }
            else if (CountdownSecond >= 10 && CountdownSecond < 60)
            {
                CountdownTime = CountdownMinute + ":" + CountdownSecond;
                Countdown_TextBox.Text = CountdownTime;
            }

            //CountdownTime = CountdownMinute + ":" + CountdownSecond;
            //Countdown_TextBox.Text = CountdownTime;
        }

        private void LogOff_Button_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void StaySignedIn_Button_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
