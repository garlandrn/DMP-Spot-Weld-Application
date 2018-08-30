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
 * Form: View PDF
 * Created By: Ryan Garland
 * Last Updated on 8/28/18
 * 
 */

namespace DMP_Spot_Weld_Application
{
    public partial class View_PDF : Form
    {
        public View_PDF()
        {
            this.ShowInTaskbar = false;
            InitializeComponent();
        }

        private void Close_Button_Click(object sender, EventArgs e)
        {
            try
            {
                User_Program.UserProgram.Enabled = true;
            }
            catch
            {

            }
            this.Close();
        }
    }
}
