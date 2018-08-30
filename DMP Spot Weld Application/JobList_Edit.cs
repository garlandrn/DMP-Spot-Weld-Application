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
    public partial class JobList_Edit : Form
    {
        public JobList_Edit()
        {
            InitializeComponent();
            Confirm_Button.DialogResult = DialogResult.Yes;
            Cancel_Button.DialogResult = DialogResult.No;
        }
    }
}
