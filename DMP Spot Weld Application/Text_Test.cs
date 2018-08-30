using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace DMP_Spot_Weld_Application
{
    public partial class Text_Test : Form
    {
        public Text_Test()
        {
            InitializeComponent();
        }

        private void Send_Button_Click(object sender, EventArgs e)
        {
            string[] test = { User_TextBox.Text, DMPID_TextBox.Text, Cell_TextBox.Text, BrakePress_TextBox.Text, ItemID_TextBox.Text, JobID_TextBox.Text, Messaging_TextBox.Text };

            File.WriteAllLines(@"\\insidedmp.com\Corporate\OH\OH Common\Engineering\Brake Press\Vision\test.txt", test);

        }

        private void Cancel_Button_Click(object sender, EventArgs e)
        {
            try
            {
                MailMessage objeto_mail = new MailMessage();
                SmtpClient client = new SmtpClient();
                client.Port = 25;
                client.Host = "ClientArray.insidedmp.com";
                client.Timeout = 10000;
                client.DeliveryMethod = SmtpDeliveryMethod.Network;
                //client.UseDefaultCredentials = false;
                //client.Credentials = new System.Net.NetworkCredential("rgarland", "ryan0505");
                objeto_mail.From = new MailAddress("rgarland@defiancemetal.com");
                objeto_mail.To.Add(new MailAddress("rgarland@defiancemetal.com"));
                objeto_mail.Subject = "";
                objeto_mail.Body = "";
                client.Send(objeto_mail);
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }
    }
}
