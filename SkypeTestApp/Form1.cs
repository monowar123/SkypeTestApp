using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using SkypeLib;
using System.IO;
using SKYPE4COMLib;

namespace SkypeTestApp
{
    public partial class Form1 : Form
    {
        SkypeHelper skypeHelper;
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            skypeHelper = new SkypeHelper();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                List<string> contacts = new List<string>();
                contacts = skypeHelper.GetAllContact();
                foreach (var id in contacts)
                {
                    listBox1.Items.Add(id);
                }
                MessageBox.Show(contacts.Count.ToString());
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }       
        }     

        private void btnMessage_Click(object sender, EventArgs e)
        {
            string id = "masud0512";
            try
            {
                Skype skype = new Skype();

                IEnumerable<SKYPE4COMLib.User> users = skype.Friends.OfType<SKYPE4COMLib.User>();

                foreach (Chat item in skype.Chats)
                {
                    listBox1.Items.Add(item.Name);
                }

                
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            skypeHelper.SendNotification();
            MessageBox.Show("Send");
        }
        
    }
}
