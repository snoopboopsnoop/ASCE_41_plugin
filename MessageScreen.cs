using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ASCE_41
{
    public partial class MessageScreen : Form
    {
        public MessageScreen()
        {
            InitializeComponent();
        }

        public MessageScreen(string text)
        {
            InitializeComponent();
            label1.Text = text;
        }

        public MessageScreen(string text, string okBtn, string cancelBtn)
        {
            InitializeComponent();
            label1.Text = text;
            button1.Text = okBtn;
            button2.Text = cancelBtn;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.OK;
            this.Close();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
            this.Close();
        }
    }
}
