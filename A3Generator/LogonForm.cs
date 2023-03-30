using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace A3Generator
{
    public partial class LogonForm : Form
    {
        public string PAT { 
            get
            {
                return this.textBox1.Text.Trim();
            }
        }

        public LogonForm()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(this.textBox1.Text.Trim()))
            {
                this.DialogResult = DialogResult.OK;
                this.Close();
            }
        }
    }
}
