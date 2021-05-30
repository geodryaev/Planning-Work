using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Planning_Work
{
    public partial class FormSettingsForSqlServer : Form
    {
        public FormSettingsForSqlServer()
        {
            InitializeComponent();
            textBox1.Text = Properties.Settings.Default.PathSqlServer;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Properties.Settings.Default.PathSqlServer = textBox1.Text;
            Properties.Settings.Default.Save();
            this.Close();
        }
    }
}
