using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace DownMailTest
{
    public partial class Logs_ : Form
    {
        public Logs_()
        {
            InitializeComponent();
        }

        public void AddTextOnRich(string text)
        {
            richTextBox2.AppendText(text + "\n");
            Application.DoEvents();
        }

        private void richTextBox2_TextChanged(object sender, EventArgs e)
        {
            richTextBox2.SelectionStart = richTextBox2.Text.Length;
            richTextBox2.ScrollToCaret();
        }

        private void Logs__Load(object sender, EventArgs e)
        {

        }

        private void Logs__FormClosing(object sender, FormClosingEventArgs e)
        {
            e.Cancel = true;
            this.Hide();
        }
    }
}
