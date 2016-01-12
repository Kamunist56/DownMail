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
    public partial class InputBox : Form
    {
        bool status;
        String temp;
        public InputBox()
        {
            InitializeComponent();
        }

        public static bool Input(string head, string mess, out string value)
        {
            InputBox ib = new InputBox();
            ib.Text = head;
            ib.label1.Text = mess;
            ib.StartPosition = FormStartPosition.CenterScreen;
            ib.ShowDialog();
            value = ib.temp;
            ib.status = true;
            return ib.status;
        }


        private void button2_Click(object sender, EventArgs e)
        {
            status = false;
            this.Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            temp = this.textBox1.Text;
            status = true;
            this.Close();
        }

    }
}
