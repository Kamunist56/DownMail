using System;
using System.Data;
using System.Windows.Forms;

namespace DownMailTest
{
    public partial class fOptions : Form
    {
        public string client;
        public string login;
        public string pass;
        public string path;
        public int port;
        public string interval;
        private WorkSQLite workSQL = new WorkSQLite(@"BoxLetters.sqlite"); //

        private void LoadData()
        {
            DataTable tb = workSQL.GetTable("SELECT login, pass, port, host FROM \"Hosts\"");
            listBox1.Items.Clear();
            for (int i = 0; i < tb.Rows.Count; i++)
            {

                listBox1.Items.Add(tb.Rows[i][0].ToString());
            }

            tb = workSQL.GetTable("SELECT path, interval FROM \"Settings\"");
            if (tb.Rows.Count > 0)
            {
                //for (int i = 0; tb.Rows.Count > i; i++)
                //{
                //    listBox1.Items.Add(tb.Rows[i][0].ToString());
                //}
                path = tb.Rows[0][0].ToString();
                interval = tb.Rows[0][1].ToString();
                textBox3.Text = path;
                maskedTextBox1.Text = interval;
            }
            workSQL.CloseConnect();

        }

        public void SetPass(string value)
        {
            textBox1.Text = value;
        }

        public void ReadParam()
        {


        }

        public void SaveParam()
        {
            if (checkSelectLogin() == false)
                return;

            string MailName = listBox1.SelectedItem.ToString();
            string host = textBox4.Text;
            string pass = textBox1.Text;
            string port = textBox2.Text;
            string path = textBox3.Text;
            string interval = maskedTextBox1.Text;

            using (DataTable table = workSQL.GetTable("SELECT id FROM \"Hosts\" where login=\"" + MailName + "\""))
            {
                if (table.Rows.Count > 0)
                {

                    workSQL.ExecuteQuery("update \"Hosts\" set pass=" + Func.AddQout(pass)
                                                             + ",port=" + Convert.ToInt64(port)
                                                             + ",host=" + Func.AddQout(host)
                                          + " where login=" + Func.AddQout(MailName));
                }
                else
                {

                    workSQL.ExecuteQuery("insert into \"Hosts\"(host,login,pass,port) values(" + Func.AddQout(host)
                                                             + "," + Func.AddQout(MailName)
                                                             + "," + Func.AddQout(pass)
                                                             + "," + Convert.ToInt64(port) + ")");

                }

                workSQL.CloseConnect();

            }

            using (DataTable table = workSQL.GetTable("SELECT id FROM \"Settings\""))
            {
                if (table.Rows.Count > 0)
                {
                    int idSet = Convert.ToInt32(table.Rows[0][0]);
                    workSQL.ExecuteQuery("update " + Func.AddQout("Settings") + " set path=" + Func.AddQout(path)
                                                                              + ",interval=" + Func.AddQout(interval)
                                                                              + "where id=" + (idSet));
                }
                else
                {
                    workSQL.ExecuteQuery("insert into " + Func.AddQout("Settings") + "(path, interval)" +
                                                        " values(" + Func.AddQout(path)
                                                        + "," + Func.AddQout(interval) + ")");
                }
                workSQL.CloseConnect();
            }



        }


        public fOptions()
        {
            InitializeComponent();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            SaveParam();
            using (Form1 Main = new Form1())
            {
                Main.CreateBase();
                Main.CheckSettings();                
            }
            Application.DoEvents();
            
        }

        private void button3_Click(object sender, EventArgs e)
        {
            String newLogin;
            if (InputBox.Input("Новый аккаунт", "Введите новый логин", out newLogin))
            {
                if (String.IsNullOrEmpty(newLogin) != true)
                {
                    listBox1.Items.Add(newLogin);
                }
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            
        }

        private void listBox1_MouseClick(object sender, MouseEventArgs e)
        {
            DataTable table = workSQL.GetTable("SELECT pass, port, host" +
                " FROM \"Hosts\" where login=\"" + listBox1.SelectedItem.ToString() + "\"");

            if (table.Rows.Count > 0)
            {
                textBox1.Text = table.Rows[0][0].ToString();
                textBox2.Text = table.Rows[0][1].ToString();
                textBox4.Text = table.Rows[0][2].ToString();
            }
            workSQL.CloseConnect();
        }

        private void fOptions_Shown(object sender, EventArgs e)
        {
            LoadData();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (FolderDialog.ShowDialog() == DialogResult.OK)
            {
                textBox3.Text = FolderDialog.SelectedPath.ToString();
            }
        }

        private void button3_Click_1(object sender, EventArgs e)
        {
            String newLogin;
            if (InputBox.Input("Новый аккаунт", "Введите новый логин", out newLogin))
            {
                listBox1.Items.Add(newLogin);
            }
        }

        private bool checkSelectLogin()
        {
            if (listBox1.SelectedItems.Count == 0)
            {
                MessageBox.Show("Нужно выбрать аккаунт");
                return false;
            }
            return true;
        }

        private void button5_Click(object sender, EventArgs e)
        {
            string host = textBox4.Text;
            string pass = textBox1.Text;
            string port = textBox2.Text;
            string path = textBox3.Text;
            string interval = maskedTextBox1.Text;
            DataTable table = workSQL.GetTable("SELECT id FROM \"Settings\"");
            if (table.Rows.Count > 0)
            {
                int idSet = Convert.ToInt32(table.Rows[0][0]);
                workSQL.ExecuteQuery("update " + Func.AddQout("Settings") + " set path=" + Func.AddQout(path)
                                                                              + ",interval=" + Func.AddQout(interval)
                                                                              + "where id=" + Convert.ToInt64(idSet));
            }

        }

        private void fOptions_Load(object sender, EventArgs e)
        {
            LoadData();
        }

        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            DataTable table = workSQL.GetTable("SELECT pass, port, host" +
                " FROM \"Hosts\" where login=" + Func.AddQout(listBox1.SelectedItem.ToString()));

            if (table.Rows.Count > 0)
            {
                textBox1.Text = table.Rows[0][0].ToString();
                textBox2.Text = table.Rows[0][1].ToString();
                textBox4.Text = table.Rows[0][2].ToString();
            }
            workSQL.CloseConnect();
        }

        private void button4_Click_1(object sender, EventArgs e)
        {
            workSQL.ExecuteQuery("Delete from Hosts where login=" + Func.AddQout(listBox1.SelectedItem.ToString()));
            LoadData();
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            if (FolderDialog.ShowDialog() == DialogResult.OK)
            {
                textBox3.Text = FolderDialog.SelectedPath;
            }
        }
    }
}
