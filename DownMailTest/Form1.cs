using System;
using System.Collections.Generic;
using System.Windows.Forms;
using OpenPop.Pop3;
using OpenPop.Mime;
using System.IO;
using System.Data;

namespace DownMailTest
{
    public partial class Form1 : Form
    {
        Pop3Client client = new Pop3Client();
        private DateTime nowDate = DateTime.Today.AddDays(-3);
        private DateTime endDate = DateTime.Today.AddDays(-3);
        private WorkSQLite workSqlite;//= new WorkSQLite(@"BoxLetters.sqlite");

        private void connect(string host, string login, string pass, int port, string dir)
        {
            try
            {
                client.Connect(host, port, true);
            }
            catch
            {
                Exception ex = new Exception();
                richTextBox2.AppendText(ex.Message);

            }
            client.Authenticate(login, pass);
            richTextBox2.AppendText("Вошел в почту\n");
            Application.DoEvents();
            List<string> msgs = client.GetMessageUids();
            richTextBox2.AppendText("Получил список id писем\n");
            Application.DoEvents();
            GetHeadMess(msgs, dir);
            richTextBox2.AppendText("Закончил\n");
        }


        private void GetHeadMess(List<string> msgs, string dir)
        {
            int i = msgs.Count;
            DateTime msgDate;
            SetDates();
            richTextBox2.AppendText("Начали смотреть че там да как...\n");
            Application.DoEvents();
            DateTime LoadDate = nowDate;

            do
            {
                richTextBox1.Clear();
                Application.DoEvents();
                OpenPop.Mime.Message msg = client.GetMessage(i);
                --i;
                string date = msg.Headers.Date;
                string subject = msg.Headers.Subject;
                string adress = msg.Headers.From.Address;
                string messId = msg.Headers.MessageId;
                string dirArchive = dir +"\\"+ Func.DirMonth() + "\\";

                if ((String.IsNullOrEmpty(subject)) & (subject == ""))
                    subject = "БезТемы";
                else { subject = Func.DelBadChars(subject); }
                
                richTextBox1.AppendText("Получили инфу по письму\n");
                Application.DoEvents();

                date = date.Remove(25, date.Length - 25);
                DateTime.TryParse(date, out msgDate);
                GetMessagersInTable();
                //прверяем диапазон письма и наличие письма в базе
                if (((msgDate.Date.CompareTo(nowDate.Date) >= 0) & (msgDate.Date.CompareTo(endDate.Date) <= 0)) & (FindMessInTable(messId)!=true))
                {                    
                    string d = msgDate.Date.ToString();
                    d = d.Remove(10, 8);
                    dirArchive = dirArchive + d + "\\";

                    //если нашли в таблице письмо с такой же темой, 
                    //то прибавим к теме нового письма адрес автора
                    if (FindSubjectInTable(subject))
                    {
                        dirArchive = dirArchive + " (" + adress + ")";
                    }

                    // проверяем длинну и создаем дирректорию
                    subject = Func.TrimSubject(dirArchive, subject); //делаем путьдо 250

                    DirectoryInfo di = new DirectoryInfo(dirArchive + subject);
                    di.Create();

                    // пишем в таблицу
                    //dataGridView1.Rows.Add(subject, adress, msgDate, messId);

                    // WorkSQLite workSQL = new WorkSQLite(@"BoxLetters.sqlite");
                    workSqlite.ExecuteQuery("insert into Messages (Subject, From_, Data, idMessage) Values("
                                            + Func.AddQout(subject) + "," + Func.AddQout(adress) + ","
                                            + Func.AddQout(msgDate.ToString()) + "," + Func.AddQout(messId) + ")");
                    GetMessagersInTable();
                    dataGridView1.Refresh();
                    //загрузка тела письма
                    richTextBox2.AppendText("Загрузка письма: " + subject + "\n");
                    Application.DoEvents();
                    LoadMess(msg, subject, dirArchive);

                    //Загрузка атачмансов
                    richTextBox1.AppendText("Проверка вложений\n");
                    Application.DoEvents();
                    DownLoadAttach(msg, subject, dirArchive);

                }

 
            }
            while ((msgDate.Date.CompareTo(nowDate.Date) >= 0) & (msgDate.Date.CompareTo(endDate.Date) <= 0));
            client.Disconnect();

        }
        private void GetMessagersInTable()
        {
            SetDates();
            string startDate = nowDate.Date.ToString();
            string endData = endDate.Date.AddDays(1).ToString();
            //startDate = startDate.Remove(10, 8);
            //endData = endDate.Date.AddDays(1);
            //    endData.Remove(10, 8);
            //endData = endData.Insert(10, " 23:59:59");

            //WorkSQLite workSqlite = new WorkSQLite(@"BoxLetters.sqlite");
            DataTable table = workSqlite.GetTable("Select Subject, From_, cast(Data as varchar) Data, idMessage"
                                                    + " From Messages"
                                                    + " Where Data between "
                                                    + Func.AddQout(startDate)
                                                    + " and " + Func.AddQout(endData)
                                                    + " Order by Data asc");
            dataGridView1.DataSource = table;
            dataGridView1.Refresh();
        }

        private void LoadMess(OpenPop.Mime.Message mess, string subject, string dirArchive)
        {

            string fileName = dirArchive + subject + "\\" + subject;
            //// ищем первую плейнтекст версию в сообщении
            richTextBox1.AppendText("Поиск текста\n");
            Application.DoEvents();
            MessagePart mpPlain = mess.FindFirstPlainTextVersion();



            if (mpPlain != null)
            {
                fileName = fileName + ".txt";
                mpPlain.Save(new FileInfo(Func.TrimSubject(fileName)));
                richTextBox1.AppendText("Сохранили текст\n");
                Application.DoEvents();
            }
            else
            {
                richTextBox1.AppendText("Смотрим аштмэйл\n");
                Application.DoEvents();
                MessagePart html = mess.FindFirstHtmlVersion();
                if (html != null)
                {
                    //html.BodyEncoding()
                    fileName = fileName + ".html";
                    html.Save(new FileInfo(Func.TrimSubject(fileName)));
                    richTextBox1.AppendText("Сохранили аштмэйл\n");
                    Application.DoEvents();
                }
            }
        }

        private void DownLoadAttach(OpenPop.Mime.Message mess, string subject, string dirArchive)
        {
            //string mesSubj = mess.Headers.Subject;
            foreach (MessagePart attachment in mess.FindAllAttachments())
            {
                if (attachment.FileName.Equals(attachment.FileName))
                {
                    // Save the raw bytes to a file
                    File.WriteAllBytes(dirArchive + subject + "\\"
                        + Func.TrimSubject(attachment.FileName), attachment.Body);
                }

            }
        }

        private void MainLoadMail()
        {
            //WorkSQLite workSqlite = new WorkSQLite(@"BoxLetters.sqlite");
            DataTable table = workSqlite.GetTable("Select path, interval from Settings");
            string path = table.Rows[0][0].ToString();
            string interval = table.Rows[0][1].ToString();

            table = workSqlite.GetTable("Select login, pass, port, host"
                                                   + " From Hosts");
            for (int i = 0; i < table.Rows.Count; i++)
            {
                string login = table.Rows[i][0].ToString();
                string pass = table.Rows[i][1].ToString();
                int port = Convert.ToInt32(table.Rows[i][2].ToString());
                string host = table.Rows[i][3].ToString();
                connect(host, login, pass, port, path);
            }
            

        }

        private bool FindMessInTable(string idMessage)
        {
            //WorkSQLite work = new WorkSQLite(@"BoxLetters.sqlite");
            DataTable tb = workSqlite.GetTable("select id from Messages where IdMessage=" + Func.AddQout(idMessage));
            if (tb.Rows.Count > 0)
            {
                return true;
            }
            else
            {
                return false;
            }
            
        }

        private bool FindSubjectInTable(string subject)
        {

            //WorkSQLite work = new WorkSQLite(@"BoxLetters.sqlite");
            DataTable tb = workSqlite.GetTable("select id from Messages where Subject=" + Func.AddQout(subject));
            if (tb.Rows.Count > 0)
            {
                return true;
            }
            return false;
        }

        private void SetFileCharset(string fileName)
        {
            string line;
            using (StreamReader sr = new StreamReader(fileName))
            {
                line = sr.ReadLine();
            }
            int ind = line.IndexOf("<meta");
            if (ind < 0)
            {
                using (StreamWriter sw = new StreamWriter(fileName))
                {
                    //sw.WriteLine("<meta charset=\"koi8-r\">");
                    sw.WriteLine(line);
                }
            }

        }

        private void SetDates()
        {
            nowDate = monthCalendar1.SelectionStart;
            endDate = monthCalendar1.SelectionEnd;
        }

        public Form1()
        {
            InitializeComponent();
        }

        private void CreateBase()
        {
            string base_ = @"BoxLetters.sqlite";
            if (File.Exists(base_)!=true)
            {
                BaseCreator baseCreator = new BaseCreator(base_);
            }
            workSqlite = new WorkSQLite(base_);
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            CreateBase();
            GetMessagersInTable();
        }


        private void button2_Click(object sender, EventArgs e)
        {
            GetMessagersInTable();


        }

        private void настройкиToolStripMenuItem_Click(object sender, EventArgs e)
        {
            fOptions options = new fOptions();
            options.Show();

        }

        private void загрузкаToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MainLoadMail();
        }

        private void загрузкаToolStripMenuItem_Click_1(object sender, EventArgs e)
        {
            MainLoadMail();
        }

        private void richTextBox2_TextChanged(object sender, EventArgs e)
        {
            richTextBox1.SelectionStart = richTextBox1.Text.Length;
            richTextBox1.ScrollToCaret();
        }

        private void тестToolStripMenuItem_Click(object sender, EventArgs e)
        {
            GetMessagersInTable();
        }

        private void удалитьПисьмоToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // WorkSQLite work = new WorkSQLite(@"BoxLetters.sqlite");
            workSqlite.ExecuteQuery("delete from Messages where idMessage=" + 
                Func.AddQout(dataGridView1[ 3,dataGridView1.CurrentRow.Index].Value.ToString()));
            GetMessagersInTable();
            dataGridView1.Refresh();
        }
    }
}