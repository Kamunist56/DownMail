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

        private void connect(string host, string login, string pass, int port)
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
            GetHeadMess(msgs);
            richTextBox2.AppendText("Закончил\n");
        }


        private void GetHeadMess(List<string> msgs)
        {
            int i = msgs.Count;
            DateTime msgDate;
            SetDates();
            richTextBox2.AppendText("Начали смотреть че там да как...\n");
            Application.DoEvents();
            bool stop = false;


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
                string dirArchive = "D:\\Temp\\" + Func.DirMonth() + "\\";


                if ((subject == null))
                    subject = "БезТемы";

                subject = Func.DelBadChars(subject);
                richTextBox1.AppendText("Получили инфу по письму\n");
                Application.DoEvents();

                date = date.Remove(25, date.Length - 25);
                DateTime.TryParse(date, out msgDate);
                GetMessagersInTable();
                if ((msgDate.Date.CompareTo(nowDate.Date) == 0))
                {
                    // если пиьмо есть в таблице то не будем его грузить
                    if (FindMessInTable(messId))
                    {
                        continue;
                    }

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

                    WorkSQLite workSQL = new WorkSQLite(@"BoxLetters.sqlite");
                    workSQL.ExecuteQuery("insert into Messages (Subject, From_, Data, idMessage) Values("
                                          + Func.AddQout(subject) + "," + Func.AddQout(adress) + ","
                                          + Func.AddQout(msgDate.ToString()) + "," + Func.AddQout(messId) + ")");
                    dataGridView1.Refresh();
                    //загрузка тела письма
                    richTextBox2.AppendText("Загрузка письма: " + subject + "\n");
                    Application.DoEvents();
                    LoadMess(msg, subject, dirArchive);

                    //Загрузка атачмансов
                    richTextBox1.AppendText("Проверка вложений\n");
                    Application.DoEvents();
                    DownLoadAttach(msg, subject, dirArchive);

                    stop = false;
                }

                if (msgDate.Date.CompareTo(nowDate.Date) == -1)
                {
                    stop = true;

                }

            }

            while ((i <= msgs.Count) && (stop == false));
            client.Disconnect();

        }

        private void GetMessagersInTable()
        {
            SetDates();
            string startDate = nowDate.Date.ToString();
            string endData = endDate.Date.ToString();
            //startDate = startDate.Remove(10, 8);
            endData = endData.Remove(10, 8);
            endData = endData.Insert(10, " 23:00:00");

            WorkSQLite workSqlite = new WorkSQLite(@"BoxLetters.sqlite");
            DataTable table = workSqlite.GetTable("Select Subject, From_, Data, idMessage"
                                                    + " From Messages"
                                                    +" Where Data between "
                                                    + Func.AddQout(startDate)
                                                    + " and " + Func.AddQout(endData));
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
            WorkSQLite workSqlite = new WorkSQLite(@"BoxLetters.sqlite");
            DataTable table = workSqlite.GetTable("Select login, pass, port, host"
                                                   + " From Hosts");
            for (int i = 0; i < table.Rows.Count; i++)
            {
                string login = table.Rows[i][0].ToString();
                string pass = table.Rows[i][1].ToString();
                int port = Convert.ToInt32(table.Rows[i][2].ToString());
                string host = table.Rows[i][3].ToString();
                connect(host, login, pass, port);
            }
        }

        private bool FindMessInTable(string idMessage)
        {
            bool est = false;
            int i = 0;

            while ((est == false) && (i < dataGridView1.RowCount - 1))
            {
                string dataIdMessage = dataGridView1.Rows[i].Cells[3].Value.ToString();
                if (dataIdMessage == idMessage)
                {
                    est = true;
                }
                i++;
            }
            return est;
        }

        private bool FindSubjectInTable(string subject)
        {
            bool est = false;
            int i = 0;

            while ((est == false) && (i < dataGridView1.RowCount - 1))
            {
                string dataSubject = dataGridView1.Rows[i].Cells[0].Value.ToString();
                if (dataSubject == subject)
                {
                    est = true;
                }
                i++;
            }
            return est;
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

        private void Form1_Load(object sender, EventArgs e)
        {

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
    }
}