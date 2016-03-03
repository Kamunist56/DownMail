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
        private DateTime nowDate;// = DateTime.Today.AddDays(-3);
        private DateTime endDate;// = DateTime.Today.AddDays(-3);
        private WorkSQLite workSqlite;
        private string ErrorLogPath = "ErrorLog.txt";
        private string WorkLog = "Log.txt";
        private int CountDownLoadMessage = 0;
        private DateTime LastStart;
        private Logs Log;


        private void connect(string host, string login, string pass, int port, string dir)
        {

            try
            {
                client.Connect(host, port, true);
                client.Authenticate(login, pass);
            }
            catch (InvalidOperationException ex)
            {
                SendOnErrorLog(login, ex.Message);
            }
            LastStart = DateTime.Now;
            toolStripStatusLabel1.Text = "Загрузка писем";
            SendOnWorkLog("Вошел в почту " + login +" Время: "+ LastStart.ToString());
            Application.DoEvents();
            List<string> msgs = client.GetMessageUids();
            SendOnWorkLog("Получил список id писем");
            Application.DoEvents(); // моргнули
            GetHeadMess(msgs, dir, login); //грузим
            richTextBox2.AppendText("Закончил\n");
            SendOnWorkLog("Закончил");
        }


        private void GetHeadMess(List<string> msgs, string dir, string login)
        {
            try
            {

                int i = msgs.Count;
                DateTime msgDate;
                SetDates();
                richTextBox2.AppendText("Начали смотреть че там да как...\n");
                Application.DoEvents();
                DateTime LoadDate = nowDate;
                string debug;

                do
                {
                    Application.DoEvents();
                    OpenPop.Mime.Message msg = client.GetMessage(i);
                    --i;
                    string date = msg.Headers.Date;
                    string subject = msg.Headers.Subject;
                    string adress = msg.Headers.From.Address;
                    string messId = msg.Headers.MessageId;
                    string dirArchive;

                    if (String.IsNullOrEmpty(date) == true)
                    {
                        msgDate = endDate = DateTime.Today.AddDays(-1);
                        continue;
                    }
                    if ((String.IsNullOrEmpty(subject)))  // если пустая тема пишем что без темы,
                        subject = "БезТемы";               //если нет убрать спец символы и пробелы в начале и в конце
                    else { subject = Func.DelBadChars(subject); }


                    Application.DoEvents();
                    // дата приходит в тестовом фомате + значение часового пояса,
                    // TryParse бывет тупит с поясом, по этому отрежу его нафиг
                    int IndexBelt = date.IndexOf('+'); // ещем плюсик или минус после них идет пояс
                    if (IndexBelt == -1)
                    {
                        IndexBelt = date.IndexOf('-');
                    }  // ну и режем
                    date = date.Remove(IndexBelt, date.Length - IndexBelt);
                    DateTime.TryParse(date, out msgDate);

                    label1.Text = "Дата проверяемого письма: " + msgDate.Date.ToString();
                    label2.Text = "Тема: " + subject;
                    
                    debug = subject + ' ' + msgDate.Date.ToString() + ' ' + date;

                    GetMessagersInTable(); // показать письма в таблице за выбранный диапазон

                    //проверка диапазона письма и наличие письма в базе
                    //while (CheckMessForDownload(msgDate, messId))
                    if ((msgDate.Date.CompareTo(nowDate.Date) >= 0) &
                        (msgDate.Date.CompareTo(endDate.Date) <= 0)
                        & (FindMessInTable(messId) == false))
                    {
                        // добавляю в имя папки число получаемого письма
                        string d = msgDate.Date.ToString();
                        d = d.Remove(10, 8); // время для имени папки не надо                   
                        dirArchive = dir + "\\" + Func.DirMonth(msgDate.Date) + "\\";// название месяца букавками
                        dirArchive = dirArchive + d + "\\";

                        //если нашел в таблице письмо с такой же темой, 
                        //то прибавим к теме нового письма адрес автора
                        if (FindSubjectInTable(subject))
                        {
                            dirArchive = dirArchive + " (" + adress + ")";
                        }

                        // проверка длинну и создание дирректории
                        subject = Func.TrimSubject(dirArchive, subject); //режу путь до 250

                        DirectoryInfo di = new DirectoryInfo(dirArchive + subject);
                        di.Create();

                        //загрузка тела письма
                        richTextBox2.AppendText("Загрузка письма: " + subject + "\n");
                        Application.DoEvents();
                        LoadMess(msg, subject, dirArchive);

                        //Загрузка атачмансов
                        Application.DoEvents();
                        DownLoadAttach(msg, subject, dirArchive);

                        // пишем в таблицу
                        workSqlite.ExecuteQuery("insert into Messages (Subject, From_, Data, idMessage, PathMessage) Values("
                                                + Func.AddQout(subject) + "," + Func.AddQout(adress) + ","
                                                + Func.AddQout(msgDate.ToString()) + "," + Func.AddQout(messId) + ","
                                                + Func.AddQout(dirArchive + subject) + ")");
                        ++CountDownLoadMessage;
                        
                        GetMessagersInTable();
                        dataGridView1.Refresh();

                    }


                }


                while ((msgDate.Date.CompareTo(nowDate.Date) != -1));
            }
            catch (InvalidOperationException ex)
            {
                SendOnErrorLog(login, ex.Message);
            }
            finally
            {
                client.Disconnect();
                toolStripStatusLabel1.Text = "Загрузка завершена";
            }



        }
        private void GetMessagersInTable()
        {
            SetDates();
            // d j,otv 
            string startDate = nowDate.Date.ToString();
            string endData = endDate.Date.AddDays(1).ToString();

            DataTable table = workSqlite.GetTable("Select Subject, From_, cast(Data as varchar) Data, idMessage "
                                                    + " From Messages"
                                                    + " Where Data between " + Func.AddQout(startDate+" 00:00:00")
                                                    + " and  " + Func.AddQout(endData+ " 23:59:59")
                                                    + " Order by Data asc");
            dataGridView1.DataSource = table;
            dataGridView1.Refresh();
        }

        private void LoadMess(OpenPop.Mime.Message mess, string subject, string dirArchive)
        {
            //  FileInfo file = new FileInfo(subject+".eml");

            // Save the full message to some file
            // mess.Save(file);

            // Now load the message again. This could be done at a later point
            // OpenPop.Mime.Message loadedMessage = mess.Load(file);
            
            string fileName = dirArchive + subject + "\\" + subject;
            //// ищем первую плейнтекст версию в сообщении
            Application.DoEvents();
            MessagePart mpPlain = mess.FindFirstPlainTextVersion();



            if (mpPlain != null)
            {
                fileName = fileName + ".txt";
                mpPlain.Save(new FileInfo(Func.TrimSubject(fileName)));
                Application.DoEvents();
            }
            else
            {
                Application.DoEvents();
                MessagePart html = mess.FindFirstHtmlVersion();
                if (html != null)
                {
                    //html.BodyEncoding()
                    fileName = fileName + ".html";
                    html.Save(new FileInfo(Func.TrimSubject(fileName)));
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
                        + Func.TrimSubject(Func.DelBadChars(attachment.FileName)), attachment.Body);
                }

            }
        }

        private void MainLoadMail()
        {
            //WorkSQLite workSqlite = new WorkSQLite(@"BoxLetters.sqlite");
            DataTable table = workSqlite.GetTable("Select path, interval from Settings");
            string path = "";
            string interval;
            if (table.Rows.Count > 0)
            {
                path = table.Rows[0][0].ToString();
                interval = table.Rows[0][1].ToString();
            }
            else
            {
                MessageBox.Show("Ни одного аккаунта не найдено");
                return;
            }


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

        private bool CheckMessForDownload(DateTime dat, string messId)
        {//если выранно 1 число, грузим те которые совпадают с этим числом
         //если выбран диапазон то грузим по диапазону
            if ((monthCalendar1.SelectionStart.Day - monthCalendar1.SelectionEnd.Day) > 1)
            {
                if (((dat.Date.CompareTo(nowDate.Date) >= 0) &
                     (endDate.Date.CompareTo(dat.Date) <= 0)) &
                     (FindMessInTable(messId) != true))
                {
                    return true;
                }
            }
            else
            {
                if ((dat.Date.CompareTo(nowDate.Date) == 0) & (dat.Date.CompareTo(nowDate.Date) == 1))
                {
                    return true;
                }

            }

            return false;
        }

        private void SendOnErrorLog(string login, string mess)
        {
            // директория для логов
            string LogFile = CreateDirLog() + "//" + ErrorLogPath;

            Func.WriteLog(LogFile, login);
            Func.WriteLog(LogFile, mess);
            Func.WriteLog(LogFile, "######################################");
        }

        private void SendOnWorkLog(string mess)
        {
            string LogFile = CreateDirLog() + "//" + WorkLog;
            Func.WriteLog(LogFile, mess);
        }

        private string CreateDirLog()
        {
            string d = DateTime.Now.Date.ToString();
            d = d.Remove(10, 8);
            string dir = "Logs\\" + d;            
            Directory.CreateDirectory(dir);
            return dir;
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

        public void CreateBase()
        {
            string base_ = @"BoxLetters.sqlite";
            if (File.Exists(base_) != true)
            {
                BaseCreator baseCreator = new BaseCreator(base_);
                baseCreator.CreateTables();
            }
            workSqlite = new WorkSQLite(base_);
        }

        private void Form1_Load(object sender, EventArgs e)
        {
           // timer1.Enabled = true;
            CreateBase();
          //  CheckSettings();            
            GetMessagersInTable();            
            label2.MaximumSize = new System.Drawing.Size(500, 0);
            label2.AutoSize = true;


        }

        public void CheckSettings()
        {       
            DataTable table = workSqlite.GetTable("Select path, interval from Settings");
            string path = "";
            string interval;
            if (table.Rows.Count > 0)
            {
                path = table.Rows[0][0].ToString();
                interval = table.Rows[0][1].ToString();
                timer1.Interval = Convert.ToInt32(interval) * 60000;
                label3.Text = "Следующий запуск: " + DateTime.Now.AddMinutes(Convert.ToInt32(interval)).ToString();
            }
            else
            {
                toolStripStatusLabel1.Text = "Настройки не заданны";
            }
            Application.DoEvents();
            
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
            richTextBox2.SelectionStart = richTextBox2.Text.Length;
            richTextBox2.ScrollToCaret();
        }

        private void тестToolStripMenuItem_Click(object sender, EventArgs e)
        {
            GetMessagersInTable();
        }

        private void удалитьПисьмоToolStripMenuItem_Click(object sender, EventArgs e)
        {
            workSqlite.ExecuteQuery("delete from Messages where idMessage=" +
                Func.AddQout(dataGridView1[3, dataGridView1.CurrentRow.Index].Value.ToString()));
            GetMessagersInTable();
            dataGridView1.Refresh();
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            monthCalendar1.SetDate(DateTime.Now);
            CheckSettings();
            MainLoadMail();

        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            timer1.Enabled = !timer1.Enabled;
        }
    }
}