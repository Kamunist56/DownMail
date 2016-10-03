using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace DownMailTest
{
    static class Program
    {
        /// <summary>
        /// Главная точка входа для приложения.
        /// </summary>
        [STAThread]
        static void Main(string[] args)
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Form1 form = new Form1();
            for (int i = 0; i<args.Length; i++)
            {
                if (args[i] == "auto")
                {
                    form.CreateBase();
                    form.SetNowDate();
                    form.CheckSettings();
                    form.MainLoadMail();
                }
            }
            
            Application.Run(form);

        }
    }
}
