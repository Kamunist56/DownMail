using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DownMailTest
{
    class Func
    {
        static public string DelBadChars(string value)
        {

            string[] arr = new string[] { "\"", @":", "?", "<", ">", "|", @"/", @"\", "\t" };
            for (int i = 0; i < arr.Length - 1; i++)
            {
                string bchar = arr[i];
                for (int t = 0; t < value.Length - 1; t++)
                {
                    int indChar = value.IndexOf(bchar);
                    if (indChar != -1)
                    {
                        value = value.Remove(indChar, 1);
                        if (indChar == 0)
                        {
                            indChar = value.IndexOf(bchar);
                            if (indChar != -1)
                            {
                                value = value.Remove(indChar, 1);

                            }
                        }
                    }

                    //} while (value.IndexOf(arr[i]) > 0);                    
                }

            }
            return value.Trim();
        }


        static public string TrimSubject(string dir, string subject)
        {//ищем последний пробел и удаляем все после него пока не станет хорошо))
            if (dir.Length + subject.Length + subject.Length > 200)
            {
                do
                {
                    int lastSpace = subject.LastIndexOf(" ");
                    if (lastSpace > 0)
                    {
                        subject = subject.Remove(lastSpace, subject.Length - lastSpace);
                    }
                }
                while (dir.Length + subject.Length + subject.Length > 200);
            }
            return subject.Trim();
        }

        static public string TrimSubject(string Filename)
        {//находим последнюю точку чтобы не удалить формат файла, и также удаляем фразы после последнего пробела до точки
            if (Filename.Length > 200)
            {
                do
                {
                    int format = Filename.LastIndexOf(".");
                    if (format > 0)
                    {
                        int lastSpace = Filename.LastIndexOf(" ");
                        if (lastSpace > 0)
                        {
                            int countdel = ((Filename.Length - lastSpace) - (Filename.Length - format));
                            Filename = Filename.Remove(lastSpace, countdel);

                        }
                    }
                }
                while (Filename.Length > 200);
            }
            return Filename.Trim();
        }

        static public string DirMonth()
        {
            DateTime Now = DateTime.Now;
            string[] month = new string[] {"Январь","Февраль", "Март", "Апрель", "Май", "Июнь",
                "Июль", "Август", "Сентябрь","Ноябрь", "Октябрь", "Декабрь"};
            return month[Now.Month - 1];
        }

        static public string AddQout(string value)
        {// задалбался я кавычки подрисовывать
            if (value != null)
            {
                value = value.Insert(0, "\"");
                value = value.Insert(value.Length, "\"");
                return value;
            }
            return "null";
        }

    }
}
