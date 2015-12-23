using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SQLite;
using System.Data;

namespace DownMailTest
{
    class WorkSQLite
    {
        private SQLiteConnection sql_con;
        private SQLiteCommand sql_cmd;
        public string pathToBase;

        public WorkSQLite(string path)
        {
            pathToBase = path;
            SetConnect();
        }

        public void SetConnect()
        {
            sql_con = new SQLiteConnection("Data Source=" + pathToBase + ";Version=3;Timeout=10;", true);

        }

        public void ExecuteQuery(string txtQuery)
        {
            SetConnect();
            sql_con.Open();
            sql_cmd = sql_con.CreateCommand();
            sql_cmd.CommandText = txtQuery;
            sql_cmd.ExecuteNonQuery();
            sql_con.Close();
        }

        public DataTable GetTable(string query)
        {
            SetConnect();
            sql_con.Open();
            sql_cmd = sql_con.CreateCommand();
            DataTable table = new DataTable();
            using (SQLiteCommand command = new SQLiteCommand(query, sql_con))
            {
                SQLiteDataReader reader = command.ExecuteReader();
                table.Load(reader);
            }
            return table;
        }

        public void CloseConnect()
        {
            sql_con.Close();
        }
    }
}
