using System.Data.SQLite;
using System.Data;

namespace DownMailTest
{
    class BaseCreater
    {
       public string pathBase;

        public BaseCreater(string pathToBase)
        {
            SetPathBase(pathToBase);            
        }

        private void SetPathBase(string path)
        {
            pathBase = path;
        }

        private void CreateBase()
        {
            SQLiteConnection.CreateFile(pathBase);
            WorkSQLite work = new WorkSQLite(pathBase);
            work.ExecuteQuery("CREATE TABLE" + Func.AddQout("Messages (") +
                               Func.AddQout("id") + " INTEGER PRIMARY KEY  AUTOINCREMENT  NOT NULL  UNIQUE ," +
                               Func.AddQout("Subject") + " VARCHAR," + Func.AddQout("From_") + " VARCHAR, " +
                               Func.AddQout("Data") + " DATETIME," + Func.AddQout("idMessage") + " VARCHAR)");
        }
    }
}
