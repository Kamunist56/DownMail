using System.Data.SQLite;

namespace DownMailTest
{
    class BaseCreator
    {
       public string pathBase;

        public BaseCreator(string pathToBase)
        {
            SetPathBase(pathToBase);
            CreateBase();
        }

        private void SetPathBase(string path)
        {
            pathBase = path;
        }

        private void CreateBase()
        {
            SQLiteConnection.CreateFile(pathBase);
        }

        public void CreateTables()
        {
            WorkSQLite work = new WorkSQLite(pathBase);

            work.ExecuteQuery("CREATE TABLE " + Func.AddQout("Messages") +"("
                                + Func.AddQout("id") + " INTEGER PRIMARY KEY  AUTOINCREMENT  NOT NULL  UNIQUE ," 
                                + Func.AddQout("Subject") + " VARCHAR," 
                                + Func.AddQout("From_") + " VARCHAR, " 
                                + Func.AddQout("Data") + " DATETIME," 
                                + Func.AddQout("idMessage") + " VARCHAR,"
                                + Func.AddQout("PathMessage") + "VARCHAR)");

            work.ExecuteQuery("CREATE TABLE " + Func.AddQout("Hosts") + "("
                                + Func.AddQout("id") + " INTEGER PRIMARY KEY  AUTOINCREMENT  NOT NULL  UNIQUE ,"
                                + Func.AddQout("login") + " VARCHAR, "
                                + Func.AddQout("pass") + " VARCHAR, "
                                + Func.AddQout("port") + " INTEGER, "
                                + Func.AddQout("host") + " VARCHAR)");

            work.ExecuteQuery("CREATE TABLE " + Func.AddQout("Settings") + "("
                                + Func.AddQout("id") + " INTEGER PRIMARY KEY  AUTOINCREMENT  NOT NULL  UNIQUE ,"
                                + Func.AddQout("path") + " VARCHAR, "
                                + Func.AddQout("interval") + " VARCHAR)");
        }
    }
}
