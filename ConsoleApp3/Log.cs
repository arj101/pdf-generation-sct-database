using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApp3
{
    sealed class Log
    {
        SqlConnection conn;
        Log() {
            string username = Environment.UserDomainName;
            string connectionString = $"Data Source={username}\\SQLEXPRESS;Initial Catalog=idpass;Integrated Security=True";
            conn = new SqlConnection(connectionString);
        }
        private static Log instance;
        public void AppendLog(string name) {
            this.conn.Open();
            using (SqlCommand cmd = new SqlCommand(@"INSERT INTO[dbo].[logfile] ([username], [date], [time]) VALUES (@Username, @Date, @Time)", conn))
            {
                cmd.Parameters.AddWithValue("@Username", name);
                cmd.Parameters.AddWithValue("@Date", DateTime.Now.Date.ToString());
                cmd.Parameters.AddWithValue("@Time", DateTime.Now.TimeOfDay.ToString());
                cmd.ExecuteNonQuery();
            }
            this.conn.Close();
        }
        public static Log Instance { get {
                if (instance == null)
                {
                    instance = new Log();
                }
                return instance;

            } }

        public static Action<string> log { get {
                if (instance == null)
                {
                    instance = new Log();
                }
                return instance.AppendLog;

            } }
    }
}
