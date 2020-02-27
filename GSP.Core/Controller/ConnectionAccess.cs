using System;
using System.Data;
using System.Data.OleDb;

namespace GSP.Core.Controller
{
    public class ConnectionAccess
    {
        public OleDbConnection OleDbConnection { get; set; }

        public string ConnectionAccessString { get; }

        public ConnectionAccess(string nameAccessDataBase)
        {
            ConnectionAccessString = $"Provider=Microsoft.Jet.OLEDB.4.0; Data Source={nameAccessDataBase};";
            OpenConnectionAccess();
        }

        public bool OpenConnectionAccess()
        {
            try
            {
                OleDbConnection = new OleDbConnection(ConnectionAccessString);
                OleDbConnection.Open();
                return true;
            }
            catch (Exception ex)
            {
                return false;
            }
        }

        public bool CloseConnectionAccess()
        {
            try
            {                
                OleDbConnection.Close();
                return true;
            }
            catch (Exception ex)
            {
                return false;
            }
        }

        public bool Injection(string query)
        {
            try
            {
                using (var cmd = new OleDbCommand() { Connection = OleDbConnection, CommandType = CommandType.Text, CommandText = query })
                {
                    cmd.ExecuteNonQuery();
                }
                return true;
            }
            catch (Exception ex)
            {
                return false;
            }
           
        }
    }
}
