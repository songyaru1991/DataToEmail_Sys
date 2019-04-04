using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using Oracle.ManagedDataAccess.Client;
using MySql.Data.MySqlClient;

namespace DAL
{
    public class DBHelper
    {
        public DataTable QueryOracle(string sql) 
        {
            string connStr = @"Data Source=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=192.168.78.154)(PORT=1521))(CONNECT_DATA=(SERVICE_NAME=SSGSDB)));User Id=swipe;Password=a2#ks#ssgs";
                using(OracleConnection conn=new OracleConnection(connStr))
                {
                    conn.Open();
                    OracleDataAdapter da = new OracleDataAdapter(sql, conn);
                    DataTable table = new DataTable();
                    da.Fill(table);
                    return table;
                }
        }

        public DataTable QueryMysql(string sql) 
        {
            string connStr = "Server=192.168.65.230;Database=swipecard;User=root; Password=foxlink";
            using (MySqlConnection conn = new MySqlConnection(connStr))
            {
                conn.Open();
                MySqlDataAdapter da = new MySqlDataAdapter(sql, conn);
                DataTable table = new DataTable();
                da.Fill(table);
                return table;
            }
        }
    }
}
