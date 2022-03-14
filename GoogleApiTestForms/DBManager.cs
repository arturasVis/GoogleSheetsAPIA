using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SqlClient;
using System.Data;
using System.Configuration;

namespace GoogleApiTestForms
{
    internal class DBManager
    {
        SqlConnection _connection;
        string sql= ConfigurationManager.ConnectionStrings["GoogleApiTestForms.Properties.Settings.gpcdbConnectionString2"].ConnectionString;

        

        public void InsertQuerry(string querry)
        {
            try
            {
                using (_connection = new SqlConnection(sql))
                using (var command = _connection.CreateCommand())
                {
                    _connection.Open();
                    command.CommandText = querry;
                    command.ExecuteNonQuery();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            
        }

        public DataTable Select(string table)
        {
            DataTable dt = new DataTable();
            string querry = $"SELECT * FROM {table}";
            using (_connection = new SqlConnection(sql))
            using (SqlDataAdapter adapter = new SqlDataAdapter(querry, _connection))
            {
                _connection.Open();
                adapter.Fill(dt);
            }
            return dt;
        }
        public DataTable SelectDate(string table,string date)
        {
            DataTable dt = new DataTable();
            string querry = $"SELECT * FROM {table} WHERE Date='{date}'";
            using (_connection = new SqlConnection(sql))
            using (SqlDataAdapter adapter = new SqlDataAdapter(querry, _connection))
            {
                _connection.Open();
                adapter.Fill(dt);
            }
            return dt;
        }
    }
}
