using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Spartan.DAL.Geral.Conexao
{
    public static class ConnectionOpen
    {
        public static SqlConnection GetOpenConnection()
        {
            var connection = new SqlConnection(ConfigurationManager.ConnectionStrings["SerasaDB"].ConnectionString);

            connection.Open();

            return connection;
        }
    }
}
