using Dapper;
using Spartan.DAL.Geral.Conexao;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Spartan.DAL.Geral.DataContextDapper
{
    public class ContextSpartan
    {
        public int ExecutarProcedure(string nomeProcedure)
        {
            using (var connection = ConnectionOpen.GetOpenConnection())
            {
                return connection.Execute(
                    nomeProcedure, commandType: CommandType.StoredProcedure);
            }
        }
    }
}
