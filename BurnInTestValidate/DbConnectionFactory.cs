using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static BurnInTestValidate.Program;

namespace BurnInTestValidate
{
    public class DbConnectionFactory
    {
        //private readonly string _connectionString;

        //public DbConnectionFactory(string connectionString)
        //{
        //    _connectionString = connectionString;
        //}
        private readonly Dictionary<DatabaseType, string> _connections;

        public DbConnectionFactory(Dictionary<DatabaseType, string> connections)
        {
            _connections = connections;
        }

        public SqlConnection CreateConnection(DatabaseType dbType)
        {
            return new SqlConnection(_connections[dbType]);
        }
        //public SqlConnection CreateConnection()
        //{
        //    return new SqlConnection(_connectionString);
        //}
    }
}
