using Dapper;
using Microsoft.Data.SqlClient;

namespace ExporterLogicLibrary
{
    public static class SqlLogic
    {
        public static List<string> GetDatabasesFromServer(string server)
        {
            List<string> output;

            using (SqlConnection cnn = GetOpenConnection(server))
            {
                output = cnn.Query<string>("SELECT name FROM sys.databases").AsList();
            }

            return output;
        }

        private static SqlConnection GetOpenConnection(string server)
        {
            if (server == string.Empty)
                throw new ArgumentException(Properties.Resources.EXP_SERVER_MISSING);

            string connectionString = $@"Data Source={server};Integrated Security=SSPI;TrustServerCertificate=true;";

            SqlConnection cnn = new(connectionString);
            cnn.Open();

            return cnn;
        }
    }
}
