using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Configuration;
using System.Collections;
using NPoco;
using System.Data.SqlClient;

namespace ExcelTestApp
{
    public class Repository
    {
        private readonly string _connectionString;

        public Repository()
        {
            _connectionString = ConfigurationManager.AppSettings["TranslationsRepositoryConnectionString"];
        }

        public List<Disease> GetRangeOfDiseases(int range)
        {
            using (var db = GetConnection())
            {
                return db.Query<Disease>($"SELECT TOP {range} * FROM Disease").ToList();
            }
        }

        public IDatabase GetConnection(bool enableAutoSelect = true)
        {
            var connection = new SqlConnection(_connectionString);

            connection.Open();

            return new Database(connection)
            {
                EnableAutoSelect = enableAutoSelect
            };
        }
    }
}
