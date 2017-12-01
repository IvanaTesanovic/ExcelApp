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

        //    State values:
        //    DoesNotApply = 0,
        //    Untranslated = 1,
        //    Submitted = 2,
        //    Accepted = 3,
        //    Denied = 4

        public List<Disease> GetRangeOfDiseases(int range)
        {
            using (var db = GetConnection())
            {
                // Add join Summary & Synonym tables to this get call

                return db.Query<Disease>($"SELECT TOP {range} * FROM Disease WHERE State = {0}").ToList();
            }
        }

        public void UpdateExportedDiseases(params string[] orphaNumbers)
        {
            using (var db = GetConnection())
            {
                db.Query<Disease>($"UPDATE Disease SET State = {2} WHERE OrphaNumber IN({orphaNumbers}) AND State != {0}");
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
