using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Configuration;
using System.Collections;
using NPoco;
using System.Data.SqlClient;
using ExcelTestApp.Entities;

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

        public List<DiseaseEntity> GetDiseases(int count)
        {
            using (var db = GetConnection())
            {
                return db.Query<DiseaseEntity>($"SELECT TOP {count} * FROM Disease WHERE State = {0} ORDER BY newid()").ToList();
            }
        }

        public void UpdateExportedDiseases(params string[] diseaseIds)
        {
            string ids = string.Join(",", diseaseIds.Select(id => "'" + id.ToString() + "'").ToArray());
            using (var db = GetConnection())
            {
                db.Query<DiseaseEntity>($"UPDATE Disease SET State = {1} WHERE Id IN({ids})");
            }
        }

        public List<DiseaseEntity> GetDiseaseByOrpha(string orpha)
        {
            using (var db = GetConnection())
            {
                return db.Query<DiseaseEntity>($"SELECT * FROM Disease WHERE OrphaNumber = {orpha}").ToList();
            }
        }

        public List<SynonymEntity> GetSynonymsByDiseaseId(string diseaseId)
        {
            using (var db = GetConnection())
            {
                return db.Fetch<SynonymEntity>($"where DiseaseId = '{diseaseId}'");
            }
        }

        public List<SummaryEntity> GetSummariesByDiseaseId(string diseaseId)
        {
            using (var db = GetConnection())
            {
                return db.Fetch<SummaryEntity>($"where DiseaseId = '{diseaseId}'");
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
