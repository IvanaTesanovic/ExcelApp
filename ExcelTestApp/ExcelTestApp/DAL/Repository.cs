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
using System.Data;

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
                return db.Query<DiseaseEntity>($"SELECT TOP {count} * FROM Disease WHERE State = 0 ORDER BY newid()").ToList();
            }
        }

        public void Update(List<DiseaseEntity> diseases)
        {
            using (var db = GetConnection())
            {
                foreach(DiseaseEntity disease in diseases)
                {
                    db.Update(disease);
                }
            }
        }

        public List<DiseaseEntity> GetDiseaseByOrpha(string orpha)
        {
            using (var db = GetConnection())
            {
                return db.Fetch<DiseaseEntity>($"WHERE OrphaNumber = {orpha}").ToList();
            }
        }

        public DiseaseEntity GetOriginalDiseaseByOrpha(string orpha)
        {
            using (var db = GetConnection())
            {
                return db.Fetch<DiseaseEntity>($"WHERE OrphaNumber = {orpha} AND IsTranslationOf is NULL").FirstOrDefault();
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

        public void InsertDisease(DiseaseEntity disease, IEnumerable<SynonymEntity> synonyms, IEnumerable<SummaryEntity> summaries)
        {
            using (var db = GetConnection())
            {
                db.BeginTransaction(IsolationLevel.ReadCommitted);

                db.Insert(disease);
                db.InsertBatch(synonyms);
                db.InsertBatch(summaries);

                db.CompleteTransaction();
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
