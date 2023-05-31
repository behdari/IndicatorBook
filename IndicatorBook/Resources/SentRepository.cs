using Dapper;
using Newtonsoft.Json;
using System.Data.SqlClient;

namespace IndicatorBook.Resources
{
    public class SentRepository
    {
        private string connectionString = "Data Source={ServerIp},1433;Initial Catalog=PhoneBookDb;Persist Security Info=True;User ID=sa;Password=12qwAS!@";

        private static SentRepository _instance;
        private static Config config;

        public static SentRepository Instance
        {
            get { return _instance ??= new SentRepository(); }
        }

        static SentRepository()
        {
            LoadJson();
        }

        static void LoadJson()
        {
            using (StreamReader r = new StreamReader("config.json"))
            {
                string json = r.ReadToEnd();
                config = JsonConvert.DeserializeObject<Config>(json);
            }
        }

        private string GetConnectionString()
        {
            var generatedConnectionString = connectionString.Replace("{ServerIp}", config.ServerIp);

            return generatedConnectionString;
        }

        public bool IsExist(SentLetter sentLetter)
        {
            return IsExist(sentLetter.File, sentLetter.Pamphleteer, sentLetter.Number);
        }

        public bool IsExist(string file, string pamphleteer, int sentLetterNumber)
        {
            var sentLetter = GetSentLetter(file, pamphleteer, sentLetterNumber);

            var isExist = sentLetter != null;

            return isExist;
        }

        public SentLetter GetSentLetterById(int id)
        {
            var existQuery = $@"
    SELECT *
  FROM [IndicatorBook].[dbo].[Sent]
  where [Id] = N'{id}'";

            SentLetter result;

            using (var connection = new SqlConnection(GetConnectionString()))
            {
                result = connection.QuerySingleOrDefault<SentLetter>(existQuery);
            }

            return result;
        }

        internal List<SentLetter> SearchByWord(string text)
        {
            var existQuery = $@"SELECT TOP (1000) [Id]
                                      ,[File]
                                      ,[Pamphleteer]
                                      ,[Number]
                                      ,[Title]
                                      ,[Description]
                                      ,[NextNumber]
                                      ,[HasAttachment]
                                      ,[Date]
                                  FROM [IndicatorBook].[dbo].[Sent]
                                  WHERE
                                      [File] like N'%{text}%'
                                      OR [Pamphleteer] like N'%{text}%'
                                      OR [Number] like N'%{text}%'
                                      OR [Title] like N'%{text}%'
                                      OR [Description] like N'%{text}%'
                                      OR [NextNumber] like N'%{text}%'
                                      OR [Date] like N'%{text}%'
                                  ORDER BY Id DESC";

            List<SentLetter> result;

            using (var connection = new SqlConnection(GetConnectionString()))
            {
                result = connection.Query<SentLetter>(existQuery).ToList();
            }

            return result;
        }

        internal void Add(SentLetter entity)
        {
            var insertQuery = $@"INSERT INTO [dbo].[Sent]
                                        ([File]
                                        ,[Pamphleteer]
                                        ,[Number]
                                        ,[Title]
                                        ,[Description]
                                        ,[NextNumber]
                                        ,[HasAttachment]
                                        ,[Date]
                                        ,[WordFile]
                                        ,[WordFileExtension])
                                  VALUES
                                        (@File
                                        ,@Pamphleteer
                                        ,@Number
                                        ,@Title
                                        ,@Description
                                        ,@NextNumber
                                        ,@HasAttachment
                                        ,@Date
                                        ,@WordFile
                                        ,@WordFileExtension)";

            using (var connection = new SqlConnection(GetConnectionString()))
            {
                var result = connection.Execute(insertQuery, new
                {
                    entity.Id,
                    entity.File,
                    entity.Pamphleteer,
                    entity.Number,
                    entity.Title,
                    entity.Description,
                    entity.NextNumber,
                    entity.HasAttachment,
                    entity.Date,
                    entity.WordFile,
                    entity.WordFileExtension
                });
            }
        }

        internal void Delete(int id)
        {
            var deleteQuery = $@"
                            DELETE FROM [dbo].[Sent]
                                  WHERE [Id] = @Id";

            using (var connection = new SqlConnection(GetConnectionString()))
            {
                var result = connection.Execute(deleteQuery, new
                {
                    id
                });
            }
        }

        internal void Update(SentLetter entity)
        {
            var insertQuery = $@"
                        UPDATE [dbo].[Sent]
                           SET [File] = @File 
                              ,[Pamphleteer] = @Pamphleteer 
                              ,[Number] = @Number 
                              ,[Title] = @Title 
                              ,[Description] = @Description 
                              ,[NextNumber] = @NextNumber 
                              ,[HasAttachment] = @HasAttachment 
                              ,[Date] = @Date 
                              ,[WordFile] = @WordFile 
                              ,[WordFileExtension] = @WordFileExtension 
                         WHERE [Id] = @Id";

            using (var connection = new SqlConnection(GetConnectionString()))
            {
                var result = connection.Execute(insertQuery, new
                {
                    entity.Id,
                    entity.File,
                    entity.Pamphleteer,
                    entity.Number,
                    entity.Title,
                    entity.Description,
                    entity.NextNumber,
                    entity.HasAttachment,
                    entity.Date,
                    entity.WordFile,
                    entity.WordFileExtension
                });
            }
        }

        internal SentLetter GetSentLetter(string file, string pamphleteer, int sentLetterNumber)
        {
            var existQuery = $@"SELECT *
                                FROM [IndicatorBook].[dbo].[Sent]
                                where [File] = N'{file}' AND
                                Pamphleteer = N'{pamphleteer}' AND
                                Number = N'{sentLetterNumber}'";

            SentLetter result;

            using (var connection = new SqlConnection(GetConnectionString()))
            {
                result = connection.QuerySingleOrDefault<SentLetter>(existQuery);
            }

            return result;
        }
    }
}
