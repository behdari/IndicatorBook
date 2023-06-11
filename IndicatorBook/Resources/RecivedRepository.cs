using Dapper;
using Newtonsoft.Json;
using System.Data.SqlClient;

namespace IndicatorBook.Resources
{
    public class RecivedRepository
    {
        private string connectionString = "Data Source={ServerIp},1433;Initial Catalog=IndicatorBook;Persist Security Info=True;User ID=sa;Password=12qwAS!@";

        private static RecivedRepository _instance;
        private static Config config;

        public static RecivedRepository Instance
        {
            get { return _instance ??= new RecivedRepository(); }
        }

        static RecivedRepository()
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

        public bool IsExist(RecivedLetter recievedLetter)
        {
            return IsExist(recievedLetter.RowNumber);
        }

        public bool IsExist(int rowNumber)
        {
            var recivedLetter = GetByRowNumber(rowNumber);

            var isExist = recivedLetter != null;

            return isExist;
        }

        public RecivedLetter GetRecivedLetterById(int id)
        {
            var existQuery = $@"SELECT *
                                FROM [IndicatorBook].[dbo].[Recived]
                                where [Id] = N'{id}'";

            RecivedLetter result;

            using (var connection = new SqlConnection(GetConnectionString()))
            {
                result = connection.QuerySingleOrDefault<RecivedLetter>(existQuery);
            }

            return result;
        }

        internal List<RecivedLetter> SearchByWord(string text)
        {
            var existQuery = $@"SELECT TOP (100) [Id]
                                      ,[RowNumber]
                                      ,[PreviousRowNumber]
                                      ,[Date]
                                      ,[LetterOwners]
                                      ,[Description]
                                      ,[HasAttachment]
                                      ,[RecivedLetterNumber]
                                      ,[RecivedLetterDate]
                                  FROM [IndicatorBook].[dbo].[Recived]
                                  WHERE
                                      [RowNumber] like N'%{text}%'
                                      OR [PreviousRowNumber] like N'%{text}%'
                                      OR [Date] like N'%{text}%'
                                      OR [LetterOwners] like N'%{text}%'
                                      OR [RecivedLetterNumber] like N'%{text}%'
                                      OR [RecivedLetterDate] like N'%{text}%'
                                      OR [Description] like N'%{text}%'
                                  ORDER BY Id DESC";

            List<RecivedLetter> result;

            using (var connection = new SqlConnection(GetConnectionString()))
            {
                result = connection.Query<RecivedLetter>(existQuery).ToList();
            }

            return result;
        }

        internal void Add(RecivedLetter entity)
        {
            var insertQuery = $@"INSERT INTO [dbo].[Recived]
                                        ([RowNumber]
                                        ,[PreviousRowNumber]
                                        ,[Date]
                                        ,[LetterOwners]
                                        ,[Description]
                                        ,[HasAttachment]
                                        ,[RecivedLetterNumber]
                                        ,[RecivedLetterDate]
                                        ,[ScanFile]
                                        ,[ScanFileExtension])
                                  VALUES
                                        (@RowNumber
                                        ,@PreviousRowNumber
                                        ,@Date
                                        ,@LetterOwners
                                        ,@Description
                                        ,@HasAttachment
                                        ,@RecivedLetterNumber
                                        ,@RecivedLetterDate
                                        ,@ScanFile
                                        ,@ScanFileExtension)";

            using (var connection = new SqlConnection(GetConnectionString()))
            {
                var result = connection.Execute(insertQuery, new
                {
                    entity.RowNumber,
                    entity.PreviousRowNumber,
                    entity.Date,
                    entity.LetterOwners,
                    entity.Description,
                    entity.HasAttachment,
                    entity.RecivedLetterNumber,
                    entity.RecivedLetterDate,
                    entity.ScanFile,
                    entity.ScanFileExtension
                });
            }
        }

        internal void Delete(int id)
        {
            var deleteQuery = $@"
                            DELETE FROM [dbo].[Recived]
                                  WHERE [Id] = @Id";

            using (var connection = new SqlConnection(GetConnectionString()))
            {
                var result = connection.Execute(deleteQuery, new
                {
                    id
                });
            }
        }

        internal void Update(RecivedLetter entity)
        {
            var insertQuery = $@"
                        UPDATE [dbo].[Recived]
                           SET [RowNumber] = @RowNumber 
                              ,[PreviousRowNumber] = @PreviousRowNumber 
                              ,[Date] = @Date 
                              ,[LetterOwners] = @LetterOwners 
                              ,[Description] = @Description 
                              ,[HasAttachment] = @HasAttachment 
                              ,[RecivedLetterNumber] = @RecivedLetterNumber 
                              ,[RecivedLetterDate] = @RecivedLetterDate 
                              ,[ScanFile] = @ScanFile 
                              ,[ScanFileExtension] = @ScanFileExtension 
                         WHERE [Id] = @Id";

            using (var connection = new SqlConnection(GetConnectionString()))
            {
                var result = connection.Execute(insertQuery, new
                {
                    entity.Id,
                    entity.RowNumber,
                    entity.PreviousRowNumber,
                    entity.Date,
                    entity.LetterOwners,
                    entity.Description,
                    entity.HasAttachment,
                    entity.RecivedLetterNumber,
                    entity.RecivedLetterDate,
                    entity.ScanFile,
                    entity.ScanFileExtension
                });
            }
        }

        internal RecivedLetter GetByRowNumber(int rowNumber)
        {
            var existQuery = $@"SELECT *
                                FROM [IndicatorBook].[dbo].[Recived]
                                where RowNumber = N'{rowNumber}'";

            RecivedLetter result;

            using (var connection = new SqlConnection(GetConnectionString()))
            {
                result = connection.QuerySingleOrDefault<RecivedLetter>(existQuery);
            }

            return result;
        }
    }
}
