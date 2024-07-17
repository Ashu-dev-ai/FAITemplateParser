using Microsoft.Azure.WebJobs;
using Microsoft.Data.SqlClient;
using Microsoft.Extensions.Logging;
using Microsoft.TeamFoundation.Common;
using System;
using System.Data;
using System.IO;
using System.Threading.Tasks;

namespace FAITemplateParser.Utilities
{
    [System.Runtime.Versioning.SupportedOSPlatform("windows")]
    public class SqlHelper
    {
        // public static readonly string AppId = Environment.GetEnvironmentVariable("AzureADAppId", EnvironmentVariableTarget.Process);
        // public static readonly string AppSecret = Environment.GetEnvironmentVariable("AADAppValue", EnvironmentVariableTarget.Process);
        public static readonly string Server = Environment.GetEnvironmentVariable("SqlServer", EnvironmentVariableTarget.Process);
        public static readonly string DB = Environment.GetEnvironmentVariable("Database", EnvironmentVariableTarget.Process);
        public static readonly string Tenant = Environment.GetEnvironmentVariable("TenantId", EnvironmentVariableTarget.Process);
        public static string CreateSqlConnection(ILogger log = null, ExecutionContext context = null)
        {
            if (log != null)
            {
                log.LogInformation($"Creating Sql Connection");
            }
            try
            {
                /*IConfidentialClientApplication app = ConfidentialClientApplicationBuilder.Create(AppId)
                     .WithAuthority("https://login.windows.net/" + Tenant)
                     .WithClientSecret(AppSecret)
                     .WithLegacyCacheCompatibility(false)
                     .WithCacheOptions(CacheOptions.EnableSharedCacheOptions)
                     .Build();*/
                string sqlConnectionString = $"Server=tcp:{Server},1433;Initial Catalog={DB};Encrypt=True;TrustServerCertificate=False;Connection Timeout=30;Authentication='Active Directory Default'";

                return sqlConnectionString;
                /*AuthenticationResult result = app.AcquireTokenForClient(
                    new string[] { "https://database.windows.net//.default" })
               .ExecuteAsync().Result;
                string ConnectionString = String.Format(@"Data Source={0}; Initial Catalog={1};", Server, DB);
                SqlConnection sqlConnection = new SqlConnection(ConnectionString);
                sqlConnection.AccessToken = result.AccessToken;
                return sqlConnection;*/
            }
            catch (Exception ex)
            {
                if (log != null)
                {
                    log.LogInformation($"error creating connection {ex.Message}");
                    string module = System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString();
                    string message = $"Exception occured while creating sql connection : Exception: {ex.Message} \r\n StackTrace : {ex.StackTrace}" +
                                        $"\r\n WorkitemId: {Globals.WorkitemId} \r\n MSPN: {Globals.MSPN} \r\n SI: {Globals.SI} " +
                                        $"\r\n Filename: {Globals.FileName} \r\n Fileurl: {Globals.Fileurl}";
                    ExceptionHandler.LogError(new LogException() { RunId = Globals.runid, Module = module, Message = message });

                }
                throw;
            }
        }
        public static bool LoadDataTableToSql(DataTable dataTable, bool truncateandload, string destinationTableName)
        {
            bool result = false;
            try
            {
                if (truncateandload)
                {
                    SqlHelper.ExecuteSqlQuery_NoResult(String.Format("TRUNCATE TABLE {0}", destinationTableName));

                }
                using (SqlConnection sqlCon = new SqlConnection(CreateSqlConnection()))
                {
                    using (SqlBulkCopy sqlBulkCopy = new SqlBulkCopy(sqlCon))
                    {
                        sqlBulkCopy.DestinationTableName = destinationTableName;
                        sqlBulkCopy.BulkCopyTimeout = 1000;
                        if (sqlCon.State == ConnectionState.Closed)
                        {
                            sqlCon.Open();
                        }
                        sqlBulkCopy.WriteToServer(dataTable);
                    }
                }
                result = true;
            }
            catch (Exception ex)
            {
                //SqlHelper.ExecuteSqlQuery_NoResult(String.Format("INSERT INTO [dbo].[FAI_Templates_ErrorLogs]([AttachmentId],[ErrorLocation],[ErrorMessage]) VALUES ({0},'{1}','{2}')", -1, String.Concat(System.Reflection.MethodBase.GetCurrentMethod().Name, " : ", destinationTableName), ex.Message.Replace("'", "''")));
                ////throw;
                string module = System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString();
                string message = $"Exception occured while writing data to table : Exception: {ex.Message} \r\n StackTrace : {ex.StackTrace}" +
                                    $"\r\n WorkitemId: {Globals.WorkitemId} \r\n MSPN: {Globals.MSPN} \r\n SI: {Globals.SI} " +
                                    $"\r\n Filename: {Globals.FileName} \r\n Fileurl: {Globals.Fileurl}";
                ExceptionHandler.LogError(new LogException() { RunId = Globals.runid, Module = module, Message = message });

            }
            return result;
        }
        public static DataTable ExecuteSqlQuery(string query, ILogger log = null, ExecutionContext context = null)
        {
            if (log != null)
            {
                log.LogInformation($"Entered ExecuteSqlQuerySection");
            }
            DataTable result = new DataTable();
            try
            {
                using (SqlConnection sqlCon = new SqlConnection(CreateSqlConnection()))
                {
                    using (SqlCommand sqlCommand = new SqlCommand(query, sqlCon))
                    {
                        using (SqlDataAdapter sqlDataAdapter = new SqlDataAdapter(sqlCommand))
                        {
                            if (sqlCon.State == ConnectionState.Closed)
                            {
                                sqlCon.Open();
                            }
                            sqlCommand.CommandTimeout = 1000;
                            sqlDataAdapter.Fill(result);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                if (log != null)
                {
                    log.LogInformation($"Error in ExecuteSqlQuerySection {ex.Message}");

                    string module = System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString();
                    string message = $"Exception occured in Execute query : Exception: {ex.Message} \r\n StackTrace : {ex.StackTrace}" +
                                        $"\r\n WorkitemId: {Globals.WorkitemId} \r\n MSPN: {Globals.MSPN} \r\n SI: {Globals.SI} " +
                                        $"\r\n Filename: {Globals.FileName} \r\n Fileurl: {Globals.Fileurl}";
                    ExceptionHandler.LogError(new LogException() { RunId = Globals.runid, Module = module, Message = message });

                }
                //throw;
            }
            return result;
        }

        public static void ExecuteSqlQuery_NoResult(string query)
        {
            try
            {
                using (SqlConnection sqlCon = new SqlConnection(CreateSqlConnection()))
                {
                    using (SqlCommand sqlCommand = new SqlCommand(query, sqlCon))
                    {
                        using (SqlDataAdapter sqlDataAdapter = new SqlDataAdapter(sqlCommand))
                        {
                            if (sqlCon.State == ConnectionState.Closed)
                            {
                                sqlCon.Open();
                            }
                            sqlCommand.CommandTimeout = 1000;
                            sqlCommand.ExecuteNonQuery();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                string module = System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString();
                string message = $"Exception occured in executing sql query : Exception: {ex.Message} \r\n StackTrace : {ex.StackTrace}" +
                                    $"\r\n WorkitemId: {Globals.WorkitemId} \r\n MSPN: {Globals.MSPN} \r\n SI: {Globals.SI} " +
                                    $"\r\n Filename: {Globals.FileName} \r\n Fileurl: {Globals.Fileurl}";
                ExceptionHandler.LogError(new LogException() { RunId = Globals.runid, Module = module, Message = message });

                //throw;
            }
        }


        /// <summary>
        /// Executes a Stored Procedure and returns dataset as result.
        /// Dataset may contain one or more tables.
        /// </summary>
        /// <param name="SPName"></param>
        /// <param name="sp_parameters"></param>
        /// <param name="appid"></param>
        /// <param name="appsecret"></param>
        /// <param name="servername"></param>
        /// <param name="dbName"></param>
        /// <param name="tenantId"></param>
        /// <param name="executionContext"></param>
        /// <param name="logger"></param>
        /// <returns></returns>
        public static DataSet ExecuteStoredProcedureandReturnDatasetResult(string SPName, Microsoft.Data.SqlClient.SqlParameter[] sp_parameters = null, string appid = null, string appsecret = null, string servername = null, string dbName = null, string tenantId = null, ExecutionContext executionContext = null, ILogger logger = null)
        {

            DataSet sqlResult = new DataSet();

            try
            {
                using (SqlConnection conn = new SqlConnection(CreateSqlConnection()))
                {
                    using (SqlCommand sqlcmd = new SqlCommand(SPName, conn))
                    {
                        int timeout = 1800;
                        sqlcmd.CommandTimeout = timeout;
                        sqlcmd.CommandType = CommandType.StoredProcedure;
                        if (conn.State == ConnectionState.Closed)
                        {
                            conn.Open();
                        }

                        if (sp_parameters != null)
                        {
                            sqlcmd.Parameters.AddRange(sp_parameters);
                        }
                        SqlDataAdapter da = new SqlDataAdapter(sqlcmd);

                        da.Fill(sqlResult);
                        da.Dispose();
                    }
                }
            }
            catch (Exception ex)
            {
                string module = System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString();
                // string message = $"Exception occured while Execute SP : Exception: {ex.Message} \r\n StackTrace : {ex.StackTrace} \r\n server: {Server} \r\n DB : {DatabaseName} \r\n SP Name : {SPName}";

                if (logger != null)
                {
                    logger.LogInformation($"Error in ExecuteStoredProcedureandReturnDatasetResult {ex.Message}");
                    string message = $"Exception occured while Executing SP : Exception: {ex.Message} \r\n StackTrace : {ex.StackTrace}" +
                                        $"\r\n WorkitemId: {Globals.WorkitemId} \r\n MSPN: {Globals.MSPN} \r\n SI: {Globals.SI} " +
                                        $"\r\n Filename: {Globals.FileName} \r\n Fileurl: {Globals.Fileurl}";
                    ExceptionHandler.LogError(new LogException() { RunId = Globals.runid, Module = module, Message = message });

                }


                throw ex;
            }
            return sqlResult;
        }

        public static async Task StreamBLOBToServer(Stream stream, int AttachmentId, string filename, string emdfileName, string filetype, string path, string _Text, string TableName = null)
        {
            string tableName = TableName.IsNullOrEmpty() ? "Stage_FAI_Templates_EmbeddedData" : TableName;
            using (SqlConnection conn = new SqlConnection(CreateSqlConnection()))
            {
                await conn.OpenAsync();
                try
                {
                    using (SqlCommand cmd = new SqlCommand(String.Format(@"INSERT INTO [dbo].[{0}] ([FileData],[AttachmentId],[AttachmentName]
                                                                     ,[FileType],[FilePath],[FileName],[FileText])
                                                                    VALUES (@bindata,@AttachmentId,@attachmentName,@filetype,@filepath,@filename,@fileText)"
                                                , tableName), conn))
                    {
                        // Add a parameter which uses the FileStream we just opened
                        // Size is set to -1 to indicate "MAX"
                        cmd.Parameters.Add("@bindata", SqlDbType.Binary, -1).Value = stream;
                        cmd.Parameters.Add("@AttachmentId", SqlDbType.Int, -1).Value = AttachmentId;
                        cmd.Parameters.Add("@filetype", SqlDbType.VarChar, 800).Value = filetype;
                        cmd.Parameters.Add("@filepath", SqlDbType.VarChar, 800).Value = path;
                        cmd.Parameters.Add("@filename", SqlDbType.VarChar, 800).Value = emdfileName;
                        cmd.Parameters.Add("@attachmentName", SqlDbType.VarChar, 800).Value = filename;
                        cmd.Parameters.Add("@fileText", SqlDbType.VarChar, 800).Value = _Text;
                        // Send the data to the server asynchronously
                        await cmd.ExecuteNonQueryAsync();
                    }
                }
                catch (Exception ex)
                {
                    //using (SqlCommand cmd = new SqlCommand("INSERT INTO [dbo].[___FWTTestResultsLogDataStorage_Errors] ([FWTAttachmentId],[ErrorMessage]) VALUES (@FWTAttachmentId,@errorMessage)", conn))
                    //{
                    //    cmd.Parameters.Add("@FWTAttachmentId", SqlDbType.Int, -1).Value = AttachmentId;
                    //    cmd.Parameters.Add("@FWTAttachmentId", SqlDbType.VarChar, 400).Value = ex.Message.ToString();
                    //    await cmd.ExecuteNonQueryAsync();
                    //}

                    string module = System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString();
                    string message = $"Exception occured while writing stream Blob to server : Exception: {ex.Message} \r\n StackTrace : {ex.StackTrace}" +
                                        $"\r\n WorkitemId: {Globals.WorkitemId} \r\n MSPN: {Globals.MSPN} \r\n SI: {Globals.SI} " +
                                        $"\r\n Filename: {Globals.FileName} \r\n Fileurl: {Globals.Fileurl}";
                    ExceptionHandler.LogError(new LogException() { RunId = Globals.runid, Module = module, Message = message });

                    //throw;
                }
            }
        }
    }
}
