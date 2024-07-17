using Azure.Storage.Blobs;
using Azure.Storage.Blobs.Models;
using FAITemplateParser.Utilities;
using Microsoft.Azure.WebJobs;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Threading.Tasks;

namespace FAITemplateParser
{
    public class FAITemplateParser
    {
        public static string tempinputfolderpath = Environment.GetEnvironmentVariable("InputPath", EnvironmentVariableTarget.Process);
        public static string tempoutputfolderpath = Environment.GetEnvironmentVariable("OutputPath", EnvironmentVariableTarget.Process);
        public static DataSet dsMasterData = null;
        public static DataSet dsMasterDataprocess = null;

        [System.Runtime.Versioning.SupportedOSPlatform("windows")]
        //[Timeout("05:00:00")]
        [FunctionName("FAIParseTemplates")]



        public async Task Run([TimerTrigger("0 */30 * * * *")] TimerInfo myTimer, ILogger log, ExecutionContext context)
        //public void Run([TimerTrigger("0 */30 * * * *")] TimerInfo myTimer, ILogger log, ExecutionContext context)
        {
            string inpath = tempinputfolderpath;
            string output = tempoutputfolderpath;

            log.LogInformation($"FAITemplateParser started at: {DateTime.Now}");
            try
            {
                //Load Sharepoint metadata for SI Sites to load FAI templates

                Globals.InitializeGlobals();
                string module = System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString();

                ExceptionHandler.LogMessage(new LogException() { RunId = Globals.runid, Module = module, Message = $"FAI Template Parser Automation Run Started", Operation = "RunStart" });


                Microsoft.Data.SqlClient.SqlParameter[] sp_params = new Microsoft.Data.SqlClient.SqlParameter[1];
                Microsoft.Data.SqlClient.SqlParameter param = new Microsoft.Data.SqlClient.SqlParameter();
                param.ParameterName = "RunId";
                param.Value = Globals.runid;
                sp_params.SetValue(param, 0);

                dsMasterData = SqlHelper.ExecuteStoredProcedureandReturnDatasetResult("[dbo].[usp_FAI_Templates_GetWorkItemsToProcess]", sp_params);




                /*param.ParameterName = "RunId";
                param.Value = Globals.runid;
                sp_params.SetValue(param, 0);
                string RunId = string.Empty;
                string WorkItemId = string.Empty;
                string MSPN = string.Empty;
                string SI = string.Empty;
                string Siteurl = string.Empty;
                string ClientId = string.Empty;
                string ClientSecretkey = string.Empty;
                string ClientSecret = Environment.GetEnvironmentVariable(ClientSecretkey, EnvironmentVariableTarget.Process);
                string SPListName = string.Empty;
                string serverRelativeFolderurl = string.Empty;
                string loadembedded = string.Empty;
                string ExcelReadOptions = string.Empty;
                SharepointFIle spfile = new SharepointFIle();
                spfile.RunId = RunId;
                spfile.Workitemid = WorkItemId;
                spfile.MSPN = MSPN;
                spfile.SI = SI;
                spfile.LoadEmbedded = loadembedded;
                spfile.ExcelReadOptions = ExcelReadOptions;

                var spfilelist = SharePointHelper.ReadFilesFromSharepointAsStream(log, Siteurl, ClientId, ClientSecret, SPListName, serverRelativeFolderurl, spfile, context: context);
*/
                // Execute SP usp_FAI_Templates_GetWorkItemsToProcess to get work items to process FAI templates
                string connectionString = Environment.GetEnvironmentVariable("AzureBlobConnectionString", EnvironmentVariableTarget.Process);
                string container = "sharepointdata";
                BlobContainerClient blobContainerClient = new BlobContainerClient(connectionString, container);

                BlobServiceClient blobServiceClient = new BlobServiceClient(connectionString);
                BlobContainerClient containerClient = blobServiceClient.GetBlobContainerClient("sharepointdata");
                var blobs = containerClient.GetBlobs();
                List<string> encountered = new List<string>();
                var fl = 0;
                foreach (BlobItem blobItem in blobs)
                {
                    string[] parts = blobItem.Name.Split('_');
                    foreach (string str in encountered)
                    {
                        if (str.Equals(parts[3]))
                        {
                            fl = 1;
                            break;
                        }
                    }
                    if (fl == 1)
                        break;
                    encountered.Add(parts[3]);
                    // string WorkItemid = spfile.Workitemid;
                    string MSPN = parts[2];
                    string SI = parts[0];
                    string FactoryCode = parts[1];
                    // string RunId = spfile.RunId;/*/*/*/*
                    string FileName = parts[3].Substring(0, parts[3].LastIndexOf('.'));
                    string FileType = "xls";
                    string Tabs = "";
                    int loadembeddedData = Convert.ToInt32("0");
                    dynamic _excelreadoptions = JsonConvert.DeserializeObject("{\"readhiddensheets\":0,\"readhiddenrows\":0,\"readhiddencolumns\":0}");
                    BlobClient blobClient = containerClient.GetBlobClient(blobItem.Name);

                    BlobDownloadInfo blobDownloadInfo = blobClient.Download();
                    DataTable dataexcelcontent = new DataTable();
                    DataTable dataexcelimagepositions = new DataTable();

                    using (MemoryStream memoryStream = new MemoryStream())
                    {
                        Console.WriteLine(DateTime.Now);
                        blobDownloadInfo.Content.CopyTo(memoryStream);
                        Console.WriteLine(DateTime.Now);
                        memoryStream.Position = 0;

                        dataexcelimagepositions = NPOIExcelDataReader.ReadExcelImages(memoryStream, SI, MSPN, FactoryCode, FileName, FileType, log, context);

                    }
                    BlobDownloadInfo _blobDownload = blobClient.Download();

                    using (MemoryStream memoryStream = new MemoryStream())
                    {
                        Console.WriteLine(DateTime.Now);
                        _blobDownload.Content.CopyTo(memoryStream);
                        Console.WriteLine(DateTime.Now);
                        memoryStream.Position = 0;


                        dataexcelcontent = ExDr.ReadExcelAndReturnResult(memoryStream, Tabs, SI, MSPN, FactoryCode, FileName, FileType, loadembeddedData, log, exlprimarycontext: context);

                        memoryStream.Dispose();
                    }
                    //}


                    if (dataexcelcontent.Rows.Count > 0)
                    {
                        SqlHelper.LoadDataTableToSql(dataexcelcontent, true, "[dbo].[stage_FAI_Templates_AttachmentData]");
                    }
                    else
                    {
                        log.LogInformation($"No data found for {FileName}");
                    }
                    if (dataexcelimagepositions.Rows.Count > 0)
                    {
                        SqlHelper.LoadDataTableToSql(dataexcelimagepositions, true, "[dbo].[stage_FAI_Templates_AttachmentImagePositionData]");
                    }
                    else
                    {
                        log.LogInformation($"No data found for {FileName}");
                    }
                    inpath = tempinputfolderpath;
                    output = tempoutputfolderpath;
                    System.GC.Collect();
                    System.GC.WaitForPendingFinalizers();
                    if (Directory.Exists(inpath)) Directory.Delete(inpath, true);
                    if (Directory.Exists(output)) Directory.Delete(output, true);

                    Console.WriteLine("\t" + blobItem.Name);


                    Microsoft.Data.SqlClient.SqlParameter[] sp_params_new = new Microsoft.Data.SqlClient.SqlParameter[4];
                    Microsoft.Data.SqlClient.SqlParameter param1 = new Microsoft.Data.SqlClient.SqlParameter();
                    param1.ParameterName = "RunId";
                    param1.Value = Globals.runid;
                    sp_params_new.SetValue(param1, 0);

                    //Microsoft.Data.SqlClient.SqlParameter[] sp_params_new = new Microsoft.Data.SqlClient.SqlParameter[1];
                    Microsoft.Data.SqlClient.SqlParameter param2 = new Microsoft.Data.SqlClient.SqlParameter();
                    param2.ParameterName = "SI";
                    param2.Value = "Lenovo";
                    sp_params_new.SetValue(param2, 1);

                    //Microsoft.Data.SqlClient.SqlParameter[] sp_params_new = new Microsoft.Data.SqlClient.SqlParameter[1];
                    Microsoft.Data.SqlClient.SqlParameter param3 = new Microsoft.Data.SqlClient.SqlParameter();
                    param3.ParameterName = "MSPN";
                    param3.Value = parts[2];
                    sp_params_new.SetValue(param3, 2);
                    Microsoft.Data.SqlClient.SqlParameter param4 = new Microsoft.Data.SqlClient.SqlParameter();
                    param4.ParameterName = "FactoryCode";
                    param4.Value = parts[1];
                    sp_params_new.SetValue(param4, 3);

                    // Execute SP usp_FAI_Templates_ProcessData

                    dsMasterDataprocess = SqlHelper.ExecuteStoredProcedureandReturnDatasetResult("[dbo].[usp_FAI_Templates_ProcessData]", sp_params_new);
                    log.LogInformation($"FAI Template Parser Automation Run Completed for : {parts[0]}, {parts[1]}");

                    // ExceptionHandler.LogMessage(new LogException() { SI = parts[0],MSPN = parts[2], Module = module, Message = "Done Processing template data" });
                }

                // dsMasterData = SqlHelper.ExecuteStoredProcedureandReturnDatasetResult("[dbo].[usp_FAI_Templates_GetWorkItemsToProcess]", sp_params);
                /*
                Globals.dtProcessedFiles = dsMasterData.Tables[1];

                //ExceptionHandler.LogMessage(new LogException() { RunId = Globals.runid, Module = module, Message = "Retrieved master workflow data" });

                log.LogInformation($"No of workitems to process: {dsMasterData.Tables[0].Rows.Count}");

                // ExceptionHandler.LogMessage(new LogException() { RunId = Globals.runid, Module = module, Message = $"No of workitems to process: {dsMasterData.Tables[0].Rows.Count}" });

                foreach (DataRow row in dsMasterData.Tables[0].Rows)
                {

                    Globals.WorkitemId = "";
                    Globals.MSPN = "";
                    Globals.SI = "";
                    Globals.FileName = "";
                    Globals.Fileurl = "";

                    /*string RunId = string.Empty;
                    string WorkItemId = string.Empty;
                    string MSPN = string.Empty;
                    string SI = string.Empty;
                    string Siteurl = string.Empty;
                    string ClientId = string.Empty;
                    string ClientSecretkey = string.Empty;
                    string ClientSecret = Environment.GetEnvironmentVariable(ClientSecretkey, EnvironmentVariableTarget.Process);
                    string SPListName = string.Empty;
                    string serverRelativeFolderurl = string.Empty;
                    string loadembedded = string.Empty;
                    string ExcelReadOptions = string.Empty;*/
                /*
                try
                {
                    RunId = row[0].ToString();
                    WorkItemId = row[1].ToString();
                    MSPN = row[2].ToString();
                    SI = row[3].ToString();
                    Siteurl = row[4].ToString();
                    ClientId = row[5].ToString();
                    ClientSecretkey = row[6].ToString();
                    ClientSecret = Environment.GetEnvironmentVariable(ClientSecretkey, EnvironmentVariableTarget.Process);
                    SPListName = row[7].ToString();
                    serverRelativeFolderurl = row[8].ToString();
                    loadembedded = row[9].ToString();
                    ExcelReadOptions = row[10].ToString();

                    Globals.WorkitemId = WorkItemId;
                    Globals.MSPN = MSPN;
                    Globals.SI = SI;
                    Globals.Fileurl = Siteurl;

                    /*SharepointFIle spfile = new SharepointFIle();
                    spfile.RunId = RunId;
                    spfile.Workitemid = WorkItemId;
                    spfile.MSPN = MSPN;
                    spfile.SI = SI;
                    spfile.LoadEmbedded = loadembedded;
                    spfile.ExcelReadOptions = ExcelReadOptions;*/
                /*
                    string message = $"Reading files from sharepoint: {Siteurl} \r\n SP List : {SPListName} \r\n SP folder : {serverRelativeFolderurl}" +
                                     $"\r\n WorkitemId: {Globals.WorkitemId} \r\n MSPN: {Globals.MSPN} \r\n SI: {Globals.SI} ";
                    // ExceptionHandler.LogMessage(new LogException() { RunId = Globals.runid, Module = module, Message = message });

                    spfilelist = SharePointHelper.ReadFilesFromSharepointAsStream(log, Siteurl, ClientId, ClientSecret, SPListName, serverRelativeFolderurl, spfile, context: context);

                    if (spfile != null)
                    {
                        foreach (SharepointFIle spfiledata in (List<SharePointHelper.SharepointFIle>)spfilelist)
                        {
                            Globals.lstspfiles.Add(spfiledata);
                        }

                    }

                    if (spfile.ErrorMessage == null)
                    {
                        Microsoft.Data.SqlClient.SqlParameter[] sp_params_new = new Microsoft.Data.SqlClient.SqlParameter[1];
                        Microsoft.Data.SqlClient.SqlParameter param1 = new Microsoft.Data.SqlClient.SqlParameter();
                        param1.ParameterName = "RunId";
                        param1.Value = Globals.runid;
                        sp_params_new.SetValue(param1, 0);

                        // Execute SP usp_FAI_Templates_ProcessData

                        dsMasterDataprocess = SqlHelper.ExecuteStoredProcedureandReturnDatasetResult("[dbo].[usp_FAI_Templates_ProcessData]", sp_params_new);

                        // ExceptionHandler.LogMessage(new LogException() { RunId = Globals.runid, Module = module, Message = "Done Processing template data" });

                    }
                    else
                    {
                        ExceptionHandler.LogError(new LogException() { RunId = Globals.runid, Module = module, Message = "Issue occured while processing FAI template data, Please check logs" });

                    }


                }
                catch (Exception ex)
                {
                    if (ex.Message.Contains("File Not Found"))
                    {
                        log.LogInformation($"Error Occured Reading SP Library files: {ex.Message} ");
                        module = System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString();
                        string message = $"Exception occured while reading Sharepoint library : Exception: {ex.Message} \r\n StackTrace : {ex.StackTrace} \r\n Sharepoint: {Siteurl} \r\n SP List : {SPListName} \r\n SP folder : {serverRelativeFolderurl}" +
                                         $"\r\n WorkitemId: {Globals.WorkitemId} \r\n MSPN: {Globals.MSPN} \r\n SI: {Globals.SI} " +
                                         $"\r\n Filename: {Globals.FileName} \r\n Fileurl: {Globals.Fileurl}";
                        // ExceptionHandler.LogError(new LogException() { RunId = Globals.runid, Module = module, Message = message });

                    }
                    else
                    {

                        log.LogInformation($"Error Occured Reading SP Library files: {ex.Message} ");
                        module = System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString();
                        string message = $"Exception occured while reading Sharepoint library : Exception: {ex.Message} \r\n StackTrace : {ex.StackTrace} \r\n Sharepoint: {Siteurl} \r\n SP List : {SPListName} \r\n SP folder : {serverRelativeFolderurl}" +
                                         $"\r\n WorkitemId: {Globals.WorkitemId} \r\n MSPN: {Globals.MSPN} \r\n SI: {Globals.SI} " +
                                         $"\r\n Filename: {Globals.FileName} \r\n Fileurl: {Globals.Fileurl}";
                        // ExceptionHandler.LogError(new LogException() { RunId = Globals.runid, Module = module, Message = message });

                    }
                    continue;
                }

            }*/

                #region old sharepoint read method

                //Parsing loaded FAI templates from SI Sharepoint

                //if (Globals.lstspfiles != null && Globals.lstspfiles.Count > 0)
                //{
                //    foreach (SharepointFIle spfiledata in Globals.lstspfiles)
                //    {
                //        if (Directory.Exists(inpath)) Directory.Delete(inpath, true);
                //        if (Directory.Exists(output)) Directory.Delete(output, true);
                //        Directory.CreateDirectory(inpath);
                //        Directory.CreateDirectory(output);

                //        string Workitemid = spfiledata.Workitemid;
                //        string MSPN = spfiledata.MSPN;
                //        string SI = spfiledata.SI;
                //        string RunId = spfiledata.RunId;
                //        string FileName = spfiledata.FileName;
                //        string FileType = spfiledata.FileType.ToString();
                //        string FileSystemObjectType = spfiledata.FileSystemObjectType;
                //        string FilePath = spfiledata.FilePath;
                //        DateTime CreatedDate = spfiledata.CreatedDate;
                //        DateTime ModifiedDate = spfiledata.ModifiedDate;
                //        string CreatedBy = spfiledata.CreatedBy;
                //        string ModifiedBy = spfiledata.ModifiedBy;
                //        MemoryStream Filestream = spfiledata.Filestream;
                //        string ErrorMessage = spfiledata.ErrorMessage;
                //        string Tabs = "";
                //        int loadembeddedData = Convert.ToInt32(spfiledata.LoadEmbedded);
                //        dynamic _excelreadoptions = JsonConvert.DeserializeObject(spfiledata.ExcelReadOptions);

                //        if (Directory.Exists(inpath)) Directory.Delete(inpath, true);
                //        if (Directory.Exists(output)) Directory.Delete(output, true);
                //        Directory.CreateDirectory(inpath);
                //        Directory.CreateDirectory(output);


                //        log.LogInformation($"Starting the load of excel for attachmentId {Workitemid} file {FileName}");

                //        DataTable dataexcelcontent = new DataTable();
                //        DataTable dataexcelimagepositions = new DataTable();

                //        if (FileType.ToString().ToUpper().EndsWith("XLS") || FileType.ToString().ToUpper().EndsWith("XLSX"))
                //        {
                //            //using (MemoryStream memoryStreamexcelimages = new MemoryStream())
                //            //{
                //            //    spfiledata.Filestream.Position = 0;
                //            //    // Copy the data from the FileStream into the MemoryStream
                //            //    spfiledata.Filestream.CopyTo(memoryStreamexcelimages);
                //            //    memoryStreamexcelimages.Position = 0;


                //            //    dataexcelcontent = ExDr.ReadExcelAndReturnResult(spfiledata.Filestream, Tabs, Workitemid, SI, MSPN, RunId, FileName, loadembeddedData, FileType, log, context);
                //            //    dataexcelimagepositions = NPOIExcelDataReader.ReadExcelImages(memoryStreamexcelimages, Workitemid, SI, MSPN, RunId, FileName, FileType, log, context);
                //            //}

                //            using (MemoryStream memoryStreamexcelimages = new MemoryStream())
                //            {
                //                using (MemoryStream fileStreamCopy = new MemoryStream())
                //                {
                //                    // Copy the data from the closed stream to the new MemoryStream
                //                    spfiledata.Filestream.CopyTo(fileStreamCopy);

                //                    // Set the position of the copied MemoryStream to 0
                //                    fileStreamCopy.Position = 0;

                //                    // Copy the data from the MemoryStream copy to the final MemoryStream
                //                    fileStreamCopy.CopyTo(memoryStreamexcelimages);
                //                    memoryStreamexcelimages.Position = 0;

                //                    // Pass the MemoryStream to the appropriate methods
                //                    dataexcelcontent = ExDr.ReadExcelAndReturnResult(memoryStreamexcelimages, Tabs, Workitemid, SI, MSPN, RunId, FileName, loadembeddedData, FileType, log, context);
                //                    dataexcelimagepositions = NPOIExcelDataReader.ReadExcelImages(memoryStreamexcelimages, Workitemid, SI, MSPN, RunId, FileName, FileType, log, context);
                //                }
                //            }
                //        }
                //        if (dataexcelcontent.Rows.Count > 0)
                //        {
                //            SqlHelper.LoadDataTableToSql(dataexcelcontent, "[dbo].[stage_FAI_Templates_AttachmentData]");
                //        }
                //        else
                //        {
                //            log.LogInformation($"No data found for {Workitemid}");
                //        }
                //        if (dataexcelimagepositions.Rows.Count > 0)
                //        {
                //            SqlHelper.LoadDataTableToSql(dataexcelimagepositions, "[dbo].[stage_FAI_Templates_AttachmentImagePositionData]");
                //        }
                //        else
                //        {
                //            log.LogInformation($"No data found for {Workitemid}");
                //        }

                //        inpath = tempinputfolderpath;
                //        output = tempoutputfolderpath;
                //        System.GC.Collect();
                //        System.GC.WaitForPendingFinalizers();
                //        if (Directory.Exists(inpath)) Directory.Delete(inpath, true);
                //        if (Directory.Exists(output)) Directory.Delete(output, true);
                //    }

                //}

                #endregion

                if (dsMasterData != null && dsMasterData.Tables.Count > 0)
                {
                    log.LogInformation($"FAI Template Parser Automation Run Completed at : {DateTime.Now}");
                    //ExceptionHandler.LogMessage(new LogException() { RunId = Globals.runid, Module = module, Message = $"FAI Template Parser Automation Run Completed", Operation = "RunEnd" });


                }


            }
            catch (Exception ex)
            {
                log.LogInformation($"{ex.Message}");

                string module = System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString();
                string message = $"Exception occured while running automation: Exception: {ex.Message} \r\n StackTrace : {ex.StackTrace} ";

                log.LogError($"{message}");

                ExceptionHandler.LogError(new LogException() { RunId = Globals.runid, Module = module, Message = message });
            }
            finally
            {
                // Globals.DisposeGlobals();

            }

        }
    }
}
