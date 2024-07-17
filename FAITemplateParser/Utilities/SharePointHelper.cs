using Microsoft.Azure.WebJobs;
using Microsoft.Extensions.Logging;
using Microsoft.SharePoint.Client;
using Microsoft.TeamFoundation.Common;
using Newtonsoft.Json;
using PnP.Framework;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;

namespace FAITemplateParser.Utilities
{
    [System.Runtime.Versioning.SupportedOSPlatform("windows")]
    public class SharePointHelper
    {
        public static readonly string _siteUrl = System.Environment.GetEnvironmentVariable("siteUrl", EnvironmentVariableTarget.Process);
        public static readonly string _clientId = System.Environment.GetEnvironmentVariable("clientId", EnvironmentVariableTarget.Process);
        public static readonly string _clientSecret = System.Environment.GetEnvironmentVariable("clientSecret", EnvironmentVariableTarget.Process);
        public static readonly string _spListName = System.Environment.GetEnvironmentVariable("spListName", EnvironmentVariableTarget.Process);
        public static readonly string _serverRelativeFolderUrl = System.Environment.GetEnvironmentVariable("serverRelativeFolderUrl", EnvironmentVariableTarget.Process);
        public static string tempinputfolderpath = Environment.GetEnvironmentVariable("InputPath", EnvironmentVariableTarget.Process);
        public static string tempoutputfolderpath = Environment.GetEnvironmentVariable("OutputPath", EnvironmentVariableTarget.Process);


        /// <summary>
        /// Read Sharepoint files as stream and their details from a Sharepoint list
        /// </summary>
        /// <param name="siteurl"></param>
        /// <param name="clientid"></param>
        /// <param name="clientsecret"></param>
        /// <param name="splistname"></param>
        /// <param name="serverrelativefolderurl"></param>
        /// <param name="log"></param>
        /// <returns></returns>
        public static List<SharepointFIle> ReadFilesFromSharepointAsStream(ILogger log, string siteurl = null, string clientid = null
                                                    , string clientsecret = null, string splistname = null, string serverrelativefolderurl = null, SharepointFIle spfile = null, string filenamepattern = null, string tablstoload = null, string resultstable = null, string Workitemid = null, string si = null, string runid = null, string mspn = null, ExecutionContext context = null)
        {
            List<SharepointFIle> SPList = new List<SharepointFIle>();
            try
            {

                string SPSiteurl = siteurl.IsNullOrEmpty() ? _siteUrl : siteurl;
                string SPClientid = clientid.IsNullOrEmpty() ? _clientId : clientid;
                string SPClientSecret = clientsecret.IsNullOrEmpty() ? _clientSecret : clientsecret;
                string SPListName = splistname.IsNullOrEmpty() ? _spListName : splistname;
                string SPServerRelativeFolderUrl = serverrelativefolderurl.IsNullOrEmpty() ? _serverRelativeFolderUrl : serverrelativefolderurl;
                string module = System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString();

                using (ClientContext clientContext = CreateSPContext(log, SPSiteurl, SPClientid, SPClientSecret))
                {
                    Web web = clientContext.Web;

                    clientContext.Load(web);
                    clientContext.Load(web.Lists);
                    clientContext.Load(web, wb => wb.ServerRelativeUrl);
                    clientContext.ExecuteQueryRetry();

                    List list = web.Lists.GetByTitle(SPListName);
                    clientContext.Load(list, l => l.RootFolder.ServerRelativeUrl);
                    clientContext.ExecuteQuery();

                    Folder folder = web.GetFolderByServerRelativeUrl(list.RootFolder.ServerRelativeUrl + SPServerRelativeFolderUrl); // "/Platform/Program Management/Serviceability/L10 Designs/"
                    clientContext.Load(folder);
                    clientContext.ExecuteQuery();


                    //Logic to handle sharepoint list threshold limits to read only 5000 items


                    List<ListItem> lstitems = new List<ListItem>();

                    ListItemCollectionPosition position = null;
                    var page = 1;

                    do
                    {
                        CamlQuery camlQuery = new CamlQuery();
                        camlQuery.ViewXml = @"<View Scope='RecursiveAll'>
                                     <Query>
                                     </Query>
                                        <RowLimit>5000</RowLimit>
                                 </View>";

                        camlQuery.FolderServerRelativeUrl = folder.ServerRelativeUrl;

                        camlQuery.ListItemCollectionPosition = position;

                        ListItemCollection listItems = list.GetItems(camlQuery);
                        clientContext.Load(listItems,
                               items => items.Include(
                                  item => item.Id,
                                  item => item.ContentType,
                                  item => item.DisplayName,
                                  item => item.FieldValuesAsText,
                                  item => item["FileRef"],
                                  item => item["File_x0020_Type"],
                                  item => item.FileSystemObjectType,
                                  item => item.HasUniqueRoleAssignments,
                                  item => item.File.ServerRelativeUrl,
                                  item => item["Modified"],
                                  item => item["Created"],
                                  item => item["Editor"],
                                  item => item["Author"],
                                  item => item.Properties,
                                  item => item.ParentList.ParentWeb.SiteUsers
                                  ),
                               items => items.ListItemCollectionPosition);
                        clientContext.ExecuteQueryRetry();

                        position = listItems.ListItemCollectionPosition;
                        foreach (ListItem listItem in listItems)
                        {
                            lstitems.Add(listItem);


                        }


                        page++;
                    }
                    while (position != null);

                    var filteredlist = (from item in lstitems
                                        where item.FileSystemObjectType == FileSystemObjectType.File
                                        select item).ToList();


                    foreach (ListItem oListItem in filteredlist)
                    {
                        try
                        {
                            if (oListItem.FileSystemObjectType == FileSystemObjectType.File)
                            {
                                Globals.FileName = oListItem["FileRef"].ToString();
                                spfile.FileName = oListItem.DisplayName;

                                if (oListItem.DisplayName.Contains("_output"))
                                {

                                    string message = $"skipping _output file : Sharepoint: {siteurl} \r\n SP List : {splistname} \r\n SP folder : {serverrelativefolderurl} \r\n SP File Name: {oListItem["FileRef"].ToString()}" +
                                         $"\r\n WorkitemId: {Globals.WorkitemId} \r\n MSPN: {Globals.MSPN} \r\n SI: {Globals.SI} " +
                                                         $"\r\n Filename: {oListItem["FileRef"].ToString()} \r\n Fileurl: {Globals.Fileurl}";
                                    ExceptionHandler.LogMessage(new LogException() { RunId = Globals.runid, Module = module, Message = message });
                                    log.LogInformation(message);
                                    continue;
                                }

                                if (oListItem.DisplayName != null)
                                {
                                    int filecount = (from f in Globals.dtProcessedFiles.AsEnumerable()
                                                     where f["FileName"].ToString() == oListItem.DisplayName && f["WorkItemid"].ToString() == Globals.WorkitemId.ToString()
                                                     select f).Count();

                                    if (filecount > 0)
                                    {
                                        string message = $"skipping file as it seems to be parsed already : Sharepoint: {siteurl} \r\n SP List : {splistname} \r\n SP folder : {serverrelativefolderurl} \r\n SP File Name: {oListItem["FileRef"].ToString()}" +
                                        $"\r\n WorkitemId: {Globals.WorkitemId} \r\n MSPN: {Globals.MSPN} \r\n SI: {Globals.SI} " +
                                                        $"\r\n Filename: {oListItem["FileRef"].ToString()} \r\n Fileurl: {Globals.Fileurl}";
                                        ExceptionHandler.LogMessage(new LogException() { RunId = Globals.runid, Module = module, Message = message });
                                        log.LogInformation(message);
                                        continue;
                                    }
                                }


                                spfile.FileSystemObjectType = oListItem.FileSystemObjectType.ToString();
                                spfile.FilePath = oListItem["FileRef"].ToString();
                                string[] splittedpath = oListItem["FileRef"].ToString().Split('.', StringSplitOptions.RemoveEmptyEntries);
                                spfile.FileType = splittedpath != null ? splittedpath[splittedpath.Length - 1] : "";

                                // Retrieve modified date and time
                                DateTime modifiedDateTime = (DateTime)oListItem["Modified"];
                                spfile.ModifiedDate = modifiedDateTime;

                                // Retrieve created date and time
                                DateTime createdDateTime = (DateTime)oListItem["Created"];
                                spfile.CreatedDate = createdDateTime;

                                // Retrieve modified by user
                                FieldUserValue modifiedByValue = oListItem["Editor"] as FieldUserValue;
                                string modifiedByUserName = modifiedByValue.LookupValue;

                                spfile.ModifiedBy = modifiedByUserName;

                                // Retrieve created by user
                                FieldUserValue createdByValue = oListItem["Author"] as FieldUserValue;
                                string createdByUserName = createdByValue.LookupValue;

                                spfile.CreatedBy = createdByUserName;

                                Microsoft.SharePoint.Client.File f1 = web.GetFileByServerRelativeUrl(oListItem["FileRef"].ToString());

                                //foreach (Microsoft.SharePoint.Client.FileVersion version in f1.Versions)
                                //{
                                //    clientContext.Load(version);
                                //    clientContext.ExecuteQuery();

                                //    Console.WriteLine($"Version Label: {version.VersionLabel}");
                                //    Console.WriteLine($"Modified By: {version.CreatedBy}");
                                //    Console.WriteLine($"Modified Date: {version.Created}");
                                //    Console.WriteLine();
                                //}


                                var stream1 = f1.OpenBinaryStream();

                                //var stream1 = f1.OpenBinaryStreamWithOptions(SPOpenBinaryOptions.GetAsZipStreamBundleFriendly);

                                clientContext.Load(f1);
                                clientContext.ExecuteQuery();

                                //ClientResult<Stream> stream = (ClientResult<Stream>)stream1;
                                if (stream1 != null && stream1.Value != null)
                                {
                                    string inpath = tempinputfolderpath;
                                    string output = tempoutputfolderpath;

                                    if (Directory.Exists(inpath)) Directory.Delete(inpath, true);
                                    if (Directory.Exists(output)) Directory.Delete(output, true);
                                    Directory.CreateDirectory(inpath);
                                    Directory.CreateDirectory(output);

                                    log.LogInformation($"Starting the load of excel for attachmentId {spfile.Workitemid} file {spfile.FileName}");

                                    DataTable dataexcelcontent = new DataTable();
                                    DataTable dataexcelimagepositions = new DataTable();

                                    using (Stream stream = stream1.Value)
                                    {
                                        string WorkItemid = spfile.Workitemid;
                                        string MSPN = spfile.MSPN;
                                        string SI = spfile.SI;
                                        string RunId = spfile.RunId;
                                        string FileName = spfile.FileName;
                                        string FileType = spfile.FileType.ToString();
                                        string FileSystemObjectType = spfile.FileSystemObjectType;
                                        string FilePath = spfile.FilePath;
                                        DateTime CreatedDate = spfile.CreatedDate;
                                        DateTime ModifiedDate = spfile.ModifiedDate;
                                        string CreatedBy = spfile.CreatedBy;
                                        string ModifiedBy = spfile.ModifiedBy;
                                        MemoryStream Filestream = spfile.Filestream;
                                        string ErrorMessage = spfile.ErrorMessage;
                                        string Tabs = "";
                                        int loadembeddedData = Convert.ToInt32(spfile.LoadEmbedded);
                                        dynamic _excelreadoptions = JsonConvert.DeserializeObject(spfile.ExcelReadOptions);


                                        using (MemoryStream mStream = new MemoryStream())
                                        {
                                            log.LogInformation($"Starting reading excel file {spfile.Workitemid} file {spfile.FileName}");

                                            stream.CopyTo(mStream);
                                            mStream.Position = 0;
                                            stream.Position = 0;

                                            dataexcelcontent = ExDr.ReadExcelAndReturnResult(mStream, Tabs, WorkItemid, SI, MSPN, RunId, FileName, loadembeddedData, FileType, log, exlprimarycontext: context);
                                            mStream.Dispose();
                                        }

                                        using (MemoryStream memoryStreamexcelimages = new MemoryStream())
                                        {
                                            log.LogInformation($"Starting reading excel image positions {spfile.Workitemid} file {spfile.FileName}");

                                            stream.Position = 0;
                                            stream.CopyTo(memoryStreamexcelimages);
                                            memoryStreamexcelimages.Position = 0;

                                            dataexcelimagepositions = NPOIExcelDataReader.ReadExcelImages(memoryStreamexcelimages, WorkItemid, SI, MSPN, RunId, FileName, FileType, log, context);
                                            memoryStreamexcelimages.Dispose();
                                        }

                                        stream.Dispose();


                                    }
                                    if (dataexcelcontent.Rows.Count > 0)
                                    {
                                        SqlHelper.LoadDataTableToSql(dataexcelcontent, true, "[dbo].[stage_FAI_Templates_AttachmentData]");
                                    }
                                    else
                                    {
                                        log.LogInformation($"No data found for {Workitemid}");
                                    }
                                    if (dataexcelimagepositions.Rows.Count > 0)
                                    {
                                        SqlHelper.LoadDataTableToSql(dataexcelimagepositions, true, "[dbo].[stage_FAI_Templates_AttachmentImagePositionData]");
                                    }
                                    else
                                    {
                                        log.LogInformation($"No data found for {Workitemid}");
                                    }


                                    inpath = tempinputfolderpath;
                                    output = tempoutputfolderpath;
                                    System.GC.Collect();
                                    System.GC.WaitForPendingFinalizers();
                                    if (Directory.Exists(inpath)) Directory.Delete(inpath, true);
                                    if (Directory.Exists(output)) Directory.Delete(output, true);


                                }

                                stream1 = null;

                                SPList.Add(spfile);
                            }

                        }
                        catch (Exception ex)
                        {
                            log.LogInformation($"Error Ocurred parsing, skipping file: {oListItem["FileRef"].ToString()} Error: {ex.Message} ");
                            //string module = System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString();
                            string message = $"Error Ocurred parsing, skipping file : Exception: {ex.Message} \r\n StackTrace : {ex.StackTrace} \r\n Sharepoint: {siteurl} \r\n SP List : {splistname} \r\n SP folder : {serverrelativefolderurl} \r\n SP File Name: {oListItem["FileRef"].ToString()}" +
                                 $"\r\n WorkitemId: {Globals.WorkitemId} \r\n MSPN: {Globals.MSPN} \r\n SI: {Globals.SI} " +
                                                 $"\r\n Filename: {oListItem["FileRef"].ToString()} \r\n Fileurl: {Globals.Fileurl}";
                            ExceptionHandler.LogError(new LogException() { RunId = Globals.runid, Module = module, Message = message });

                            spfile.ErrorMessage = String.Concat("Module: ", module, "Message: ", message);

                            SPList.Add(spfile);

                            continue;
                        }

                    }

                    return SPList;

                }




                //var filteredlist = (from item in lstitems
                //                    where item["FileRef"].ToString().Contains(@"/teams/DCIS/Shared Documents/Platform/Program Management/Serviceability/L10 Designs")
                //                    && item.FileSystemObjectType == FileSystemObjectType.File
                //                    && item["FileRef"].ToString().Contains("CSR")
                //                    select item).ToList();






            }
            catch (Exception ex)
            {
                log.LogInformation($"Error Occured Reading SP Library files: {ex.Message} ");
                string module = System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString();
                string message = $"Exception occured while reading Sharepoint library : Exception: {ex.Message} \r\n StackTrace : {ex.StackTrace} \r\n Sharepoint: {siteurl} \r\n SP List : {splistname} \r\n SP folder : {serverrelativefolderurl}" + $"\r\n WorkitemId: {Globals.WorkitemId} \r\n MSPN: {Globals.MSPN} \r\n SI: {Globals.SI} " + $"\r\n Filename: {Globals.FileName} \r\n Fileurl: {Globals.Fileurl}";
                ExceptionHandler.LogError(new LogException() { RunId = Globals.runid, Module = module, Message = message });

                throw;
            }

        }

        /// <summary>
        /// Create SharePoint Client context
        /// </summary>
        /// <param name="siteurl"></param>
        /// <param name="clientid"></param>
        /// <param name="clientsecret"></param>
        /// <param name="log"></param>
        /// <returns></returns>
        public static ClientContext CreateSPContext(ILogger log, string siteurl = null, string clientid = null, string clientsecret = null)
        {
            ClientContext spctx = null;

            string SPSiteurl = siteurl.IsNullOrEmpty() ? _siteUrl : siteurl;
            string SPClientid = clientid.IsNullOrEmpty() ? _clientId : clientid;
            string SPClientSecret = clientsecret.IsNullOrEmpty() ? _clientSecret : clientsecret;

            try
            {
                log.LogInformation($"Creating connection context to Sharepoint - {SPSiteurl}");

                var authManager = new AuthenticationManager();

                spctx = authManager.GetACSAppOnlyContext(SPSiteurl, SPClientid, SPClientSecret);

                log.LogInformation($"Client Context creation successful.");
            }
            catch (Exception ex)
            {
                string module = System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString();
                string message = $"Exception occured while Creating Sharepoint context : Exception: {ex.Message} \r\n StackTrace : {ex.StackTrace} \r\n Sharepoint: {siteurl}\r\n WorkitemId: {Globals.WorkitemId} \r\n MSPN: {Globals.MSPN} \r\n SI: {Globals.SI} " +
                                             $"\r\n Filename: {Globals.FileName} \r\n Fileurl: {Globals.Fileurl}"; ;
                ExceptionHandler.LogError(new LogException() { RunId = Globals.runid, Module = module, Message = message });

                throw;
            }

            return spctx;
        }



        /// <summary>
        /// Type for Sharepoint file data
        /// </summary>

        public class SharepointFIle
        {
            public string Workitemid { get; set; }
            public string MSPN { get; set; }
            public string SI { get; set; }
            public string RunId { get; set; }
            public string FileName { get; set; }
            public object FileType { get; set; }
            public string FileSystemObjectType { get; set; }
            public string FilePath { get; set; }
            public DateTime CreatedDate { get; set; }
            public DateTime ModifiedDate { get; set; }
            public string CreatedBy { get; set; }
            public string ModifiedBy { get; set; }
            public MemoryStream Filestream { get; set; }
            public string ErrorMessage { get; set; }
            public string LoadEmbedded { get; set; }
            public string ExcelReadOptions { get; set; }

        }
    }
}
