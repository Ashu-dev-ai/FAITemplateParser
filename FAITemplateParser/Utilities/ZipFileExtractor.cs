using Microsoft.Azure.WebJobs;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;
using SevenZipExtractor;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;

namespace FAITemplateParser.Utilities
{
    public class ZipFileExtractor
    {
        [System.Runtime.Versioning.SupportedOSPlatform("windows")]
        public static DataTable ReadExcelInZipAndReturnResult(MemoryStream memoryStream, string RunId, string WorkItemId, string MSPN, string SI, string FileName, string TabNames, int _loadembedded, string _FileType, ILogger _log, ExecutionContext executionContext)
        {
            //string dllPath = Environment.Is64BitProcess ? "C:\\home\\site\\wwwroot\\bin\\x64\\7z.dll" : "C:\\home\\site\\wwwroot\\bin\\x86\\7z.dll";
            string dllPath = Environment.Is64BitProcess ? "C:\\Users\\dobakkia\\.nuget\\packages\\sevenzipextractor\\1.0.16\\build\\x64\\7z.dll" : "C:\\Users\\dobakkia\\.nuget\\packages\\sevenzipextractor\\1.0.16\\build\\x86\\7z.dll";
            var currentDirectory = executionContext.FunctionDirectory;
            //var env = Environment.Is64BitProcess;

            memoryStream.Position = 0;
            SevenZipFormat sevenZipFormat = sevenzFormat(_FileType);

            DataTable result = new DataTable();
            result.Columns.Add("RunId");
            result.Columns.Add("WorkItemId");
            result.Columns.Add("MSPN");
            result.Columns.Add("SI");
            result.Columns.Add("FileName");
            result.Columns.Add("TabName");
            result.Columns.Add("Data");
            try
            {
                using (var zip = new ArchiveFile(memoryStream, sevenZipFormat, dllPath))
                {
                    foreach (var entry in zip.Entries)
                    {
                        if (entry.Size > 0 && (entry.FileName.StartsWith("PAT") || entry.FileName.StartsWith("PQT")) && (entry.FileName.ToUpper().EndsWith(".XLS") || entry.FileName.ToUpper().EndsWith(".XLSX") || entry.FileName.ToUpper().EndsWith(".XLSM")))
                        {
                            using (MemoryStream m2 = new MemoryStream()) // entry.Open())
                            {
                                string _unzippedFileType = "";
                                _unzippedFileType = entry.FileName.ToUpper().EndsWith(".XLS") ? "XLS" : "XLSX";
                                m2.Position = 0;
                                MemoryStream copyofm2 = new MemoryStream();
                                m2.CopyTo(copyofm2);
                                m2.Position = 0;
                                copyofm2.Position = 0;
                                entry.Extract(m2);
                                string JsonData = "";

                                DataSet dataSet = ExDr.ReadExcelAndReturnDataset(m2, RunId, WorkItemId, MSPN, SI, FileName, _loadembedded, _unzippedFileType, _log, executionContext);
                                if (String.IsNullOrEmpty(TabNames))
                                {
                                    TabNames = String.Join(",", dataSet.Tables.Cast<DataTable>().Select(t => t.TableName).ToList());
                                }
                                foreach (string Tab in TabNames.Split(","))
                                {
                                    var datatable = dataSet.Tables[Tab];
                                    JsonData = JsonConvert.SerializeObject(datatable);
                                    result.Rows.Add(new object[] { RunId, WorkItemId, MSPN, SI, entry.FileName, Tab, JsonData });
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                //SqlHelper.ExecuteSqlQuery_NoResult(String.Format("INSERT INTO [dbo].[FAI_Templates_ErrorLogs]([AttachmentId],[ErrorLocation],[ErrorMessage]) VALUES ({0},'{1}','{2}')", WorkItemId, System.Reflection.MethodBase.GetCurrentMethod().Name, ex.Message.Replace("'", "''")));
                _log.LogInformation($"Zip loader error: {ex.Message}");
                string module = System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString();
                string message = $"Exception occured extracting zip file : Exception: {ex.Message} \r\n StackTrace : {ex.StackTrace}" +
                                    $"\r\n WorkitemId: {Globals.WorkitemId} \r\n MSPN: {Globals.MSPN} \r\n SI: {Globals.SI} " +
                                    $"\r\n Filename: {Globals.FileName} \r\n Fileurl: {Globals.Fileurl}";
                ExceptionHandler.LogError(new LogException() { RunId = Globals.runid, Module = module, Message = message });

                //throw;
            }

            return result;
        }
        public static SevenZipFormat sevenzFormat(string _fileType)
        {
            var map = new Dictionary<string, SevenZipFormat>()
                        {
                            {"7Z", SevenZipFormat.SevenZip},
                            {"ZIP", SevenZipFormat.Zip},
                            {"RAR", SevenZipFormat.Rar}
                        };
            SevenZipFormat format;
            return map.TryGetValue(_fileType.ToUpper(), out format) ? format : SevenZipFormat.Zip;
        }
    }
}
