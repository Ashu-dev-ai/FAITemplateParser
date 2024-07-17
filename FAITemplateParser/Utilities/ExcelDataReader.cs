using ExcelDataReader;
using Microsoft.Azure.WebJobs;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;

namespace FAITemplateParser.Utilities
{
    [System.Runtime.Versioning.SupportedOSPlatform("windows")]
    public class ExDr
    {

        public static DataSet ReadExcelAndReturnDataset(MemoryStream memoryStream, string MSPN, string SI, string FileName, string FileType, int _loadembedded, ILogger xlseclog, ExecutionContext xlseccontext)
        {
            DataSet result = new DataSet();
            try
            {
                memoryStream.Position = 0;
                System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
                if (_loadembedded == 1)
                {

                    MemoryStream newMemoryStream = new MemoryStream();
                    memoryStream.CopyTo(newMemoryStream);
                    memoryStream.Position = 0;
                    newMemoryStream.Position = 0;
                    EmbeddedDataReaderHelper.ReadEmbeddedData(newMemoryStream, MSPN, SI, FileName, FileType, xlseclog, xlseccontext);

                }
                //int readhiddenrows = 0;
                //int readhiddencolumns = 0;
                //int readhiddensheets = 0;
                //memoryStream.Position = 0;

                // Create a new DataSet
                //DataSet dataSet = new DataSet();

                //// Open the Excel workbook using ClosedXML
                //using (var workBook = new XLWorkbook(memoryStream))
                //{
                //    // Loop through all worksheets in the workbook
                //    foreach (IXLWorksheet workSheet in workBook.Worksheets)
                //    {
                //        // Create a DataTable for each worksheet
                //        DataTable dataTable = new DataTable(workSheet.Name);

                //        // Load data from the worksheet into the DataTable
                //        foreach (IXLRow row in workSheet.RowsUsed())
                //        {
                //            DataRow dataRow = dataTable.NewRow();
                //            foreach (IXLCell cell in row.CellsUsed())
                //            {
                //                dataRow[cell.Address.ColumnNumber - 1] = cell.Value;
                //            }
                //            dataTable.Rows.Add(dataRow);
                //        }

                //        // Add the DataTable to the DataSet
                //        dataSet.Tables.Add(dataTable);
                //    }
                //}



                using (IExcelDataReader reader = ExcelReaderFactory.CreateReader(memoryStream, new ExcelReaderConfiguration()
                {
                    FallbackEncoding = System.Text.Encoding.GetEncoding(1252)
                }))
                {
                    var conf = new ExcelDataSetConfiguration
                    {
                        FilterSheet = (tableReader, sheetIndex) => tableReader.VisibleState == "visible",
                        ConfigureDataTable = _ => new ExcelDataTableConfiguration
                        {
                            UseHeaderRow = false,
                            FilterRow = (rowReader) => rowReader.RowHeight > 0,
                            FilterColumn = (rowReader, columnIndex) => rowReader.GetColumnWidth(columnIndex) > 0,
                            EmptyColumnNamePrefix = "Column"
                        }
                    };
                    result = reader.AsDataSet(conf);
                    reader.Close();
                }

                memoryStream.Dispose();

            }
            catch (Exception ex)
            {
                //xlseclog.LogInformation(String.Format("INSERT INTO [dbo].[FAI_Templates_ErrorLogs]([AttachmentId],[ErrorLocation],[ErrorMessage]) VALUES ({0},'{1}','{2}')", WorkItemId, System.Reflection.MethodBase.GetCurrentMethod().Name, ex.Message.Replace("'", "''")));
                //SqlHelper.ExecuteSqlQuery_NoResult(String.Format("INSERT INTO [dbo].[FAI_Templates_ErrorLogs]([AttachmentId],[ErrorLocation],[ErrorMessage]) VALUES ({0},'{1}','{2}')", WorkItemId, System.Reflection.MethodBase.GetCurrentMethod().Name, ex.Message.Replace("'", "''")));

                string module = System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString();
                string message = $"Exception occured while reading excel data : Exception: {ex.Message} \r\n StackTrace : {ex.StackTrace}" +
                                    $"\r\n WorkitemId: {Globals.WorkitemId} \r\n MSPN: {Globals.MSPN} \r\n SI: {Globals.SI} " +
                                    $"\r\n Filename: {Globals.FileName} \r\n Fileurl: {Globals.Fileurl}";
                ExceptionHandler.LogError(new LogException() { RunId = Globals.runid, Module = module, Message = message });

                throw;
            }
            return result;
        }
        public static System.Data.DataTable ReadExcelAndReturnResult(MemoryStream memoryStream, string TabNames, string SI, string MSPN, string FactoryCode, string FileName, string FileType, int _loadembedded, ILogger exlprimarylog, ExecutionContext exlprimarycontext)
        {
            memoryStream.Position = 0;

            System.Data.DataTable result = new System.Data.DataTable();
            // result.Columns.Add("RunId", typeof(Guid));
            // result.Columns.Add("WorkItemId");
            result.Columns.Add("MSPN");
            result.Columns.Add("FactoryCode");
            result.Columns.Add("SI");
            result.Columns.Add("FileName");
            result.Columns.Add("TabName");
            result.Columns.Add("Data");
            List<string> tablenames = new List<string>();

            using (DataSet dataSet = ReadExcelAndReturnDataset(memoryStream, MSPN, SI, FileName, FileType, _loadembedded, exlprimarylog, exlprimarycontext))
            {
                string JsonData = "";
                try
                {
                    if (String.IsNullOrEmpty(TabNames))
                    {
                        tablenames = dataSet.Tables.Cast<System.Data.DataTable>().Select(t => t.TableName).ToList();
                    }
                    else if (!String.IsNullOrEmpty(TabNames))
                    {
                        tablenames = TabNames.Split(",").ToList();
                    }

                    foreach (string Tab in tablenames)
                    {
                        var datatable = dataSet.Tables[Tab];
                        JsonData = JsonConvert.SerializeObject(datatable);
                        result.Rows.Add(new object[] { MSPN, FactoryCode, SI, FileName, Tab, JsonData });
                    }
                }
                catch (Exception ex)
                {
                    //exlprimarylog.LogInformation(String.Format("INSERT INTO [dbo].[FAI_Templates_ErrorLogs]([AttachmentId],[ErrorLocation],[ErrorMessage]) VALUES ({0},'{1}','{2}')", WorkItemId, System.Reflection.MethodBase.GetCurrentMethod().Name, ex.Message.Replace("'", "''")));
                    //SqlHelper.ExecuteSqlQuery_NoResult(String.Format("INSERT INTO [dbo].[FAI_Templates_ErrorLogs]([AttachmentId],[ErrorLocation],[ErrorMessage]) VALUES ({0},'{1}','{2}')", WorkItemId, System.Reflection.MethodBase.GetCurrentMethod().Name, ex.Message.Replace("'", "''")));
                    string module = System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString();
                    string message = $"Exception occured while reading excel data : Exception: {ex.Message} \r\n StackTrace : {ex.StackTrace}" +
                                        $"\r\n WorkitemId: {Globals.WorkitemId} \r\n MSPN: {Globals.MSPN} \r\n SI: {Globals.SI} " +
                                        $"\r\n Filename: {Globals.FileName} \r\n Fileurl: {Globals.Fileurl}";
                    ExceptionHandler.LogError(new LogException() { RunId = Globals.runid, Module = module, Message = message });

                    throw;
                }

            }
            memoryStream.Dispose();
            return result;
        }


    }
}
