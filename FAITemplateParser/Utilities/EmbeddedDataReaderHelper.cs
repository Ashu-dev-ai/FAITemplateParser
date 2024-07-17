using Microsoft.Azure.WebJobs;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlTypes;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Xml;

namespace FAITemplateParser.Utilities
{
    [System.Runtime.Versioning.SupportedOSPlatform("windows")]
    public class EmbeddedDataReaderHelper
    {
        public static string tempinputfolderpath = Environment.GetEnvironmentVariable("InputPath", EnvironmentVariableTarget.Process);
        public static string tempoutputfolderpath = Environment.GetEnvironmentVariable("OutputPath", EnvironmentVariableTarget.Process);
        [System.Runtime.Versioning.SupportedOSPlatform("windows")]
        public static void ReadEmbeddedData(MemoryStream embbeddedStream, string RunId, string WorkItemId, string MSPN, string SI, string filename, string fileType, ILogger _embdlog, ExecutionContext _embdcontext)
        {
            embbeddedStream.Position = 0;

            using (MemoryStream mstream = new MemoryStream())
            {
                embbeddedStream.CopyTo(mstream);
                mstream.Position = 0;

                if (fileType.ToLower() == "xls")
                {
                    MemoryStream newembbeddedStream = XLVersionModifier.ConvertXLVersionTo2016(embbeddedStream);
                    newembbeddedStream.Position = 0;
                    embbeddedStream = new MemoryStream();
                    embbeddedStream = newembbeddedStream;
                }
                DataTable resultmetadata = new DataTable();
                resultmetadata.Columns.Add("RunId", typeof(Guid));
                resultmetadata.Columns.Add("WorkItemId");
                resultmetadata.Columns.Add("MSPN");
                resultmetadata.Columns.Add("SI");
                resultmetadata.Columns.Add("FileName");
                resultmetadata.Columns.Add("MetadataJson");
                resultmetadata.Columns.Add("Timestamp");
                string embeddedmetadata = ExtractMetadata(mstream);
                mstream.Dispose();
                string timestamp = DateTime.Now.ToString();
                resultmetadata.Rows.Add(new object[] { RunId, WorkItemId, MSPN, SI, filename, embeddedmetadata, timestamp });
                SqlHelper.LoadDataTableToSql(resultmetadata, true, "Stage_FAI_Templates_EmbeddedMetadata");
            }

            ExtractEmbeddeddata(embbeddedStream, RunId, WorkItemId, MSPN, SI, filename, _embdlog, _embdcontext);

        }
        public static string ExtractMetadata(MemoryStream sourcestream)
        {
            string result = "";
            try
            {
                sourcestream.Position = 0;
                sourcestream.Seek(0, SeekOrigin.Begin);

                //string tempFilePath = Path.GetTempFileName();
                //File.WriteAllBytes(tempFilePath, sourcestream.ToArray());

                //try
                //{
                //    using (Package toExtract = Package.Open(tempFilePath, FileMode.Open, FileAccess.Read))
                //    {
                //        // Perform operations with the Package
                //    }
                //}
                //catch (Exception ex)
                //{
                //    Console.WriteLine("Error opening Package: " + ex.Message);
                //    // Handle the exception as needed
                //}
                //finally
                //{
                //    File.Delete(tempFilePath); // Clean up the temporary file
                //}


                // Read and validate the data
                byte[] readData = new byte[sourcestream.Length];
                int bytesRead = sourcestream.Read(readData, 0, readData.Length);
                using (MemoryStream copyStream = new MemoryStream())
                {
                    sourcestream.Position = 0;
                    sourcestream.CopyTo(copyStream);
                    copyStream.Position = 0;

                    try
                    {
                        List<myoject> list = new List<myoject>();
                        using (Package toExtract = Package.Open(copyStream, FileMode.Open, FileAccess.Read))
                        {
                            var workbookPart = toExtract.GetParts().Where(x => x.Uri.OriginalString.ToLower().EndsWith(".xml.rels"));
                            foreach (var ppart in workbookPart)
                            {
                                if (!ppart.Uri.OriginalString.ToLower().Contains("external"))
                                {
                                    XmlDocument doc = new XmlDocument();
                                    doc.Load(toExtract.GetPart(ppart.Uri).GetStream(FileMode.Open, FileAccess.Read));
                                    XmlNodeList xNode = doc.SelectNodes("/*[local-name()]/*[local-name()]");
                                    foreach (XmlNode eee in xNode)
                                    {
                                        if (!eee.Attributes["Target"].InnerText.ToLower().EndsWith(".vml") && !eee.Attributes["Target"].InnerText.ToLower().EndsWith(".emf") &&
                                            !eee.Attributes["Target"].InnerText.ToLower().Contains("printersettings") && !eee.Attributes["Target"].InnerText.ToLower().Contains("externallink") &&
                                            !ppart.Uri.OriginalString.ToLower().Contains("workbook"))
                                        {
                                            myoject temp = new myoject();
                                            temp.Name = ppart.Uri.OriginalString.Split("/").Last().Replace(".xml", "").Replace(".rels", "");
                                            temp.Description = eee.Attributes["Target"].InnerText.Split("/").Last().Replace(".xml", "");
                                            temp.path = eee.Attributes["Target"].InnerText;
                                            list.Add(temp);
                                        }
                                    }
                                }
                            }
                            Uri myUri = new Uri("/xl/workbook.xml", UriKind.Relative);

                            var sheetnamePart = toExtract.GetPart(myUri);
                            var sn = sheetnamePart.GetStream(FileMode.Open, FileAccess.Read);
                            XmlDocument sheetdoc = new XmlDocument();
                            sheetdoc.Load(sn);
                            XmlNamespaceManager snsmgr = new XmlNamespaceManager(sheetdoc.NameTable);
                            snsmgr.AddNamespace("ns", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
                            snsmgr.AddNamespace("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
                            XmlNodeList sxNode = sheetdoc.SelectNodes("ns:workbook/ns:sheets/ns:sheet", snsmgr);
                            foreach (XmlNode seee in sxNode)
                            {
                                myoject stemp = new myoject();
                                stemp.Name = seee.Attributes["r:id"].InnerText.Replace("rId", "sheet");
                                stemp.Description = seee.Attributes["name"].InnerText;
                                stemp.path = "SheetNames";
                                list.Add(stemp);
                            }
                            toExtract.Close();
                            result = JsonConvert.SerializeObject(list);
                        }
                    }
                    catch (Exception ex)
                    {
                        string module = System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString();
                        string message = $"Exception occured while opening package : Exception: {ex.Message} \r\n StackTrace : {ex.StackTrace}" +
                                            $"\r\n WorkitemId: {Globals.WorkitemId} \r\n MSPN: {Globals.MSPN} \r\n SI: {Globals.SI} " +
                                            $"\r\n Filename: {Globals.FileName} \r\n Fileurl: {Globals.Fileurl}";
                        ExceptionHandler.LogError(new LogException() { RunId = Globals.runid, Module = module, Message = message });

                    }
                }

            }
            catch (Exception ex)
            {

                string module = System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString();
                string message = $"Exception occured while opening package : Exception: {ex.Message} \r\n StackTrace : {ex.StackTrace}" +
                                    $"\r\n WorkitemId: {Globals.WorkitemId} \r\n MSPN: {Globals.MSPN} \r\n SI: {Globals.SI} " +
                                    $"\r\n Filename: {Globals.FileName} \r\n Fileurl: {Globals.Fileurl}";
                ExceptionHandler.LogError(new LogException() { RunId = Globals.runid, Module = module, Message = message });

            }

            return result;
        }
        [System.Runtime.Versioning.SupportedOSPlatform("windows")]
        public static void ExtractEmbeddeddata(MemoryStream sourcestream, string RunId, string WorkItemId, string MSPN, string SI, string attachmentfilename, ILogger _xlembdlog, ExecutionContext _xlembdcontext)
        {
            DataTable resultsdatatable = new DataTable();
            resultsdatatable.Columns.Add("RunId", typeof(Guid));
            resultsdatatable.Columns.Add("WorkItemId");
            resultsdatatable.Columns.Add("MSPN");
            resultsdatatable.Columns.Add("SI");
            resultsdatatable.Columns.Add("AttachmentFileName");
            resultsdatatable.Columns.Add("FileName");
            resultsdatatable.Columns.Add("FileType");
            resultsdatatable.Columns.Add("FilePath");
            resultsdatatable.Columns.Add("FileText");

            DataTable resultsimgtable = new DataTable();
            resultsimgtable.Columns.Add("RunId", typeof(Guid));
            resultsimgtable.Columns.Add("WorkItemId");
            resultsimgtable.Columns.Add("MSPN");
            resultsimgtable.Columns.Add("SI");
            resultsimgtable.Columns.Add("AttachmentName");
            resultsimgtable.Columns.Add("FileName");
            resultsimgtable.Columns.Add("FileType");
            resultsimgtable.Columns.Add("FilePath");
            resultsimgtable.Columns.Add("FileData", typeof(SqlBinary));

            string inpath = tempinputfolderpath;
            string output = tempoutputfolderpath;
            List<myoject> list = new List<myoject>();

            sourcestream.Position = 0;

            using (Package toExtract = Package.Open(sourcestream, FileMode.Open, FileAccess.Read))
            {
                foreach (PackagePart pPart in toExtract.GetParts())
                {
                    _xlembdlog.LogInformation($"Starting {pPart.Uri.OriginalString}");
                    inpath = tempinputfolderpath;
                    if (pPart.Uri.ToString().Contains("embeddings", StringComparison.InvariantCultureIgnoreCase) && pPart.Uri.ToString().EndsWith(".bin", StringComparison.InvariantCultureIgnoreCase))
                    {
                        PackagePart embeddingPart = toExtract.GetPart(pPart.Uri);
                        if (embeddingPart != null)
                        {
                            Stream s = pPart.GetStream();
                            inpath = inpath + pPart.Uri.ToString().Substring(pPart.Uri.ToString().LastIndexOf("/") + 1);
                            System.IO.FileStream writeStream = new System.IO.FileStream(inpath, FileMode.Create, FileAccess.Write);
                            ReadWriteStream(s, writeStream);
                            if (Ole10Native.ExtractFile(inpath, output) == false)
                            {
                                Stream otherstream = pPart.GetStream();
                                MemoryStream msbin = new MemoryStream();
                                otherstream.CopyTo(msbin);
                                _xlembdlog.LogInformation($"PDF extraction");
                                try
                                {
                                    msbin.Position = 0;
                                    string filetype = PDFExtractor.extractPDF(msbin, "%PDF") ? "PDF" : "Unknown";
                                    resultsdatatable.Rows.Add(new object[] { RunId, WorkItemId, MSPN, SI, attachmentfilename, pPart.Uri.OriginalString.Split("/").Last(), filetype, pPart.Uri.OriginalString, "" });
                                }
                                catch (Exception pdfexception)
                                {
                                    _xlembdlog.LogInformation($"PDF extraction {pdfexception.Message}");
                                    string module = System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString();
                                    string message = $"Exception occured while extracting pdf : Exception: {pdfexception.Message} \r\n StackTrace : {pdfexception.StackTrace}" +
                                                        $"\r\n WorkitemId: {Globals.WorkitemId} \r\n MSPN: {Globals.MSPN} \r\n SI: {Globals.SI} " +
                                                        $"\r\n Filename: {Globals.FileName} \r\n Fileurl: {Globals.Fileurl}";
                                    ExceptionHandler.LogError(new LogException() { RunId = Globals.runid, Module = module, Message = message });


                                }
                            }
                            writeStream.Close();
                            writeStream.Dispose();
                            s.Close();
                            s.Dispose();
                        }
                    }
                    else if (pPart.Uri.ToString().Contains("embeddings", StringComparison.InvariantCultureIgnoreCase) &&
                              (pPart.Uri.ToString().EndsWith(".xlsm", StringComparison.InvariantCultureIgnoreCase) ||
                                pPart.Uri.ToString().EndsWith(".xlsx", StringComparison.InvariantCultureIgnoreCase) ||
                                pPart.Uri.ToString().EndsWith(".xls", StringComparison.InvariantCultureIgnoreCase)
                              )
                            )
                    {
                        string xlsmfilename = "";
                        string xlsmfilepath = pPart.Uri.OriginalString;
                        string csvData = "";
                        PackagePart embeddingPart = toExtract.GetPart(pPart.Uri);
                        if (embeddingPart != null)
                        {
                            using (Stream embeddedxlsmstream = embeddingPart.GetStream())
                            {
                                using (MemoryStream embeddedxlsmmemorystream = new MemoryStream())
                                {
                                    embeddedxlsmstream.CopyTo(embeddedxlsmmemorystream);
                                    DataTable csvTable = ExDr.ReadExcelAndReturnResult(embeddedxlsmmemorystream, "", WorkItemId, SI, MSPN, RunId, attachmentfilename, 0, "XLSX", _xlembdlog, _xlembdcontext);
                                    xlsmfilename = csvTable.Rows[0][2].ToString();
                                    csvData = csvTable.Rows[0][3].ToString();
                                }
                            }

                        }
                        resultsdatatable.Rows.Add(new object[] { RunId, WorkItemId, MSPN, SI, attachmentfilename, xlsmfilename, "xlsm", xlsmfilepath, csvData });
                    }
                    else if (pPart.Uri.OriginalString.EndsWith(".jpeg", StringComparison.InvariantCultureIgnoreCase) ||
                     pPart.Uri.OriginalString.EndsWith(".jpg", StringComparison.InvariantCultureIgnoreCase) ||
                     pPart.Uri.OriginalString.EndsWith(".png", StringComparison.InvariantCultureIgnoreCase)
                    )
                    {
                        PackagePart embeddingPart = toExtract.GetPart(pPart.Uri);
                        using (Stream embeddedimgstream = embeddingPart.GetStream())
                        {
                            MemoryStream ms = new MemoryStream();
                            embeddedimgstream.CopyTo(ms);
                            ms.Position = 0;
                            MemoryStream msnew = new MemoryStream();
                            try
                            {
                                msnew = ImageCompressorHelper.CompressImageData(ms);
                            }
                            catch (Exception imgex)
                            {
                                _xlembdlog.LogInformation($"Image Processing error {imgex.Message}");
                                string module = System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString();
                                string message = $"Exception occured while reading image file : Exception: {imgex.Message} \r\n StackTrace : {imgex.StackTrace}" +
                                                    $"\r\n WorkitemId: {Globals.WorkitemId} \r\n MSPN: {Globals.MSPN} \r\n SI: {Globals.SI} " +
                                                    $"\r\n Filename: {Globals.FileName} \r\n Fileurl: {Globals.Fileurl}";
                                ExceptionHandler.LogError(new LogException() { RunId = Globals.runid, Module = module, Message = message });

                            }
                            msnew.Position = 0;
                            resultsimgtable.Rows.Add(new object[] { RunId, WorkItemId, MSPN, SI, attachmentfilename, pPart.Uri.OriginalString.Split("/").Last(), pPart.Uri.OriginalString.Split(".").Last(), pPart.Uri.OriginalString, msnew.ToArray() });
                        }
                    }
                }
            }
            output = tempoutputfolderpath;
            foreach (var f in Directory.GetFiles(output))
            {
                var dt = File.OpenRead(f);
                string filename = Path.GetFileName(dt.Name);
                string filepath = filename.Contains('}') ? filename.Substring(0, filename.IndexOf("}")) : filename;
                filepath = filepath.Replace("{", "").Replace("}", "");
                filename = filename.Contains('}') ? filename.Substring(filename.IndexOf("}") + 1) : filename.Replace(Path.GetExtension(dt.Name), "");
                string filetype = Path.GetExtension(dt.Name);
                string contents;
                using (var sr = new StreamReader(dt))
                {
                    contents = sr.ReadToEnd();
                }
                resultsdatatable.Rows.Add(new object[] { RunId, WorkItemId, MSPN, SI, attachmentfilename, filename, filetype, filepath, contents });
            }
            _xlembdlog.LogInformation($"No Of Rows in result {resultsdatatable.Rows.Count}");
            output = tempoutputfolderpath;
            System.GC.Collect();
            System.GC.WaitForPendingFinalizers();

            sourcestream.Dispose();

            if (resultsdatatable.Rows.Count > 0)
            {
                SqlHelper.LoadDataTableToSql(resultsdatatable, true, "Stage_FAI_Templates_EmbeddedData");
            }
            if (resultsimgtable.Rows.Count > 0)
            {
                SqlHelper.LoadDataTableToSql(resultsimgtable, true, "Stage_FAI_Templates_ImageData");
            }
        }
        private static void ReadWriteStream(Stream readStream, Stream writeStream)
        {
            int Length = 256;
            Byte[] buffer = new Byte[Length];
            int bytesRead = readStream.Read(buffer, 0, Length);
            // write the required bytes
            while (bytesRead > 0)
            {
                writeStream.Write(buffer, 0, bytesRead);
                bytesRead = readStream.Read(buffer, 0, Length);
            }
            readStream.Close();
            writeStream.Close();
        }
        public class myoject
        {
            public string Name { get; set; }
            public string Description { get; set; }
            public string path { get; set; }
        }
    }
}
