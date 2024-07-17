using Microsoft.Azure.WebJobs;
using Microsoft.Extensions.Logging;
using NPOI.XSSF.UserModel;
using System;
using System.IO;


namespace FAITemplateParser.Utilities
{
    [System.Runtime.Versioning.SupportedOSPlatform("windows")]
    public class NPOIExcelDataReader
    {
        public static System.Data.DataTable ReadExcelImages(MemoryStream memoryStream, string SI, string MSPN, string FactoryCode, string attachmentname, string fileType, ILogger xlseclog, ExecutionContext xlseccontext)
        {
            memoryStream.Position = 0;
            System.Data.DataTable result = new System.Data.DataTable();
            //result.Columns.Add("RunId",typeof(Guid));
            //result.Columns.Add("Workitemid");
            result.Columns.Add("SI");
            result.Columns.Add("MSPN");
            result.Columns.Add("FactoryCode");
            result.Columns.Add("FileName");
            result.Columns.Add("TabName");
            result.Columns.Add("PictureName");
            result.Columns.Add("PictureType");
            result.Columns.Add("ColumnFrom");
            result.Columns.Add("RowFrom");
            result.Columns.Add("ColumnTo");
            result.Columns.Add("RowTo");
            result.Columns.Add("ColumnFromCellAddress");
            result.Columns.Add("RowFromCellAddress");
            result.Columns.Add("ColumnToCellAddress");
            result.Columns.Add("RowToCellAddress");
            result.Columns.Add("Data", typeof(byte[]));

            try
            {
                memoryStream.Position = 0;
                Console.WriteLine(memoryStream.Length);
                var workbook = new XSSFWorkbook(memoryStream);
                int numsheets = workbook.NumberOfSheets;

                for (int i = 0; i < numsheets; i++)
                {
                    var worksheet = workbook.GetSheetAt(i) as XSSFSheet;

                    var drawing = worksheet.GetDrawingPatriarch() as XSSFDrawing;
                    if (drawing != null)
                    {

                        var shapes = drawing.GetShapes();

                        foreach (var shape in shapes)
                        {
                            if (shape is XSSFPicture)
                            {
                                var picture = (XSSFPicture)shape;

                                var col1 = -1;
                                var row1 = -1;
                                var col2 = -1;
                                var row2 = -1;

                                var pictureData = picture.PictureData.Data;
                                var picturename = picture.PictureData != null ? picture.PictureData.MimeType : "";
                                var picturetype = picture.PictureData != null ? picture.PictureData.PictureType.ToString() : "";
                                var picturePosition = picture.ClientAnchor;
                                var colfromcelladdr = String.Empty;
                                int rowfromcelladdr = -1;
                                var coltocelladdr = String.Empty;
                                int rowtocelladdr = -1;

                                if (picturePosition is XSSFClientAnchor)
                                {
                                    try
                                    {
                                        col1 = picturePosition.Col1 != null && picturePosition?.Col1 >= 0 ? picturePosition.Col1 : -1;
                                        row1 = picturePosition.Row1 != null && picturePosition?.Row1 >= 0 ? picturePosition.Row1 : -1;
                                        col2 = picturePosition.Col2 != null && picturePosition?.Col2 >= 0 ? picturePosition.Col2 : -1;
                                        row2 = picturePosition.Row2 != null && picturePosition?.Row2 >= 0 ? picturePosition.Row2 : -1;

                                        colfromcelladdr = GetCellAddress(col1, row1);
                                        coltocelladdr = GetCellAddress(col2, row2);
                                        rowfromcelladdr = row1 + 1;
                                        rowtocelladdr = row2 + 1;
                                    }
                                    catch (Exception ex)
                                    {
                                        if (col1 >= 0 && row1 >= 0)
                                        {
                                            colfromcelladdr = GetCellAddress(col1, row1);
                                            rowfromcelladdr = row1 + 1;
                                        }

                                        if (col2 >= 0 && row2 >= 0)
                                        {
                                            coltocelladdr = GetCellAddress(col2, row2);
                                            rowtocelladdr = row2 + 1;
                                        }

                                        result.Rows.Add(new object[] {  SI, MSPN,FactoryCode, attachmentname, worksheet.SheetName, picturename
                                                            , picturetype, col1, row1, col2, row2, colfromcelladdr, rowfromcelladdr
                                                            , coltocelladdr, rowtocelladdr, pictureData });

                                        string module = System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString();
                                        string message = $"Exception occured while mapping cell positions : Exception: {ex.Message} \r\n StackTrace : {ex.StackTrace}" +
                                                            $"\r\n WorkitemId: {Globals.WorkitemId} \r\n MSPN: {Globals.MSPN} \r\n SI: {Globals.SI} " +
                                                            $"\r\n Filename: {Globals.FileName} \r\n Fileurl: {Globals.Fileurl}";
                                        //ExceptionHandler.LogError(new LogException() { RunId = Globals.runid, Module = module, Message = message });


                                        continue;
                                    }

                                }

                                result.Rows.Add(new object[] { SI, MSPN,FactoryCode, attachmentname, worksheet.SheetName, picturename
                                                            , picturetype, col1, row1, col2, row2, colfromcelladdr, rowfromcelladdr
                                                            , coltocelladdr, rowtocelladdr, pictureData });
                            }

                        }

                    }
                }
            }

            catch (Exception ex)
            {
                //xlseclog.LogInformation(String.Format("INSERT INTO [dbo].[FAI_Templates_ErrorLogs]([AttachmentId],[ErrorLocation],[ErrorMessage]) VALUES ({0},'{1}','{2}')", Workitemid, System.Reflection.MethodBase.GetCurrentMethod().Name, ex.Message.Replace("'", "''")));
                //SqlHelper.ExecuteSqlQuery_NoResult(String.Format("INSERT INTO [dbo].[FAI_Templates_ErrorLogs]([AttachmentId],[ErrorLocation],[ErrorMessage]) VALUES ({0},'{1}','{2}')", Workitemid, System.Reflection.MethodBase.GetCurrentMethod().Name, ex.Message.Replace("'", "''")));
                string module = System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString();
                string message = $"Exception occured while reading excel images : Exception: {ex.Message} \r\n StackTrace : {ex.StackTrace}" +
                                    $"\r\n WorkitemId: {Globals.WorkitemId} \r\n MSPN: {Globals.MSPN} \r\n SI: {Globals.SI} " +
                                    $"\r\n Filename: {Globals.FileName} \r\n Fileurl: {Globals.Fileurl}";
                ExceptionHandler.LogError(new LogException() { RunId = Globals.runid, Module = module, Message = message });

                //throw;
            }

            return result;
        }

        private static string GetImageCellAddress(XSSFClientAnchor anchor)
        {
            int dx1 = anchor.Dx1 / 256; // top-left x coordinate in units of 1/256th of a character width
            int dy1 = anchor.Dy1 / 256; // top-left y coordinate in units of 1/256th of a character height
            int dx2 = anchor.Dx2 / 256; // bottom-right x coordinate in units of 1/256th of a character width
            int dy2 = anchor.Dy2 / 256; // bottom-right y coordinate in units of 1/256th of a character height

            int col1 = anchor.Col1 + dx1; // top-left column index
            int row1 = anchor.Row1 + dy1; // top-left row index
            int col2 = anchor.Col1 + dx2; // bottom-right column index
            int row2 = anchor.Row1 + dy2; // bottom-right row index

            string cell1 = GetCellAddress(col1, row1); // get cell address for top-left corner
            string cell2 = GetCellAddress(col2, row2); // get cell address for bottom-right corner

            return $"{cell1}:{cell2}"; // return cell range
        }

        private static string GetCellAddress(int colIndex, int rowIndex)
        {
            int div = colIndex + 1;
            string columnString = string.Empty;
            int mod = 0;

            while (div > 0)
            {
                mod = (div - 1) % 26;
                columnString = (char)(65 + mod) + columnString;
                div = (int)((div - mod) / 26);
            }

            return $"{columnString}";
        }
    }
}
