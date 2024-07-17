using Spire.Xls;
using System.IO;

namespace FAITemplateParser.Utilities
{
    public class XLVersionModifier
    {
        public static MemoryStream ConvertXLVersionTo2016(MemoryStream memoryStream)
        {
            MemoryStream result = new MemoryStream();

            Spire.Xls.Workbook workbook = new Spire.Xls.Workbook();
            workbook.LoadFromStream(memoryStream);
            workbook.SaveToStream(result, FileFormat.Version2016);
            result.Position = 0;
            return result;
        }
    }
}
