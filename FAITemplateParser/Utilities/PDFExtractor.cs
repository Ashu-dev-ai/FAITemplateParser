using System;
using System.IO;
using System.Linq;
using System.Text;


namespace FAITemplateParser.Utilities
{
    public class PDFExtractor
    {
        public static bool extractPDF(MemoryStream memoryStream, string searchstring)
        {
            bool resultstring = false;
            byte[] bytes = Encoding.ASCII.GetBytes(searchstring);
            memoryStream.Position = 0;
            int location = FindPosition(memoryStream, bytes);
            if (location > 1)
            {
                resultstring = true;
            }
            return resultstring;
        }
        public static int FindPosition(MemoryStream stream, byte[] byteSequence)
        {
            if (byteSequence.Length > stream.Length)
                return -1;

            byte[] buffer = new byte[byteSequence.Length];

            BufferedStream bufStream = new BufferedStream(stream, byteSequence.Length);
            int i;
            while ((i = bufStream.Read(buffer, 0, byteSequence.Length)) == byteSequence.Length)
            {
                if (byteSequence.SequenceEqual(buffer))
                    return Convert.ToInt32(bufStream.Position - byteSequence.Length);
                else
                    bufStream.Position -= byteSequence.Length - PadLeftSequence(buffer, byteSequence);
            }
            bufStream.Close();
            bufStream.Dispose();
            return -1;
        }
        private static int PadLeftSequence(byte[] bytes, byte[] seqBytes)
        {
            int i = 1;
            while (i < bytes.Length)
            {
                int n = bytes.Length - i;
                byte[] aux1 = new byte[n];
                byte[] aux2 = new byte[n];
                Array.Copy(bytes, i, aux1, 0, n);
                Array.Copy(seqBytes, aux2, n);
                if (aux1.SequenceEqual(aux2))
                    return i;
                i++;
            }
            return i;
        }
    }
}
