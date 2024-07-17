using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;

namespace FAITemplateParser.Utilities
{
    [System.Runtime.Versioning.SupportedOSPlatform("windows")]
    public class ImageCompressorHelper
    {
        [System.Runtime.Versioning.SupportedOSPlatform("windows")]
        public static MemoryStream CompressImageData(MemoryStream imagememorystream)
        {
            MemoryStream result = new MemoryStream();
            try
            {
                using (Image Oldimage = Image.FromStream(imagememorystream))
                {
                    using (Image nimage = compressImage(Oldimage, 100, 100, 35))
                    {
                        nimage.Save(result, ImageFormat.Jpeg);
                    }
                }
                if (imagememorystream.ToArray().Length > result.ToArray().Length)
                {
                    return result;
                }
                else
                {
                    return imagememorystream;
                }
            }
            catch (Exception ex)
            {
                string module = System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString();
                string message = $"Exception occured compressing image : Exception: {ex.Message} \r\n StackTrace : {ex.StackTrace}" +
                                    $"\r\n WorkitemId: {Globals.WorkitemId} \r\n MSPN: {Globals.MSPN} \r\n SI: {Globals.SI} " +
                                    $"\r\n Filename: {Globals.FileName} \r\n Fileurl: {Globals.Fileurl}";
                ExceptionHandler.LogError(new LogException() { RunId = Globals.runid, Module = module, Message = message });

                return imagememorystream;
            }
        }
        [System.Runtime.Versioning.SupportedOSPlatform("windows")]
        private static Image compressImage(Image OldImage, int newWidth, int newHeight,
                            int newQuality)   // set quality to 1-100, eg 50
        {
            using (Image memImage = new Bitmap(OldImage, OldImage.Width, OldImage.Height))
            {
                ImageCodecInfo myImageCodecInfo;
                System.Drawing.Imaging.Encoder myEncoder;
                EncoderParameter myEncoderParameter;
                EncoderParameters myEncoderParameters;
                string _mimeType = "image/jpeg";
                myImageCodecInfo = GetEncoderInfo(_mimeType);
                myEncoder = System.Drawing.Imaging.Encoder.Quality;
                myEncoderParameters = new EncoderParameters(1);
                myEncoderParameter = new EncoderParameter(myEncoder, newQuality);
                myEncoderParameters.Param[0] = myEncoderParameter;

                MemoryStream memStream = new MemoryStream();
                memImage.Save(memStream, myImageCodecInfo, myEncoderParameters);
                Image newImage = Image.FromStream(memStream);
                ImageAttributes imageAttributes = new ImageAttributes();
                using (Graphics g = Graphics.FromImage(newImage))
                {
                    g.InterpolationMode =
                      System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;  //**
                    g.DrawImage(newImage, new Rectangle(System.Drawing.Point.Empty, newImage.Size), 0, 0,
                      newImage.Width, newImage.Height, System.Drawing.GraphicsUnit.Pixel, imageAttributes);
                }
                return newImage;
            }
        }
        [System.Runtime.Versioning.SupportedOSPlatform("windows")]
        private static ImageCodecInfo GetEncoderInfo(String mimeType)
        {
            ImageCodecInfo[] encoders;
            encoders = ImageCodecInfo.GetImageEncoders();
            foreach (ImageCodecInfo ici in encoders)
                if (ici.MimeType == mimeType) return ici;

            return null;
        }
    }
}
