using System.Collections.Generic;
using System.Data;

namespace FAITemplateParser.Utilities
{
    [System.Runtime.Versioning.SupportedOSPlatform("windows")]
    public class Globals
    {

        public static string runid = "";
        public static string WorkitemId = "";
        public static string MSPN = "";
        public static string SI = "";
        public static string FileName = "";
        public static string Fileurl = "";

        public static string feedname = "";

        public static List<SharePointHelper.SharepointFIle> lstspfiles = new List<SharePointHelper.SharepointFIle>();

        public static DataTable dtProcessedFiles = new DataTable();

        public static void InitializeGlobals()
        {
            lstspfiles.Clear();
            //dtProcessedFiles.Clear();

            if (runid == "")
            {
                System.Guid guid = System.Guid.NewGuid();
                runid = guid.ToString();
            }
        }

        public static void DisposeGlobals()
        {
            runid = "";
            WorkitemId = "";
            MSPN = "";
            SI = "";
            FileName = "";
            Fileurl = "";

            lstspfiles.Clear();
            dtProcessedFiles.Clear();
            dtProcessedFiles = null;
        }
    }
}
