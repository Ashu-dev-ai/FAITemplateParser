using System.Reflection;

namespace FAITemplateParser.Utilities
{
    public class ExceptionHandler
    {
        public static void LogError(LogException exp)
        {
            exp.ExceptionLevel = "Error";
            LogExceptiontoDB(exp);

        }

        public static void LogWarning(LogException exp)
        {
            exp.ExceptionLevel = "Warning";
            LogExceptiontoDB(exp);
        }

        public static void LogMessage(LogException exp)
        {
            exp.ExceptionLevel = "Message";
            LogExceptiontoDB(exp);
        }

        public static void LogSuccess(LogException exp)
        {
            exp.ExceptionLevel = "Success";
            LogExceptiontoDB(exp);

        }

        public static void LogExceptiontoDB(LogException exp)
        {

            int counter = 0;

            var props = typeof(LogException).GetProperties();

            Microsoft.Data.SqlClient.SqlParameter[] sp_params = new Microsoft.Data.SqlClient.SqlParameter[props.Length];

            foreach (var p in props)
            {
                Microsoft.Data.SqlClient.SqlParameter param = new Microsoft.Data.SqlClient.SqlParameter();
                param.ParameterName = p.Name;
                PropertyInfo property = exp.GetType().GetProperty(p.Name);
                param.Value = property.GetValue(exp, null);

                sp_params.SetValue(param, counter);
                counter++;
            }

            var result = SqlHelper.ExecuteStoredProcedureandReturnDatasetResult("usp_FAItemplateParserAutomationLogger", sp_params);

        }
    }
}
