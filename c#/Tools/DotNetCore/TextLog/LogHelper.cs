using Newtonsoft.Json;
using System;
using System.IO;

namespace DotNetCore.TextLog
{
    class LogHelper
    {
        public const string LOG_DIR = "Logs";
        public const string LOG_SUFFIX = "log";

        public static void WriteLog(Exception e)
        {
            WriteLog(JsonConvert.SerializeObject(e));
        }

        public static void WriteLog(string log)
        {
            var dataDir = AppDomain.CurrentDomain.BaseDirectory;
            if (!Directory.Exists(Path.Combine(dataDir, LogHelper.LOG_DIR)))
            {
                Directory.CreateDirectory(Path.Combine(dataDir, LogHelper.LOG_DIR));
            }

            var currentMonthLogPath = Path.Combine(dataDir, LogHelper.LOG_DIR, LogHelper.GetCurrentMonthString());
            var currentMonthLogFile = currentMonthLogPath + "." + LogHelper.LOG_SUFFIX;
            if (!File.Exists(currentMonthLogPath + "." + LogHelper.LOG_SUFFIX))
            {
                using (var writer = File.Create(currentMonthLogFile)) { }
            }

            using (var writer = File.AppendText(currentMonthLogFile))
            {
                writer.WriteLine(GetLogString(log));
            }
        }

        private static string GetCurrentMonthString()
        {
            var currentDate = DateTime.Now;
            return string.Format("{0}-{1}", currentDate.Year, currentDate.Month);
        }

        private static string GetLogString(string log)
        {
            return string.Format("{0} - {1}", DateTime.Now.ToString(), log);
        }
    }
}
