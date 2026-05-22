using System;
using System.IO;

namespace EliteSheets.Services
{
    public static class Logger
    {
        private static readonly object _lock = new object();

        public static void Log(string message, Exception ex = null)
        {
            try
            {
                string appDataPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
                string logDirectory = Path.Combine(appDataPath, "RKTools", "Logs");
                Directory.CreateDirectory(logDirectory);

                string dateStr = DateTime.Now.ToString("yyyyMMdd");
                string logFilePath = Path.Combine(logDirectory, $"EliteSheets_{dateStr}.log");

                string timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff");
                string level = ex != null ? "ERROR" : "INFO";

                string logEntry = $"[{timestamp}] [{level}] {message}";
                if (ex != null)
                {
                    logEntry += $"\nException: {ex.GetType().Name}\nMessage: {ex.Message}\nStackTrace: {ex.StackTrace}";
                }

                lock (_lock)
                {
                    File.AppendAllText(logFilePath, logEntry + Environment.NewLine);
                }
            }
            catch
            {
                // Fallback: If we fail to log to the file, we can't do much without causing issues in the main application.
                // Outputting to debug is safe.
                System.Diagnostics.Debug.WriteLine("Logger failed to write log entry.");
            }
        }
    }
}
