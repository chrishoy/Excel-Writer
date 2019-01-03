namespace ExcelWriter
{
    using System;
    using System.Diagnostics;

    public class DebugLogger : ILogger
    {
        public void Log(LogType logType, string assembly, string message, string typeName)
        {
            Debug.Print("{0}: {1}, {2}", logType, assembly, message);
        }

        public void Log(LogType logType, string assembly, string message, string typeName, string message2)
        {
            Debug.Print("{0}: {1}, {2}, {3}", logType, assembly, message, message2);
        }


        public void Log(LogType logType, string assembly, string message, Exception ex)
        {
            Debug.Print("{0}: {1}, {2}", logType, assembly, message);
            Debug.Print("{0}", ex);
        }
    }

    public enum LogType
    {
        Trace,
        Info,
        Fatal,
    }
}
