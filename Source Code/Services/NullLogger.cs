using System;

namespace ExcelWriter.Interface
{
    public class NullLogger : ILogger
    {
        public void Log(LogType logType, string assembly, string message, string typeName)
        {
        }

        public void Log(LogType logType, string assembly, string message, string typeName, string message2)
        {
        }

        public void Log(LogType logType, string assembly, string message, Exception ex)
        {
        }
    }
}
