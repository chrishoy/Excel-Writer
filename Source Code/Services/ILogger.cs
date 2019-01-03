using System;

namespace ExcelWriter
{
    public interface ILogger
    {
        void Log(LogType logType, string assembly, string message, string typeName);
        void Log(LogType logType, string assembly, string message, string typeName, string message2);
        void Log(LogType logType, string assembly, string message, Exception ex);
    }
}
