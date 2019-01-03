using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelWriter
{
    public interface ILogger
    {
        void Log(LogType logType, string assembly, string message, string typeName);
        void Log(LogType logType, string assembly, string message, string typeName, string message2);
        void Log(LogType logType, string assembly, string message, Exception ex);
    }
}
