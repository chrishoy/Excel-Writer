using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

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
