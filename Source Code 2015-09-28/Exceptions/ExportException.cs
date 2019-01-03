using System;
using System.Runtime.Serialization;

namespace ExcelWriter
{
    public class ExportException : Exception
    {
        public ExportException() : base()
        { }
        
        public ExportException(string message) : base(message)
        { }

        protected ExportException(SerializationInfo info, StreamingContext context) : base(info, context)
        { }

        public ExportException(string message, Exception innerException)
            : base(message, innerException)
        { }
    }
}
