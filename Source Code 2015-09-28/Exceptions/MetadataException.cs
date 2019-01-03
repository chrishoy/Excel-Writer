using System;
using System.Runtime.Serialization;

namespace ExcelWriter
{
    public class MetadataException : Exception
    {
        public MetadataException() : base()
        { }
        
        public MetadataException(string message) : base(message)
        { }

        protected MetadataException(SerializationInfo info, StreamingContext context) : base(info, context)
        { }
        
        public MetadataException(string message, Exception innerException) : base(message, innerException)
        { }
    }
}
