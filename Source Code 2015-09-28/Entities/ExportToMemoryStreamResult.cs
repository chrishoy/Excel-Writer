namespace ExcelWriter
{
    using System;
    using System.IO;

    public class ExportToMemoryStreamResult
    {
        public Exception Error { get; set; }
        public MemoryStream MemoryStream { get; set; }
    }
}
