namespace ExcelWriter
{
    using System.IO;

    /// <summary>
    /// Represents the staus of the result
    /// </summary>
    public enum MemoryStreamResultStatus
    {
        Unknown,
        Success,
        Failure,
    }

    /// <summary>
    /// Represents a <see cref="MemoryStream"/> created to return report information to client and/or other services. 
    /// </summary>
    public class MemoryStreamResult
    {
        /// <summary>
        /// Gets or sets a <see cref="MemoryStream"/> which contains the report.
        /// </summary>
        public MemoryStream MemoryStream { get; set; }

        /// <summary>
        /// Gets or sets any error which may have been encountered during the report generation process.
        /// </summary>
        public string ErrorMessage { get; set; }

        /// <summary>
        /// Gets or sets the staus of the <see cref="MemoryStreamResult"/>
        /// </summary>
        public MemoryStreamResultStatus Status { get; set; }
    }
}
