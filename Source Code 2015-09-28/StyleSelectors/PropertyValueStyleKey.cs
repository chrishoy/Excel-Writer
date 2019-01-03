namespace ExcelWriter
{
    using System;

    /// <summary>
    /// Represents a style key to be returned when a property value matches a specified value.
    /// </summary>
    public class PropertyValueStyleKey
    {
        /// <summary>
        /// Gets or sets the value to be matched
        /// </summary>
        public object Value { get; set; }

        /// <summary>
        /// Gets or sets the style key to be returned when the value matches.
        /// </summary>
        public string StyleKey { get; set; }
    }
}
