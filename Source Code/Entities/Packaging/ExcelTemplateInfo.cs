using DocumentFormat.OpenXml.Packaging;
using System;

namespace ExcelWriter
{
    /// <summary>
    /// Represents information about a template
    /// </summary>
    internal sealed class ExcelTemplateInfo
    {
        public ExcelTemplateInfo(string templateId, string resourceString, string templateFileName)
        {
            if (string.IsNullOrEmpty(templateId))
            {
                throw new ArgumentNullException("templateId");
            }
            if (string.IsNullOrEmpty(resourceString))
            {
                throw new ArgumentNullException("resourceString");
            }
            if (string.IsNullOrEmpty(templateFileName))
            {
                throw new ArgumentNullException("templateFileName");
            }
            //if (excelMapStyles == null)
            //{
            //    throw new ArgumentNullException("excelMapStyles");
            //}

            this.TemplateId = templateId;
            this.ResourceString = resourceString;
            this.TemplateFileName = templateFileName;
            //this.ExcelMapStyles = excelMapStyles;
        }

        /// <summary>
        /// 
        /// </summary>
        public string TemplateId { get; private set; }

        /// <summary>
        /// 
        /// </summary>
        public string ResourceString { get; private set; }

        /// <summary>
        /// 
        /// </summary>
        public string TemplateFileName { get; private set; }

        ///// <summary>
        ///// Dictionary of <see cref="ExcelMapStyle"/>s that relate to this template.
        ///// </summary>
        //internal ExcelMapStylesDictionary ExcelMapStyles { get; private set; }
    }
}
