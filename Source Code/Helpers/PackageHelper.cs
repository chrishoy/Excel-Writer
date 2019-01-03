namespace ExcelWriter
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Linq;
    using System.Text;
    using System.Windows.Markup;
    using System.Xml;

    /// <summary>
    /// Helper class for package management.
    /// </summary>
    internal static class PackageHelper
    {
        #region Private Fields

        /// <summary>
        /// Used for general thread locking.
        /// </summary>
        private static object padLock = new object();

        #endregion Private Fields

        #region Public Methods

        /// <summary>
        /// Checks to see if the file has been copied locally (to a temp folder).<br/>
        /// If it has then returns a path to this file. If not then pulls down from source location and then returns path to the copy.<br/>
        /// This is an attempt to prevent contention when multiple users attempt to use the same source package.
        /// Simply uses name and CreationTime for match...
        /// </summary>
        /// <param name="sourceFilePath">Source path and file name.</param>
        /// <returns>A full path to a copy in the local temp folder.</returns>
        public static string CheckAndCopyFileLocally(string sourceFilePath)
        {
            string localCopyFilePath = string.Format(@"{0}{1}", Path.GetTempPath(), Path.GetFileName(sourceFilePath));

            lock (padLock)
            {
                DateTime sourceFileCreationTime = File.GetCreationTime(sourceFilePath);

                if (!File.Exists(localCopyFilePath))
                {
                    // No local copy, take one and return path
                    {
                        File.Copy(sourceFilePath, localCopyFilePath);
                        File.SetAttributes(localCopyFilePath, FileAttributes.Normal);
                        File.SetCreationTime(localCopyFilePath, sourceFileCreationTime);
                    }
                }
                else if (File.GetCreationTime(localCopyFilePath) != sourceFileCreationTime)
                {
                    // Local copy exists.... check creation time
                    // Delete original copy
                    File.SetAttributes(localCopyFilePath, FileAttributes.Normal);
                    File.Delete(localCopyFilePath);

                    // Create a new copy and return path
                    File.Copy(sourceFilePath, localCopyFilePath);
                    File.SetAttributes(localCopyFilePath, FileAttributes.Normal);
                    File.SetCreationTime(localCopyFilePath, sourceFileCreationTime);
                }
            }

            return localCopyFilePath;
        }

        /// <summary>
        /// Clones an instance of a XAML element.
        /// </summary>
        /// <typeparam name="T">The type of the element being cloned</typeparam>
        /// <param name="source">The value ot be cloned</param>
        /// <returns>A cloned copy</returns>
        public static T CloneXamlInstance<T>(T source)
        {
            string xml = XamlWriter.Save(source);
            T clone;

            using (var sr = new StringReader(xml))
            {
                using (var xr = new XmlTextReader(sr))
                {
                    clone = (T)XamlReader.Load(xr);
                }
            }

            return clone;
        }

        #endregion Public Methods
    }
}
