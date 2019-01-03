namespace ExcelWriter
{
    using System;
    using System.Collections.Generic;
    using System.Diagnostics;
    using System.IO;
    using System.IO.Packaging;
    using System.Linq;
    using System.Net;
    using System.Xml;
    using System.Xml.Linq;

    public sealed class ResourcePackage
    {
        private static object padLock = new object();

        private ResourceStore resourceStore;

        private const string FilesPathPart = "TemplateFiles";
        private const string MetadataPathPart = "Metadata";
        private const string UriDelim = @"/";

        /// <summary>
        /// Initializes a new instance of the <see cref="ResourcePackage"/> class.
        /// </summary>
        public ResourcePackage()
        {
            this.resourceStore = new ResourceStore();
        }

        /// <summary>
        /// Packs the specified directory.
        /// </summary>
        /// <param name="directory">The directory.</param>
        /// <returns></returns>
        public static string Pack(string directory)
        {
            try
            {
                writeOutput = false;

                if (string.IsNullOrEmpty(directory))
                {
                    throw new ArgumentNullException("directory");
                }

                var di = new DirectoryInfo(directory);
                if (!di.Exists)
                {
                    throw new MetadataException(string.Format("Directory not found <{0}>", directory));
                }

                Output(string.Format("Starting packaging of <{0}>", directory));

                var filePath = string.Concat(directory, ".zip"); // excel template metadata
                FileInfo fi = new FileInfo(filePath);
                if (fi.Exists)
                {
                    Output(string.Format("Deleting existing package <{0}>", filePath));
                    fi.Delete();
                }

                // check there's a metadata directory
                DirectoryInfo metadataDirInfo = new DirectoryInfo(string.Concat(directory, System.IO.Path.DirectorySeparatorChar, MetadataPathPart));
                if (!metadataDirInfo.Exists)
                {
                    throw new Exception(string.Format("Directory <{0}> not found", metadataDirInfo.FullName));
                }

                Output(string.Format("Metadata directory found: <{0}>", metadataDirInfo.FullName));

                // and check there's a template files directory
                DirectoryInfo filesDirInfo = new DirectoryInfo(string.Concat(directory, System.IO.Path.DirectorySeparatorChar, FilesPathPart));
                if (!filesDirInfo.Exists)
                {
                    throw new Exception(string.Format("Directory <{0}> not found", filesDirInfo.FullName));
                }

                Output(string.Format("TemplateFiles directory found: <{0}>", filesDirInfo.FullName));

                using (var zip = ZipPackage.Open(filePath, System.IO.FileMode.CreateNew, System.IO.FileAccess.ReadWrite))
                {
                    // read all the templates files into our internal store
                    foreach (var f in metadataDirInfo.GetFiles())
                    {
                        //string uriString = string.Format("{0}{1}{0}{2}", UriDelim, MetadataPathPart, f.Name.Replace(f.Extension, null));
                        string uriString = string.Format("{0}{1}{0}{2}", UriDelim, MetadataPathPart, f.Name);

                        Uri partUri = new Uri(uriString, UriKind.Relative);
                        PackagePart part = zip.CreatePart(partUri, System.Net.Mime.MediaTypeNames.Application.Zip, CompressionOption.Normal);

                        byte[] data = File.ReadAllBytes(f.FullName);
                        part.GetStream().Write(data, 0, data.Length);

                        Output(string.Format("Metadata part created: Uri <{0}> Size <{1}>", uriString, data.Length));
                    }

                    foreach (var f in filesDirInfo.GetFiles())
                    {
                        //string uriString = string.Format("{0}{1}{0}{2}", UriDelim, FilesPathPart, f.Name.Replace(f.Extension, null));
                        string uriString = string.Format("{0}{1}{0}{2}", UriDelim, FilesPathPart, f.Name);
                        Uri partUri = new Uri(uriString, UriKind.Relative);
                        PackagePart part = zip.CreatePart(partUri, System.Net.Mime.MediaTypeNames.Application.Zip, CompressionOption.Normal);

                        byte[] data = File.ReadAllBytes(f.FullName);
                        part.GetStream().Write(data, 0, data.Length);

                        Output(string.Format("TemplateFile part created: Uri <{0}> Size <{1}>", uriString, data.Length));
                    }
                }

                fi = new FileInfo(filePath);
                if (!fi.Exists)
                {
                    throw new MetadataException("New file cannot be found");
                }

                Output(string.Format("Completed packaging of <{0}>. File size <{1}>", directory, fi.Length));

                return filePath;
            }
            catch (Exception ex)
            {
                Output(ex.ToString());
                throw;
            }
            finally
            {
                writeOutput = false;
            }
        }

        /// <summary>
        /// Opens the specified URI.
        /// </summary>
        /// <param name="uri">The URI.</param>
        /// <returns></returns>
        public static ResourcePackage Open(Uri uri)
        {
            string tempFilePath = Path.GetTempFileName();
            try
            {
                WebClient webClient = new WebClient();
                webClient.Credentials = CredentialCache.DefaultCredentials;
                webClient.DownloadFile(uri, tempFilePath);

                return Open(tempFilePath);
            }
            catch (Exception ex)
            {
                throw new MetadataException(string.Format("Failing to open package <{0}>", uri.AbsoluteUri), ex);
            }
            finally
            {
                File.Delete(tempFilePath);
            }
        }

        /// <summary>
        /// Opens the specified file name.
        /// </summary>
        /// <param name="fileName">Name of the file.</param>
        /// <returns></returns>
        public static ResourcePackage Open(string fileName)
        {
            string errors = null;
            return Open(fileName, out errors);
        }

        /// <summary>
        /// Opens the specified file name.
        /// </summary>
        /// <param name="fileName">Name of the file.</param>
        /// <param name="errors">The errors.</param>
        /// <returns></returns>
        public static ResourcePackage Open(string fileName, out string errors)
        {
            // avoid contention whilst opening file, should be v quick anyway
            lock (padLock)
            {
                try
                {
                    errors = null;
                    writeOutput = false;

                    if (string.IsNullOrEmpty(fileName))
                    {
                        throw new ArgumentNullException("fileName");
                    }

                    var fi = new FileInfo(fileName);
                    if (!fi.Exists)
                    {
                        throw new MetadataException(string.Format("File not found <{0}>", fileName));
                    }

                    Output(string.Format("Starting open of package <{0}>", fileName));

                    ResourcePackage package = new ResourcePackage();

                    List<PackagePart> metadataParts = new List<PackagePart>();
                    List<PackagePart> fileParts = new List<PackagePart>();

                    using (var zip = ZipPackage.Open(fileName, FileMode.Open, FileAccess.Read))
                    {
                        Output("Zip opened");

                        foreach (var part in zip.GetParts())
                        {
                            string uriString = part.Uri.OriginalString;

                            if (uriString.StartsWith(string.Concat(UriDelim, MetadataPathPart)))
                            {
                                metadataParts.Add(part);

                                Output(string.Format("Found Metadata part: Uri <{0}>", uriString));
                            }
                            else if (uriString.StartsWith(string.Concat(UriDelim, FilesPathPart)))
                            {
                                fileParts.Add(part);

                                Output(string.Format("Found TemplateFile part: Uri <{0}>", uriString));
                            }
                        }

                        // read all the templates files into our internal store
                        foreach (var m in metadataParts)
                        {
                            long length = m.GetStream().Length;
                            byte[] data = new byte[length];
                            m.GetStream().Read(data, 0, (int)length);

                            Output(string.Format("Loading Metadata part: Uri <{0}> Size {1}", m.Uri.OriginalString, length));

                            bool hasBomb = false;
                            if (data.Length > 2 && data[0] == 0xEF && data[1] == 0xBB && data[2] == 0xBF)
                            {
                                hasBomb = true;
                                Output("Stripping Byte Order Mark");
                            }

                            string resourceString = System.Text.UTF8Encoding.UTF8.GetString(hasBomb ? data.Skip(3).ToArray() : data);

                            string error = null;
                            if (!package.TryLoadResourceString(m.Uri.OriginalString, resourceString, out error))
                            {
                                if (string.IsNullOrEmpty(errors))
                                {
                                    errors = error;
                                }
                                else
                                {
                                    errors += error;
                                }
                                errors += Environment.NewLine;
                            }
                        }

                        foreach (var f in fileParts)
                        {
                            string name = f.Uri.OriginalString.Replace(string.Concat(UriDelim, FilesPathPart, UriDelim), null);
                            long length = f.GetStream().Length;
                            byte[] data = new byte[length];
                            f.GetStream().Read(data, 0, (int)length);

                            Output(string.Format("Loading TemplateFile part: Uri <{0}> Size {1}", f.Uri.OriginalString, length));

                            package.LoadDesignerFile(name, data);
                        }
                    }

                    package.ValidateInternal();

                    return package;
                }
                catch (Exception ex)
                {
                    Output(ex.ToString());
                    throw;
                }
                finally
                {
                    writeOutput = false;
                }
            }
        }

        /// <summary>
        /// Validates the specified file name.
        /// </summary>
        /// <param name="fileName">Name of the file.</param>
        public static void Validate(string fileName)
        {
            try
            {
                writeOutput = false;

                var package = Open(fileName);
                package.ValidateInternal();
            }
            catch (Exception ex)
            {
                Output(ex.ToString());
                throw;
            }
            finally
            {
                writeOutput = false;
            }
        }

        /// <summary>
        /// Exposed resource store.
        /// Used to get Resources and designer files
        /// </summary>
        public ResourceStore ResourceStore { get { return this.resourceStore; } }

        /// <summary>
        /// Flushes this instance.
        /// </summary>
        public void Flush()
        {
            this.resourceStore.Flush();
        }

        #region Privates

        /// <summary>
        /// Check all the template files referenced in the templates metadata have been loaded
        /// </summary>
        private void ValidateInternal()
        {
            Output("Validating package");
            this.resourceStore.Validate();
        }

        /// <summary>
        /// Tries the load template resource string.
        /// </summary>
        /// <param name="resourceString">The resource string.</param>
        /// <param name="error">The error.</param>
        /// <returns>True of no errors loading else false</returns>
        private bool TryLoadResourceString(string uri, string resourceString, out string error)
        {
            try
            {
                error = null;

                // create a resource store based on this metadata string
                var rs = ResourceStore.Parse(resourceString, uri);

                // then merge into our master store that encompasses the entire package
                this.resourceStore.Merge(rs);
            }
            catch (MetadataException mex)
            {
                Output(true, mex.Message);

                error = mex.Message;
                return false;
            }
            catch (Exception ex)
            {
                string errorMessage = string.Format("<{0}> - Failed to load template resource string <{1}>", uri, ex);
                Output(true, errorMessage);

                error = errorMessage;
                return false;
            }
            return true;
        }

        /// <summary>
        /// Loads the template file.
        /// </summary>
        /// <param name="fileName">Name of the file.</param>
        /// <param name="data">The data.</param>
        private void LoadDesignerFile(string fileName, byte[] data)
        {
            this.resourceStore.AddDesignerFileData(fileName, data);
        }

        #endregion

        #region Diagnostics

        static bool writeOutput;
        private static void Output(string message)
        {
            Output(writeOutput, message);
        }

        private static void Output(bool write, string message)
        {
            if (!write)
            {
                return;
            }

            message = string.Format("{0} : {1}", DateTime.Now.ToString("HH:mm:ss FFFF"), message);

            Console.WriteLine(message);
            Debug.WriteLine(message);
        }

        #endregion
    }
}
