namespace ExcelWriter
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.IO.Packaging;
    using System.Linq;
    using System.Net;
    using System.Text;
    using System.Xml;
    using System.Windows.Markup;
    using System.Xml.Linq;
    using System.Diagnostics;

    public sealed class DocumentMetadataPackage
    {
        private static object padLock = new object();

        private const string FilesPathPart = "TemplateFile";
        private const string MetadataPathPart = "Metadata";
        private const string UriDelim = @"/";

        public DocumentMetadataPackage()
        { }

        #region Public Properties

        public DocumentMetadataBase DocumentMetadata { get; private set; }

        #endregion Public Properties

        public static string Pack(string directory)
        {
            try
            {
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

                bool haveFiles = false;

                // and check there's a template files directory
                DirectoryInfo filesDirInfo = new DirectoryInfo(string.Concat(directory, System.IO.Path.DirectorySeparatorChar, FilesPathPart));
                if (filesDirInfo.Exists)
                {
                    haveFiles = true;
                    Output(string.Format("TemplateFiles directory found: <{0}>", filesDirInfo.FullName));
                }

                using (var zip = ZipPackage.Open(filePath, System.IO.FileMode.CreateNew, System.IO.FileAccess.ReadWrite))
                {

                    var metadataArray = metadataDirInfo.GetFiles();
                    if (metadataArray.Count() == 0)
                    {
                        throw new MetadataException("No metadata files found");
                    }
                    else if (metadataArray.Count() > 1)
                    {
                        throw new MetadataException("Only only metadata file allowed in a ExportMetadataPackage");
                    }

                    // process the metadata into part
                    var metadataFile = metadataArray[0];

                    string uriString = string.Format("{0}{1}{0}{2}", UriDelim, MetadataPathPart, metadataFile.Name);

                    Uri partUri = new Uri(uriString, UriKind.Relative);
                    PackagePart part = zip.CreatePart(partUri, System.Net.Mime.MediaTypeNames.Application.Zip, CompressionOption.Normal);

                    byte[] data = File.ReadAllBytes(metadataFile.FullName);
                    part.GetStream().Write(data, 0, data.Length);

                    Output(string.Format("Metadata part created: Uri <{0}> Size <{1}>", uriString, data.Length));

                    if (haveFiles)
                    {
                        bool haveTemplateFile = false;
                        var filesArray = filesDirInfo.GetFiles();
                        if (filesArray.Count() > 1)
                        {
                            throw new MetadataException("Only only template file allowed in a ExportMetadataPackage");
                        }
                        else if (filesArray.Count() == 1)
                        {
                            haveTemplateFile = true;
                        }

                        // process the template file if there is one
                        if (haveTemplateFile)
                        {
                            var templateFile = filesArray[0];

                            uriString = string.Format("{0}{1}{0}{2}", UriDelim, FilesPathPart, templateFile.Name);
                            partUri = new Uri(uriString, UriKind.Relative);
                            part = zip.CreatePart(partUri, System.Net.Mime.MediaTypeNames.Application.Zip, CompressionOption.Normal);

                            data = File.ReadAllBytes(templateFile.FullName);
                            part.GetStream().Write(data, 0, data.Length);

                            Output(string.Format("TemplateFile part created: Uri <{0}> Size <{1}>", uriString, data.Length));
                        }
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
        }

        public static DocumentMetadataPackage Open(Uri uri)
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

        public static DocumentMetadataPackage Open(string fileName)
        {
            // avoid contention whilst opening file, should be v quick anyway
            lock (padLock)
            {
                try
                {
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

                    DocumentMetadataPackage package = new DocumentMetadataPackage();

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

                        if (metadataParts.Count == 0)
                        {
                            throw new MetadataException("No metadata files found");
                        }
                        else if (metadataParts.Count > 1)
                        {
                            throw new MetadataException("Only one metadata file allowed in a ExportMetadataPackage");
                        }
                        else
                        {
                            // metadata process
                            // read all the templates files into our internal store
                            var m = metadataParts[0];

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

                            UTF8Encoding encoding = new UTF8Encoding(true);
                            string resourceString = encoding.GetString(hasBomb ? data.Skip(3).ToArray() : data);

                            Output("Stripping Byte Order Mark");

                            package.DocumentMetadata = DocumentMetadataBase.Deserialize(resourceString);

                            Output("Export metadata deserialized");
                        }

                        // template file process
                        if (!string.IsNullOrEmpty(package.DocumentMetadata.TemplateFileName))
                        {
                            Output(string.Format("Checking for existence of the TemplateFilePath <{0}>", package.DocumentMetadata.TemplateFileName));

                            if (fileParts.Count == 0)
                            {
                                throw new MetadataException(string.Format("No file parts found and <{0}> expected", package.DocumentMetadata.TemplateFileName));
                            }
                            else if (fileParts.Count > 1)
                            {
                                throw new MetadataException(string.Format("More than 1 file parts found and only <{0}> expected", package.DocumentMetadata.TemplateFileName));
                            }

                            var f = fileParts[0];

                            string name = f.Uri.OriginalString.Replace(string.Concat(UriDelim, FilesPathPart, UriDelim), null);
                            long length = f.GetStream().Length;

                            package.DocumentMetadata.HasTemplate = true;
                            package.DocumentMetadata.TemplateData = new byte[length];
                            f.GetStream().Read(package.DocumentMetadata.TemplateData, 0, (int)length);

                            Output(string.Format("Loaded TemplateFile data : Uri <{0}> Size {1}", f.Uri.OriginalString, length));
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
            }
        }

        public static void Validate(string fileName)
        {
            var package = Open(fileName);
            package.ValidateInternal();
        }

        #region Privates

        /// <summary>
        /// Check all the template files referenced in the templates metadata have been loaded
        /// </summary>
        private void ValidateInternal()
        {
            Output("Package validation to be implemented");
        }
        
        #endregion

        #region Diagnostics

        private static void Output(string message)
        {
            Debug.WriteLine(string.Format("{0} : {1}", DateTime.Now.ToString("HH:mm:ss FFFF"), message));
        }

        #endregion
    }
}
