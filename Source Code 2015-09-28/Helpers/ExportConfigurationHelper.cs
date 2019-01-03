using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Configuration;

namespace Gam.MM.Framework.Export.Map
{
    public static class ExportConfigurationHelper
    {
        private static Gam.Framework.Configuration.ConfigurationManager manager;
        private static string configurationFilePath;

        public static string ConfigurationFilePath
        {
            get
            {
                Initialize();
                return configurationFilePath;
            }
            set { configurationFilePath = value; }
        }

        /// <summary>
        /// Gets value stored in application settings section 'ExportPackageDirUri'.<br/>
        /// This is the OLD location of export template packages.
        /// </summary>
        [Obsolete("This is the OLD location of export template packages. Use GetExportPackageDirPath instead.")]
        public static string GetExportPackageDirUri
        {
            get
            {
                Initialize();

                AppSettingsSection appSettings = manager.AppSettings;
                if (appSettings != null)
                {
                    return appSettings.Settings["ExportPackageDirUri"].Value;
                }
                return null;
            }
        }

        /// <summary>
        /// Gets value stored in application settings section 'ExportPackageDirPath'.<br/>
        /// This is the NEW location of export template packages.
        /// </summary>
        public static string GetExportPackageDirPath
        {
            get
            {
                Initialize();

                AppSettingsSection appSettings = manager.AppSettings;
                if (appSettings != null)
                {
                    return appSettings.Settings["ExportPackageDirPath"].Value;
                }
                return null;
            }
        }

        public static string GetDocumentsDirPath
        {
            get
            {
                Initialize();

                AppSettingsSection appSettings = manager.AppSettings;
                if (appSettings != null)
                {
                    return appSettings.Settings["DocumentsDirPath"].Value;
                }
                return null;
            }
        }

        /// <summary>
        /// Gets the name of the package within the Export Package Directory, which will
        /// be used to load a libary ox export tempaltes.
        /// </summary>
        public static string GetExportTemplateLibraryPackageName
        {
            get
            {
                Initialize();

                AppSettingsSection appSettings = manager.AppSettings;
                if (appSettings != null)
                {
                    return appSettings.Settings["ExportTemplatesLibraryPackage"].Value;
                }
                return null;
            }
        }

        public static string ResourcesPackageName
        {
            get
            {
                Initialize();

                AppSettingsSection appSettings = manager.AppSettings;
                if (appSettings != null)
                {
                    return appSettings.Settings["ResourcesPackage"].Value;
                }
                return null;
            }
        }

        /// <summary>
        /// General config appSettings value reader.
        /// </summary>
        /// <param name="key"></param>
        /// <returns></returns>
        public static string GetAppSetting(string key)
        {
            Initialize();

            AppSettingsSection appSettings = manager.AppSettings;
            if (appSettings != null)
            {
                return appSettings.Settings[key].Value;
            }
            return null;
        }

        private static void Initialize()
        {
            if (configurationFilePath == null)
            {
                ConfigurationFilePath = System.Reflection.Assembly.GetExecutingAssembly().Location + ".config";
                if (new System.IO.FileInfo(ConfigurationFilePath).Exists == false)
                {
                    throw new InvalidOperationException(string.Format("Unable to find config file <{0}>", ConfigurationFilePath));
                }
            }

            if (manager == null)
            {
                var helper = new Gam.Framework.Configuration.ConfigurationManagerHelper();
                helper.ConfigurationFilePath = configurationFilePath;
                manager = helper.ConfigManager;
            }
        }
    }
}
