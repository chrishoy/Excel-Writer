using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Reflection;
using System.Xml;
using System.Windows.Markup;
using System.IO;

namespace Gam.MM.Framework.Export.Map
{
    /// <summary>
    /// Base class for Maps.
    /// Provides a style that can be applied to the map which dictates background colour and borders.
    /// </summary>
    public abstract class BaseMap : IExcelPreparable
    {
        #region Private Fields

        private object dataContext;
        private string cellStyleKey;
        private string cellStyleSelectorKey;
        private string mapId;
        private string templateId;
        private string key;

        #endregion Private Fields

        #region Public Properties

        /// <summary>
        /// Replacement for CellStyle (non-dependency property)
        /// </summary>
        public string CellStyleKey
        {
            get { return this.cellStyleKey; }
            set { this.cellStyleKey = value; }
        }

        /// <summary>
        /// The key to a CellStyleSelector that can be used to choose CellStyleKeys at runtime
        /// </summary>
        public string CellStyleSelectorKey
        {
            get { return this.cellStyleSelectorKey; }
            set { this.cellStyleSelectorKey = value; }
        }

        /// <summary>
        /// Get/set the context of the data for this instance (SimpleBindable)
        /// </summary>
        public object DataContext
        {
            get { return this.dataContext; }
            set { this.dataContext = value; }
        }

        /// <summary>
        /// Uniquely identifies the map - (CH perhaps?)
        /// </summary>
        [Obsolete("Check the use of MapId and similarities with Key")]
        public string MapId
        {
            get { return this.mapId; }
            set { this.mapId = value; }
        }

        /// <summary>
        /// Identifies this element as a potential template, making it re-assignable to different data parts via metadata mapping.<br/>
        /// </summary>
        public string TemplateId
        {
            get { return this.templateId; }
            set { this.templateId = value; }
        }

        /// <summary>
        /// Identifies this element as a potentially re-usable resource which can be looked up and implemented by Key<br/>
        /// </summary>
        public string Key
        {
            get { return this.key; }
            set { this.key = value; }
        }

        #endregion Public Properties

        #region Public Methods

        /// <summary>
        /// Finds the first instance of an element of a specified type derived from <see cref="BaseMap"/> in this <see cref="BaseMap"/><br/>
        /// This includes this instance.
        /// </summary>
        /// <typeparam name="T">The type of <see cref="BaseMap"/> that we wish to find the first instance of</typeparam>
        /// <returns>The first instance of type <typeparamref name="T"/> found in the hierarchy.</returns>
        public abstract T FirstDescendentOfType<T>() where T : BaseMap;

        #endregion Public Methods

        /// <summary>
        /// Evaluates the path specified in a <see cref="BindingExtension"/> to a property of the assigned DataContext of the supplied <see cref="BaseMap"/>.
        /// </summary>
        /// <param name="map"></param>
        /// <param name="propertyPath"></param>
        /// <returns></returns>
        internal static object EvaluatePath(object src, BindingExtension bindingExtension)
        {
            if (src == null || bindingExtension == null || string.IsNullOrEmpty(bindingExtension.Path)) return null;

            try
            {
                return GetPropValue(src, bindingExtension.Path);
            }
            catch
            {
                // Really, should raise some sort of error as the property path is invalid on a null - but this is what WPF bindings do...
                System.Diagnostics.Debug.Print(string.Format("Map.EvaluatePath failed: {0}, Source={1}", bindingExtension.ToString(), src));
                return null;
            }
        }

        /// <summary>
        /// Parses and evaluates a property path, returning a value using reflection.
        /// </summary>
        /// <param name="src"></param>
        /// <param name="propertyPath"></param>
        /// <returns></returns>
        internal static object GetPropValue(object src, string propertyPath)
        {
            if (src == null)
            {
                // Really, should raise some sort of error as the property path is invalid on a null - but this is what WPF bindings do...
                System.Diagnostics.Debug.Print(string.Format("Map.GetPropValue failed: PropertyPath='{0}', Source={1}", propertyPath, src));
                return null;
            }

            // Find property path separator '.'
            int propPathSeparatorPos = propertyPath.IndexOf('.');
            if (propPathSeparatorPos < 0)
            {
                // Isn't a separator so direct evaluation
                return GetKeyIndexedPropertyValue(src, propertyPath);
            }

            // Evaluate up to separator and make recursive evaluation call.
            object propertyValue = GetKeyIndexedPropertyValue(src, propertyPath.Substring(0, propPathSeparatorPos));
            return GetPropValue(propertyValue, propertyPath.Substring(propPathSeparatorPos + 1));
        }

        #region Private Helpers

        /// <summary>
        /// Gets a property value from either a straight property, or from a key indexed property (e.g. Values[SomeKey])
        /// </summary>
        /// <param name="src"></param>
        /// <param name="keyIndexedPropertyName"></param>
        /// <returns></returns>
        private static object GetKeyIndexedPropertyValue(object src, string keyIndexedPropertyName)
        {
            object value = null;

            int keyStart = keyIndexedPropertyName.IndexOf('[');
            if (keyStart >= 0)
            {
                // The requested property value is key-indexed.
                int keyEnd = keyIndexedPropertyName.IndexOf(']', keyStart);
                if (keyEnd <= keyStart) throw new InvalidOperationException(string.Format("Unable to get value from '{0}'", keyIndexedPropertyName));

                string propertyName = keyIndexedPropertyName.Substring(0, keyStart);
                string key = keyIndexedPropertyName.Substring(keyStart + 1, keyEnd - keyStart - 1);

                System.Reflection.PropertyInfo pi = src.GetType().GetProperty(propertyName);
                object collection = pi.GetValue(src, null);

                // Probably a whole load of type checks to be done here, but for the moment, I'm going to 
                // assume (probably incorrectly) that most people use generic dictionaries for collections.

                // Check if IDictionary
                if (collection is System.Collections.IDictionary)
                {
                    // Generic Dictionary ?
                    if (pi.PropertyType.IsGenericType)
                    {
                        value = GetValueFromGenericDictionary(collection, key);
                    }
                    else
                    {
                        // Write out to debug window (similar to Binding)
                        System.Diagnostics.Debug.Print(string.Format("Map.GetPropValue failed: Non-Generic Dictionary Key Indexed PropertyPath='{0}', Source='{1}' not cuurrently supported", keyIndexedPropertyName, src));
                    }
                }
                else if (collection is System.Collections.IList)
                {
                    // If IList, try read by indexer
                    int index;
                    if (int.TryParse(key, out index))
                    {
                        var list = collection as System.Collections.IList;
                        if (list.Count > index)
                        {
                            value = list[index];
                        }
                        else 
                        {
                            // Write out to debug window (similar to Binding)
                            System.Diagnostics.Debug.Print(string.Format("Map.GetPropValue failed: Index out of range PropertyPath='{0}', Source='{1}'", keyIndexedPropertyName, src));
                        }
                    }
                }
                else
                {
                    // Write out to debug window (similar to Binding)
                    System.Diagnostics.Debug.Print(string.Format("Map.GetPropValue failed: Non Generic Collection Key Indexed PropertyPath='{0}', Source='{1}' not cuurrently supported", keyIndexedPropertyName, src));
                }
            }
            else
            {
                // The requested property is non key-indexed. Return the straight property.
                System.Reflection.PropertyInfo propertyInfo = src.GetType().GetProperty(keyIndexedPropertyName);
                if (propertyInfo != null)
                {
                    value = propertyInfo.GetValue(src, null);
                }
                else
                {
                    System.Diagnostics.Debug.Print(string.Format("Map.GetPropValue failed: Unable to find PropertyPath='{0}', Source='{1}'. Please check this exists not cuurrently supported", keyIndexedPropertyName, src));
                }
            }

            // Return the determined value
            return value;
        }

        /// <summary>
        /// Reads a value from a key indexed dictionary.
        /// </summary>
        /// <param name="key">The specified property (read from Path - eg. SomeDictionary[SomeKey] : SomeDictionary is the property)</param>
        /// <param name="collection">Collection from which value is to be read</param>
        /// <returns></returns>
        private static object GetValueFromGenericDictionary(object collection, string key)
        {
            var collectionType = collection.GetType();

            // Generic dictionary in the form Dictionary<key,value> (first generic arg is key type)
            Type[] arguments = collectionType.GetGenericArguments();
            Type keyType = arguments[0];

            // Use TryGetValue as accessing indexer via reflection is a mystery to me...
            MethodInfo tryGetValueMethod = collectionType.GetMethod("TryGetValue");

            // Convert the supplied key to the required key type
            object genericKey = System.Convert.ChangeType(key, keyType);
            object valueToRead = null;

            object[] parameterList = new object[] { genericKey, valueToRead };
            tryGetValueMethod.Invoke(collection, parameterList);

            // Return the read value (valueToRead)
            return parameterList[1];
        }

        #endregion Private Helpers
    }
}
