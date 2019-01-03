namespace ExcelWriter
{
    using System;
    using System.Reflection;
    using System.Windows.Data;
    using System.Xml;
    using System.Xml.Linq;
    using System.Xml.XPath;

    /// <summary>
    /// Container which holds a binding extension and its associated evaluated result.
    /// </summary>
    internal class BindingContainer
    {
        #region Private Fields

        private BindingExtension binding;
        private object evaluatedResult;

        #endregion Private Fields

        #region Construction

        /// <summary>
        /// Initializes a new instance of the <see cref="BindingContainer" /> class.<br/>
        /// No public ctor. Can only create from a binding using Create.
        /// </summary>
        /// <param name="binding">The <see cref="BindingExtension"/> which this container will used to evaluate and store values.</param>
        private BindingContainer(BindingExtension binding)
        {
            if (binding == null)
            {
                throw new ArgumentNullException("binding");
            }

            this.binding = binding;
            this.evaluatedResult = new UnevaluatedBindingResult();
        }

        #endregion Construction

        #region Public Methods

        /// <summary>
        /// If the supplied value is a <see cref="BindingExtension"/>, creates a <see cref="BindingContainer"/> to hold evaluations,
        /// otherwise simply returns the value.
        /// </summary>
        /// <param name="value">The value to be tested</param>
        /// <returns>A <see cref="BindingContainer"/> to wrap the <see cref="BindingExtension"/> or the value itself</returns>
        public static object CreateIfRequired(object value)
        {
            var extension = value as BindingExtension;
            if (extension != null)
            {
                return new BindingContainer(extension);
            }

            return value;
        }

        /// <summary>
        /// If the supplied value is a <see cref="BindingContainer"/>, evaluates the binding and returns the value,
        /// otherwise simply returns the value.
        /// </summary>
        /// <param name="value">The value or <see cref="BindingContainer"/> to be evaluated</param>
        /// <param name="dataContext">The DataContext used for binding evaluations</param>
        /// <returns>The evaluated value.</returns>
        public static object EvaluateIfRequired(object value, object dataContext)
        {
            if (value is BindingContainer)
            {
                // If BindingContainer result not evaluated, evaluate and set.
                var bindingContainer = (BindingContainer)value;
                if (bindingContainer.evaluatedResult is UnevaluatedBindingResult)
                {
                    object result = Evaluate(dataContext, bindingContainer.binding);

                    // Apply StringFormat if set on binding.
                    if (result is string && !string.IsNullOrEmpty(bindingContainer.binding.StringFormat))
                    {
                        try
                        {
                            result = string.Format(bindingContainer.binding.StringFormat, result);
                        }
                        catch
                        {
                            // Really, should raise some sort of error binding is invalid - but this is what WPF bindings do...
                            System.Diagnostics.Debug.Print(
                                "BindingContainer.Evaluate failed: {0}, StringFormat='{1}', Source={2}", 
                                    bindingContainer.binding, bindingContainer.binding.StringFormat, value);
                        }
                    }

                    bindingContainer.evaluatedResult = result;
                }

                return bindingContainer.evaluatedResult;
            }

            return value;
        }

        /// <summary>
        /// If the supplied value is a <see cref="BindingContainer"/>, returns the source binding,
        /// otherwise returns the supplied value.
        /// </summary>
        /// <param name="value">The value to be tested</param>
        /// <returns>Either the Value itself, or the binding behind the container.</returns>
        public static object GetSourceBindingOrValue(object value)
        {
            if (value is BindingContainer)
            {
                var bindingContainer = (BindingContainer)value;
                return bindingContainer.binding;
            }

            return value;
        }

        /// <summary>
        /// Takes an object (which may be the result of a <see cref="BindingContainer"/> evaluation), returns a string value.
        /// </summary>
        /// <param name="value">The value to be converted</param>
        /// <returns>A string representation of the value</returns>
        public static string ConvertToString(object value)
        {
            return value == null ? null : (string)value;
        }

        /// <summary>
        /// Takes an object (which may be the result of a <see cref="BindingContainer"/> evaluation), returns a nullable boolean value.
        /// </summary>
        /// <param name="value">The value to be converted</param>
        /// <returns>A nullable boolean representation of the value</returns>
        public static bool? ConvertToNullableBoolean(object value)
        {
            if (value is string)
            {
                return Boolean.Parse(value.ToString().ToLowerInvariant());
            }
            return value == null ? null : (bool?)value;
        }

        /// <summary>
        /// Takes an object (which may be the result of a <see cref="BindingContainer"/> evaluation), returns a nullable boolean value.
        /// </summary>
        /// <param name="value">The value to be converted</param>
        /// <returns>A nullable boolean representation of the value</returns>
        public static double? ConvertToNullableDouble(object value)
        {
            if (value == null)
            {
                return null;
            }

            if (value is double? || value is double)
            {
                return (double?)value;
            }
            else
            {
                double result;
                if (double.TryParse(value.ToString(), out result))
                {
                    return result;
                }
                return null;
            }
        }

        /// <summary>
        /// Takes an object (which may be the result of a <see cref="BindingContainer"/> evaluation),
        /// returns a nullable <see cref="System.Windows.Media.Color"/>.
        /// </summary>
        /// <param name="value">The value to be converted</param>
        /// <returns>A nullable colour representation of the value</returns>
        public static System.Windows.Media.Color? ConvertToNullableColour(object value)
        {
            if (value is System.Windows.Media.Color?)
            {
                return (System.Windows.Media.Color?)value;
            }
            else if (value is System.Windows.Media.Color)
            {
                return new System.Windows.Media.Color?((System.Windows.Media.Color)value);
            }
            if (value is System.Drawing.Color)
            {
                var colourValue = (System.Drawing.Color)value;
                return new System.Windows.Media.Color?(System.Windows.Media.Color.FromArgb(colourValue.A, colourValue.R, colourValue.G, colourValue.B));
            }
            else if (value is System.Drawing.Color?)
            {
                var colourValue = (System.Drawing.Color?)value;
                if (colourValue == null)
                {
                    return null;
                }
                else
                {
                    return new System.Windows.Media.Color?(System.Windows.Media.Color.FromArgb(colourValue.Value.A,
                                                                                               colourValue.Value.R,
                                                                                               colourValue.Value.G,
                                                                                               colourValue.Value.B));
                }
            }
            else if (value is string)
            {
                var colourValue = (string)value;
                try
                {
                    var colour = (System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString(colourValue);
                    return colour;
                }
                catch (Exception ex)
                {
                    // Write out to debug window... I suppose this is a binding/parse error of some description
                    System.Diagnostics.Debug.Print("Can not convert '{0}' to a 'System.Windows.Media.Color'. Error={1}", colourValue, ex.Message);
                    return null;
                }
            }
            return null;
        }

        #endregion Public Methods

        /// <summary>
        /// Parses and evaluates a property path, returning a value using reflection.
        /// </summary>
        /// <param name="src">The source onbject</param>
        /// <param name="propertyPath">The path to the property to be evaluated</param>
        /// <returns>The evaluated property value</returns>
        public static object GetPropValue(object src, string propertyPath)
        {
            if (src == null)
            {
                // Really, should raise some sort of error as the property path is invalid on a null - but this is what WPF bindings do...
                System.Diagnostics.Debug.Print("Map.GetPropValue failed: PropertyPath='{0}'", propertyPath);
                return null;
            }

            if (string.IsNullOrEmpty(propertyPath))
            {
                // If we do not specify a property path, then return the source.
                return src;
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

        /// <summary>
        /// Gets a property value be evaluating an XPath expression on a source XML object.
        /// Supported source objects are XmlDataProvider, XmlDocument, XDocument.
        /// </summary>
        /// <param name="src">The source element</param>
        /// <param name="xpath">The XPath expression into the source element.</param>
        /// <returns>The evaluated value</returns>
        public static object GetPropValueFromXPath(object src, string xpath)
        {
            // if the source is null then return null
            if (src == null)
            {
                // could raise some sort of error as the property path is invalid on a null - but this is what WPF bindings do...
                System.Diagnostics.Debug.Print("Map.GetPropValue failed: XPath='{0}'", xpath);
                return null;
            }

            try
            {
                // try to load the source in a XDocument
                XDocument doc;                

                if (src is XmlDataProvider)
                {
                    using (var nodeReader = new XmlNodeReader(((XmlDataProvider)src).Document))
                    {
                        nodeReader.MoveToContent();
                        doc = XDocument.Load(nodeReader);
                    }
                }
                else if (src is XmlDocument)
                {
                    using (var nodeReader = new XmlNodeReader((XmlDocument)src))
                    {
                        nodeReader.MoveToContent();
                        doc = XDocument.Load(nodeReader);
                    }
                }
                else if (src is XDocument)
                {
                    doc = (XDocument)src;
                }
                else
                {
                    System.Diagnostics.Debug.Print("XPath only supported for src of XmlDataProvider, XmlDocument or XDocument: XPath='{0}', Source={1}", xpath, src);
                    return null;
                }

                var namespaceManager = new XmlNamespaceManager(new NameTable());
                var value = doc.XPathSelectElement(xpath, namespaceManager).Value;

                // will return a date time instance is it is one
                return TryForDate(value);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.Print("BindingContainer.GetPropValueFromXPath failed: XPath={0}, Source={1}, Error={2}", xpath, src, ex.Message);
                return null;
            }
        }

        #region Private Helpers

        /// <summary>
        /// Evaluates the path or xpath specified in a <see cref="BindingExtension"/> to a property of the assigned DataContext of the supplied <see cref="BaseMap"/>.
        /// </summary>
        /// <param name="src">The supplied source object, usually the DataContext of a <see cref="BaseMap"/> derived element.</param>
        /// <param name="bindingExtension">The binding on which the source object</param>
        /// <returns>The evaluated value</returns>
        private static object Evaluate(object src, BindingExtension bindingExtension)
        {
            if (src == null || bindingExtension == null)
            {
                return null;
            }

            try
            {
                if (!string.IsNullOrEmpty(bindingExtension.XPath))
                {
                    return GetPropValueFromXPath(src, bindingExtension.XPath);
                }
                else
                {
                    return GetPropValue(src, bindingExtension.Path);
                }
            }
            catch
            {
                // Really, should raise some sort of error as the property path is invalid on a null - but this is what WPF bindings do...
                System.Diagnostics.Debug.Print("Map.EvaluatePath failed: {0}, Source={1}", bindingExtension, src);
                return null;
            }
        }

        /// <summary>
        /// Gets a property value from either a straight property, or from a key indexed property (e.g. Values[SomeKey])
        /// </summary>
        /// <param name="src">The source object</param>
        /// <param name="keyIndexedPropertyName">The property, or key indexed property.</param>
        /// <returns>The evaluated value</returns>
        private static object GetKeyIndexedPropertyValue(object src, string keyIndexedPropertyName)
        {
            const string Proc = "Map.BindingContainer.GetKeyIndexedPropertyValue";
            object value = null;

            int keyStart = keyIndexedPropertyName.IndexOf('[');
            if (keyStart >= 0)
            {
                // The requested property value is key-indexed.
                int keyEnd = keyIndexedPropertyName.IndexOf(']', keyStart);
                if (keyEnd <= keyStart)
                {
                    throw new InvalidOperationException(string.Format("Unable to get value from '{0}'", keyIndexedPropertyName));
                }

                string propertyName = keyIndexedPropertyName.Substring(0, keyStart);
                string key = keyIndexedPropertyName.Substring(keyStart + 1, keyEnd - keyStart - 1);

                PropertyInfo pi = src.GetType().GetProperty(propertyName);
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
                        // At the moment, we can't determine the type of the key of a non-generic dictionary
                        // so we are limited to using string keys on non-generic dictionary. Unfortunatetly, an incorrectly
                        // typed lookup on a non-generic dictionary seems to give us null, ie. no error, so we can't trap it...!!!
                        value = ((System.Collections.IDictionary)collection)[key];
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
                            System.Diagnostics.Debug.Print("{0} failed: Index out of range PropertyPath='{1}', Source='{2}'", Proc, keyIndexedPropertyName, src);
                        }
                    }
                }
                else
                {
                    // Write out to debug window (similar to Binding)
                    System.Diagnostics.Debug.Print("{0} failed: Non Generic Collection Key Indexed PropertyPath='{1}', Source='{2}' not cuurrently supported", Proc, keyIndexedPropertyName, src);
                }
            }
            else
            {
                // The requested property is non key-indexed. Return the straight property.
                System.Reflection.PropertyInfo propertyInfo = src.GetType().GetProperty(keyIndexedPropertyName);
                if (propertyInfo == null)
                {
                    System.Diagnostics.Debug.Write(string.Format("{0} failed: Could not find property '{1}', Source={2}", Proc, keyIndexedPropertyName, src));
                }
                else
                {
                    value = propertyInfo.GetValue(src, null);
                }
            }

            // Return the determined value
            return value;
        }

        /// <summary>
        /// Reads a value from a key indexed dictionary.
        /// </summary>
        /// <param name="collection">Collection from which value is to be read</param>
        /// <param name="key">The specified property (read from Path - eg. SomeDictionary[SomeKey] : SomeDictionary is the property)</param>
        /// <returns>The evaluated value</returns>
        private static object GetValueFromGenericDictionary(object collection, string key)
        {
            var collectionType = collection.GetType();

            // Generic dictionary in the form Dictionary<key,value> (first generic arg is key type)
            Type[] arguments = collectionType.GetGenericArguments();
            Type keyType = arguments[0];

            // Use TryGetValue as accessing indexer via reflection is a mystery to me...
            MethodInfo tryGetValueMethod = collectionType.GetMethod("TryGetValue");

            // Convert the supplied key to the required key type
            object genericKey = Convert.ChangeType(key, keyType);
            object valueToRead = null;

            object[] parameterList = new object[] { genericKey, valueToRead };
            tryGetValueMethod.Invoke(collection, parameterList);

            // Return the read value (valueToRead)
            return parameterList[1];
        }

        /// <summary>
        /// Attempt to parse the content to a DateTime and return the DateTime object if it is one.
        /// </summary>
        private static object TryForDate(string content)
        {
            DateTime result;
            if (content != null && DateTime.TryParse(content, out result))
            {
                return result;
            }
            return content;
        }

        #endregion Private Helpers
    }
}
