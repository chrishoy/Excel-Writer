namespace ExcelWriter
{
    using System.Collections.Generic;
    using System.Linq;
    using System.Windows.Markup;

    /// <summary>
    /// Selects a cell style based on a specified property value of a row.
    /// </summary>
    [ContentProperty("ValueStyleKeys")]
    public class PropertyValueCellStyleSelector : CellStyleSelector
    {
        #region Private Fields

        private const string proc = "PropertyValueCellStyleSelector";
        private List<PropertyValueStyleKey> valueStyleKeys;

        #endregion Private Fields

        #region Construction

        /// <summary>
        /// 
        /// </summary>
        public PropertyValueCellStyleSelector()
        {
            this.valueStyleKeys = new List<PropertyValueStyleKey>();
        }

        #endregion Construction

        /// <summary>
        /// Gets or sets a list of style keys that are to be used to matched property value.
        /// </summary>
        public List<PropertyValueStyleKey> ValueStyleKeys
        {
            get { return this.valueStyleKeys; }
        }

        /// <summary>
        /// Gets or sets the name of the property which should be monitored.
        /// </summary>
        public string PropertyName { get; set; }

        /// <summary>
        /// Key of style to be selected when HiererchyLevel is out of range or can not be ascertained
        /// </summary>
        public string DefaultStyleKey { get; set; }

        /// <summary>
        /// Select the key of the style to be used when the item matches a value.
        /// </summary>
        /// <param name="item">The supplied item</param>
        /// <returns>The key of the selected style</returns>
        public override string SelectCellStyleKey(object item)
        {
            if (item != null && !string.IsNullOrEmpty(PropertyName))
            {
                object value = ReadPropertyValue(item, PropertyName);
                if (value != null)
                {
                    // Match the value, return the key
                    var match = this.valueStyleKeys.FirstOrDefault(v => v.Value.ToString() == value.ToString());
                    if (match != null) return match.StyleKey;
                }
            }
            return this.DefaultStyleKey;
        }

        /// <summary>
        /// Reads a property value, or writes to Debug window.
        /// </summary>
        /// <param name="src"></param>
        /// <param name="propertyName"></param>
        /// <returns></returns>
        private static object ReadPropertyValue(object src, string propertyName)
        {

            // The requested property is non key-indexed. Return the straight property.
            System.Reflection.PropertyInfo propertyInfo = src.GetType().GetProperty(propertyName);
            if (propertyInfo == null)
            {
                System.Diagnostics.Debug.Print("{0} failed: Could not find property '{1}', Source={2}", proc, propertyName, src);
                return null;
            }

            return propertyInfo.GetValue(src, null);
        }
    }
}
