namespace ExcelWriter
{
    using System;
    using System.Collections;
    using System.Collections.Generic;
    using System.IO;
    using System.Windows.Markup;
    using System.Xml;

    /// <summary>
    /// Managed collection of <see cref="Style"/> derived objects
    /// </summary>
    public sealed class ResourceCollection : IEnumerable<IResource>, IList
    {
        #region Private Fields

        private List<object> list;
        private Dictionary<string, IResource> dictionary;

        #endregion Private Fields

        #region Construction

        public ResourceCollection()
        {
            this.dictionary = new Dictionary<string, IResource>();
            this.list = new List<object>();

            this.MergeResources = new List<ResourceMetadata>();
        }

        #endregion Construction

        public List<ResourceMetadata> MergeResources { get; set; }

        #region Public Methods

        /// <summary>
        /// Add a <see cref="StyleBase"/> derived class instance to the collection
        /// </summary>
        /// <param name="value"></param>
        public void Add(IResource value)
        {
            this.dictionary.Add(value.Key, value);
            this.list.Add(value);
        }

        /// <summary>
        /// Returns a <see cref="StyleBase"/> derived class instance by key.
        /// </summary>
        /// <param name="key"></param>
        /// <returns>Instance of a a <see cref="StyleBase"/> derived class, or null if not found.</returns>
        public IResource FindByKey(string key)
        {
            if (key == null) return null;

            if (this.dictionary.ContainsKey(key))
            {
                return this.dictionary[key];
            }
            return null;
        }

        /// <summary>
        /// Updates this collection with the supplied style, applying taking into consideration the BasedOn property of the supplied value.
        /// </summary>
        /// <param name="value"></param>
        internal void MergeCloneStyle(StyleBase value)
        {
            if (!string.IsNullOrEmpty(value.BasedOnKey))
            {
                // Look up mandatory existing style
                StyleBase baseStyle = this.dictionary[value.BasedOnKey] as StyleBase;
                if (baseStyle == null) 
                {
                    throw new MetadataException(string.Format("Supplied BasedOnKey is not a StyleBase <{0}>", value.BasedOnKey));
                }

                StyleBase newStyle;

                // And update with properties of the supplied style
                if (value is Style)
                {
                    if (baseStyle is Style)
                    {
                        // ExcelMapStyle based on ExcelMapStyle
                        newStyle = Style.CreateMergedStyle(baseStyle, (Style)value);
                        this.Add(newStyle);
                    }
                    else if (baseStyle is CellStyle)
                    {
                        // ExcelMapStyle can't be based on a ExcelCellMapStyle
                        throw new InvalidOperationException(string.Format("Can't base ExcelMapStyle '{0}'  on ExcelCellMapStyle '{1}' (not yet anyway)", value.Key, baseStyle.Key));
                    }
                    else
                    {
                         // Whoops.. not an accounted for style
                        throw new InvalidOperationException(string.Format("Style Key='{0}' BasedOnStyle is not valid style", baseStyle.Key));
                    }

                }
                else if (value is CellStyle)
                {
                    if (baseStyle is Style)
                    {
                        // ExcelCellMapStyle based on ExcelMapStyle
                        newStyle = CellStyle.CreateMergedStyle(baseStyle, (CellStyle)value);
                    }
                    else if (baseStyle is CellStyle)
                    {
                        // Straight Copy & Merge
                        newStyle = CellStyle.CreateMergedStyle(baseStyle, (CellStyle)value);
                        this.Add(newStyle);
                    }
                    else
                    {
                         // Whoops.. not an accounted for style
                        throw new InvalidOperationException(string.Format("Style Key='{0}' BasedOnStyle is not valid style", baseStyle.Key));
                    }
                }
                else
                {
                    // Whoops.. not an accounted for style
                    throw new InvalidOperationException(string.Format("Style Key='{0}' is not valid style", value.Key));
                }
            }
            else
            {
                // Clone and update (Potentially override at some later date)
                this.Add(value.Clone());
            }
        }

        /// <summary>
        /// Creates an instance of a <see cref="StylesCollection"/> from a supplied xaml string
        /// </summary>
        /// <param name="value">A xaml string representation of the <see cref="StylesCollection"/></param>
        /// <returns>An instant of a <see cref="StylesCollection"/></returns>
        public static ResourceCollection Deserialize(string value)
        {
            using (var sr = new StringReader(value))
            {
                using (var xr = new XmlTextReader(sr))
                {
                    return (ResourceCollection)XamlReader.Load(xr);
                }
            }
        }

        #endregion Public Methods

        #region IEnumerable<IResource>, IList members

        public IEnumerator<IResource> GetEnumerator()
        {
            IEnumerator ie = dictionary.GetEnumerator();
            while (ie.MoveNext())
            {
                yield return ((KeyValuePair<string, IResource>)ie.Current).Value;
            } 
        }

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return this.GetEnumerator();
        }

        public int Add(object value)
        {
            this.Add((IResource)value);
            return this.list.Count - 1;
        }

        public void Clear()
        {
            this.dictionary.Clear();
            this.list.Clear();
        }

        public bool Contains(object value)
        {
            return this.list.Contains(value);
        }

        public int IndexOf(object value)
        {
            return this.list.IndexOf(value);
        }

        public void Insert(int index, object value)
        {
            var newValue = (IResource)value;
            if (this.dictionary.ContainsKey(newValue.Key))
            {
                throw new System.ArgumentException("An element with the same key already exists.");
            }
            this.list.Insert(index, newValue);
            this.dictionary.Add(newValue.Key, newValue);
        }

        public bool IsFixedSize
        {
            get { return false; }
        }

        public bool IsReadOnly
        {
            get { return false; }
        }

        public void Remove(object value)
        {
            var oldValue = (IResource)value;
            this.list.Remove(oldValue);
            this.dictionary.Remove(oldValue.Key);
        }

        public void RemoveAt(int index)
        {
            object value = this.list[index];
            var oldValue = (IResource)value;
            this.list.Remove(oldValue);
            this.dictionary.Remove(oldValue.Key);
        }

        public object this[int index]
        {
            get
            {
                return this.list[index];
            }
            set
            {
                var originalValue = (IResource)this[index];
                var newValue = (IResource)value;
                this.dictionary.Remove(originalValue.Key);
                this.list[index] = value;
                this.dictionary.Add(newValue.Key, newValue);
            }
        }

        public void CopyTo(Array array, int index)
        {
            throw new NotImplementedException();
        }

        public int Count
        {
            get { return this.list.Count; }
        }

        public bool IsSynchronized
        {
            get { throw new NotImplementedException(); }
        }

        public object SyncRoot
        {
            get { throw new NotImplementedException(); }
        }

        #endregion IEnumerable<IResource>, IList members
    }
}
