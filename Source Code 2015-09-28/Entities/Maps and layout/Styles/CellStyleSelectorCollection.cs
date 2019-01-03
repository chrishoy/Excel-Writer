using System;
using System.IO;
using System.Windows;
using System.Windows.Markup;
using System.Xml;
using System.Collections.Generic;
using System.Collections;
using System.Linq;

namespace ExcelWriter
{
    /// <summary>
    /// Managed collection of <see cref="Style"/> derived objects
    /// </summary>
    public sealed class CellStyleSelectorCollection : IEnumerable<CellStyleSelector>, IList
    {
        #region Private Fields

        private List<object> list;
        private Dictionary<string, CellStyleSelector> dictionary;

        #endregion Private Fields

        #region Construction

        public CellStyleSelectorCollection()
        {
            this.dictionary = new Dictionary<string, CellStyleSelector>();
            this.list = new List<object>();
        }

        #endregion Construction

        #region Public Methods

        /// <summary>
        /// Add a <see cref="CellStyleSelector"/> derived class instance to the collection
        /// </summary>
        /// <param name="value"></param>
        public void Add(CellStyleSelector value)
        {
            this.dictionary.Add(value.Key, value);
            this.list.Add(value);
        }

        /// <summary>
        /// Returns a <see cref="CellStyleSelector"/> derived class instance by key.
        /// </summary>
        /// <param name="key"></param>
        /// <returns>Instance of a <see cref="CellStyleSelector"/> derived class, or null if not found.</returns>
        public CellStyleSelector FindByKey(string key)
        {
            if (key == null) return null;

            if (this.dictionary.ContainsKey(key))
            {
                return this.dictionary[key];
            }
            return null;
        }

        /// <summary>
        /// Creates an instance of a <see cref="CellStyleSelectorCollection"/> from a supplied xaml string
        /// </summary>
        /// <param name="value">A xaml string representation of the <see cref="CellStyleSelectorCollection"/></param>
        /// <returns>An instant of a <see cref="CellStyleSelectorCollection"/></returns>
        public static CellStyleSelectorCollection Deserialize(string value)
        {
            using (var sr = new StringReader(value))
            {
                using (var xr = new XmlTextReader(sr))
                {
                    return (CellStyleSelectorCollection)XamlReader.Load(xr);
                }
            }
        }

        #endregion Public Methods

        #region IEnumerable<CellStyleSelector>, IList members

        public IEnumerator<CellStyleSelector> GetEnumerator()
        {
            IEnumerator ie = dictionary.GetEnumerator();
            while (ie.MoveNext())
            {
                yield return ((KeyValuePair<string, CellStyleSelector>)ie.Current).Value;
            }
        }

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return this.GetEnumerator();
        }

        public int Add(object value)
        {
            this.Add((CellStyleSelector)value);
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
            var newValue = (CellStyleSelector)value;
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
            var oldValue = (CellStyleSelector)value;
            this.list.Remove(oldValue);
            this.dictionary.Remove(oldValue.Key);
        }

        public void RemoveAt(int index)
        {
            object value = this.list[index];
            var oldValue = (CellStyleSelector)value;
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
                var originalValue = (CellStyleSelector)this[index];
                var newValue = (CellStyleSelector)value;
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

        #endregion IEnumerable<ExcelMapStyleBase>, IList members
    }
}
