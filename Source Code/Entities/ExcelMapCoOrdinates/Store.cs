namespace ExcelWriter
{
    using System;
    using System.Collections;
    using System.Collections.Generic;

    /// <summary>
    /// Represents a store of <see cref="Store{T}" />s which can be accesses by Id
    /// </summary>
    /// <typeparam name="T">An object which exposes a <see cref="IStorable"/> interface.</typeparam>
    internal class Store<T> : IEnumerable<T> where T : IStorable
    {
        #region Private Fields

        private Dictionary<int, T> dictionary;

        #endregion Private Fields

        #region Construction

        /// <summary>
        /// Initializes a new instance of the <see cref="Store{T}" /> class.
        /// </summary>
        public Store()
        {
            this.dictionary = new Dictionary<int, T>();
        }

        #endregion Construction

        #region Public Methods

        /// <summary>
        /// Indexer, keyed by <see cref="{T}"/>.Id
        /// </summary>
        /// <param name="id">The id of the <see cref="{T}"/> required</param>
        /// <returns>A <see cref="{T}"/></returns>
        public T GetById(int id)
        {
            return this.dictionary[id];
        }

        /// <summary>
        /// Adds a <see cref="{T}"/> to this list. If already there (by Id) raises an exception.
        /// </summary>
        /// <param name="item">The <see cref="{T}"/> to be added</param>
        public void Add(T item)
        {
            if (item == null)
            {
                throw new ArgumentNullException("item");
            }

            this.dictionary.Add(item.Id, item);
        }

        /// <summary>
        /// Adds an item to this list that is not already there (by Id)
        /// </summary>
        /// <param name="item">A <see cref="{T}"/></param>
        public void AddDistinct(T item)
        {
            if (!this.dictionary.ContainsKey(item.Id))
            {
                this.dictionary.Add(item.Id, item);
            }
        }

        /// <summary>
        /// Adds items to this list that are not already there (by Id)
        /// </summary>
        /// <param name="items">List of <see cref="{T}"/></param>
        public void AddDistinct(IEnumerable items)
        {
            foreach (KeyValuePair<int, T> kvp in items)
            {
                T item = kvp.Value;

                if (!this.dictionary.ContainsKey(item.Id))
                {
                    this.dictionary.Add(item.Id, item);
                }
            }
        }

        /// <summary>
        /// Removes an item from this <see cref="Store{T}" />.
        /// </summary>
        /// <param name="item">The <see cref="{T}"/> to be removed</param>
        public void Remove(T item)
        {
            if (this.dictionary.ContainsKey(item.Id))
            {
                this.dictionary.Remove(item.Id);
            }
        }


        #endregion Public Methods

        #region IEnumerable<T>

        /// <summary>
        ///  Returns an enumerator that iterates through a collection.
        /// </summary>
        /// <returns>An System.Collections.IEnumerator object that can be used to iterate through the collection.</returns>
        IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return this.dictionary.GetEnumerator();
        }

        /// <summary>
        ///  Returns an enumerator that iterates through a collection.
        /// </summary>
        /// <returns>An System.Collections.IEnumerator object that can be used to iterate through the collection.</returns>
        IEnumerator<T> IEnumerable<T>.GetEnumerator()
        {
            foreach (T value in this.dictionary.Values)
            {
                yield return value;
            }
        }

        #endregion IEnumerable<T>
    }
}
