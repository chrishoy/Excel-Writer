namespace ExcelWriter
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text;

    /// <summary>
    /// Extension methods for <see cref="ExcelMapCoOrdinate"/> derived classes.
    /// </summary>
    internal static class ExcelMapCoOrdinateExtensions
    {
        /// <summary>
        /// Adds a keyed element to a list of keyed elements in a <see cref="ExcelMapCoOrdinate"/> derived class.
        /// </summary>
        /// <param name="key"></param>
        /// <param name="element"></param>
        internal static void AddKeyedElement(this ExcelMapCoOrdinate map, string key, BaseMap element)
        {
            if (map.KeyedElements == null)
            {
                map.KeyedElements = new List<KeyValuePair<string, BaseMap>>();
            }
            map.KeyedElements.Add(new KeyValuePair<string, BaseMap>(key, element));
        }

        /// <summary>
        /// Finds the first keyed element of a specified <see cref="BaseMap"/> derived type which matches a specified key.<br/>
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="key"></param>
        /// <returns></returns>
        internal static T FirstKeyedElementOfType<T>(this ExcelMapCoOrdinate map, string key) where T : BaseMap
        {
            if (map.KeyedElements != null)
            {
                foreach (KeyValuePair<string, BaseMap> kvp in map.KeyedElements.FindAll(kvp => kvp.Key.CompareTo(key) == 0))
                {
                    if (kvp.Value is T)
                    {
                        return (T)kvp.Value;
                    }
                }
            }
            return null;
        }

        /// <summary>
        /// Finds the first instance of an element of a specified type derived from <see cref="BaseMap"/> in this <see cref="BaseMap"/><br/>
        /// which has a specified key. This includes this instance.
        /// </summary>
        /// <param name="excelMap">The <see cref="ExcelMapCoOrdinate"/> derived entity to search</param>
        /// <param name="key">The key of the typed item that we require</param>
        /// <typeparam name="T">The type of <see cref="BaseMap"/> that we wish to find the first instance of</typeparam>
        /// <returns>The first instance of type <typeparamref name="T"/> found in the hierarchy, or null if none found.</returns>
        internal static T FirstDescendentKeyedElementOfType<T>(this ExcelMapCoOrdinate excelMap, string key) where T : BaseMap
        {
            // Attempts to find the first keyed element within this map
            T item = excelMap.FirstKeyedElementOfType<T>(key);
            if (item != null)
            {
                return item;
            }
            else if (excelMap is ExcelMapCoOrdinateContainer)
            {
                var container = excelMap as ExcelMapCoOrdinateContainer;

                foreach (var mapKeyValue in container.Cells)
                {
                    T cellItem = mapKeyValue.Value.FirstDescendentKeyedElementOfType<T>(key);
                    if (cellItem != null)
                    {
                        return cellItem;
                    }
                }
            }
            return null;
        }

        /// <summary>
        /// Finds the first instance of an element of a specified type derived from <see cref="BaseMap"/> in this <see cref="BaseMap"/><br/>
        /// which has a specified key, includint this instance.<br/>
        /// If none found, goes to parent (Ancestor), and tries again, until we either reach the root or find a descendent with the correct type and key.
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="excelMap"></param>
        /// <param name="key"></param>
        /// <returns></returns>
        internal static T FirstAncendentKeyedElementOfType<T>(this ExcelMapCoOrdinate excelMap, string key) where T : BaseMap
        {
            // Attempts to find the first keyed element within this map, including descendents
            T item = excelMap.FirstDescendentKeyedElementOfType<T>(key);
            if (item != null)
            {
                return item;
            }
            else if (excelMap.Parent != null)
            {
                return excelMap.Parent.FirstAncendentKeyedElementOfType<T>(key);
            }
            return null;
        }
    }
}
