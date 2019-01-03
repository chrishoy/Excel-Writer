using System.Windows;
using System.Collections.Generic;
using System.Linq;

namespace ExcelWriter
{
    /// <summary>
    /// Represents a collection of <see cref="Map"/> derived entities, each of which can be used for defining the structure of export data.
    /// </summary>
    public sealed class MapCollection : List<BaseMap>
    {
        /// <summary>
        /// Finds a map in this collection with a matching key. If more than one exists, returns null
        /// </summary>
        /// <param name="key"></param>
        /// <returns></returns>
        internal BaseMap FindByKey(string key)
        {
            return this.SingleOrDefault(x => key.Equals(x.Key));
        }
    }
}
