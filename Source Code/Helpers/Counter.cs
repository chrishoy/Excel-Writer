namespace ExcelWriter
{
    /// <summary>
    /// Simple counter used to generate id's for items stored in a store.
    /// </summary>
    internal static class Counter
    {
        #region Private Fields

        private static int id;

        #endregion Private Fields

        /// <summary>
        /// Gets the next id
        /// </summary>
        /// <returns>The next id</returns>
        public static int GetNextId()
        {
            if (id == int.MaxValue)
            {
                id = 1;
            }
            else
            {
                id++;
            }

            // Return id, then increment
            return id;
        }

        /// <summary>
        /// Resets the counter
        /// </summary>
        public static void Reset()
        {
            id = 0;
        }
    }
}
