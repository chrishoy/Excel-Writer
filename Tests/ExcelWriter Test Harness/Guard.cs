namespace ExcelWriter.TestHarness
{
    using System;

    /// <summary>
    /// Basic assertions
    /// </summary>
    internal static class Guard
    {
        /// <summary>
        /// Thorws <see cref="ArgumentNullException"/>if argument null.
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="arg">The argument.</param>
        /// <param name="name">The name.</param>
        /// <exception cref="ArgumentNullException"></exception>
        public static void IsNotNull<T>(T arg, string name) where T : class
        {
            if (arg == null)
            {
                throw new ArgumentNullException(name);
            }
        }
    }
}
