namespace ExcelWriter
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text;

    /// <summary>
    /// Represents a list of options for representing data when exporting to charts
    /// </summary>
    public class ChartOptions : List<ChartOptionBase>
    {
        /// <summary>
        /// Gets a specied option in this list of <see cref="ChartOptionBase"/>s.<br/>
        /// Returns null if none found in collection
        /// </summary>
        /// <typeparam name="T">Type of the option that is required</typeparam>
        /// <returns>Option or null</returns>
        public T GetOptionOrDefault<T>() where T: ChartOptionBase
        {
            foreach (ChartOptionBase option in this)
            {
                if (option is T)
                {
                    return option as T;
                }
            }
            return default(T);
        }

        /// <summary>
        /// Adds or updates an option to the list of options.<br/>
        /// If the option already exists, then it is updated with the supplied version.
        /// </summary>
        /// <typeparam name="T">The type of option to update</typeparam>
        /// <param name="option">The option to be added/updated with</param>
        public void UpsertOption<T>(T option) where T: ChartOptionBase
        {
            T existingOption = this.GetOptionOrDefault<T>();
            if (existingOption != null)
            {
                int index = this.IndexOf(existingOption);
                this.RemoveAt(index);
                this.Insert(index, option);
            }
            else
            {
                this.Add(option);
            }
        }

        /// <summary>
        /// Removes an option from the list of options.<br/>
        /// </summary>
        /// <typeparam name="T">The type of option to be removed</typeparam>
        /// <param name="option">The option to be removed</param>
        public void RemoveOption<T>() where T : ChartOptionBase
        {
            T existingOption = this.GetOptionOrDefault<T>();
            if (existingOption != null)
            {
                int index = this.IndexOf(existingOption);
                this.RemoveAt(index);
            }
        }
    }
}
