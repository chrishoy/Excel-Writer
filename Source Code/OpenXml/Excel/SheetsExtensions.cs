namespace ExcelWriter.OpenXml.Excel
{
    using System.Linq;
    using DocumentFormat.OpenXml;
    using DocumentFormat.OpenXml.Spreadsheet;

    /// <summary>
    /// Extension methods for <see cref="Sheets"/>
    /// </summary>
    public static class SheetsExtensions
    {
        /// <summary>
        /// Gets the next sheet identifier.
        /// </summary>
        /// <param name="sheets">The sheets.</param>
        /// <returns></returns>
        public static UInt32Value GetNextSheetId(this Sheets sheets)
        {
            uint next = (uint)sheets.ChildElements.Count + 1;

            bool keepTrying = true;
            while (keepTrying)
            {
                bool match = (from s in sheets.Descendants<DocumentFormat.OpenXml.Spreadsheet.Sheet>()
                              where s.SheetId.HasValue
                              && s.SheetId.Value.CompareTo(next) == 0
                              select s).Any();
                if (!match)
                {
                    keepTrying = false;
                    break;
                }
                next++;
            }
            return next;        
        }

        /// <summary>
        /// Determines whether [has sheet with name] [the specified sheet name].
        /// </summary>
        /// <param name="sheets">The sheets.</param>
        /// <param name="sheetName">Name of the sheet.</param>
        /// <returns></returns>
        public static bool HasSheetWithName(this DocumentFormat.OpenXml.Spreadsheet.Sheets sheets, string sheetName) 
        {
            return (from s in sheets.Descendants<DocumentFormat.OpenXml.Spreadsheet.Sheet>()
                    where s.Name != null && s.Name.HasValue && s.Name.Value == sheetName
                    select s).Any();
        }

        /// <summary>
        /// Gets the name of the safe sheet.
        /// </summary>
        /// <param name="sheets">The sheets.</param>
        /// <param name="proposedName">Name of the proposed.</param>
        /// <returns></returns>
        public static string GetSafeSheetName(this Sheets sheets, string proposedName) 
        {
            if (sheets.HasSheetWithName(proposedName))
            {
                int number = 0;
                string stem = GetNameStemAndNumericTrailer(proposedName, out number);
                number++;

                string trailer = string.Concat("~", number.ToString());

                // Work out max permitted stem length
                if ((stem.Length + trailer.Length) > Constants.SheetNameMaxLength)
                {
                    stem = stem.Substring(0, Constants.SheetNameMaxLength - trailer.Length);
                }
                string proposedNewName = string.Concat(stem, trailer);

                if (sheets.HasSheetWithName(proposedNewName))
                {
                    // Recurse until we get a safe name
                    return GetSafeSheetName(sheets, proposedNewName);
                }
                return proposedNewName;
            }
            else
            {
                return proposedName;
            }
        }

        /// <summary>
        /// We are adopting the ~ notation to indicate that this is a duplicated name.
        /// (ie. 'Name 1', 'Name 1~1', 'Name 1~2' etc)
        /// </summary>
        /// <param name="name">The supplied name (eg. 'Name 1~2'</param>
        /// <param name="number">Returns the number after the last ~</param>
        /// <returns>
        /// The stem (before the last ~)
        /// </returns>
        private static string GetNameStemAndNumericTrailer(string name, out int number)
        {
            int finalValue = 0;

            // Find the last '~'
            int idx = name.LastIndexOf('~');
            if (idx > 0 && idx < name.Length)
            {
                int.TryParse(name.Substring(idx + 1), out finalValue);
                number = finalValue;
                return name.Substring(0, idx);
            }
            number = finalValue;
            return name;
        }
    }
}
