namespace ExcelWriter.OpenXml.Excel
{
    using System.Linq;
    using DocumentFormat.OpenXml.Spreadsheet;
    using DocumentFormat.OpenXml.Packaging;
    using System.Text.RegularExpressions;
    using DocumentFormat.OpenXml;

    /// <summary>
    /// 
    /// </summary>
    public static class WorkbookExtensions
    {
        #region Public Static Methods


        /// <summary>
        /// Adds the name of the defined.
        /// </summary>
        /// <param name="workbook">The workbook.</param>
        /// <param name="sheetName">Name of the sheet.</param>
        /// <param name="name">The name.</param>
        /// <param name="columnIndex">Index of the column.</param>
        /// <param name="rowIndex">Index of the row.</param>
        /// <param name="count">The count.</param>
        public static void AddDefinedName(
            this Workbook workbook,
            string sheetName,
            string name,
            uint columnIndex,
            uint rowIndex,
            int count)
        {
            AddDefinedName(workbook, sheetName, name, columnIndex, rowIndex, 1, count);
        }

        /// <summary>
        /// Creates a new named range with a name in the format [SHEET_NAME]_[NAME].
        /// Any non-alphanumeric characters are replaced with an underscore.
        /// If an existing named range exists with the same name then an underscore
        /// followed by a number is added.
        /// </summary>
        /// <param name="workbook">The workbook.</param>
        /// <param name="sheetName">Name of the sheet.</param>
        /// <param name="name">The name of the named range.</param>
        /// <param name="columnIndex">Index of the column.</param>
        /// <param name="rowIndex">Index of the row.</param>
        /// <param name="columnCount">The number of columns in the named range.</param>
        /// <param name="rowCount">The number of rows in the named range.</param>
        public static void AddDefinedName(
            this Workbook workbook,
            string sheetName,
            string name,
            uint columnIndex,
            uint rowIndex,
            int columnCount,
            int rowCount)
        {
            if (workbook.DefinedNames == null)
            {
                workbook.DefinedNames = new DefinedNames();
            }

            string formattedName = FormatName(workbook.DefinedNames, name);
            string start = CellExtensions.GetCellFormula(columnIndex, rowIndex);
            string finish = CellExtensions.GetCellFormula(
                (uint)(columnIndex + columnCount - 1), 
                (uint)(rowIndex + rowCount - 1));

            workbook.DefinedNames.Append(
                new DefinedName()
                {
                    Name = formattedName,
                    Text = string.Format("\'{0}\'!{1}:{2}", sheetName, start, finish)
                });
        }

        /// <summary>
        /// Returns a worksheet part from the workbookpart for the supplied sheet name
        /// </summary>
        /// <param name="workbook">The workbook part</param>
        /// <param name="sheetName">The name of the sheet to be returned</param>
        /// <returns>
        /// The worksheet part matching the supplied name
        /// </returns>
        public static WorksheetPart GetWorksheetPartByName(this Workbook workbook, string sheetName)
        {
            // get the sheet with the matching name
            var match = workbook.GetSheetByName(sheetName);

            // if there is one then return the part with that id
            if (match != null)
            {
                return workbook.WorkbookPart.GetPartById(match.Id) as WorksheetPart;
            }
            // otherwise return null
            return null;
        }

        /// <summary>
        /// Returns a worksheet part from the workbookpart for the supplied sheet name
        /// </summary>
        /// <param name="workbook">The workbook part</param>
        /// <param name="sheetName">The name of the sheet to be returned</param>
        /// <returns>
        /// The worksheet part matching the supplied name
        /// </returns>
        public static EnumValue<SheetStateValues> GetWorksheetStateName(this Workbook workbook, string sheetName)
        {
            // get the sheet with the matching name
            var match = workbook.GetSheetByName(sheetName);

            // if there is one then return the part with that id
            if (match != null)
            {
                return match.State;
            }
            // otherwise return null
            return new EnumValue<SheetStateValues>();
        }


        /// <summary>
        /// Gets the name of the sheet by.
        /// </summary>
        /// <param name="workbook">The workbook.</param>
        /// <param name="sheetName">Name of the sheet.</param>
        /// <returns></returns>
        public static DocumentFormat.OpenXml.Spreadsheet.Sheet GetSheetByName(this Workbook workbook, string sheetName)
        {
            return (from sheet in workbook.Descendants<DocumentFormat.OpenXml.Spreadsheet.Sheet>()
                    where sheet.Name.HasValue
                    && sheet.Name.Value.CompareTo(sheetName) == 0
                    select sheet).FirstOrDefault();
        }

        /// <summary>
        /// used the defined name within the workbook to hide rows of the DefinedName (Range of cells)
        /// </summary>
        /// <param name="workbook">The workbook.</param>
        /// <param name="dn">The dn.</param>
        public static void HideRowsOfDefinedName(this Workbook workbook, DefinedName dn)
        {
            string SheetName = "";
            uint RowStart = 0;
            uint RowEnd = 0;
            
            if (BreakDownDefinedName(workbook, dn, ref SheetName, ref RowStart, ref RowEnd))
            {
                WorksheetPart ws = workbook.GetWorksheetPartByName(SheetName);
                if (ws != null)
                {
                    ws.Worksheet.HideRows(RowStart, RowEnd);
                }
            }
        }

        /// <summary>
        /// used the defined name within the workbook to show rows of the DefinedName (Range of cells)
        /// </summary>
        /// <param name="workbook">The workbook.</param>
        /// <param name="dn">The dn.</param>
        public static void ShowRowsOfDefinedName(this Workbook workbook, DefinedName dn)
        {
            string SheetName = "";
            uint RowStart = 0;
            uint RowEnd = 0;

            if (BreakDownDefinedName(workbook, dn, ref SheetName, ref RowStart, ref RowEnd))
            {
                WorksheetPart ws = workbook.GetWorksheetPartByName(SheetName);
                if (ws != null)
                {
                    ws.Worksheet.ShowRows(RowStart, RowEnd);
                }
            }
        }

        /// <summary>
        /// Gets the open XML worksheet object using the sheet name and calls the HideRows extended method on that sheet.
        /// </summary>
        /// <param name="workbook">The workbook.</param>
        /// <param name="SheetName">Name of the sheet.</param>
        /// <param name="RowStart">The row start.</param>
        /// <param name="RowEnd">The row end.</param>
        public static void HideRangeOfWorksheet(this Workbook workbook, string SheetName, uint RowStart, uint RowEnd)
        {
            WorksheetPart Sht = workbook.GetWorksheetPartByName(SheetName);

            Sht.Worksheet.HideRows(RowStart, RowEnd);
        }

        /// <summary>
        /// The defined name range comprises sheet and excel reference. This method decomposes this into contituents relevent for rows.
        /// Maybe requires extension to columns in the future.
        /// </summary>
        /// <param name="workbook">The workbook.</param>
        /// <param name="dn">The dn.</param>
        /// <param name="SheetName">Name of the sheet.</param>
        /// <param name="RowStart">The row start.</param>
        /// <param name="RowEnd">The row end.</param>
        /// <returns></returns>
        public static bool BreakDownDefinedName(this Workbook workbook, DefinedName dn, ref string SheetName, ref uint RowStart, ref uint RowEnd)
        {
            string[] DefinedNameRefElements = dn.Text.Split('!');

            SheetName = DefinedNameRefElements[0];

            if (SheetName.StartsWith("'") & SheetName.EndsWith("'"))
            {
                SheetName = SheetName.Substring(1, SheetName.Length - 2);
            }

            if (DefinedNameRefElements.Length > 0)
            {
                string[] DefinedNameCellElements = DefinedNameRefElements[1].Split(':');

                bool ok = true;

                if (DefinedNameCellElements.Length > 0)
                {
                    ok = TryGetRowIndex(DefinedNameCellElements[0], out RowStart);

                    if (ok && DefinedNameCellElements.Length > 1)
                    {
                        ok = TryGetRowIndex(DefinedNameCellElements[1], out RowEnd);
                        return ok;
                    }
                }
            }
            return false;
        }

        /// <summary>
        /// Breaks the name of down defined.
        /// </summary>
        /// <param name="workbook">The workbook.</param>
        /// <param name="dn">The dn.</param>
        /// <param name="SheetName">Name of the sheet.</param>
        /// <param name="RowStart">The row start.</param>
        /// <param name="RowEnd">The row end.</param>
        /// <param name="ColStart">The col start.</param>
        /// <param name="ColEnd">The col end.</param>
        /// <returns></returns>
        public static bool BreakDownDefinedName(this Workbook workbook, DefinedName dn, ref string SheetName, ref uint RowStart, ref uint RowEnd, ref uint ColStart, ref uint ColEnd)
        {
            bool val = BreakDownDefinedName(workbook, dn, ref SheetName, ref RowStart, ref RowEnd);

            // Example "Disclaimer!$A$3:$AW$35"

            string RangePart = dn.Text.Split('!')[1]; // Table the string to right of the !

            string StartPart = RangePart.Split(':')[0]; // the LHS of the :
            string EndPart = RangePart.Split(':')[1];   // the RHS of the :

            string StartColLetter = StartPart.Split('$')[1].ToString();
            string EndColLetter = EndPart.Split('$')[1].ToString();

            char[] StartColChars = StartColLetter.ToCharArray();
            char[] EndColChars = EndColLetter.ToCharArray();

            ColStart = Helpers.ColLetterNumber(StartColChars);
            ColEnd = Helpers.ColLetterNumber(EndColChars);
            
            return val;
        }

        /// <summary>
        /// Gets an instance of a <see cref="DefinedName" /> named range in a workbook.
        /// </summary>
        /// <param name="workbook">The workbook.</param>
        /// <param name="namedRangeName">Name of the named range.</param>
        /// <returns></returns>
        public static DefinedName GetDefinedNameByName(this Workbook workbook, string namedRangeName)
        {
            if (workbook.DefinedNames != null)
            {
                foreach (DefinedName dn in workbook.DefinedNames)
                {
                    if (dn.Name.HasValue)
                    {
                        string name = dn.Name.Value;
                        if (name.CompareTo(namedRangeName) == 0)
                        {
                            return dn;
                        }
                    }
                }
            }
            return null;
        }

        #endregion

        #region Private Static Methods

        /// <summary>
        /// Formats the name.
        /// </summary>
        /// <param name="name">The name.</param>
        /// <returns></returns>
        private static string FormatName(string name)
        { 
            bool isFirst = true;
            string[] nameCharacters = name.ToCharArray().Select(c =>
            {
                string result = "_";

                if (char.IsLetter(c) || (!isFirst && char.IsNumber(c)))
                {
                    result = c.ToString();
                }

                isFirst = false;

                return result;
            }).ToArray();

            return string.Join(string.Empty, nameCharacters);
        }

        /// <summary>
        /// Formats the name.
        /// </summary>
        /// <param name="definedNames">The defined names.</param>
        /// <param name="name">The name.</param>
        /// <returns></returns>
        private static string FormatName(DefinedNames definedNames, string name)
        {
            if (string.IsNullOrEmpty(name))
            {
                name = "Range";
            }

            name = FormatName(name);

            // Check for duplicate names.
            string numberedName = name;
            for (int i = 1; definedNames.OfType<DefinedName>().Any(d => string.Equals(d.Name, numberedName)); i++)
            {
                numberedName = string.Format("{0}_{1}", name, i);
            }

            return numberedName;
        }

        /// <summary>
        /// Decomposes the cell name into a row index (ie. D12 -&gt; 12) As we currently wish only to hide rows the column letter is
        /// irrelevent.
        /// </summary>
        /// <param name="cellName">Name of the cell.</param>
        /// <param name="value">The value.</param>
        /// <returns></returns>
        private static bool TryGetRowIndex(string cellName, out uint value)
        {
            value = 0;

            // Create a regular expression to match the row index portion the cell name.
            var regex = new Regex(@"\d+");
            Match match = regex.Match(cellName);

            uint result;
            
            if( uint.TryParse(match.Value, out result))
            {
                value = result;
                return true;
            }
            return false;
        }


        #endregion
    }
}
