namespace ExcelWriter
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text;
    using OpenXml.Excel;

    using OpenXmlSpreadsheet = DocumentFormat.OpenXml.Spreadsheet;
    using DocumentFormat.OpenXml.Packaging;

    /// <summary>
    /// Helper to manage writing of rows, columns and cell to Excel.
    /// </summary>
    internal class OpenXmlExcelWriteManager
    {
        #region Private Fields

        private OpenXmlSpreadsheet.Worksheet worksheet;
        private OpenXmlSpreadsheet.SheetData sheetData;
        private Dictionary<uint, string> columnLetterStore;
        private Dictionary<uint, OpenXmlSpreadsheet.Column> excelColumns;
        private Dictionary<uint, OpenXmlSpreadsheet.Row> excelRows;

        #endregion Private Fields

        #region Construction

        /// <summary>
        /// Initializes a new instance of the <see cref="OpenXmlExcelWriteManager" /> class.
        /// </summary>
        public OpenXmlExcelWriteManager(OpenXmlSpreadsheet.Worksheet worksheet)
        {
            if (worksheet == null)
            {
                throw new ArgumentNullException("worksheet");
            }

            this.worksheet = worksheet;
            this.sheetData = worksheet.GetFirstChild<OpenXmlSpreadsheet.SheetData>();

            this.columnLetterStore = new Dictionary<uint, string>();
            this.excelColumns = new Dictionary<uint, OpenXmlSpreadsheet.Column>();
            this.excelRows = new Dictionary<uint, OpenXmlSpreadsheet.Row>();
        }

        #endregion Construction

        #region Public Properties

        #endregion Public Properties

        /// <summary>
        /// Sets the data type and value of the cell according to the type of the supplied value
        /// </summary>
        /// <param name="cell">The cell into which the value is to be written</param>
        /// <param name="value">The value to be written into the cell</param>
        public void SetCellValue(OpenXmlSpreadsheet.Cell cell, object value)
        {

            // Convert the supplied type to some underlying type that Excel can interpret.
            value = ToUnderlyingTypeValue(value);

            // Convert this underlying type to an Excel string
            string excelString = ToExcelString(value);

            // Set the DataType only if we are setting the value to a string.
            if ((value is string) || (value is bool) || (value is bool?))
            {
                cell.DataType = new DocumentFormat.OpenXml.EnumValue<OpenXmlSpreadsheet.CellValues>(OpenXmlSpreadsheet.CellValues.String);
            }
            else
            {
                cell.DataType = null;
            }

            cell.CellValue = new OpenXmlSpreadsheet.CellValue(excelString);

            // Preserve white space in our cell values if it is a string.
            if (value is string)
            {
                cell.CellValue.Space = DocumentFormat.OpenXml.SpaceProcessingModeValues.Preserve;
            }
        }

        /// <summary>
        /// Gets a row at a specified row index
        /// </summary>
        /// <param name="rowIndex"></param>
        /// <returns></returns>
        public OpenXmlSpreadsheet.Row GetRow(uint rowIndex)
        {
            return this.excelRows[rowIndex];
        }

        /// <summary>
        /// Convert the supplied value to some type that can be interpreted in Excel.
        /// </summary>
        /// <param name="value"></param>
        /// <returns></returns>
        private static object ToUnderlyingTypeValue(object value)
        {
            if (value is System.DBNull)
            {
                return null;
            }

            // Do some basic conversion from nullables
            if (value is DateTime?)
            {
                DateTime? newValue = (DateTime?)value;
                return newValue.GetValueOrDefault();
            }
            else if (value is double?)
            {
                double? newValue = (double?)value;
                return newValue.GetValueOrDefault();
            }
            else if (value is bool?)
            {
                bool? boolean = (bool?)value;
                return boolean.GetValueOrDefault();
            }

            else if (value is DateTime)
            {
                return value;
            }
            else if (value is double)
            {
                return value;
            }
            else if (value is bool)
            {
                return value;
            }
            else if (value is string)
            {
                return value;
            }

            // Finally, attempt conversion to a double (general numeric conversion)
            try
            {
                return System.Convert.ToDouble(value);
            }
            catch (InvalidCastException)
            {
                // Just convert to a string and write out some debug information.
                //System.Diagnostics.Debug.Write(string.Format("Value '{0}' is not a type which can be converted for export to Excel", value));
                return null; // value.ToString();
            }
        }

        /// <summary>
        /// Convert the supplie underlying type to Excel cell value
        /// </summary>
        /// <param name="value">The value to be converted</param>
        /// <returns>An excel formated string representation of a value.</returns>
        private static string ToExcelString(object value)
        {
            // Convert the underlying to a string and set the type
            string excelString = string.Empty;
            if (value is DateTime)
            {
                DateTime dateTime = (DateTime)value;
                excelString = dateTime.ToOADate().ToString(System.Globalization.CultureInfo.InvariantCulture);
            }
            else if (value is double)
            {
                double doubleValue = (double)value;
                if (double.IsNaN(doubleValue) || double.IsInfinity(doubleValue))
                {
                    excelString = string.Empty;
                }
                else
                {
                    excelString = value + string.Empty;
                }
            }
            else if (value is bool)
            {
                excelString = value.ToString().ToUpper();
            }
            else
            {
                excelString = value + string.Empty;
            }
            return excelString;
        }

        #region Public Methods

        /// <summary>
        /// Looks up the column letter required for a specified worksheet column index (1=A)
        /// </summary>
        /// <param name="columnIndex">The index of the column</param>
        /// <returns>The letter representing the column</returns>
        public string GetColumnLetter(uint columnIndex)
        {
            string columnLetter = null;
            if (this.columnLetterStore.ContainsKey(columnIndex))
            {
                columnLetter = this.columnLetterStore[columnIndex];
            }
            else
            {
                columnLetter = CellExtensions.GetColumnLetter(columnIndex);
                this.columnLetterStore.Add(columnIndex, columnLetter);
            }

            return columnLetter;
        }

        #endregion Public Methods

        /// <summary>
        /// Removes <see cref="OpenXmlSpreadsheet.Columns"/> collection, all <see cref="OpenXmlSpreadsheet.Column"/>s, all <see cref="OpenXmlSpreadsheet.MergeCells"/>
        /// and all <see cref="OpenXmlSpreadsheet.Row"/s> from the supplied <see cref="WorksheetPart"/>
        /// </summary>
        /// <param name="worksheetPart"></param>
        public void EmptyWorksheet()
        {
            OpenXmlSpreadsheet.Columns columns = this.worksheet.GetFirstChild<OpenXmlSpreadsheet.Columns>();

            // Clear all contents of the sheet if this is an existing sheet.
            this.worksheet.RemoveAllChildren<OpenXmlSpreadsheet.MergeCells>();

            this.sheetData.RemoveAllChildren<OpenXmlSpreadsheet.Row>();
            if (columns != null)
            {
                columns.RemoveAllChildren<OpenXmlSpreadsheet.Column>();
            }
            this.worksheet.RemoveAllChildren<OpenXmlSpreadsheet.Columns>();
        }

        public void AddColumn(double? columnWidth, bool isHidden)
        {
            // Add the column and set collated ionformation
            var newColumn = this.sheetData.AddColumn();

            // Set the width?
            if (columnWidth.HasValue)
            {
                newColumn.Width = new DocumentFormat.OpenXml.DoubleValue(columnWidth.Value);
            }

            // Hide?
            if (isHidden)
            {
                newColumn.Hidden = new DocumentFormat.OpenXml.BooleanValue(true);
            }

            // Add in to dictionary
            excelColumns.Add((uint)excelColumns.Count + 1, newColumn);
        }

        public void AddRow(double? rowHeight, bool isHidden)
        {
            // Add the row and set collated information
            var newRow = sheetData.AddRow();

            // Set the height?
            if (rowHeight.HasValue)
            {
                newRow.Height = new DocumentFormat.OpenXml.DoubleValue(rowHeight);
                newRow.CustomHeight = new DocumentFormat.OpenXml.BooleanValue(true);
            }

            // Hide?
            if (isHidden)
            {
                newRow.Hidden = new DocumentFormat.OpenXml.BooleanValue(true);
            }

            // Add in to dictionary
            excelRows.Add((uint)excelRows.Count + 1, newRow);
        }
    }
}
