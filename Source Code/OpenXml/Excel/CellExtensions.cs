namespace ExcelWriter.OpenXml.Excel
{
    using System;
    using DocumentFormat.OpenXml;
    using DocumentFormat.OpenXml.Spreadsheet;
    using System.Text.RegularExpressions;

    public static class CellExtensions
    {
        #region Public Static Methods

        /// <summary>
        /// Gets the cell formula for the specified column and row indices.
        /// </summary>
        /// <param name="columnIndex">Index of the column.</param>
        /// <param name="rowIndex">Index of the row.</param>
        /// <returns>The cell formula.</returns>
        public static string GetCellFormula(uint columnIndex, uint rowIndex)
        {
            return string.Format("${0}${1}", GetColumnLetter(columnIndex), rowIndex);
        }

        /// <summary>
        /// Gets the cell reference for the specified the column and row indices.
        /// </summary>
        /// <param name="columnIndex">Index of the column.</param>
        /// <param name="rowIndex">Index of the row.</param>
        /// <returns>The cell reference.</returns>
        public static string GetCellReference(uint columnIndex, uint rowIndex)
        {
            return string.Format("{0}{1}", GetColumnLetter(columnIndex), rowIndex);
        }

        public static string GetCellReference(string columnLetter, uint rowIndex)
        {
            return string.Format("{0}{1}", columnLetter, rowIndex);
        }

        /// <summary>
        /// Gets the column letter using the column index.
        /// </summary>
        /// <param name="columnIndex">Index of the column.</param>
        /// <returns>The column letter.</returns>
        public static string GetColumnLetter(uint columnIndex)
        {
            uint dividend = columnIndex;
            string columnName = string.Empty;

            while (dividend > 0)
            {
                uint modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo).ToString() + columnName;
                dividend = (uint)((dividend - modulo) / 26);
            }

            return columnName;
        }

        /// <summary>
        /// Given a cell reference (eg. A10), parses the specified cell to get the column reference (eg. A)
        /// </summary>
        /// <param name="cellName"></param>
        /// <returns></returns>
        public static string GetColumnReference(string cellName)
        {
            // Create a regular expression to match the column name portion of the cell name.
            Regex regex = new Regex("[A-Za-z]+");
            Match match = regex.Match(cellName);
            return match.Value;
        }

        /// <summary>
        /// Given a cell reference (eg. A10), parses the specified cell to get the row index (eg. 10).
        /// </summary>
        /// <param name="cellReference"></param>
        /// <returns></returns>
        public static uint GetRowIndex(string cellReference)
        {
            // Create a regular expression to match the row index portion the cell name.
            Regex regex = new Regex("\\d+");
            Match match = regex.Match(cellReference);
            return uint.Parse(match.Value);
        }

        /// <summary> 
        /// Given the full cell name, it will return the zero based column index. 
        /// Note: This method will only handle columns with a length of up to two (ie. A to Z and AA to ZZ).  
        /// A length of three can be implemented when needed. 
        /// </summary> 
        /// <param name="cellReference">Cell reference (eg. A1)</param> 
        /// <returns>Zero based index if the conversion was successful; otherwise null</returns> 
        public static uint? GetColumnIndex(string cellReference)
        {
            string columnReference = GetColumnReference(cellReference);
            return GetColumnIndexFromReference(columnReference);
        }

        /// <summary> 
        /// Given just the column name (no row index), it will return the zero based column index. 
        /// Note: This method will only handle columns with a length of up to two (ie. A to Z and AA to ZZ).  
        /// A length of three can be implemented when needed. 
        /// </summary> 
        /// <param name="columnReference">Column Name (ie. A or AB)</param> 
        /// <returns>Zero based index if the conversion was successful; otherwise null</returns> 
        public static uint? GetColumnIndexFromReference(string columnReference)
        {
            uint? columnIndex = null;

            char[] colLetters = GetColumnReference(columnReference).ToCharArray();

            if (colLetters.Length <= 2)
            {
                int index = 0;
                foreach (char col in colLetters)
                {
                    uint indexValue = Convert.ToUInt32(col) - 64;

                    if (indexValue >= 0)
                    {
                        // The first letter of a two digit column needs some extra calculations 
                        if (index == 0 && colLetters.Length == 2)
                        {
                            columnIndex = columnIndex == null ? indexValue * 26 : columnIndex + (indexValue * 26);
                        }
                        else
                        {
                            columnIndex = columnIndex == null ? indexValue : columnIndex + indexValue;
                        }
                    }

                    index++;
                }
            }

            return columnIndex;
        }

        /// <summary>
        /// Sets the cell reference using the column and row indices.
        /// </summary>
        /// <param name="cell">The cell.</param>
        /// <param name="columnIndex">Index of the column.</param>
        /// <param name="rowIndex">Index of the row.</param>
        public static void SetCellReference(this DocumentFormat.OpenXml.Spreadsheet.Cell cell, uint columnIndex, uint rowIndex)
        {
            cell.CellReference = GetCellReference(columnIndex, rowIndex);
        }

        public static void SetCellReference(this DocumentFormat.OpenXml.Spreadsheet.Cell cell, string columnLetter, uint rowIndex)
        {
            cell.CellReference = GetCellReference(columnLetter, rowIndex);
        }

        /// <summary>
        /// Sets the CellValue and DataValue to an Excel recognized string. 
        /// Special formatting is used for DateTime, Double and Boolean values.
        /// </summary>
        /// <param name="cell">The cell.</param>
        /// <param name="value">The value.</param>
        public static void SetCellValue(this DocumentFormat.OpenXml.Spreadsheet.Cell cell, object value)
        {
            string excelString = string.Empty;

            if (value is DateTime?)
            {
                DateTime? dateTime = (DateTime?)value;
                value = dateTime.GetValueOrDefault();
            }
            else if (value is double?)
            {
                double? boolean = (double?)value;
                value = boolean.GetValueOrDefault();
            }
            else if (value is bool?)
            {
                bool? boolean = (bool?)value;
                value = boolean.GetValueOrDefault();
            }

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

            // Set the DataType only if we are setting the value to a string.
            if ((value is string) || (value is bool) || (value is bool?))
            {
                cell.DataType = new EnumValue<CellValues>(CellValues.String);
            }
            else
            {
                cell.DataType = null;
            }

            cell.CellValue = new CellValue(excelString);

            // Preserve white space in our cell values if it is a string.
            if (value is string)
            {
                cell.CellValue.Space = SpaceProcessingModeValues.Preserve;
            }
        }

        #endregion
    }
}
