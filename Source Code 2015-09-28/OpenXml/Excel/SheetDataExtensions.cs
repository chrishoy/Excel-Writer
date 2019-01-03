namespace ExcelWriter.OpenXml.Excel
{
    using System;
    using System.Linq;
    using DocumentFormat.OpenXml;
    using DocumentFormat.OpenXml.Spreadsheet;
    using System.Collections.Generic;

    /// <summary>
    /// Extension methods for <see cref="SheetData"/>
    /// </summary>
    public static class SheetDataExtensions
    {
        #region Public Static Methods

        /// <summary>
        /// Adds a column with the specified fixed width.
        /// </summary>
        /// <param name="sheetData">The sheet data.</param>
        /// <param name="width">The width.</param>
        /// <returns>
        /// The new column.
        /// </returns>
        public static Column AddColumn(this SheetData sheetData, double width)
        {
            var worksheet = (Worksheet)sheetData.Parent;
            var columns = worksheet.GetFirstChild<Columns>();
            if (columns == null)
            {
                columns = new Columns();
                worksheet.InsertBefore<Columns>(columns, sheetData);
            }

            uint columnIndex = (uint)sheetData.GetColumnCount() + 1;

            var column = new Column
            {
                CustomWidth = new BooleanValue(true),
                Max = columnIndex,
                Min = columnIndex,
                Width = new DoubleValue(width),
            };

            columns.Append(column);

            return column;
        }

        /// <summary>
        /// Adds a column.
        /// </summary>
        /// <param name="sheetData">The sheet data.</param>
        /// <returns>
        /// The new column.
        /// </returns>
        public static Column AddColumn(this SheetData sheetData)
        {
            var worksheet = (Worksheet)sheetData.Parent;
            var columns = worksheet.GetFirstChild<Columns>();
            if (columns == null)
            {
                columns = new Columns();
                worksheet.InsertBefore(columns, sheetData);
            }

            uint columnIndex = (uint)sheetData.GetColumnCount() + 1;
            var column = new Column
            {
                Max = columnIndex,
                Min = columnIndex,
                Width = new DoubleValue(20D),
            };

            columns.Append(column);

            return column;
        }

        /// <summary>
        /// Adds a row.
        /// </summary>
        /// <param name="sheetData">The sheet data.</param>
        /// <returns>
        /// The new row.
        /// </returns>
        public static Row AddRow(this SheetData sheetData)
        {
            uint rowIndex = (uint)sheetData.GetRowCount() + 1;
            var row = new Row
            {
                RowIndex = rowIndex,
            };
            
            sheetData.Append(row);

            return row;
        }

        /// <summary>
        /// Adds a row.
        /// </summary>
        /// <param name="sheetData">The sheet data.</param>
        /// <param name="height">The height.</param>
        /// <returns>
        /// The new row.
        /// </returns>
        public static Row AddRow(this SheetData sheetData, double height)
        {
            uint rowIndex = (uint)sheetData.GetRowCount() + 1;
            var row = new Row
            {
                RowIndex = rowIndex,
                Height = new DoubleValue(height),
                CustomHeight = new BooleanValue(true),
            };

            sheetData.Append(row);

            return row;
        }

        /// <summary>
        /// Sets the height of a row in a worksheet.
        /// </summary>
        /// <param name="sheetData">The sheet data.</param>
        /// <param name="rowIndex">Index of the row.</param>
        /// <param name="height">The height.</param>
        [Obsolete("Original call supplying height as uint should have been double. Change type of height parameter to double.")]
        public static void SetRowHeight(this SheetData sheetData, uint rowIndex, uint height)
        {
            SetRowHeight(sheetData, rowIndex, Convert.ToDouble(height));
        }

        /// <summary>
        /// Sets the height of a row in a worksheet
        /// </summary>
        /// <param name="sheetData">The sheet data.</param>
        /// <param name="rowIndex">Index of the row.</param>
        /// <param name="height">The height.</param>
        public static void SetRowHeight(this SheetData sheetData, uint rowIndex, double height)
        {
            IEnumerable<Row> rows = sheetData.OfType<Row>();
            Row row = rows.FirstOrDefault(r=>r.RowIndex == rowIndex);
            if (row != null)
            {
                row.Height = new DoubleValue(height);
                row.CustomHeight = new BooleanValue(true);
            }
        }

        /// <summary>
        /// Sets the height of a row in a worksheet
        /// </summary>
        /// <param name="sheetData">The sheet data.</param>
        /// <param name="rowIndex">Index of the row.</param>
        /// <param name="hidden">if set to <c>true</c> [hidden].</param>
        public static void SetRowHidden(this SheetData sheetData, uint rowIndex, bool hidden)
        {
            IEnumerable<Row> rows = sheetData.OfType<Row>();
            Row row = rows.FirstOrDefault(r => r.RowIndex == rowIndex);
            if (row != null)
            {
                row.Hidden = new BooleanValue(hidden);
            }
        }

        /// <summary>
        /// Merges the range of cells into one cell.
        /// </summary>
        /// <param name="sheetData">The sheet data.</param>
        /// <param name="startCell">The start cell (top left).</param>
        /// <param name="endCell">The end cell (bottom right).</param>
        public static void MergeCells(this SheetData sheetData, Cell startCell, Cell endCell)
        {
            var worksheet = (Worksheet)sheetData.Parent;

            MergeCells mergeCells = worksheet.OfType<MergeCells>().FirstOrDefault();
            if (mergeCells == null)
            {
                mergeCells = new MergeCells();
                worksheet.InsertAfter(mergeCells, sheetData);
            }

            mergeCells.Append(
                new MergeCell
                {
                    Reference = string.Format("{0}:{1}", startCell.CellReference.Value, endCell.CellReference.Value)
                });
        }

        /// <summary>
        /// Gets the cell at the specified coordinates or creates it if a cell does not exist there.
        /// </summary>
        /// <param name="sheetData">The sheet data.</param>
        /// <param name="columnIndex">Index of the column.</param>
        /// <param name="rowIndex">Index of the row.</param>
        /// <returns>
        /// The new cell.
        /// </returns>
        public static Cell GetCell(this SheetData sheetData, uint columnIndex, uint rowIndex)
        {
            while (rowIndex > sheetData.OfType<Row>().Count())
            {
                sheetData.AddRow();
            }

            var row = (Row)sheetData.ElementAt((int)rowIndex - 1);
            
            while (columnIndex > row.OfType<Cell>().Count())
            {
                var paddingCell = new Cell();
                paddingCell.SetCellReference((uint)row.OfType<Cell>().Count() + 1, rowIndex);
                row.Append(paddingCell);
            }

            Cell cell = (Cell)row.ElementAt((int)columnIndex - 1);

            return cell;
        }

        /// <summary>
        /// Gets a cell if exists in the worksheet. Returns null if cell does not exist.
        /// </summary>
        /// <param name="sheetData">The sheet data.</param>
        /// <param name="columnIndex">Index of the column.</param>
        /// <param name="rowIndex">Index of the row.</param>
        /// <returns></returns>
        public static Cell GetCellIfExists(this SheetData sheetData, uint columnIndex, uint rowIndex)
        {
            Row row = GetRowIfExists(sheetData, rowIndex);
            if (row != null)
            {
                string cellReference = CellExtensions.GetCellReference(columnIndex, rowIndex);
                return row.OfType<Cell>().FirstOrDefault(c => c.CellReference == cellReference);
            }
            return null;
        }

        /// <summary>
        /// Gets a row if it exists in a worksheet. Returns null if row does not exist.
        /// </summary>
        /// <param name="sheetData">The sheet data.</param>
        /// <param name="rowIndex">Index of the row.</param>
        /// <returns></returns>
        public static Row GetRowIfExists(this SheetData sheetData, uint rowIndex)
        {
            return sheetData.OfType<Row>().FirstOrDefault(r => r.RowIndex == rowIndex);
        }

        /// <summary>
        /// Gets the column count.
        /// </summary>
        /// <param name="sheetData">The sheet data.</param>
        /// <returns>
        /// The number of columns.
        /// </returns>
        public static int GetColumnCount(this SheetData sheetData)
        {
            var worksheet = (Worksheet)sheetData.Parent;
            var columns = worksheet.GetFirstChild<Columns>();

            return columns.OfType<Column>().Count();
        }

        /// <summary>
        /// Gets the row count.
        /// </summary>
        /// <param name="sheetData">The sheet data.</param>
        /// <returns>
        /// The number of rows.
        /// </returns>
        public static int GetRowCount(this SheetData sheetData)
        {
            return sheetData.OfType<Row>().Count();
        }

        /// <summary>
        /// Sets the cell contents and style. Handles formatting of the column automatically.
        /// </summary>
        /// <param name="sheetData">The sheet data.</param>
        /// <param name="columnIndex">Index of the column.</param>
        /// <param name="rowIndex">Index of the row.</param>
        /// <param name="value">The value.</param>
        /// <returns></returns>
        public static Cell SetCell(this SheetData sheetData, uint columnIndex, uint rowIndex, object value)
        {
            Cell cell = sheetData.GetCell(columnIndex, rowIndex);
            cell.SetDataType(value);
            cell.SetCellValue(value);

            return cell;
        }

        /// <summary>
        /// Sets the type of the data.
        /// </summary>
        /// <param name="cell">The cell.</param>
        /// <param name="value">The value.</param>
        public static void SetDataType(this Cell cell, object value) 
        {
            if ((value is string) || (value is bool) || (value is bool?))
            {
                cell.DataType = new EnumValue<CellValues>(CellValues.String);
            }
        }

        /// <summary>
        /// Sets the cell contents and style. Handles formatting of the column automatically.
        /// </summary>
        /// <param name="sheetData">The sheet data.</param>
        /// <param name="columnIndex">Index of the column.</param>
        /// <param name="rowIndex">Index of the row.</param>
        /// <param name="value">The value.</param>
        /// <param name="styleIndex">Index of the style.</param>
        /// <returns></returns>
        public static Cell SetCell(this SheetData sheetData, uint columnIndex, uint rowIndex, object value, uint styleIndex)
        {
            Cell cell = SetCell(sheetData, columnIndex, rowIndex, value);
            cell.StyleIndex = styleIndex;

            return cell;
        }

        #endregion
    }
}
