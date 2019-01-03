namespace ExcelWriter
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Windows;

    using Constants;

    /// <summary>
    /// Represents information which enables a table column headings to be mapped to Excel.<br />
    /// Holds the column and any column header which may have been applied to that column.
    /// </summary>
    internal class TableColumnsInfo
    {
        #region Private Fields

        /// <summary>
        /// The column infos
        /// </summary>
        private readonly TableColumnInfo[] columnInfos;

        /// <summary>
        /// The group header row infos
        /// </summary>
        private SortedDictionary<int, GroupHeaderRowInfo> groupHeaderRowInfos;
        /// <summary>
        /// The table
        /// </summary>
        private Table table;
        /// <summary>
        /// The column count
        /// </summary>
        private int columnCount;
        /// <summary>
        /// The column width divisor
        /// </summary>
        private double columnWidthDivisor;

        #endregion Private Fields

        #region Construction

        /// <summary>
        /// Initializes a new instance of the <see cref="TableColumnsInfo" /> class.
        /// </summary>
        /// <param name="table">The table which is used to seed this information</param>
        /// <param name="legacyProcess">if set to <c>true</c> [legacy process].</param>
        /// <exception cref="ArgumentNullException">table</exception>
        public TableColumnsInfo(Table table, bool legacyProcess)
        {
            if (table == null)
            {
                throw new ArgumentNullException("table");
            }

            // Why we have to provide 5 here I don't know (it's a divisor for table column widths)
            // .... But will have to stay for the old templates.
            this.columnWidthDivisor = legacyProcess ? ExcelConstants.LegacyColumnWidthDivisor : 1d;
            this.groupHeaderRowInfos = new SortedDictionary<int, GroupHeaderRowInfo>();

            this.table = table;
            this.columnCount = table.TableData.Columns.Count;

            // Build an array of information relating to the table columns and headers.
            this.columnInfos = new TableColumnInfo[this.columnCount];
            this.AssignColumns(table);

            // Add in header information
            this.AssignGroupHeaders(table);
        }

        #endregion Construction

        #region Public Properties

        /// <summary>
        /// Gets the header row infos.
        /// </summary>
        /// <value>
        /// The header row infos.
        /// </value>
        public IEnumerable<GroupHeaderRowInfo> HeaderRowInfos
        {
            get { return this.groupHeaderRowInfos.Values; }
        }

        /// <summary>
        /// Gets the number of columns that this table represents.
        /// </summary>
        /// <value>
        /// The column count.
        /// </value>
        public int ColumnCount
        {
            get { return this.columnCount; }
        }

        /// <summary>
        /// Gets the table which is represented by this <see cref="TableColumnsInfo" />.
        /// </summary>
        /// <value>
        /// The table.
        /// </value>
        public Table Table    
        {
            get { return this.table; }
        }

        /// <summary>
        /// Gets an array of <see cref="TableColumnInfo" />s, each elements representing a <see cref="TableColumn" /> and <see cref="TableColumnHeader" />.
        /// </summary>
        /// <value>
        /// The column infos.
        /// </value>
        public TableColumnInfo[] ColumnInfos
        {
            get { return this.columnInfos; }
        }

        #endregion Public Properties

        #region Private Helpers

        /// <summary>
        /// Assign columns and column values to an array of columns for this table
        /// </summary>
        /// <param name="table">The table.</param>
        private void AssignColumns(Table table)
        {
            for (int columnIdx = 0; columnIdx < this.columnCount; columnIdx++)
            {
                TableColumn column = table.TableData.Columns[columnIdx];
                bool isLastColumn = columnIdx == (this.columnCount - 1);

                // Set ParentDataContext for value resolution
                column.ParentDataContext = table.TableData.DataContext;
                if (column.DataContext == null)
                {
                    column.DataContext = column.ParentDataContext;
                }

                this.columnInfos[columnIdx] = new TableColumnInfo
                {
                    Column = column,
                    Width = column.Width / this.columnWidthDivisor,
                    Hidden = BindingContainer.ConvertToNullableBoolean(column.ColumnIsHidden).GetValueOrDefault(false),
                    ColumnSpan = column.ColumnSpan,
                    IsLastColumn = isLastColumn,
                    SpanLastColumn = isLastColumn && table.SpanLastColumn,
                };
            }
        }

        /// <summary>
        /// Assigns the group headers defined on a table to the columns.
        /// </summary>
        /// <param name="table">The table.</param>
        private void AssignGroupHeaders(Table table)
        {
            // Reads the number of column groups that are specified in the column headers
            // These are additional grouped columns above the column headers, one for each level
            var columnHeaderGroups = GetColumnGroupHeaders(table.ColumnHeaders);

            // Group Column Header Levels
            foreach (var columnHeaderLevel in columnHeaderGroups)
            {
                var rowInfo = new GroupHeaderRowInfo();
                rowInfo.Level = columnHeaderLevel.Key;

                foreach (TableColumnHeader columnHeader in columnHeaderLevel)
                {
                    // Update height
                    if (columnHeader.Height.HasValue)
                    {
                        // If header height not yet assigned, or is smaller than current then assign.
                        if (!rowInfo.Height.HasValue || rowInfo.Height.Value < columnHeader.Height.Value)
                        {
                            rowInfo.Height = columnHeader.Height;
                        }
                    }

                    // Check hidden state (need to change this!!! Don't use Visibility)
                    if (!rowInfo.Hidden)
                    {
                        rowInfo.Hidden = columnHeader.Visibility != Visibility.Visible;
                    }

                    for (int columnNo = columnHeader.Start; columnNo <= columnHeader.Finish; columnNo++)
                    {
                        if (columnNo <= this.columnInfos.Length)
                        {
                            TableColumnInfo colInfo = this.columnInfos[columnNo - 1];
                            colInfo.AddHeader(columnHeaderLevel.Key, columnHeader);
                        }
                    }
                }

                // Add it to the sorted list.
                this.groupHeaderRowInfos.Add(rowInfo.Level, rowInfo);
            }
        }

        /// <summary>
        /// Determine the number of column groups that are specified in the column headers
        /// These are additional grouped columns above the column headers, one for each level.
        /// </summary>
        /// <param name="columnHeaders">The column headers.</param>
        /// <returns></returns>
        private static IOrderedEnumerable<IGrouping<int, TableColumnHeader>> GetColumnGroupHeaders(TableColumnHeaderCollection columnHeaders)
        {
            var columnHeaderGroups = columnHeaders.GroupBy(x => x.Level).OrderBy(x => x.Key);
            return columnHeaderGroups;
        }

        #endregion Private Helpers
    }
}
