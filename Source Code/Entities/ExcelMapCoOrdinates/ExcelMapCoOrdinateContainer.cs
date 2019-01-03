namespace ExcelWriter
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text;
    using System.Windows;
    
    /// <summary>
    /// Represents a container for rows and columns so then can be grouped adn stacked into a worksheet.<br/>
    /// Used for managing <see cref="StackPanel"/> to Excel Workbook translation.
    /// </summary>
    internal class ExcelMapCoOrdinateContainer : ExcelMapCoOrdinate
    {
        #region Private Fields

        // Create container for Columns, Rows and Cells
        private ExcelMapCoOrdinateCellList cells = new ExcelMapCoOrdinateCellList();

        private string containerType;               // Mainly used for debugging used in ToString()
        private uint currentColumnIndex;
        private uint currentRowIndex;

        private uint mapRowCount;
        private uint mapColumnCount;

        #endregion Private Fields
        
        #region Construction

        /// <summary>
        /// Constructs a co-ordinate which based on a cell located by row and column 
        /// </summary>
        /// <param name="rowIndex"></param>
        /// <param name="columnIndex"></param>
        public ExcelMapCoOrdinateContainer(uint rowIndex, uint columnIndex, string containerType) : base()
        {
            this.SetCurrentColumn(columnIndex);
            this.SetCurrentRow(rowIndex);
            this.containerType = containerType;
        }

        #endregion Construction

        #region Public Properties

        /// <summary>
        /// Current Excel column index for writing a value into a cell.<br/>
        /// Note that this is 1 based.
        /// </summary>
        public uint CurrentColumnIndex
        {
            get { return this.currentColumnIndex; }
        }

        /// <summary>
        /// Current Excel row index for writing a value into a cell<br/>
        /// Note that this is 1 based.
        /// </summary>
        public uint CurrentRowIndex
        {
            get { return this.currentRowIndex; }
        }

        /// <summary>
        /// The number of worksheet rows that this entity represents
        /// </summary>
        public override uint MapRowCount
        {
            get { return this.mapRowCount; }
        }

        /// <summary>
        /// The number of worksheet columns that this entity represents
        /// </summary>
        public override uint MapColumnCount
        {
            get { return this.mapColumnCount; }
        }

        #endregion Public Properties

        #region Internal Properties

        /// <summary>
        /// Used primarily for debugging, so we can tract what is in what...
        /// </summary>
        internal string ContainerType
        {
            get { return this.containerType; }
        }

        /// <summary>
        /// Gets the cells that this container contains.
        /// </summary>
        internal ExcelMapCoOrdinateCellList Cells
        {
            get { return this.cells; }
            set { this.cells = value; }
        }

        #endregion Internal Properties

        #region Public Methods

        /// <summary>
        /// Sets the CurrentRowIndex to the specified value.
        /// </summary>
        /// <param name="newIndex"></param>
        public void SetCurrentRow(uint newIndex)
        {
            if (newIndex > this.currentRowIndex)
            {
                if (newIndex > this.mapRowCount) this.mapRowCount = newIndex;
            }
            this.currentRowIndex = newIndex;
        }

        /// <summary>
        /// Sets the CurrentColumnIndex to the specified value.
        /// </summary>
        /// <param name="newIndex"></param>
        public void SetCurrentColumn(uint newIndex)
        {
            if (newIndex > this.mapColumnCount)
            {
                if (newIndex > this.mapColumnCount) this.mapColumnCount = newIndex;
            }
            this.currentColumnIndex = newIndex;
        }

        /// <summary>
        /// Moves the CurrentRowIndex to the next row.
        /// </summary>
        public void MoveToNextRow()
        {
            this.SetCurrentRow(this.CurrentRowIndex + 1);
        }

        /// <summary>
        /// Moves the CurrentColumnIndex to the next row.
        /// </summary>
        public void MoveToNextColumn()
        {
            this.SetCurrentColumn(this.CurrentColumnIndex + 1);
        }

        /// <summary>
        /// Moves the CurrentRowIndex to the first row (if any).
        /// </summary>
        public void MoveToFirstRow()
        {
            if (this.MapRowCount > 0)
            {
                this.SetCurrentRow(1);
            }
        }

        /// <summary>
        /// Moves the CurrentColumnIndex to the first column (if any).
        /// </summary>
        public void MoveToFirstColumn()
        {
            if (this.MapColumnCount > 0)
            {
                this.SetCurrentColumn(1);
            }
        }

        /// <summary>
        /// Moves the CurrentRowIndex to the last row (if any).
        /// </summary>
        public void MoveToLastRow()
        {
            if (this.MapRowCount > 0)
            {
                this.SetCurrentRow(this.MapRowCount);
            }
        }

        /// <summary>
        /// Moves the CurrentColumnIndex to the last column (if any).
        /// </summary>
        public void MoveToLastColumn()
        {
            if (this.MapColumnCount > 0)
            {
                this.SetCurrentColumn(this.MapColumnCount);
            }
        }

        #endregion Public Methods

        #region Internal Methods

        /// <summary>
        /// Sets an <see cref="ExelMapCoOrdinate"/> at the current location.
        /// </summary>
        /// <param name="excelMapCoOrdinate"></param>
        internal void SetExcelMapCoOrdinate(ExcelMapCoOrdinate exelMapCoOrdinate)
        {
            if (this.currentRowIndex < 1) throw new InvalidOperationException("CurrentRow can not be < 1");
            if (this.currentColumnIndex < 1) throw new InvalidOperationException("CurrentColumn can not be < 1");

            var key = new System.Drawing.Point((int)this.currentColumnIndex, (int)this.currentRowIndex);
            if (this.cells.ContainsKey(key))
            {
                this.cells[key] = exelMapCoOrdinate;
            }
            else
            {
                this.cells.Add(key, exelMapCoOrdinate);
            }
            exelMapCoOrdinate.SetParent(this);
        }

        /// <summary>
        /// Build and return a list of <see cref="ExcelDefinedNameInfo"/> which represents
        /// worksheet defined names that are present in this container.
        /// </summary>
        /// <param name="definedNameList"></param>
        internal void UpdateDefinedNameList(ref List<ExcelDefinedNameInfo> definedNameList, string sheetName)
        {
            // Update list if this container has a defined name.
            ExcelMapCoOrdinate.UpdateDefinedNameList(this, sheetName, ref definedNameList);

            // Check contained maps.
            foreach (var cell in this.cells)
            {
                if (cell.Value is ExcelMapCoOrdinateCell)
                {
                    // Each cell can have a defined name
                    ExcelMapCoOrdinate.UpdateDefinedNameList(cell.Value, sheetName, ref definedNameList);
                }
                else if (cell.Value is ExcelMapCoOrdinateContainer)
                {
                    // Each container can have a defined name
                    (cell.Value as ExcelMapCoOrdinateContainer).UpdateDefinedNameList(ref definedNameList, sheetName);
                }
            }
        }

        #endregion Internal Methods

        #region Overrides

        /// <summary>
        /// Gets the 1-based index of excel row where this entity starts in the excel worksheet
        /// </summary>
        public override int ExcelRowStart { get; set; }

        /// <summary>
        /// Gets the 1-based index of excel row where this entity ends in the excel worksheet
        /// </summary>
        public override int ExcelRowEnd { get; set; }

        /// <summary>
        /// Gets the 1-based index of excel column where this entity starts in the excel worksheet
        /// </summary>
        public override int ExcelColumnStart { get; set; }

        /// <summary>
        /// Gets the 1-based index of excel column where this entity ends in the excel worksheet
        /// </summary>
        public override int ExcelColumnEnd { get; set; }

        /// <summary>
        /// Returns a string represenation of this object instance
        /// </summary>
        /// <returns>A string representation of this object instance.</returns>
        public override string ToString()
        {
            return string.Format("ExcelMapCoOrdinateContainer[{0},Id={1}]:CurrentCell=R{2}:C{3},RowCols={4},{5},WorksheetRowCol=R{6}C{7}",
                                     this.containerType,
                                     this.Id,
                                     this.currentRowIndex,
                                     this.currentColumnIndex,
                                     this.mapRowCount,
                                     this.mapColumnCount,
                                     this.ExcelRowStart,
                                     this.ExcelColumnStart);
        }

        /// <summary>
        /// Returns the cell at an explicit column and row index.<br/>
        /// This takes into consideration that there may be nested elements.
        /// </summary>
        /// <param name="columnIndex"></param>
        /// <param name="rowIndex"></param>
        /// <returns></returns>
        public ExcelMapCoOrdinateCell GetCell(uint columnIndex, uint rowIndex)
        {
            var cellListKey = new System.Drawing.Point((int)columnIndex, (int)rowIndex);
            if (this.cells.ContainsKey(cellListKey))
            {
                ExcelMapCoOrdinate element = this.cells[cellListKey];
            }
            return null;
        }

        /// <summary>
        /// Builds a model of the rows within this entity.
        /// </summary>
        internal override RowOrColumnsModel BuildRowsModel()
        {
            var rowsModel = new RowOrColumnsModel(true);

            // Consider all elements in the container on a column-by-column basis.
            for (uint colIdx = 1; colIdx <= this.mapColumnCount; colIdx++)
            {
                // Create a RowsModel for the column
                var columnRowsModel = new RowOrColumnsModel(true);

                for (uint rowIdx = 1; rowIdx <= this.mapRowCount; rowIdx++)
                {
                    // Use System.Drawing.Point as key as it has a very efficient hashing algorithm
                    var cellListKey = new System.Drawing.Point((int)colIdx, (int)rowIdx);
                    if (this.cells.ContainsKey(cellListKey))
                    {
                        ExcelMapCoOrdinate element = this.cells[cellListKey];
                        RowOrColumnsModel elementRowsModel = element.BuildRowsModel();

                        columnRowsModel.AppendModel(elementRowsModel);
                    }
                }

                // Merge the columns model generated for the row with the model for the container.
                rowsModel.MergeModel(columnRowsModel);

                // Add this container into the column (if not already there)
                RowOrColumnInfo rowInfo = rowsModel.First;
                while (rowInfo != null)
                {
                    rowInfo.AddMap(this);
                    rowInfo = rowInfo.Next;
                }
            }

            return rowsModel;
        }

        /// <summary>
        /// Builds a model of the columns within this entity.
        /// </summary>
        internal override RowOrColumnsModel BuildColumnsModel()
        {
            var columnsModel = new RowOrColumnsModel(false);

            // Consider all elements in the container on a row-by-row basis.
            for (uint rowIdx = 1; rowIdx <= this.mapRowCount; rowIdx++)
            {
                // Create a ColumnsModel for the row
                var rowColumnsModel = new RowOrColumnsModel(false);

                for (uint colIdx = 1; colIdx <= this.mapColumnCount; colIdx++)
                {
                    // Use System.Drawing.Point as key as it has a very efficient hashing algorithm
                    var cellListKey = new System.Drawing.Point((int)colIdx, (int)rowIdx);
                    if (this.cells.ContainsKey(cellListKey))
                    {
                        ExcelMapCoOrdinate element = this.cells[cellListKey];
                        RowOrColumnsModel elementColumnsModel = element.BuildColumnsModel();
                        rowColumnsModel.AppendModel(elementColumnsModel);
                    }
                }

                // Merge the columns model generated for the row with the model for the container.
                columnsModel.MergeModel(rowColumnsModel);

                // Add this container into the column (if not already there)
                RowOrColumnInfo columnInfo = columnsModel.First;
                while (columnInfo != null)
                {
                    columnInfo.AddMap(this);
                    columnInfo = columnInfo.Next;
                }
            }

            return columnsModel;
        }

        /// <summary>
        /// Adds all the entites that this element holds, plus this element itself,<br/>
        /// to the supplied <see cref="LayeredCellsDictionary">cellsDictionary</see><br/>
        /// for all worksheet row/column index postions that this entity covers.
        /// </summary>
        /// <param name="cellsDictionary">The <see cref="LayeredCellsDictionary"/> to be updated</param>
        internal override void UpdateLayeredCells(ref LayeredCellsDictionary cellsDictionary)
        {
            // Consider each element in maps
            for (uint colIdx = 1; colIdx <= this.mapColumnCount; colIdx++)
            {
                for (uint rowIdx = 1; rowIdx <= this.mapRowCount; rowIdx++)
                {
                    // Use System.Drawing.Point as key as it has a very efficient hashing algorithm
                    var cellListKey = new System.Drawing.Point((int)colIdx, (int)rowIdx);
                    if (this.cells.ContainsKey(cellListKey))
                    {
                        ExcelMapCoOrdinate element = this.cells[cellListKey];
                        element.UpdateLayeredCells(ref cellsDictionary);
                    }
                }
            }

            // Add this as a layer
            base.UpdateLayeredCells(ref cellsDictionary);
        }

        #endregion Overrides
    }
}
