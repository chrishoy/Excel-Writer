namespace ExcelWriter
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text;

    /// <summary>
    /// Represents a placeholder for a set of cells (entity) that is to be written into an Excel worksheet.
    /// </summary>
    internal abstract class ExcelMapCoOrdinate : IStorable
    {
        #region Private fields

        private ExcelMapCoOrdinate parent;
        private List<StyleBase> styles;
        private List<KeyValuePair<string, BaseMap>> keyedElements;

        private RowOrColumnInfoStore rowStore = new RowOrColumnInfoStore();
        private RowOrColumnInfoStore columnStore = new RowOrColumnInfoStore();

        #endregion Private fields

        #region Construction

        /// <summary>
        /// Prevents a default instance of the <see cref="ColumnInfo" /> class from being created.
        /// </summary>
        protected ExcelMapCoOrdinate()
        {
            this.Id = Counter.GetNextId();
        }

        #endregion Construction

        #region Public Properties

        /// <summary>
        /// Uniquely identifies this instance of a <see cref="ExcelMapCoOrdinate"/> derrived class.
        /// </summary>
        public int Id { get; private set; }

        /// <summary>
        /// Gets or sets the width assigned to the co-ordinate.<br/>
        /// If no widths assinged then this will be null;
        /// </summary>
        public double? AssignedWidth { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether this <see cref="ExcelMapCoOrdinate"/> should be column-wise hidden.
        /// </summary>
        public bool ColumnIsHidden { get; set; }

        /// <summary>
        /// Gets or sets the height assigned to the co-ordinate.<br/>
        /// If no heights assinged then this will be null;
        /// </summary>
        public double? AssignedHeight { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether this <see cref="ExcelMapCoOrdinate"/> should be row-wise hidden.
        /// </summary>
        public bool RowIsHidden { get; set; }

        /// <summary>
        /// Gets the parent which contains this element.
        /// </summary>
        public ExcelMapCoOrdinate Parent
        {
            get { return this.parent; }
        }

        /// <summary>
        /// Gets the number of rows for this entity.
        /// </summary>
        public abstract uint MapRowCount { get; }

        /// <summary>
        /// Gets the number of column maps within this entity
        /// </summary>
        public abstract uint MapColumnCount { get; }

        /// <summary>
        /// Gets the 1-based index of excel row where this entity starts in the excel worksheet
        /// </summary>
        public abstract int ExcelRowStart { get; set; }

        /// <summary>
        /// Gets the 1-based index of excel row where this entity ends in the excel worksheet
        /// </summary>
        public abstract int ExcelRowEnd { get; set; }

        /// <summary>
        /// Gets the 1-based index of excel column where this entity starts in the excel worksheet
        /// </summary>
        public abstract int ExcelColumnStart { get; set; }

        /// <summary>
        /// Gets the 1-based index of excel column where this entity ends in the excel worksheet
        /// </summary>
        public abstract int ExcelColumnEnd { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether the last column in this entity will be spanned to the last column of the container element
        /// when exported to an Excel worksheet.
        /// </summary>
        public bool SpanLastColumn { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether the last row in this entity will be spanned to the last row of the container element
        /// when exported to an Excel worksheet.
        /// </summary>
        public bool SpanLastRow { get; set; }

        /// <summary>
        /// Gets or sets a defined name for the content of this cell/container.<br/>
        /// This maps to range names in Excel.
        /// </summary>
        public string DefinedName { get; set; }

        /// <summary>
        /// Gets a list of styles associated with this map
        /// </summary>
        public IEnumerable<StyleBase> Styles
        {
            get { return this.styles; }
        }

        #endregion Public Properties

        #region Internal Properties

        /// <summary>
        /// Gets a list of <see cref="ColumnInfo">Column</see>s.
        /// </summary>
        internal RowOrColumnInfoStore Columns 
        {
            get { return this.columnStore; }
        }

        /// <summary>
        /// Gets a list of <see cref="RowInfo">Row</see>s.
        /// </summary>
        internal RowOrColumnInfoStore Rows
        {
            get { return this.rowStore; }
        }

        /// <summary>
        /// Gets or sets the <see cref="BaseMap"/> derived elements that have been referenced by key.
        /// </summary>
        internal List<KeyValuePair<string, BaseMap>> KeyedElements
        {
            get { return this.keyedElements; }
            set { this.keyedElements = value; }
        }

        #endregion Internal Properties

        #region Public Methods

        /// <summary>
        /// Updates a supplied list with defined name information if there is any.
        /// </summary>
        /// <param name="mapCoOrdinate">A <see cref="ExcelMapCoOrdinate"/></param>
        /// <param name="sheetName">A worksheet name</param>
        /// <param name="values">The <see cref="List<ExcelDefinedNameInfo>"/> to be updated</param>
        internal static void UpdateDefinedNameList(ExcelMapCoOrdinate mapCoOrdinate, string sheetName, ref List<ExcelDefinedNameInfo> values)
        {
            if (values == null)
            {
                throw new ArgumentNullException("values");
            }

            if (mapCoOrdinate == null)
            {
                return;
            }

            if (!string.IsNullOrEmpty(mapCoOrdinate.DefinedName))
            {
                var info = new ExcelDefinedNameInfo
                {
                    DefinedName = mapCoOrdinate.DefinedName,
                    StartRowIndex = (uint)mapCoOrdinate.ExcelRowStart,
                    EndRowIndex = (uint)mapCoOrdinate.ExcelRowEnd,
                    StartColumnIndex = (uint)mapCoOrdinate.ExcelColumnStart,
                    EndColumnIndex = (uint)mapCoOrdinate.ExcelColumnEnd,
                    SheetName = sheetName,
                };

                values.Add(info);
            }
        }

        /// <summary>
        /// Gets the excel worksheet end column index for this entity by first checking if the entity is marked to span
        /// to the bounds of the parent container the bounds, and if no spanning is reqruied, the bounds of this entity.
        /// </summary>
        /// <returns>The end worksheet column index</returns>
        internal uint GetEndColumnIndex()
        {
            if (this.SpanLastColumn && this.Parent != null)
            {
                return this.Parent.GetEndColumnIndex();
            }

            return (uint)this.ExcelColumnEnd;
        }

        /// <summary>
        /// Gets the excel worksheet end row index for this entity by first checking if the entity is marked to span
        /// to the bounds of the parent container the bounds, and if no spanning is reqruied, the bounds of this entity.
        /// </summary>
        /// <returns>The end worksheet row index</returns>
        internal uint GetEndRowIndex()
        {
            if (this.SpanLastRow && this.Parent != null)
            {
                return this.Parent.GetEndRowIndex();
            }

            return (uint)this.ExcelRowEnd;
        }

        /// <summary>
        /// Gets the <see cref="ExcelMapCoOrdinate"/> which is at the highest level in the tree from this position.<br/>
        /// This will be the top-most element in the tree (ie. the container which represents the worksheet itself).
        /// </summary>
        /// <returns></returns>
        internal ExcelMapCoOrdinate GetRoot()
        {
            // The root element is that element that has no parent
            if (this.Parent == null)
            {
                return this;
            }
            else
            {
                // Recurse up the tree
                return this.Parent.GetRoot();
            }
        }

        #endregion Public Methods

        #region Internal Methods

        /// <summary>
        /// Builds a model of the columns contained within this entity.
        /// </summary>
        /// <returns></returns>
        internal abstract RowOrColumnsModel BuildColumnsModel();

        /// <summary>
        /// Builds a model of the rows contained within this entity.
        /// </summary>
        /// <returns></returns>
        internal abstract RowOrColumnsModel BuildRowsModel();

        /// <summary>
        /// Adds a style to an internal list of styles
        /// </summary>
        /// <param name="style">A <see cref="StyleBase"/> derived style</param>
        internal void AddStyle(StyleBase style)
        {
            if (style != null)
            {
                if (this.styles == null)
                {
                    this.styles = new List<StyleBase>();
                }
                this.styles.Add(style);
            }
        }

        /// <summary>
        /// Sets the parent of this <see cref="ExcelMapCoOrdinate"/>
        /// </summary>
        /// <param name="parent">The parent/container <see cref="ExcelMapCoOrdinate"/></param>
        internal void SetParent(ExcelMapCoOrdinate parent)
        {
            if (parent == null) throw new ArgumentNullException("parent");
            this.parent = parent;
        }

        /// <summary>
        /// Adds this element to the <see cref="LayeredCellsDictionary">cellsDictionary</see>
        /// for all worksheet row/column index postions that this entity covers.
        /// </summary>
        /// <param name="cellsDictionary">The <see cref="LayeredCellsDictionary"/> to be updated</param>
        internal virtual void UpdateLayeredCells(ref LayeredCellsDictionary cellsDictionary)
        {
            // Update the ExcelRowsColumnsDictionary
            uint endColumnIndex = this.GetEndColumnIndex();
            uint endRowIndex = this.GetEndRowIndex();

            for (uint colIdx = (uint)this.ExcelColumnStart; colIdx <= endColumnIndex; colIdx++)
            {
                for (uint rowIdx = (uint)this.ExcelRowStart; rowIdx <= endRowIndex; rowIdx++)
                {
                    cellsDictionary.Upsert(new System.Drawing.Point((int)colIdx, (int)rowIdx), this);
                }
            }
        }

        #endregion Internal Methods
    }
}
