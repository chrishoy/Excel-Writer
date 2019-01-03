using System;
namespace ExcelWriter
{
    /// <summary>
    /// Represents information about a row or a column which is to be exported to Excel.
    /// </summary>
    internal class RowOrColumnInfo : IStorable
    {
        #region Private Fields

        private MapStore maps = new MapStore();
        private int excelIndex;

        #endregion Private Fields

        #region Construction

        /// <summary>
        /// Initializes a new instance of the <see cref="RowOrColumnInfo" /> class.
        /// </summary>
        /// <param name="isRow">True if this information related to a row, otherwise relates to a column</param>
        private RowOrColumnInfo(bool isRow)
        {
            this.IsRow = isRow;
            this.Id = Counter.GetNextId();
        }

        #endregion Construction

        #region Public Properties

        /// <summary>
        /// Gets the 1-based index of the row or column in Excel that this <see cref="RowOrColumnInfo"/> represents
        /// </summary>
        public int ExcelIndex
        {
            get
            {
                if (this.excelIndex <= 0)
                {
                    throw new InvalidOperationException("At attempt has been made to read an index that has not been set to +'ve value.");
                }
                else
                {
                    return this.excelIndex;
                }
            }
            set
            {
                if (value <= 0)
                {
                    throw new InvalidOperationException("An attempt has been made to set the Excel Row or Column index to a non +'ve value");
                }
                this.excelIndex = value;
            }
        }

        /// <summary>
        /// Gets a value indicating whether this information relates to a row (or a column)
        /// </summary>
        public bool IsRow { get; private set; }

        /// <summary>
        /// Gets a value which uniquely identifies this instance of a <see cref="RowOrColumnInfo"/> class.
        /// </summary>
        public int Id { get; private set; }

        /// <summary>
        /// Gets or sets the width of the column when exported to Excel.
        /// </summary>
        public double? HeightOrWidth { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether this column is hidden when exported to Excel.
        /// </summary>
        public bool Hidden { get; set; }

        /// <summary>
        /// Gets or sets a pointer to the next <see cref="RowOrColumnInfo"/>
        /// </summary>
        public RowOrColumnInfo Next { get; set; }

        /// <summary>
        /// Gets a <see cref="MapStore"/> which contains a set of <see cref="ExcelMapCoOrdinate">maps</see> associated with this column.
        /// </summary>
        public MapStore Maps
        {
            get { return this.maps; }
        }

        #endregion Public Properties

        #region Public Methods

        /// <summary>
        /// Creates and returns a new instance of the <see cref="RowOrColumnInfo" /> class.
        /// </summary>
        /// <param name="isRow">If true, the <see cref="RowOrColumnInfo" /> created relates a row, otherwise a column.</param>
        /// <returns>A new instance of a <see cref="RowOrColumnInfo"/> or </returns>
        public static RowOrColumnInfo Create(bool isRow)
        {
            return new RowOrColumnInfo(isRow);
        }

        /// <summary>
        /// Creates and returns a new instance of the <see cref="RowOrColumnInfo" /> class.
        /// </summary>
        /// <param name="map">The <see cref="ExcelMapCoOrdinate"/> which is the source for this column</param>
        /// <param name="isRow">If true, the <see cref="RowOrColumnInfo" /> created relates a row, otherwise a column.</param>
        /// <returns>A new instance of a <see cref="RowOrColumnInfo"/> class</returns>
        public static RowOrColumnInfo Create(ExcelMapCoOrdinate map, bool isRow)
        {
            var info = new RowOrColumnInfo(isRow);

            if (isRow)
            {
                info.HeightOrWidth = map.AssignedHeight;
                info.Hidden = map.RowIsHidden;
            }
            else
            {
                info.HeightOrWidth = map.AssignedWidth;
                info.Hidden = map.ColumnIsHidden;
            }

            info.AddMap(map);

            return info;
        }

        /// <summary>
        /// Adds a map to this <see cref="RowOrColumnInfo"/>
        /// </summary>
        /// <param name="map">An <see cref="ExcelMapCoOrdinate"/> to be associated with this column</param>
        public void AddMap(ExcelMapCoOrdinate map)
        {
            // Adds map to the column (if not already there)
            this.Maps.AddDistinct(map);

            // Add the column to the map
            if (this.IsRow)
            {
                map.Rows.AddDistinct(this);
            }
            else
            {
                map.Columns.AddDistinct(this);
            }
        }

        /// <summary>
        /// Adds all of the <see cref="ExcelMapCoOrdinate">maps</see> associated with a <see cref="RowOrColumnInfo"/>
        /// </summary>
        /// <param name="info">A <see cref="RowOrColumnInfo"/></param>
        public void AddMaps(RowOrColumnInfo info)
        {
            // Add RowOrColumnInfo maps to the column (if not already there)
            this.maps.AddDistinct(info.Maps);

            if (this.IsRow)
            {
                // Add the column to the maps
                foreach (ExcelMapCoOrdinate map in info.Maps)
                {
                    map.Rows.AddDistinct(this);
                }
            }
            else
            {
                // Add the column to the maps
                foreach (ExcelMapCoOrdinate map in info.Maps)
                {
                    map.Columns.AddDistinct(this);
                }
            }
        }

        /// <summary>
        /// Moves all of the <see cref="ExcelMapCoOrdinate">maps</see> associated with a 
        /// <see cref="RowOrColumnInfo"/> to this <see cref="RowOrColumnInfo"/>.
        /// </summary>
        /// <param name="info">A <see cref="RowOrColumnInfo"/></param>
        public void MoveMaps(RowOrColumnInfo info)
        {
            // Add RowOrColumnInfo maps to the column (if not already there)
            this.maps.AddDistinct(info.Maps);

            if (this.IsRow)
            {
                // Add the column to the maps
                foreach (ExcelMapCoOrdinate map in info.Maps)
                {
                    map.Rows.AddDistinct(this);
                    map.Rows.Remove(info);
                }
            }
            else
            {
                // Add the column to the maps
                foreach (ExcelMapCoOrdinate map in info.Maps)
                {
                    map.Columns.AddDistinct(this);
                    map.Columns.Remove(info);
                }
            }
        }

        /// <summary>
        /// Gets a string representation of this object instance.
        /// </summary>
        /// <returns>A string representation of this object instance</returns>
        public override string ToString()
        {
            if (this.IsRow)
            {
                return string.Format("RowInfo[Id={0}]:Height={1},Hidden={2}", this.Id, this.HeightOrWidth, this.Hidden);
            }
            else
            {
                return string.Format("ColumnInfo[Id={0}]:Width={1},Hidden={2}", this.Id, this.HeightOrWidth, this.Hidden);
            }
        }

        #endregion Public Methods
    }
}
