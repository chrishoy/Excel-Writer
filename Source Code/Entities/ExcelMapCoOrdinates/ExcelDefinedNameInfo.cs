using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelWriter
{
    /// <summary>
    /// Maintains information about a defined name that is to be written into an excel worksheet
    /// </summary>
    internal class ExcelDefinedNameInfo
    {
        #region Private Fields

        private string sheetName;
        private string definedName;

        private uint startRowIndex;
        private uint endRowIndex;

        private uint startColumnIndex;
        private uint endColumnIndex;

        #endregion Private Fields

        #region Construction

        #endregion Construction

        #region Public Properties

        public string SheetName
        {
            get { return this.sheetName; }
            set { this.sheetName = value; }
        }

        public string DefinedName
        {
            get { return this.definedName; }
            set { this.definedName = value; }
        }

        public uint StartRowIndex
        {
            get { return this.startRowIndex; }
            set { this.startRowIndex = value; }
        }

        public uint EndRowIndex
        {
            get { return this.endRowIndex; }
            set { this.endRowIndex = value; }
        }

        //public uint RowIndex
        //{
        //    get { return this.rowIndex; }
        //    set { this.rowIndex = value; }
        //}

        //public uint RowCount
        //{
        //    get { return this.rowCount; }
        //    set { this.rowCount = value; }
        //}

        public uint StartColumnIndex
        {
            get { return this.startColumnIndex; }
            set { this.startColumnIndex = value; }
        }

        public uint EndColumnIndex
        {
            get { return this.endColumnIndex; }
            set { this.endColumnIndex = value; }
        }

        #endregion Public Properties

        #region Public Methods

        ///// <summary>
        ///// Updates a supplied list with defined name information if there is any.
        ///// </summary>
        ///// <param name="mapCoOrdinateContainer"></param>
        ///// <param name="sheetName"></param>
        ///// <param name="values"></param>
        //public static void UpdateList(ExcelMapCoOrdinateContainer mapCoOrdinateContainer, string sheetName, ref List<ExcelDefinedNameInfo> values)
        //{
        //    if (values == null) throw new ArgumentNullException("values");
        //    if (mapCoOrdinateContainer == null) return;

        //    if (!string.IsNullOrEmpty(mapCoOrdinateContainer.DefinedName))
        //    {
        //        var info = new ExcelDefinedNameInfo
        //                        {
        //                            DefinedName = mapCoOrdinateContainer.DefinedName,
        //                            RowIndex = mapCoOrdinateContainer.StartRowIndex,
        //                            RowCount = mapCoOrdinateContainer.TotalRowCount,
        //                            ColumnIndex = mapCoOrdinateContainer.StartColumnIndex,
        //                            ColumnCount = mapCoOrdinateContainer.TotalColumnCount,
        //                            SheetName = sheetName,
        //                        };

        //        values.Add(info);
        //    }
        //}

        #endregion Public Methods
    }
}
