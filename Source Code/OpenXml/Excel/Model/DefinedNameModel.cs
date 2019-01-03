namespace ExcelWriter.OpenXml.Excel.Model
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using DocumentFormat.OpenXml.Spreadsheet;

    /// <summary>
    /// Encapsulates information about an DefinedName (range) in Excel.
    /// </summary>
    public class DefinedNameModel
    {
        /// <summary>
        /// Returns an instance of a <see cref="DefinedNameModel"/> for a defined name (named range) with in a specified workbook.<br/>
        /// TODO: Does not take defined name scope into consideration yet. Investigate Worksheet scope defined names.
        /// </summary>
        /// <param name="wb"></param>
        /// <param name="definedName"></param>
        /// <returns></returns>
        public static DefinedNameModel GetDefinedNameModel(Workbook wb, string definedName)
        {
            if (wb == null) throw new ArgumentNullException("wb");
            if (string.IsNullOrEmpty(definedName)) throw new ArgumentNullException("definedName");

            DefinedNameModel definedNameModel = null;
            DefinedName dn = wb.GetDefinedNameByName(definedName);
            if (dn != null)
            {
                definedNameModel = new DefinedNameModel(wb, dn);
                definedNameModel.Name = definedName;
            }
            return definedNameModel;
        }


        #region Constructor

        /// <summary>
        /// Private ctor. Prevents public construction.<br/>
        /// Loads the model from the supplied <see cref="DefinedName"/>
        /// </summary>
        /// <param name="wb">The workbool</param>
        /// <param name="definedName">The named range</param>
        private DefinedNameModel(Workbook wb, DefinedName definedName)
        {
            this.IsDefined = wb.BreakDownDefinedName(definedName, ref worksheetName, ref rowStart, ref rowEnd, ref colStart, ref colEnd);

            if (this.IsDefined)
            {
                var wsPart = wb.GetWorksheetPartByName(worksheetName);
                this.Worksheet = wsPart.Worksheet;
                this.SheetData = this.Worksheet.GetFirstChild<SheetData>();

                //get cells to be cloned according to the specified rows and columns
                this.Cells = this.SheetData.Descendants<DocumentFormat.OpenXml.Spreadsheet.Cell>().Where(c =>
                                                CellExtensions.GetRowIndex(c.CellReference) >= rowStart &&
                                                CellExtensions.GetRowIndex(c.CellReference) <= rowEnd &&
                                                CellExtensions.GetColumnIndex(c.CellReference) >= colStart &&
                                                CellExtensions.GetColumnIndex(c.CellReference) <= colEnd)
                                                .ToList<DocumentFormat.OpenXml.Spreadsheet.Cell>();
            }
        }

        #endregion

        private bool isDefined;
        private string name;

        private uint colStart;
        private uint colEnd;

        private uint rowStart;
        private uint rowEnd;

        private string worksheetName;
        private IList<DocumentFormat.OpenXml.Spreadsheet.Cell> cells;
        private Worksheet worksheet;
        private SheetData sheetData;

        public bool IsDefined
        {
            get { return isDefined; }
            private set { isDefined = value; }
        }

        public string Name
        {
            get { return this.name; }
            private set { this.name = value; }
        }

        public uint ColStart
        {
            get { return this.colStart; }
            private set { this.colStart = value; }
        }

        public uint ColEnd
        {
            get { return this.colEnd; }
            private set { this.colEnd = value; }
        }

        public int ColCount
        {
            get { return (int)(this.colEnd + 1 - this.colStart); }
        }

        public uint RowStart
        {
            get { return this.rowStart; }
            private set { this.rowStart = value; }
        }

        public uint RowEnd
        {
            get { return this.rowEnd; }
            private set { this.rowEnd = value; }
        }

        public int RowCount
        {
            get { return (int)(this.rowEnd + 1 - this.rowStart); }
        }

        public IList<DocumentFormat.OpenXml.Spreadsheet.Cell> Cells
        {
            get { return this.cells; }
            private set { this.cells = value; }
        }

        public string WorksheetName
        {
            get { return this.worksheetName; }
            private set { this.worksheetName = value; }
        }

        public Worksheet Worksheet
        {
            get { return this.worksheet; }
            private set { this.worksheet = value; }
        }

        public SheetData SheetData
        {
            get { return this.sheetData; }
            private set { this.sheetData = value; }
        }

        /// <summary>
        /// Get a cell within the range given a 0 based row index and 0 based column index within that range.
        /// </summary>
        /// <param name="row"></param>
        /// <param name="col"></param>
        /// <returns></returns>
        public Cell GetCell(uint col, uint row)
        {
            string cellReference = CellExtensions.GetCellReference((this.ColStart + col), (this.RowStart + row));
            return this.Cells.FirstOrDefault(c => c.CellReference == cellReference);
        }

        /// <summary>
        /// Sets the value of a cell within a range given a 0 based row index and 0 based column index within that range. 
        /// </summary>
        /// <param name="rowIdx"></param>
        /// <param name="colIdx"></param>
        /// <param name="value"></param>
        public void SetCell(uint col, uint row, object value)
        {
            this.SheetData.SetCell((col - this.ColStart), (row - this.RowStart), value);
        }

        private static Row GetRow(SheetData sheetData, uint rowIndex)
        {
            return sheetData.OfType<Row>().FirstOrDefault(r => r.RowIndex == rowIndex);
        }
    }
}
