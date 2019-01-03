namespace ExcelWriter.OpenXml.Excel.Model
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text;

    using DocumentFormat.OpenXml;
    using DocumentFormat.OpenXml.Packaging;
    using DocumentFormat.OpenXml.Spreadsheet;

    using Drawing = DocumentFormat.OpenXml.Drawing;
    using DrawingCharts = DocumentFormat.OpenXml.Drawing.Charts;
    using DrawingSpreadsheet = DocumentFormat.OpenXml.Drawing.Spreadsheet;

    /// <summary>
    /// Base class which encapsulates information about a model object with an Excel worksheet.
    /// </summary>
    public abstract class ModelBase
    {
        #region Private Fields

        private Worksheet worksheet;
        private DrawingSpreadsheet.TwoCellAnchor anchor;
        private ExcelPositionalInfo positionalInfo;

        #endregion Private Fields

        #region Construction

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="anchor"></param>
        public ModelBase(DrawingSpreadsheet.TwoCellAnchor anchor)
        {
            this.Anchor = anchor;
            this.PositionalInfo = new ExcelPositionalInfo();

            this.PositionalInfo.From.Row.Index = uint.Parse(anchor.FromMarker.RowId.Text);
            this.PositionalInfo.From.Row.OffsetInEmus = uint.Parse(anchor.FromMarker.RowOffset.Text);
            this.PositionalInfo.From.Column.Index = uint.Parse(anchor.FromMarker.ColumnId.Text);
            this.PositionalInfo.From.Column.OffsetInEmus = uint.Parse(anchor.FromMarker.ColumnOffset.Text);

            this.PositionalInfo.To.Row.Index = uint.Parse(anchor.ToMarker.RowId.Text);
            this.PositionalInfo.To.Row.OffsetInEmus = uint.Parse(anchor.ToMarker.RowOffset.Text);
            this.PositionalInfo.To.Column.Index = uint.Parse(anchor.ToMarker.ColumnId.Text);
            this.PositionalInfo.To.Column.OffsetInEmus = uint.Parse(anchor.ToMarker.ColumnOffset.Text);
        }

        #endregion Construction

        #region Public Properties

        /// <summary>
        /// Gets the <see cref="Worksheet"/> on which the chart resides
        /// </summary>
        public Worksheet Worksheet
        {
            get { return this.worksheet; }
            protected set { this.worksheet = value; }
        }

        /// <summary>
        /// Gets the <see cref="DrawingSpreadsheet.TwoCellAnchor"/> under which this chart resides
        /// </summary>
        protected DrawingSpreadsheet.TwoCellAnchor Anchor
        {
            get { return this.anchor; }
            set { this.anchor = value; }
        }

        /// <summary>
        /// Gets or sets information about model placement within the host worksheet.
        /// </summary>
        public ExcelPositionalInfo PositionalInfo
        {
            get { return positionalInfo; }
            protected set { positionalInfo = value; }
        }

        /// <summary>
        /// Gets or sets the width of the model in English Metric Units
        /// </summary>
        public long WidthInEmus { get; protected set; }

        /// <summary>
        /// Gets or sets the width of the model in English Metric Units
        /// </summary>
        public long HeightInEmus { get; protected set; }

        #endregion Public Properties

        #region Public Methods

        /// <summary>
        /// Moves the picture into position within its worksheet.
        /// </summary>
        /// <param name="positionalInfo">A <see cref="ExcelPositionalInfo"/></param>
        public void SizeAndMove(ExcelPositionalInfo positionalInfo)
        {
            // From Marker
            this.Anchor.FromMarker.RowId.Text = positionalInfo.From.Row.Index.ToString();
            this.Anchor.FromMarker.RowOffset.Text = positionalInfo.From.Row.OffsetInEmus.ToString();
            this.Anchor.FromMarker.ColumnId.Text = positionalInfo.From.Column.Index.ToString();
            this.Anchor.FromMarker.ColumnOffset.Text = positionalInfo.From.Column.OffsetInEmus.ToString();

            // To Marker
            this.Anchor.ToMarker.RowId.Text = positionalInfo.To.Row.Index.ToString();
            this.Anchor.ToMarker.RowOffset.Text = positionalInfo.To.Row.OffsetInEmus.ToString();
            this.Anchor.ToMarker.ColumnId.Text = positionalInfo.To.Column.Index.ToString();
            this.Anchor.ToMarker.ColumnOffset.Text = positionalInfo.To.Column.OffsetInEmus.ToString(); ;
        }

        #endregion Public Methods
    }
}