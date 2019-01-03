namespace ExcelWriter.OpenXml.Excel.Model
{
    using System;
    using System.Collections.Generic;
    using System.Linq;

    using OpenXmlPackaging = DocumentFormat.OpenXml.Packaging;
    using DocumentFormat.OpenXml.Spreadsheet;
    using DrawingSpreadsheet = DocumentFormat.OpenXml.Drawing.Spreadsheet;
    
    /// <summary>
    /// Represents an entity which encapsulates information about an Shape (or set of shaped) in an Excel worksheet.
    /// </summary>
    public class ShapeModel : ModelBase
    {
        #region Private Fields

        private bool isValid;
        private string name;

        private DrawingSpreadsheet.Shape shape;
        private OpenXmlPackaging.DrawingsPart drawingsPart;

        #endregion Private Fields

        #region Construction

        /// <summary>
        /// Initializes a new instance of the <see cref="ShapeModel" /> class.<br/>
        /// Private ctor. Prevents public construction.
        /// </summary>
        /// <param name="ws">The <see cref="Worksheet"/></param>
        /// <param name="anchor">A <see cref="DrawingSpreadsheet.TwoCellAnchor"/></param>
        private ShapeModel(Worksheet ws, DrawingSpreadsheet.TwoCellAnchor anchor)
            : base(anchor)
        {
            if (ws == null) throw new ArgumentNullException("ws");
            if (anchor == null) throw new ArgumentNullException("anchor");

            this.Worksheet = ws;
            this.drawingsPart = ws.WorksheetPart.DrawingsPart;
            this.Shape = anchor.Descendants<DrawingSpreadsheet.Shape>().FirstOrDefault();

            // Shape properties and extents....
            if (this.Shape != null)
            {
                DocumentFormat.OpenXml.Drawing.Extents extents = this.Shape.Descendants<DocumentFormat.OpenXml.Drawing.Extents>().FirstOrDefault();
                if (extents != null)
                {
                    if (extents.Cx.HasValue)
                    {
                        this.WidthInEmus = extents.Cx.Value;
                    }

                    if (extents.Cy.HasValue)
                    {
                        this.HeightInEmus = extents.Cy.Value;
                    }
                }
            }

            // Get the name of the Shape
            this.Name = GetShapeName(this.Shape);

            this.IsValid = true;
        }

        #endregion Construction

        #region Public Properties

        /// <summary>
        /// Gets a value indicating whether more than one shape exists on the worksheet with the same name.
        /// </summary>
        public bool HasMoreThanOneInstance { get; private set; }

        /// <summary>
        /// Gets a value indicating whether the model is valid with the context of the worksheet.
        /// </summary>
        public bool IsValid
        {
            get { return this.isValid; }
            private set { this.isValid = value; }
        }

        /// <summary>
        /// Gets the Name of the shape
        /// </summary>
        public string Name
        {
            get { return this.name; }
            private set { this.name = value; }
        }

        /// <summary>
        /// Gets the shape
        /// </summary>
        public DrawingSpreadsheet.Shape Shape 
        {
            get { return this.shape; }
            private set { this.shape = value; }
        }

        /// <summary>
        /// Gets the <see cref="Worksheet"/> on which the shape resides.
        /// </summary>
        public DrawingSpreadsheet.WorksheetDrawing Drawing
        {
            get
            {
                return this.drawingsPart == null ? null : this.drawingsPart.WorksheetDrawing; 
            }
        }

        #endregion Public Properties

        #region Public Methods

        /// <summary>
        /// Returns an instance of a <see cref="ShapeModel"/> for the first shape with a specified name in a worksheet.<br/>
        /// </summary>
        /// <param name="ws">The <see cref="Worksheet"/> in which the chart resides</param>
        /// <param name="id">The shape id (name in sheet)</param>
        /// <returns>A model which represents the first shape which has a specified name in the worksheet.</returns>
        public static ShapeModel GetShapeModel(Worksheet ws, string id)
        {
            Guard.IsNotNull(ws, "ws");
            Guard.IsNotNullOrEmpty(id, "id");

            ShapeModel model = null;

            // Get all drawing shapes in the worksheet with a specified name
            IEnumerable<DrawingSpreadsheet.Shape> shapes = GetShapesWithName(ws.WorksheetPart, id);

            int countOfShapes = shapes.Count();
            if (countOfShapes > 0)
            {
                DrawingSpreadsheet.TwoCellAnchor anchor = shapes.First().Ancestors<DrawingSpreadsheet.TwoCellAnchor>().FirstOrDefault();
                model = new ShapeModel(ws, anchor);
                model.HasMoreThanOneInstance = countOfShapes > 1;
            }

            return model;
        }

        /// <summary>
        /// Returns an instance of a <see cref="ShapeModel"/> for the first shape with a specified name in a worksheet.<br/>
        /// </summary>
        /// <param name="wsPart">The <see cref="OpenXmlPackaging.WorksheetPart"/> in which the chart resides</param>
        /// <param name="id">The shape id (name in sheet)</param>
        /// <returns>A model which represents the first shape which has a specified name in the worksheet.</returns>
        [Obsolete("This method is to be removed. Use 'ShapeModel GetShapeModel(Worksheet ws, string id)' instead")]
        public static ShapeModel GetShapeModel(OpenXmlPackaging.WorksheetPart wsPart, string id)
        {
            Guard.IsNotNull(wsPart, "wsPart");
            Guard.IsNotNullOrEmpty(id, "id");

            ShapeModel model = null;

            // Get all drawing shapes in the worksheet with a specified name
            IEnumerable<DrawingSpreadsheet.Shape> shapes = GetShapesWithName(wsPart, id);

            int countOfShapes = shapes.Count();
            if (countOfShapes > 0)
            {
                DrawingSpreadsheet.TwoCellAnchor anchor = shapes.First().Ancestors<DrawingSpreadsheet.TwoCellAnchor>().FirstOrDefault();
                model = new ShapeModel(wsPart.Worksheet, anchor);
                model.HasMoreThanOneInstance = countOfShapes > 1;
            }

            return model;
        }

        /// <summary>
        /// Creates a deep copy of this <see cref="ShapeModel"/> and a new shape in the worksheet.
        /// </summary>
        /// <returns>The <see cref="ShapeModel"/> that represents the shape</returns>
        public ShapeModel Clone()
        {
            return this.Clone(this.Worksheet);
        }

        /// <summary>
        /// Creates a deep copy of this <see cref="ShapeModel"/> and associated shape in the worksheet.
        /// </summary>
        /// <param name="targetWorksheet">The worksheet into which the clone will be placed. If null, the cloned <see cref="ShapeModel"/> will be based on the original <see cref="Worksheet"/>/></param>
        /// <returns>The <see cref="ShapeModel"/> that represents the shape</returns>
        public ShapeModel Clone(Worksheet targetWorksheet)
        {
            // If no target worksheet is supplied, clone in sit (ie. on the current worksheet)
            Worksheet cloneToWorksheet = targetWorksheet == null ? this.Worksheet : targetWorksheet;

            if (cloneToWorksheet.WorksheetPart.DrawingsPart == null)
            {
                var drawingsPart = cloneToWorksheet.WorksheetPart.AddNewPart<OpenXmlPackaging.DrawingsPart>();
                drawingsPart.WorksheetDrawing = new DrawingSpreadsheet.WorksheetDrawing();

                // if a drawings part is being created then we need to add a Drawing to the end of the targetworksheet
                DocumentFormat.OpenXml.Spreadsheet.Drawing drawing = new DocumentFormat.OpenXml.Spreadsheet.Drawing()
                {
                    Id = cloneToWorksheet.WorksheetPart.GetIdOfPart(cloneToWorksheet.WorksheetPart.DrawingsPart)
                };

                cloneToWorksheet.Append(drawing);
            }

            // Clone the anchor for the template chart to get a new anchor + shape
            DrawingSpreadsheet.TwoCellAnchor clonedAnchor = (DrawingSpreadsheet.TwoCellAnchor)this.Anchor.CloneNode(true);
            DrawingSpreadsheet.Shape clonedShape = clonedAnchor.Descendants<DrawingSpreadsheet.Shape>().FirstOrDefault();

            // Insert the cloned anchor.
            cloneToWorksheet.WorksheetPart.DrawingsPart.WorksheetDrawing.Append(clonedAnchor);

            // Insert the cloned anchor.
            return new ShapeModel(cloneToWorksheet, clonedAnchor);
        }

        /// <summary>
        /// Moves the shape into position within its worksheet.<br/>
        /// NB! For a shape the size and extents (dx and dy from cell left and top) remain unchanged...
        /// </summary>
        /// <param name="rowIndex"></param>
        /// <param name="columnIndex"></param>
        public void Move(uint rowIndex, uint columnIndex)
        {
            // From Marker
            this.Anchor.FromMarker.RowId.Text = rowIndex.ToString();
            this.Anchor.FromMarker.ColumnId.Text = columnIndex.ToString();

            // To Marker
            this.Anchor.ToMarker.RowId.Text = rowIndex.ToString();
            this.Anchor.ToMarker.ColumnId.Text = columnIndex.ToString();
        }

        /// <summary>
        /// Removes all references to the shape from the worksheet.<br/>
        /// This will invalidate this <see cref="ShapeModel"/>, i.e. errors will be raised if an attempt is made to use the invalid model.
        /// </summary>
        public void RemoveShape()
        {
            // Remove the Anchor from the WorksheetDrawing
            if (this.Anchor != null && this.Anchor.Parent != null)
            {
                this.Anchor.Remove();
            }

            this.Worksheet = null;
            this.Anchor = null;
            this.shape = null;
            this.drawingsPart = null;
            this.IsValid = false;
        }

        #endregion Public Methods

        #region Private Helpers

        /// <summary>
        /// Gets shapes within a worksheet that have a specified name (id)
        /// </summary>
        /// <param name="worksheetPart">The <see cref="OpenXmlPackaging.WorksheetPart"/></param>
        /// <param name="name">The id/name required</param>
        /// <returns></returns>
        private static IEnumerable<DrawingSpreadsheet.Shape> GetShapesWithName(OpenXmlPackaging.WorksheetPart worksheetPart, string name)
        {
            IEnumerable<DrawingSpreadsheet.Shape> shapes = worksheetPart.DrawingsPart.WorksheetDrawing.Descendants<DrawingSpreadsheet.Shape>()
                                                          .Where(s => s.NonVisualShapeProperties.NonVisualDrawingProperties.Name == name);
            return shapes;
        }

        /// <summary>
        /// Gets the name of a supplied shape in a worksheet
        /// </summary>
        /// <param name="shape"></param>
        /// <returns></returns>
        private static string GetShapeName(DrawingSpreadsheet.Shape shape)
        {
            return shape.NonVisualShapeProperties.NonVisualDrawingProperties.Name;
        }

        #endregion Private Helpers
    }
}
