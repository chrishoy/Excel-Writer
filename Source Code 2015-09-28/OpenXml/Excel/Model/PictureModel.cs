namespace ExcelWriter.OpenXml.Excel.Model
{
    using System.Collections.Generic;
    using System.Linq;

    using DocumentFormat.OpenXml.Packaging;
    using DocumentFormat.OpenXml.Spreadsheet;

    using Drawing = DocumentFormat.OpenXml.Drawing;
    using DrawingPictures = DocumentFormat.OpenXml.Drawing.Pictures;
    using DrawingSpreadsheet = DocumentFormat.OpenXml.Drawing.Spreadsheet;

    /// <summary>
    /// Encapsulates information about an Excel image object.
    /// </summary>
    public class PictureModel : ModelBase
    {
        #region Local Fields

        private string imageId;
        private ImagePart imagePart;

        #endregion Local Fields

        #region Construction

        /// <summary>
        /// Initializes a new instance of the <see cref="PictureModel"/> class.
        /// </summary>
        /// <param name="imagePart">The <see cref="ImagePart"/> on which the model is based</param>
        /// <param name="anchor">The <see cref="DrawingSpreadsheet.TwoCellAnchor"/> which hosts the image</param>
        private PictureModel(ImagePart imagePart, DrawingSpreadsheet.TwoCellAnchor anchor) : base(anchor)
        {
            this.imagePart = imagePart;

            // Grab the picture and store the height and width
            DrawingSpreadsheet.Picture picture = anchor.Descendants<DrawingSpreadsheet.Picture>().FirstOrDefault();

            // Shape properties and extents....
            Drawing.Extents extents = picture.Descendants<Drawing.Extents>().FirstOrDefault();
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

        #endregion

        #region Pubic Properties

        /// <summary>
        /// Gets a value indicating whether the model is valid with the context of the worksheet.
        /// </summary>
        public bool IsValid { get; private set; }

        /// <summary>
        /// Gets the picture Id
        /// </summary>
        public string ImageId
        {
            get { return this.imageId; }
            private set { this.imageId = value; }
        }

        /// <summary>
        /// Gets a reference to the <see cref="ImagePart"/> which this <see cref="PictureModel"/> represents.
        /// </summary>
        public ImagePart ImagePart
        {
            get { return this.imagePart; }
        }


        #endregion Public Properties

        #region Static Public Methods

        /// <summary>
        /// Returns an instance of a <see cref="PictureModel"/> for the first image with a specified picture id in a worksheet.
        /// </summary>
        /// <param name="ws">The <see cref="Worksheet"/> in which the image resides</param>
        /// <param name="id">The image id</param>
        /// <returns>The <see cref="PictureModel"/> that represents the image</returns>
        public static PictureModel GetPictureModel(Worksheet ws, string id)
        {
            Guard.IsNotNull(ws, "ws");
            Guard.IsNotNullOrEmpty(id, "id");

            PictureModel pictureModel = null;
            ImagePart imagePart = GetImagePart(id, ws.WorksheetPart);

            if (imagePart != null)
            {
                // Get the Anchor that host the Graphic
                DrawingSpreadsheet.Picture pic = GetHostingPicture(id, ws.WorksheetPart);
                DrawingSpreadsheet.TwoCellAnchor anchor = pic.Ancestors<DrawingSpreadsheet.TwoCellAnchor>().FirstOrDefault();

                pictureModel = new PictureModel(imagePart, anchor)
                {
                    Worksheet = ws,
                    ImageId = id,
                    IsValid = true,
                };
            }

            return pictureModel;
        }

        #endregion Static Public Methods

        #region Public Methods

        /// <summary>
        /// Creates a deep copy of this <see cref="PictureModel"/> and associated chart in the worksheet.
        /// </summary>
        /// <returns>The <see cref="PictureModel"/> that represents the chart</returns>
        public PictureModel Clone()
        {
            return this.Clone(this.Worksheet);
        }

        /// <summary>
        /// Creates a deep copy of this <see cref="PictureModel"/> and associated chart in the worksheet.
        /// </summary>
        /// <param name="targetWorksheet">The worksheet into which the clone will be placed. If null, the cloned <see cref="PictureModel"/> will be based on the original <see cref="Worksheet"/>/></param>
        /// <returns>The <see cref="PictureModel"/> that represents the chart</returns>
        public PictureModel Clone(Worksheet targetWorksheet)
        {
            // If no target worksheet is supplied, clone in situ (ie. on the current worksheet)
            Worksheet cloneToWorksheet = targetWorksheet == null ? this.Worksheet : targetWorksheet;

            // Name of the source and target worksheet (for debugging)
            string sourceWorksheetName = this.Worksheet.WorksheetPart.GetSheetName();
            string targetWorksheetName = cloneToWorksheet.WorksheetPart.GetSheetName();

            System.Diagnostics.Debug.Print("PictureModel - Cloning chart on worksheet '{0}' into '{1}'", sourceWorksheetName, targetWorksheetName);

            // Create a DrawingPart in the target worksheet if it does not already exist
            if (cloneToWorksheet.WorksheetPart.DrawingsPart == null)
            {
                var drawingsPart = cloneToWorksheet.WorksheetPart.AddNewPart<DrawingsPart>();
                drawingsPart.WorksheetDrawing = new DrawingSpreadsheet.WorksheetDrawing();

                // if a drawings part is being created then we need to add a Drawing to the end of the targetworksheet
                var drawing = new DocumentFormat.OpenXml.Spreadsheet.Drawing()
                {
                    Id = cloneToWorksheet.WorksheetPart.GetIdOfPart(cloneToWorksheet.WorksheetPart.DrawingsPart)
                };

                cloneToWorksheet.Append(drawing);
            }

            // Take copy elements
            ImagePart imagePart2 = cloneToWorksheet.WorksheetPart.DrawingsPart.AddImagePart(this.imagePart.ContentType);
            imagePart2.FeedData(this.imagePart.GetStream());

            // Clone the anchor for the template image to get a new image anchor
            var anchor2 = (DrawingSpreadsheet.TwoCellAnchor)this.Anchor.CloneNode(true);

            // Insert the cloned anchor into the worksheet drawing of the DrawingsPart.
            cloneToWorksheet.WorksheetPart.DrawingsPart.WorksheetDrawing.Append(anchor2);

            // Update the BlipFill in the Anchor 2 (TwoCellAnchor -> GraphicFrame -> Graphic -> GraphicData -> Picture -> BlipFill)
            DrawingSpreadsheet.BlipFill blipFill = anchor2.Descendants<DrawingSpreadsheet.BlipFill>().FirstOrDefault();
            blipFill.Blip.Embed = cloneToWorksheet.WorksheetPart.DrawingsPart.GetIdOfPart(imagePart2);

            // Wrap and return as a model
            PictureModel chartModel = new PictureModel(imagePart2, anchor2)
            {
                Worksheet = cloneToWorksheet,
                ImageId = this.ImageId,
                IsValid = true,
            };

            return chartModel;
        }

        /// <summary>
        /// Moves the picture into position within its worksheet.
        /// </summary>
        /// <param name="fromRowIndex">The worksheet row where the picture will start</param>
        /// <param name="fromColumnIndex">The worksheet column where the picture will start</param>
        /// <param name="toRowIndex">The worksheet row when the picture will end</param>
        /// <param name="toColumnIndex">The worksheet column where the picture will end</param>
        public void Move(uint fromRowIndex, uint fromColumnIndex, uint toRowIndex, uint toColumnIndex)
        {
            // From Marker
            this.Anchor.FromMarker.RowId.Text = fromRowIndex.ToString();
            this.Anchor.FromMarker.RowOffset.Text = "0";
            this.Anchor.FromMarker.ColumnId.Text = fromColumnIndex.ToString();
            this.Anchor.FromMarker.ColumnOffset.Text = "0";

            // To Marker
            this.Anchor.ToMarker.RowId.Text = toRowIndex.ToString();
            this.Anchor.ToMarker.RowOffset.Text = "0";
            this.Anchor.ToMarker.ColumnId.Text = toColumnIndex.ToString();
            this.Anchor.ToMarker.ColumnOffset.Text = "0";
        }

        /// <summary>
        /// Removes all references to the picture from the worksheet.<br/>
        /// This will invalidate this <see cref="PictureModel"/>, i.e. errors will be raised if an attempt is made to use the invalid model.
        /// </summary>
        public void RemovePicture()
        {
            try
            {
                this.Worksheet.WorksheetPart.DrawingsPart.DeletePart(this.imagePart);

                IEnumerable<ImagePart> imageParts = this.Worksheet.WorksheetPart.DrawingsPart.GetPartsOfType<ImagePart>();

                // Remove the Anchor from the WorksheetDrawing
                this.Anchor.Remove();
                this.Anchor = null;

                this.imagePart = null;
                this.Worksheet = null;

                this.IsValid = false;
            }
            catch (System.InvalidOperationException)
            {
                // Do nothing, the part had already beed destroyed, probably by another model based on the same image
            }
        }

        #endregion Public Methods

        #region Private Helpers

        /// <summary>
        /// Gets the identified <see cref="ImagePart"/> on a <see cref="WorksheetPart"/>
        /// </summary>
        /// <param name="id">The id of the image</param>
        /// <param name="wp">The <see cref="WorksheetPart"/></param>
        /// <returns>The identified <see cref="ImagePart"/></returns>
        private static ImagePart GetImagePart(string id, WorksheetPart wp)
        {
            DrawingSpreadsheet.Picture sourcePic = GetHostingPicture(id, wp);

            // we have the graphics frame with data, so now pull out the chart part
            if (sourcePic != null && sourcePic.BlipFill != null && sourcePic.BlipFill.Blip != null)
            {
                var sourceBlip = sourcePic.BlipFill.Blip;
                if (sourceBlip != null && sourceBlip.Embed.HasValue)
                {
                    return (ImagePart)wp.DrawingsPart.GetPartById(sourceBlip.Embed.Value);
                }
            }

            return null;
        }

        private static DrawingSpreadsheet.Picture GetHostingPicture(string id, WorksheetPart wp)
        {
            DrawingSpreadsheet.Picture sourcePic = null;
            if (wp.DrawingsPart != null)
            {
                // we need to pull out the graphic frame that matches the supplied name
                foreach (var pic in wp.DrawingsPart.WorksheetDrawing.Descendants<DrawingSpreadsheet.Picture>())
                {
                    // need to check it has the various properties
                    if (pic.NonVisualPictureProperties != null &&
                        pic.NonVisualPictureProperties.NonVisualDrawingProperties != null &&
                        pic.NonVisualPictureProperties.NonVisualDrawingProperties.Name.HasValue)
                    {
                        // and then try and match
                        if (id.CompareTo(pic.NonVisualPictureProperties.NonVisualDrawingProperties.Name.Value) == 0)
                        {
                            sourcePic = pic;
                            break;
                        }
                    }
                }
            }

            return sourcePic;
        }

        #endregion
    }
}
