namespace ExcelWriter
{
    using System;
    using System.Linq;
    using System.Text;

    using OpenXml.Excel.Model;

    using Drawing = DocumentFormat.OpenXml.Drawing;
    using DrawingSpreadsheet = DocumentFormat.OpenXml.Drawing.Spreadsheet;
    using OpenXmlSpreadsheet = DocumentFormat.OpenXml.Spreadsheet;

    public sealed partial class ExportGenerator
    {
        #region Shape processing

        /// <summary>
        /// Processes the supplied shape, setting properties on the shape as required.
        /// </summary>
        /// <param name="shape">The <see cref="Shape"/> which is to be represented in the worksheet</param>
        /// <param name="shapeModel">The <see cref="ShapeModel"/> which models the instance of the <see cref="DrawingSpreadsheet.Shape">OpenXML shape</see> in the Excel workbook</param>
        private static void ProcessDynamicShape(Shape shape, ShapeModel shapeModel)
        {
            if (shapeModel == null) throw new ArgumentNullException("shapeModel");

            System.Windows.Media.Color? fillColour = BindingContainer.ConvertToNullableColour(shape.FillColour);

            if (fillColour.HasValue)
            {
                UpdateShapeFillColour(shapeModel.Shape, fillColour.Value);
            }
        }

        /// <summary>
        /// Sets the color of the series using a solidcolor brush
        /// If a null brush is supplied any color is removed so the color will be automatic
        /// </summary>
        /// <param name="shape">The <see cref="DrawingSpreadsheet.Shape"/> to be updated</param>
        /// <param name="colour">The <see cref="System.Drawing.Color">colour</see> to be used as the shape fill colour</param>
        private static void UpdateShapeFillColour(DrawingSpreadsheet.Shape shape, System.Windows.Media.Color colour)
        {
            if (shape == null)
            {
                throw new ArgumentNullException("shape");
            }

            // Get the ShapeProperties
            DrawingSpreadsheet.ShapeProperties shapeProperties = shape.ShapeProperties;
            if (shapeProperties != null)
            {
                // Try and find a SolidFill, and remove it.
                Drawing.SolidFill solidFill = shapeProperties.Elements<Drawing.SolidFill>().FirstOrDefault();
                if (solidFill != null)
                {
                    solidFill.Remove();
                }

                var newSolidFill = new Drawing.SolidFill();

                StringBuilder hexString = new StringBuilder();
                hexString.Append(colour.R.ToString("X").PadLeft(2, '0'));
                hexString.Append(colour.G.ToString("X").PadLeft(2, '0'));
                hexString.Append(colour.B.ToString("X").PadLeft(2, '0'));

                var hexColour = new Drawing.RgbColorModelHex()
                {
                    Val = hexString.ToString()
                };

                //var outlineNoFill = new Drawing.Outline();
                //outlineNoFill.Append(new Drawing.NoFill());

                newSolidFill.Append(hexColour);

                // Append the SolidFill after the PresetGeometry (preset grometry being the geometry behind the shape)
                Drawing.PresetGeometry presetGeometry = shapeProperties.Elements<Drawing.PresetGeometry>().FirstOrDefault();
                presetGeometry.InsertAfterSelf(newSolidFill);
            }

            //shape.Append(shapeProperties);
        }

        #endregion
    }
}
