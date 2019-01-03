namespace ExcelWriter.OpenXml.Excel
{
    using System.Linq;
    using DocumentFormat.OpenXml.Packaging;
    using DrawingSpreadsheet = DocumentFormat.OpenXml.Drawing.Spreadsheet;
    using DocumentFormat.OpenXml.Spreadsheet;

    public static class WorksheetPartExtensions
    {
        public static Sheet GetSheet(this WorksheetPart workSheetPart)
        {
            var spreadSheet = workSheetPart.OpenXmlPackage as SpreadsheetDocument;
            if (spreadSheet != null)
            {
                var id = spreadSheet.WorkbookPart.GetIdOfPart(workSheetPart);

                return (from s in  spreadSheet.WorkbookPart.Workbook.Descendants<Sheet>()
                        where s.Id.HasValue && s.Id.Value.CompareTo(id) == 0
                        select s).FirstOrDefault();
            }
            return null;
        }

        public static int GetSheetIndex(this WorksheetPart workSheetPart)
        {
            var spreadSheet = workSheetPart.OpenXmlPackage as SpreadsheetDocument;
            if (spreadSheet != null)
            {
                var id = spreadSheet.WorkbookPart.GetIdOfPart(workSheetPart);

                var sheet = (from s in spreadSheet.WorkbookPart.Workbook.Descendants<Sheet>()
                            where s.Id.HasValue && s.Id.Value.CompareTo(id) == 0
                            select s).FirstOrDefault();

                if (sheet != null)
                {
                    return spreadSheet.WorkbookPart.Workbook.Sheets.ToList().IndexOf(sheet);
                }
            }
            return -1;
        }

        /// <summary>
        /// Gets the name of a supplied <see cref="WorksheetPart"/>
        /// </summary>
        /// <param name="workSheetPart">The <see cref="WorksheetPart"/></param>
        /// <returns>The name of the sheet, or null if not found.</returns>
        public static string GetSheetName(this WorksheetPart workSheetPart)
        {
            var spreadSheet = workSheetPart.OpenXmlPackage as SpreadsheetDocument;
            if (spreadSheet != null)
            {
                var id = spreadSheet.WorkbookPart.GetIdOfPart(workSheetPart);

                var sheet = (from s in spreadSheet.WorkbookPart.Workbook.Descendants<Sheet>()
                             where s.Id.HasValue && s.Id.Value.CompareTo(id) == 0
                             select s).FirstOrDefault();

                return sheet.Name;
            }
            return null;
        }

        public static string GetPartId(this WorksheetPart workSheetPart)
        {
            var spreadSheet = workSheetPart.OpenXmlPackage as SpreadsheetDocument;
            if (spreadSheet != null)
            {
                return spreadSheet.WorkbookPart.GetIdOfPart(workSheetPart);
            }
            return null;
        }

        /// <summary>
        /// Updates the sources.
        /// This ensures that any formula in table or charts are updated to reflect and worksheet name changes
        /// or changes in numbers of rows bound
        /// </summary>
        /// <param name="sheet">The sheet.</param>
        /// <param name="oldSourceName">Old name of the source.</param>
        /// <param name="newSourceName">New name of the source.</param>
        /// <param name="tableRowCount">The table row count.</param>
        public static void UpdateSources(this WorksheetPart sheet, string oldSourceName, string newSourceName, int? tableRowCount)
        {
            // if there's sheet data, update any formula on it
            var sheetData = sheet.Worksheet.GetFirstChild<SheetData>();
            if (sheetData != null)
            {
                foreach (var f in sheetData.Descendants<CellFormula>())
                {
                    f.Text = Helpers.UpdateFormula(f.Text, oldSourceName, newSourceName, null);
                }
            }
            // if there are drawings
            if (sheet.DrawingsPart != null)
            {
                // process all charts, this may need extending for other types
                foreach (ChartPart chartPart in sheet.DrawingsPart.GetPartsOfType<ChartPart>())
                {
                    chartPart.UpdateSources(oldSourceName, newSourceName, tableRowCount);
                }
            }
        }

        /// <summary>
        /// Tries to find a shape with the matching title attribute
        /// </summary>
        /// <param name="reportWorksheetPart">The report worksheet part.</param>
        /// <param name="shapeTitle">The shape title.</param>
        /// <returns>A Spreadsheet Shape</returns>
        public static DrawingSpreadsheet.Shape GetShapeByName(this WorksheetPart worksheetPart, string shapeName)
        {
            if (string.IsNullOrEmpty(shapeName) || worksheetPart == null || worksheetPart.DrawingsPart == null)
            {
                return null;
            }

            foreach (var shape in worksheetPart.DrawingsPart.WorksheetDrawing.Descendants<DrawingSpreadsheet.Shape>())
            {
                var nvdp = shape.NonVisualShapeProperties.GetFirstChild<DrawingSpreadsheet.NonVisualDrawingProperties>();
                if (nvdp != null)
                {
                    if (nvdp.Name.HasValue && nvdp.Name.Value.CompareTo(shapeName) == 0)
                    {
                        return shape;
                    }
                }
            }
            return null;
        }
    }
}   
   
