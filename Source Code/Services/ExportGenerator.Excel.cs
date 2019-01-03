namespace ExcelWriter
{
    using System;
    using System.Collections.Generic;
    using System.Drawing;
    using System.Globalization;
    using System.IO;
    using System.Linq;
    using System.Reflection;
    using System.Windows.Documents;

    using DocumentFormat.OpenXml;
    using DocumentFormat.OpenXml.Packaging;

    using OpenXml.Excel;
    using OpenXml.Excel.Model;

    using DrawingSpreadsheet = DocumentFormat.OpenXml.Drawing.Spreadsheet;
    using OpenXmlSpreadsheet = DocumentFormat.OpenXml.Spreadsheet;

    /// <summary>
    /// 
    /// </summary>
    /// <content>
    /// Responsible for creating excel documents
    /// </content>
    public sealed partial class ExportGenerator
    {
        #region Public Methods

        /// <summary>
        /// If required clones a sheet from the template into the output document
        /// </summary>
        /// <param name="document">the target document</param>
        /// <param name="newSheetName">the name of the new worksheet to create</param>
        /// <param name="templateDocument">the template document containing the template sheet</param>
        /// <param name="templateSheetName">the name of the template worksheet to clone if necessary</param>
        /// <returns>
        /// A new or existing WorksheetPart
        /// </returns>
        public static WorksheetPart TryCloneSheet(SpreadsheetDocument document, string newSheetName, SpreadsheetDocument templateDocument, string templateSheetName)
        {
            // TBD. Add sheet length validation in here, max currently 31 chars
            if (string.IsNullOrEmpty(newSheetName))
            {
                return null;
            }

            // try and get the sheet from th target document first, if its found there then go no further
            WorksheetPart newWSPart = document.WorkbookPart.Workbook.GetWorksheetPartByName(newSheetName);

            // if sheet isnt found in the target and there's a name to look for in the template
            // then try attempt to clone from the template
            if (newWSPart == null && !string.IsNullOrEmpty(templateSheetName))
            {
                // if its not found in the target doc, then go ahead and get the template sheets..
                var templateWSPart = templateDocument.WorkbookPart.Workbook.GetWorksheetPartByName(templateSheetName);

                // if found...
                if (templateWSPart != null)
                {
                    var state = templateDocument.WorkbookPart.Workbook.GetWorksheetStateName(templateSheetName);

                    // ...clone it into the target sheet with the provided name
                    newWSPart = document.WorkbookPart.InsertClonedWorksheetPart(templateWSPart, state, newSheetName);
                }
            }

            return newWSPart;
        }

        /// <summary>
        /// Merges indexed styles for a worksheet between source and target.
        /// As they're referenced as indices, style 1 in source may be different to 1 in target.
        /// This routine handles this by creating a new style if necessary, or finding the id of the
        /// target style if it already exists/
        /// </summary>
        /// <param name="newTargetWorksheet">The new target document</param>
        /// <param name="sourceDocument">The source document</param>
        /// <param name="targetDocument">The target document</param>
        /// <param name="mergeStyles">A dictionary of styles</param>
        public static void MergeResourcesAndStyles(WorksheetPart newTargetWorksheet,
                                                   SpreadsheetDocument sourceDocument,
                                                   SpreadsheetDocument targetDocument,
                                                   Dictionary<uint, uint> mergeStyles = null)
        {
            // init this if its null
            if (mergeStyles == null)
            {
                mergeStyles = new Dictionary<uint, uint>();
            }

            // get the stylesheets up front, these handle all styles
            OpenXmlSpreadsheet.Stylesheet sourceStylesheet = sourceDocument.WorkbookPart.GetPartsOfType<WorkbookStylesPart>().First().Stylesheet;
            OpenXmlSpreadsheet.Stylesheet targetStylesheet = targetDocument.WorkbookPart.GetPartsOfType<WorkbookStylesPart>().First().Stylesheet;

            // merge any columns styles
            var columns = newTargetWorksheet.Worksheet.GetFirstChild<OpenXmlSpreadsheet.Columns>();
            if (columns != null)
            {
                foreach (var c in columns.Elements<OpenXmlSpreadsheet.Column>())
                {
                    if (c.Style != null && c.Style.HasValue)
                    {
                        c.Style = Helpers.GetOrCreateCellStyleIndex(c.Style, sourceStylesheet, targetStylesheet, mergeStyles);
                    }
                }
            }

            // then for the data in the sheets
            var sheetData = newTargetWorksheet.Worksheet.GetFirstChild<OpenXmlSpreadsheet.SheetData>();
            if (sheetData != null)
            {
                // merge rows
                foreach (var r in sheetData.Elements<OpenXmlSpreadsheet.Row>())
                {
                    if (r.StyleIndex != null && r.StyleIndex.HasValue)
                    {
                        r.StyleIndex = Helpers.GetOrCreateCellStyleIndex(r.StyleIndex, sourceStylesheet, targetStylesheet, mergeStyles);
                    }

                    // and each cell in row
                    foreach (var cell in r.Descendants<OpenXmlSpreadsheet.Cell>())
                    {
                        if (cell.DataType != null && cell.DataType.HasValue && cell.DataType.Value == OpenXmlSpreadsheet.CellValues.SharedString && cell.CellValue != null)
                        {
                            int stringId = 0;
                            if (int.TryParse(cell.CellValue.Text, out stringId))
                            {
                                // we need to handle shared strings, again these are referenced by index
                                var newIndex = Helpers.GetOrCreateSharedStringIndex(stringId, sourceDocument, targetDocument);
                                if (newIndex != -1)
                                {
                                    cell.CellValue.Text = newIndex.ToString(CultureInfo.InvariantCulture);
                                }
                            }
                        }

                        if (cell.StyleIndex != null && cell.StyleIndex.HasValue)
                        {
                            cell.StyleIndex = Helpers.GetOrCreateCellStyleIndex(cell.StyleIndex, sourceStylesheet, targetStylesheet, mergeStyles);
                        }
                    }
                }
            }

            // finally work though any conditional formatting rule
            foreach (var cfr in newTargetWorksheet.Worksheet.Descendants<OpenXmlSpreadsheet.ConditionalFormattingRule>())
            {
                if (cfr.FormatId.HasValue)
                {
                    cfr.FormatId = Helpers.CreateDifferentialFormatIndex(cfr.FormatId, sourceStylesheet, targetStylesheet);
                }
            }
        }

        ///// <summary>
        ///// Sets the height of a row to fit the text it contains.
        ///// </summary>
        ///// <param name="sd">The sheet</param>
        ///// <param name="rowStart">The start row</param>
        ///// <param name="fontFamily">The font family</param>
        ///// <param name="emsize">The size in emus</param>
        ///// <param name="lineWrapWidth">The width of the line-wrap</param>
        ///// <param name="widthMultiplier">A width multiplier</param>
        ///// <param name="heightMultiplier">A height multiplier</param>
        //public static void AutoRowResizeText(OpenXmlSpreadsheet.SheetData sd, uint rowStart, string fontFamily, decimal emsize, decimal lineWrapWidth, decimal widthMultiplier, decimal heightMultiplier)
        //{
        //    foreach (OpenXmlSpreadsheet.Row row in sd.Elements<OpenXmlSpreadsheet.Row>())
        //    {
        //        if (row.RowIndex >= rowStart)
        //        {
        //            decimal height = WorksheetExtensions.RowHeight(row, fontFamily, emsize, lineWrapWidth, widthMultiplier, heightMultiplier);

        //            row.Height = new DoubleValue((double)height);
        //            row.CustomHeight = new BooleanValue(true);
        //        }
        //    }
        //}

        #endregion Public Methods

        #region Privates static possibly - for move to openxml?

        /// <summary>
        /// The "Table" is a DefinedName (Range of cells) on the worksheet with a "Table" Suffix. This procedure examined the
        /// DefinedName, evaluates the DataPart to which it is associated and hides the rows of the DefinedName if the DataPart
        /// does not exist in the sets collection.
        /// </summary>
        /// <param name="workbook">A <see cref="OpenXmlSpreadsheet.Workbook" /></param>
        private static void HideTablesWithNoData(OpenXmlSpreadsheet.Workbook workbook)
        {
            if (workbook == null || workbook.DefinedNames == null)
            {
                return;
            }

            const string RangeNameSuffix = "Table";

            // Check all of the named ranges set up in the workbook
            foreach (OpenXmlSpreadsheet.DefinedName destRange in workbook.DefinedNames)
            {
                if (destRange.Name.HasValue)
                {
                    // If the named range ends with 'TableRows' then find the named
                    // range which will be used as the data source and basis of the row hiding
                    string destRangeName = destRange.Name.Value;
                    if (destRangeName.EndsWith(RangeNameSuffix))
                    {
                        string dataRangeName = destRangeName.Remove(destRangeName.Length - RangeNameSuffix.Length);
                        OpenXmlSpreadsheet.DefinedName dataRange = GetNamedRange(workbook, dataRangeName);

                        if (dataRange == null)
                        {
                            workbook.HideRowsOfDefinedName(destRange);
                        }
                    }
                }
            }
        }

        /// <summary>
        /// The "TableRows" is a DefineName (Range of Cells) on a worksheet that cover the potential cells that can be populated with
        /// data from DataParts. This will not include totals or headings. The procedure closes up the "gap" following the last item
        /// of data and the totals.
        /// </summary>
        /// <param name="workbook">A <see cref="OpenXmlSpreadsheet.Workbook" /></param>
        private static void HideTableRowsWithNoData(OpenXmlSpreadsheet.Workbook workbook)
        {
            if (workbook == null || workbook.DefinedNames == null)
            {
                return;
            }

            const string RangeNameSuffix = "TableRows";

            // Check all of the named ranges set up in the workbook
            foreach (OpenXmlSpreadsheet.DefinedName destRange in workbook.DefinedNames)
            {
                if (destRange.Name.HasValue)
                {
                    // If the named range ends with 'TableRows' then find the named
                    // range which will be used as the data source and basis of the row hiding
                    string destRangeName = destRange.Name.Value;
                    if (destRangeName.EndsWith(RangeNameSuffix))
                    {
                        string dataRangeName = destRangeName.Remove(destRangeName.Length - RangeNameSuffix.Length);
                        OpenXmlSpreadsheet.DefinedName dataRange = GetNamedRange(workbook, dataRangeName);

                        if (dataRange != null)
                        {
                            // Look up where data is sourced from
                            string dataSheetName = string.Empty;
                            uint dataRowStart = 0;
                            uint dataRowEnd = 0;
                            workbook.BreakDownDefinedName(dataRange, ref dataSheetName, ref dataRowStart, ref dataRowEnd);

                            // Look up where rows are to be hidden from
                            string destSheetName = string.Empty;
                            uint destRowStart = 0;
                            uint destRowEnd = 0;
                            workbook.BreakDownDefinedName(destRange, ref destSheetName, ref destRowStart, ref destRowEnd);

                            // Worksheet where rows are to be hidden
                            OpenXmlSpreadsheet.Worksheet ws = GetWorksheet(workbook, destSheetName);

                            // Calculate the number of rows to hide in the destination range (and where to hide from)
                            uint destRowHideStart = destRowStart + (dataRowEnd - dataRowStart);
                            if (destRowHideStart < destRowEnd)
                            {
                                ws.HideRows(destRowHideStart, destRowEnd);
                            }
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Gets an instance of a named range in a workbook
        /// </summary>
        /// <param name="wb">The <see cref="OpenXmlSpreadsheet.Workbook" /></param>
        /// <param name="namedRangeName">The requested named range</param>
        /// <returns>
        /// A <see cref="OpenXmlSpreadsheet.DefinedName" />
        /// </returns>
        private static OpenXmlSpreadsheet.DefinedName GetNamedRange(OpenXmlSpreadsheet.Workbook wb, string namedRangeName)
        {
            foreach (OpenXmlSpreadsheet.DefinedName dn in wb.DefinedNames)
            {
                if (dn.Name.HasValue)
                {
                    string name = dn.Name.Value;
                    if (name.CompareTo(namedRangeName) == 0)
                    {
                        return dn;
                    }
                }
            }

            return null;
        }

        ///// <summary>
        ///// This procedure looks for all defined names with the suffix "InlineText" and uses this as the destination range for the values. The source data
        ///// range is determined by replacing "InlineText" with "Data" in the defined name. From here the source data (in RTF XAML) is broken down into paragraphs
        ///// and runs and added as Inline Strings to the merged cells within the destination range.
        ///// </summary>
        ///// <param name="workbook">A <see cref="OpenXmlSpreadsheet.Workbook"/></param>
        //private static void ProcessXAMLRTFToInlineText(OpenXmlSpreadsheet.Workbook workbook)
        //{
        //    if (workbook == null || workbook.DefinedNames == null)
        //    {
        //        return;
        //    }

        //    const string RangeNameSuffix = "InlineText";

        //    // Check all of the named ranges set up in the workbook
        //    foreach (OpenXmlSpreadsheet.DefinedName destRange in workbook.DefinedNames)
        //    {
        //        if (destRange.Name.HasValue)
        //        {
        //            string destRangeName = destRange.Name.Value;
        //            if (destRangeName.EndsWith(RangeNameSuffix))
        //            {
        //                string dataRangeName = destRangeName.Remove(destRangeName.Length - RangeNameSuffix.Length);
        //                OpenXmlSpreadsheet.DefinedName dataRange = GetNamedRange(workbook, dataRangeName);

        //                // Look up where data is sourced from
        //                string dataSheetName = string.Empty;
        //                uint dataRowStart = 0;
        //                uint dataRowEnd = 0;
        //                workbook.BreakDownDefinedName(dataRange, ref dataSheetName, ref dataRowStart, ref dataRowEnd);

        //                string destSheetName = string.Empty;
        //                uint destRowStart = 0;
        //                uint destRowEnd = 0;
        //                uint destColStart = 0;
        //                uint destColEnd = 0;
        //                workbook.BreakDownDefinedName(destRange, ref destSheetName, ref destRowStart, ref destRowEnd, ref destColStart, ref destColEnd);

        //                // Where the RTF XAML resides on the hidden sheet.
        //                //WorksheetPart sourceWorksheetPart = workbook.GetWorksheetPartByName(dataSheetName);

        //                // The page wide location of the resultant paragraphs. one excel row per paragraph. 
        //                //WorksheetPart targetWorksheetPart = workbook.GetWorksheetPartByName(destSheetName);

        //                //OpenXmlSpreadsheet.SheetData sourcesheetData = sourceWorksheetPart.Worksheet.GetFirstChild<OpenXmlSpreadsheet.SheetData>();
        //                //OpenXmlSpreadsheet.SheetData targetsheetData = targetWorksheetPart.Worksheet.GetFirstChild<OpenXmlSpreadsheet.SheetData>();

        //               // OpenXmlSpreadsheet.Cell sourceXAMLCell = sourcesheetData.GetCell(1, dataRowEnd); // There is only one row it is always the first column.
        //                //OpenXmlSpreadsheet.Cell targetXAMLCell = targetsheetData.GetCell(destColStart, destRowStart);

        //                // Use the cell on the hidden data sheet as source for the XAML reader
        //                //Section rtfSection = XamlSectionDocumentReader(sourceXAMLCell.CellValue.InnerText);

        //                // The fonts required are configurable. (They should really be from the RTF XAML).
        //                //string fontName = ExportConfigurationHelper.GetAppSetting("ExportInlineTextFontName");
        //                //decimal fontSize = System.Convert.ToDecimal(ExportConfigurationHelper.GetAppSetting("ExportInlineTextFontSize"));

        //                // An intermediate structure
        //                //List<OpenXmlSpreadsheet.InlineString> ilsTextCollection = Helpers.ConvertParagraphList(rtfSection.Blocks, fontName, fontSize);

        //                // New styles are added as part of the InlistString apllication. The stylesheet is required for ths.
        //                //OpenXmlSpreadsheet.Stylesheet ss = workbook.WorkbookPart.WorkbookStylesPart.Stylesheet;

        //                // Write the lines into sequential rows in the sheet.
        //                //targetWorksheetPart.Worksheet.ApplyInlineStringList(targetsheetData, ss, ilsTextCollection, destColStart, destRowStart, destColEnd - destColStart);

        //                //decimal avgColWidth = System.Convert.ToDecimal(ExportConfigurationHelper.GetAppSetting("InlineTextAvgColWidth"));
        //                //decimal widthMultiplier = System.Convert.ToDecimal(ExportConfigurationHelper.GetAppSetting("InlineTextWidthMultiplier"));
        //                //decimal heightMultiplier = System.Convert.ToDecimal(ExportConfigurationHelper.GetAppSetting("InlineTextHeightMultiplier"));

        //                // Inline Strings do not have an auto resize. This must be done by examining the string and determining the number of lines.
        //                //AutoRowResizeText(targetsheetData, destRowStart, fontName, fontSize, avgColWidth * ((destColEnd - destColStart) + 1), widthMultiplier, heightMultiplier);
        //            }
        //        }
        //    }
        //}

        /// <summary>
        /// Gets an instance of a <see cref="OpenXmlSpreadsheet.Worksheet" /> in a <see cref="OpenXmlSpreadsheet.Workbook" />
        /// </summary>
        /// <param name="wb">The <see cref="OpenXmlSpreadsheet.Workbook" /></param>
        /// <param name="sheetName">The requested sheet name</param>
        /// <returns>
        /// A <see cref="OpenXmlSpreadsheet.Worksheet" />
        /// </returns>
        private static OpenXmlSpreadsheet.Worksheet GetWorksheet(OpenXmlSpreadsheet.Workbook wb, string sheetName)
        {
            WorksheetPart wsp = null;
            OpenXmlSpreadsheet.Worksheet ws = null;

            try
            {
                wsp = wb.GetWorksheetPartByName(sheetName);
                if (wsp != null)
                {
                    ws = wsp.Worksheet;
                }
            }
            catch (Exception)
            {
            }

            return ws;
        }

        /// <summary>
        /// Generates model-based elements such as Chart, Shape and Picture, within the supplied <see cref="WorksheetPart" />
        /// </summary>
        /// <param name="map">The root of the map which contains the elements to be written into the worksheet</param>
        /// <param name="worksheetPart">The <see cref="WorksheetPart" /> into which the elements are to be written</param>
        /// <param name="resourceStore">A <see cref="ResourceStore" /> containing template and id driven resources to be used</param>
        private static void GenerateModelBasedElements(BaseMap map, WorksheetPart worksheetPart, ResourceStore resourceStore)
        {
            // Create an instance of a converter that can make sense of Excel colun widths
            // NB! This is not a simple task, and the converter will be extended to convert row heights, which is much simpler...!
            // The converter is required for positioninig model based elements (charts, pictures, shapes etc)
            // The font information should really be read from the 'Normal' style defined in the output workbook...!!! Hmm....
            var dimensionConverter = new ExcelDimensionConverter("Calibri", 11.0f);

            // Generate the charts and shapes into the worksheet
            GenerateCharts(map, worksheetPart.Worksheet, resourceStore, dimensionConverter);
            GenerateShapes(map, worksheetPart.Worksheet, resourceStore, dimensionConverter);
            GeneratePictures(map, worksheetPart.Worksheet, resourceStore, dimensionConverter);
        }

        /// <summary>
        /// Generates all of the shapes that have been defined in the <see cref="Sheet" /> into the supplied <see cref="OpenXmlSpreadsheet.Worksheet" />.
        /// </summary>
        /// <param name="map">The XAML defined <see cref="BaseMap" /></param>
        /// <param name="targetWorksheet">The <see cref="OpenXmlSpreadsheet.Worksheet">worksheet in which the chart has been placed</see>.<br />
        /// If null, the template chart sheet is used.</param>
        /// <param name="resourceStore">The <see cref="ResourceStore" /> where models are held</param>
        /// <param name="excelDimensionConverter">The excel diemnsion converter</param>
        private static void GenerateShapes(BaseMap map, OpenXmlSpreadsheet.Worksheet targetWorksheet, ResourceStore resourceStore, ExcelDimensionConverter excelDimensionConverter)
        {
            // Process all of the shapes which have been defined/dynamically created in the sheet.
            List<Shape> shapes = map.AllDescendentsOfType<Shape>();
            foreach (Shape shape in shapes)
            {
                // Look up the ChartModel from resources.
                string shapeTemplateKey = BindingContainer.ConvertToString(shape.ShapeTemplateKey);
                TempDiagnostics.Output(string.Format("Attempting to generate shape based on ShapeTemplate Key='{0}'", shapeTemplateKey));

                ShapeModel shapeModel = null;
                resourceStore.ShapeModelDictionary.TryGetValue(shapeTemplateKey, out shapeModel);

                if (shapeModel != null)
                {
                    // Clone the ShapeModel, which will, in-turn, clone the shape,
                    // Find out where it is to go in the worksheet (a placeholder will have
                    // been created into which the shape can go), and move the cloned shape there.

                    // If no worksheet supplied, clone the shape within the model's worksheet.
                    if (targetWorksheet == null)
                    {
                        targetWorksheet = shapeModel.Worksheet;
                    }

                    string targetSheetName = targetWorksheet.WorksheetPart.GetSheetName();

                    TempDiagnostics.Output(string.Format("Cloning Shape into sheet '{0}'", targetSheetName));
                    ShapeModel shapeModelClone = shapeModel.Clone(targetWorksheet);
                    ExcelMapCoOrdinatePlaceholder place = shape.MapPlaceholder;

                    // Get some metrics from the placeholder so we can calculate the various dimensions.
                    uint startRowIndex = (uint)place.ExcelRowStart - 1;
                    uint startColumnIndex = (uint)place.ExcelColumnStart - 1;
                    var sizeAndPosition = new ExcelPositionalInfo(startRowIndex, startColumnIndex, startRowIndex, startColumnIndex);

                    // Calculate, and apply, DX and DY offsets for sizing based on the actual cell width and heights specified in the placeholder.
                    // This will update the endRowIndex and endColumn In
                    UpdateSizeAndPosition(shape, shapeModelClone, place, ref sizeAndPosition, excelDimensionConverter);

                    // Pove into position
                    shapeModelClone.SizeAndMove(sizeAndPosition);

                    // Update the cloned shape
                    ProcessDynamicShape(shape, shapeModelClone);
                }
            }
        }

        /// <summary>
        /// Generates all of the pictures that have been defined in the <see cref="Sheet" /> into the supplied <see cref="OpenXmlSpreadsheet.Worksheet" />.
        /// </summary>
        /// <param name="sheet">The XAML defined <see cref="BaseMap" /> which represents the sheet.</param>
        /// <param name="targetWorksheet">The <see cref="OpenXmlSpreadsheet.Worksheet">worksheet in which the chart has been placed</see>.<br />
        /// If null, the template chart sheet is used.</param>
        /// <param name="resourceStore">The <see cref="ResourceStore" /> where models are held</param>
        /// <param name="excelDimensionConverter">The excel dimension converter.</param>
        private static void GeneratePictures(BaseMap sheet, OpenXmlSpreadsheet.Worksheet targetWorksheet, ResourceStore resourceStore, ExcelDimensionConverter excelDimensionConverter)
        {
            // Process all of the pictures which have been defined/dynamically created in the sheet.
            List<Picture> pictures = sheet.AllDescendentsOfType<Picture>();
            foreach (Picture picture in pictures)
            {
                // Look up the PictureModel from resources.
                string pictureTemplateKey = BindingContainer.ConvertToString(picture.PictureTemplateKey);
                TempDiagnostics.Output(string.Format("Attempting to generate picture based on PictureTemplate Key='{0}'", pictureTemplateKey));

                PictureModel pictureModel = null;
                resourceStore.PictureModelDictionary.TryGetValue(pictureTemplateKey, out pictureModel);

                if (pictureModel != null)
                {
                    // Clone the PictureModel, which will, in-turn, clone the picture,
                    // Find out where it is to go in the worksheet (a placeholder will have
                    // been created into which the picture can go), and move the cloned picture there.

                    // If no worksheet supplied, clone the picture within the model's worksheet.
                    if (targetWorksheet == null)
                    {
                        targetWorksheet = pictureModel.Worksheet;
                    }

                    string targetSheetName = targetWorksheet.WorksheetPart.GetSheetName();
                    TempDiagnostics.Output(string.Format("Cloning Picture into sheet '{0}'", targetSheetName));

                    PictureModel pictureModelClone = pictureModel.Clone(targetWorksheet);
                    ExcelMapCoOrdinatePlaceholder placeholder = picture.MapPlaceholder;

                    // Get some metrics from the placeholder so we can calculate the various dimensions.
                    uint startRowIndex = (uint)placeholder.ExcelRowStart - 1;
                    uint startColumnIndex = (uint)placeholder.ExcelColumnStart - 1;
                    var sizeAndPosition = new ExcelPositionalInfo(startRowIndex, startColumnIndex, startRowIndex, startColumnIndex);

                    // Calculate, and apply, DX and DY offsets for sizing based on the actual cell width and heights specified in the placeholder.
                    // This will update the endRowIndex and endColumn In sizeAndPosition
                    UpdateSizeAndPosition(picture, pictureModelClone, placeholder, ref sizeAndPosition, excelDimensionConverter);

                    // Pove into position
                    pictureModelClone.SizeAndMove(sizeAndPosition);

                    // Update the cloned picture (nothing in the picture can be changed at the moment)
                    ////ProcessDynamicPicture(picture, pictureModelClone);
                }
            }
        }

        /// <summary>
        /// Updates the sizeAndPosition
        /// </summary>
        /// <param name="element">The element containing the map which can be positioned</param>
        /// <param name="model">The model containing the element to be positioned</param>
        /// <param name="placeholder">A placeholder in co-ordinate map</param>
        /// <param name="sizeAndPosition">Size and position information</param>
        /// <param name="excelDimensionConverter">A converter which converts Excel dimensions</param>
        private static void UpdateSizeAndPosition(PositionableMap element,
                                                  ModelBase model,
                                                  ExcelMapCoOrdinatePlaceholder placeholder,
                                                  ref ExcelPositionalInfo sizeAndPosition,
                                                  ExcelDimensionConverter excelDimensionConverter)
        {
            // Find the root container (i.e. the container that represents the sheet)
            ExcelMapCoOrdinate sheetContainer = placeholder.GetRoot();
            int excelStartRow = placeholder.ExcelRowStart;
            int excelStartColumn = placeholder.ExcelColumnStart;

            UpdatePositionalInfo(ref sizeAndPosition, sheetContainer, element, model, excelStartRow, excelStartColumn, excelDimensionConverter);
        }

        /// <summary>
        /// Gets the first row or column index that the supplied position (representing EMUs from the top-left of a container) will fall.
        /// </summary>
        /// <param name="position">Positionable information to be applied to the element.</param>
        /// <param name="sheet">The root element (the sheet) which is used to get row/column positional references.</param>
        /// <param name="element">The positionable element containg the item to be positioned.</param>
        /// <param name="model">The model which will be used to create the content of the positionable element.</param>
        /// <param name="placeholderRowIndex">The index of the row</param>
        /// <param name="placeholderColumnIndex">The index of the column</param>
        /// <param name="edc">A converter which knows how to covert row and column dimensions.</param>
        private static void UpdatePositionalInfo(ref ExcelPositionalInfo position,
                                                 ExcelMapCoOrdinate sheet,
                                                 PositionableMap element,
                                                 ModelBase model,
                                                 int placeholderRowIndex,
                                                 int placeholderColumnIndex,
                                                 ExcelDimensionConverter edc)
        {
            // Get the RowOrColumnInfo which represents the first row and column in the sheet.
            RowOrColumnInfo rowInfo = GetRowOrColumn1(sheet.Rows);
            RowOrColumnInfo columnInfo = GetRowOrColumn1(sheet.Columns);

            // Get row where placeholder starts
            var rowsOrderedByIndex = new List<RowOrColumnInfo>();
            while (rowInfo != null)
            {
                rowsOrderedByIndex.Add(rowInfo);
                rowInfo = rowInfo.Next;
            }

            // Get column where placeholder starts
            var columnsOrderedByIndex = new List<RowOrColumnInfo>();
            while (columnInfo != null)
            {
                columnsOrderedByIndex.Add(columnInfo);
                columnInfo = columnInfo.Next;
            }

            // Determine if we are explicitly controlling the height/width/offset via Element properties,
            // or simply using values in the Excel template.

            // Vertical Offset + Height
            double? verticalOffsetInPoints = BindingContainer.ConvertToNullableDouble(element.Placement.VerticalOffset);
            int verticalOffsetInPixels = verticalOffsetInPoints.HasValue
                ? edc.HeightToPixels(verticalOffsetInPoints.Value)
                : edc.EmusToPixels(model.PositionalInfo.From.Row.OffsetInEmus);

            double? heightInPoints = BindingContainer.ConvertToNullableDouble(element.Placement.Height);
            int heightInPixels = heightInPoints.HasValue
                ? edc.HeightToPixels(heightInPoints.Value)
                : edc.EmusToPixels(model.HeightInEmus);

            // Horizontal Offset + Width
            double? horizontalOffsetInPoints = BindingContainer.ConvertToNullableDouble(element.Placement.HorizontalOffset);
            int horizontalOffsetInPixels = horizontalOffsetInPoints.HasValue
                ? edc.WidthToPixels(horizontalOffsetInPoints.Value)
                : edc.EmusToPixels(model.PositionalInfo.From.Column.OffsetInEmus);

            double? widthInExcelWidthUnits = BindingContainer.ConvertToNullableDouble(element.Placement.Width);

            int widthInPixels = widthInExcelWidthUnits.HasValue
                ? edc.WidthToPixels(widthInExcelWidthUnits.Value)
                : edc.EmusToPixels(model.WidthInEmus);

            // Update row info
            // Offset according to specified value if set on the element
            position.From.Row = GetFromPosition(true, rowsOrderedByIndex, placeholderRowIndex, verticalOffsetInPixels, edc);

            RowOrColumnInfo fromRow = rowsOrderedByIndex[(int)position.From.Row.Index];
            position.To.Row = GetToPosition(true, fromRow, position.From.Row.OffsetInEmus, heightInPixels, edc);

            // Update column info
            // Offset according to specified value if set on the element
            position.From.Column = GetFromPosition(false, columnsOrderedByIndex, placeholderColumnIndex, horizontalOffsetInPixels, edc);

            RowOrColumnInfo fromColumn = columnsOrderedByIndex[(int)position.From.Column.Index];
            position.To.Column = GetToPosition(false, fromColumn, position.From.Column.OffsetInEmus, widthInPixels, edc);
        }

        /// <summary>
        /// Gets the position of an offset (in pixels) from a supplied 'First Row or Column'
        /// </summary>
        /// <param name="isRow">Row (true) or Columns (false)</param>
        /// <param name="rowsOrColumns">List of row/columns</param>
        /// <param name="rowOrColumnIndex">The index of the row/column where the search starts.</param>
        /// <param name="startOffsetInPixels">Offset to find from the row/column index</param>
        /// <param name="edc">Converter to convert widths to actual widths</param>
        /// <returns>
        /// A structure containing a row/column index and offset in EMUs
        /// </returns>
        /// <exception cref="InvalidOperationException">negative placement not currently supported. Watcht this space...</exception>
        private static IndexOffset GetFromPosition(bool isRow, List<RowOrColumnInfo> rowsOrColumns, int rowOrColumnIndex, int startOffsetInPixels, ExcelDimensionConverter edc)
        {
            long runningHeightOrWidthInEmus = 0;
            int idx = rowOrColumnIndex - 1;
            RowOrColumnInfo rowOrColumn = rowsOrColumns[idx];

            // Maybe you have to consider padding when dealing with columns!!!
            long startOffsetInEmus = edc.PixelsToEmus(startOffsetInPixels);

            long remainingInEmus = startOffsetInEmus;

            while (idx >= 0 && idx < rowsOrColumns.Count)
            {
                rowOrColumn = rowsOrColumns[idx];

                // Determine whether the current remaining Emus would place the position in the previous, current, or next row/col
                if (remainingInEmus < 0)
                {
                    // We need to move back 1 column and re-calculate the remaining pixels
                    throw new InvalidOperationException("negative placement not currently supported. Watcht this space...");
                }
                else
                {
                    if (rowOrColumn.Hidden)
                    {
                        idx++;
                    }
                    else
                    {
                        if (rowOrColumn.HeightOrWidth.HasValue)
                        {
                            long heightOrWidthInEmus;

                            if (isRow)
                            {
                                var heightInPixels = edc.OpenXmlHeightToPixels(rowOrColumn.HeightOrWidth.Value);
                                heightOrWidthInEmus = edc.PixelsToEmus(heightInPixels);
                            }
                            else
                            {
                                var widthInPixels = edc.OpenXmlWidthToPixels(rowOrColumn.HeightOrWidth.Value);
                                heightOrWidthInEmus = edc.PixelsToEmus(widthInPixels);
                            }

                            runningHeightOrWidthInEmus += heightOrWidthInEmus;

                            if (runningHeightOrWidthInEmus > startOffsetInEmus)
                            {
                                // Have we exceeded the requested offset position.
                                idx = -1;
                            }
                            else
                            {
                                remainingInEmus -= heightOrWidthInEmus;
                                idx++;
                            }
                        }
                        else
                        {
                            // Encountered a non-assigned height or width, so we limit
                            // the height or width to the end of this cell.
                            idx = -1;
                        }
                    }
                }
            }

            // Build and return a structure for the calculated index + offset;
            return new IndexOffset((uint)rowOrColumn.ExcelIndex - 1, remainingInEmus);
        }

        /// <summary>
        /// Gets the position of an offset (in pixels) from a supplied 'First Row or Column'
        /// </summary>
        /// <param name="isRow">Row (true) or Columns (false)</param>
        /// <param name="fromRowOrColumn">From row or column.</param>
        /// <param name="startOffsetInEmus">The start offset in emus.</param>
        /// <param name="heightOrWidthInPixels">The height or width in pixels.</param>
        /// <param name="edc">Converter to convert widths to actual widths</param>
        /// <returns>
        /// A structure containing a row/column index and offset in EMUs
        /// </returns>
        private static IndexOffset GetToPosition(bool isRow, RowOrColumnInfo fromRowOrColumn, long startOffsetInEmus, int heightOrWidthInPixels, ExcelDimensionConverter edc)
        {
            long runningOffsetInEmus = 0;
            RowOrColumnInfo rowOrColumn = fromRowOrColumn;
            RowOrColumnInfo lastRowOrColumn = null;

            // You have to consider padding when dealing with columns
            long endOffsetInEmus = startOffsetInEmus + edc.PixelsToEmus(heightOrWidthInPixels);

            long remainingOffsetInEmus = endOffsetInEmus;

            while (rowOrColumn != null)
            {
                // Store the column bwing tested
                lastRowOrColumn = rowOrColumn;

                if (rowOrColumn.Hidden)
                {
                    rowOrColumn = rowOrColumn.Next;
                }
                else
                {
                    if (rowOrColumn.HeightOrWidth.HasValue)
                    {
                        long rowOrColumnHeightOrWidthInEmus;
                        int rowOrColumnHeightOrWidthInPixels;

                        if (isRow)
                        {
                            rowOrColumnHeightOrWidthInPixels = edc.HeightToPixels(rowOrColumn.HeightOrWidth.Value);
                            rowOrColumnHeightOrWidthInEmus = edc.PixelsToEmus(rowOrColumnHeightOrWidthInPixels);
                        }
                        else
                        {
                            rowOrColumnHeightOrWidthInPixels = edc.WidthToPixels(rowOrColumn.HeightOrWidth.Value);
                            rowOrColumnHeightOrWidthInEmus = edc.PixelsToEmus(rowOrColumnHeightOrWidthInPixels);
                        }

                        runningOffsetInEmus += rowOrColumnHeightOrWidthInEmus;

                        if (runningOffsetInEmus > endOffsetInEmus)
                        {
                            // Have we exceeded the requested offset position, no more columns
                            rowOrColumn = null;
                        }
                        else
                        {
                            rowOrColumn = rowOrColumn.Next;
                            if (rowOrColumn != null)
                            {
                                remainingOffsetInEmus -= rowOrColumnHeightOrWidthInEmus;
                            }
                        }
                    }
                    else
                    {
                        // Encountered a non-assigned height or width, so we limit
                        // the height or width to the end of this cell.
                        rowOrColumn = null;
                    }
                }
            }

            // Build and return a structure for the calculated index + offset;
            return new IndexOffset((uint)lastRowOrColumn.ExcelIndex - 1, remainingOffsetInEmus);
        }

        /// <summary>
        /// Gets the row or column1.
        /// </summary>
        /// <param name="rowsOrColumns">The rows or columns.</param>
        /// <returns></returns>
        private static RowOrColumnInfo GetRowOrColumn1(RowOrColumnInfoStore rowsOrColumns)
        {
            return rowsOrColumns.FirstOrDefault(c => c.ExcelIndex == 1);
        }

        /// <summary>
        /// Generates all of the charts that have been defined in the <see cref="Sheet" /> into the supplied <see cref="OpenXmlSpreadsheet.Worksheet" />.
        /// </summary>
        /// <param name="map">The XAML defined <see cref="BaseMap" /></param>
        /// <param name="targetWorksheet">The <see cref="OpenXmlSpreadsheet.Worksheet">worksheet in which the chart has been placed</see>. If null, the template chart sheet is used.</param>
        /// <param name="resourceStore">The <see cref="ResourceStore" /> where models are held</param>
        /// <param name="excelDimensionConverter">Used for making sense of column (and later row) heights in Excel.</param>
        private static void GenerateCharts(BaseMap map, OpenXmlSpreadsheet.Worksheet targetWorksheet, ResourceStore resourceStore, ExcelDimensionConverter excelDimensionConverter)
        {
            // Process all of the charts which have been defined/dynamically created in the sheet.
            List<Chart> charts = map.AllDescendentsOfType<Chart>();
            foreach (Chart chart in charts)
            {
                // If the chart was processed, there will be a placeholder to receive the chart
                if (chart.MapPlaceholder != null)
                {
                    ExcelMapCoOrdinatePlaceholder place = chart.MapPlaceholder;

                    // Look up the ChartModel from resources.
                    string chartTemplateKey = BindingContainer.ConvertToString(chart.ChartTemplateKey);
                    TempDiagnostics.Output(string.Format("Attempting to generate chart based on ChartTemplate Key='{0}'", chartTemplateKey));

                    ChartModel chartModel = null;
                    resourceStore.ChartModelDictionary.TryGetValue(chartTemplateKey, out chartModel);

                    if (chartModel != null)
                    {
                        // Clone the ChartModel, which will, in-turn, clone the chart,
                        // Find out where it is to go in the worksheet (a placeholder will have
                        // been created into which the chart can go), and move the cloned chart there.

                        // If no worksheet supplied, clone the chart within the model's worksheet.
                        if (targetWorksheet == null)
                        {
                            targetWorksheet = chartModel.Worksheet;
                        }

                        string targetSheetName = targetWorksheet.WorksheetPart.GetSheetName();

                        TempDiagnostics.Output(string.Format("Cloning Chart into sheet '{0}'", targetSheetName));
                        ChartModel chartModelClone = chartModel.Clone(targetWorksheet);

                        uint startRowIndex = (uint)place.ExcelRowStart - 1;
                        uint startColumnIndex = (uint)place.ExcelColumnStart - 1;

                        uint endRowIndex = place.GetEndRowIndex();
                        uint endColumnIndex = place.GetEndColumnIndex();

                        chartModelClone.Move(startRowIndex, startColumnIndex, endRowIndex, endColumnIndex);

                        // Set the title on the cloned chart
                        if (chart.Title != null)
                        {
                            chartModelClone.SetTitle(chart.Title.ToString());
                        }

                        // Update the chart with TableData.
                        if (chart.TableData != null)
                        {
                            // Update the cloned chart.
                            string tableDataSheet = targetWorksheet.WorksheetPart.GetSheetName();
                            ProcessDynamicChart(tableDataSheet, chart.TableData, chartModelClone);
                        }
                    }
                }
            }
        }

        #endregion Privates possibly - for move to openxml?

        #region Private non-static methods

        /// <summary>
        /// Always returns an instance, if not throws an exception.
        /// </summary>
        /// <param name="assemblyInfo">The assembly information string</param>
        /// <returns>
        /// A <see cref="IDocumentCustomProcess" />
        /// </returns>
        /// <exception cref="ExportException">
        /// </exception>
        private IDocumentCustomProcess SafeLoadDocumentCustomProcess(string assemblyInfo)
        {
            try
            {
                var parts = assemblyInfo.Split(';');
                if (parts.Length > 1)
                {
                    var part1 = parts[0];
                    var fullNamespace = part1.Replace("clr-namespace:", string.Empty);

                    var part2 = parts[1];
                    var assemblyName = part2.Replace("assembly=", string.Empty);

                    var assembly = Assembly.Load(assemblyName);
                    var type = assembly.GetType(fullNamespace);

                    object instance = Activator.CreateInstance(type, null);
                    var documentCustomProcess = instance as IDocumentCustomProcess;
                    if (documentCustomProcess == null)
                    {
                        throw new ExportException(string.Format("Custom process does not implement IDocumentCustomProcess <{0}>", assemblyInfo));
                    }

                    return documentCustomProcess;
                }

                throw new ExportException(string.Format("Unable to load custom process <{0}>", assemblyInfo));
            }
            catch (Exception ex)
            {
                throw new ExportException(string.Format("Error during load of custom assembly info <{0}>", assemblyInfo), ex);
            }
        }

        #region Export internal

        /// <summary>
        /// Creates a spreadsheet doc
        /// If a templateFilePath is supplied in the metadata it will be based on that, otherwise it'll be blank.
        /// For each entry in the set then data will be pumped into the data sheet.
        /// If a presentation sheet is defined then any dynamic chart series generation will be done,
        /// as well as any updates to formula or references, to ensure they're all pointing at the new data sheet.
        /// If mappings are defined they will be processed, this consists of moving a Drawing or Range of cells to
        /// a new location in the workbook. Generally this will be used for arranging parts into a predefined report.
        /// </summary>
        /// <param name="exportParameters">A set of parameters</param>
        /// <param name="metadata">Some <see cref="Book" /></param>
        /// <param name="sets">A list of <see cref="ExportTripleSet" />s</param>
        /// <param name="dataParts">A list of <see cref="IDataPart" />s</param>
        /// <param name="templatePackage">The template package.</param>
        /// <returns>
        /// A <see cref="MemoryStream" />
        /// </returns>
        /// <exception cref="ExportException">
        /// </exception>
        /// <exception cref="InvalidOperationException">You need at least one Map</exception>
        private MemoryStream ExcelExportInternal(ExportParameters exportParameters,
                                                 Book metadata,
                                                 List<ExportTripleSet> sets,
                                                 IEnumerable<IDataPart> dataParts,
                                                 ExcelTemplatePackage templatePackage)
        {
            // if this is set then make sure its tidied up at the end....
            bool initSheets = false;

            // base the spreadsheet based on a blank workbook, this may be overridable in the future
            MemoryStream stream = new MemoryStream();

            // if a template file (Excel Workbook) is provided, use this is a basis for the output
            if (metadata.HasTemplate)
            {
                stream.Write(metadata.TemplateData, 0, metadata.TemplateData.Length);
            }
            else
            {
                // otherwise use our blank sheet
                initSheets = true;
                stream.Write(Properties.Resources.BlankWorkbook, 0, Properties.Resources.BlankWorkbook.Length);
            }

            using (SpreadsheetDocument outputDocument = SpreadsheetDocument.Open(stream, true, new OpenSettings { AutoSave = true }))
            {
                // this is used during output of data to the worksheet - A StylesManager manages the styles in the document (workbook)
                var stylesManager = new ExcelStylesManager(outputDocument);

                BeginExport(initSheets, outputDocument);

                // if theres a pre process assembly defined call it
                if (!string.IsNullOrEmpty(metadata.PreProcessAssemblyInfo))
                {
                    var customDocProcess = this.SafeLoadDocumentCustomProcess(metadata.PreProcessAssemblyInfo);
                    customDocProcess.PreProcess(outputDocument, exportParameters, metadata, dataParts);
                }

                // build a dictionary of placeholders
                // used below during mapping, will all be removed after set processing
                Dictionary<string, MappingPlaceholder> placeholders = new Dictionary<string, MappingPlaceholder>();
                if (metadata.MappingPlaceholders != null)
                {
                    foreach (var dp in metadata.MappingPlaceholders)
                    {
                        if (!placeholders.ContainsKey(dp.Id))
                        {
                            placeholders.Add(dp.Id, dp);
                        }
                    }
                }

                // also build a dictionary of placeholder sets
                Dictionary<string, MappingPlaceholderSet> placeholderSets = new Dictionary<string, MappingPlaceholderSet>();
                if (metadata.MappingPlaceholderSets != null)
                {
                    foreach (var set in metadata.MappingPlaceholderSets)
                    {
                        if (!placeholderSets.ContainsKey(set.Id))
                        {
                            placeholderSets.Add(set.Id, set);
                        }
                    }
                }

                // Start off the timings
                TempDiagnostics.Output("==== Output tripple sets ===", true);

                // for each set
                foreach (var set in sets)
                {
                    IDataPart dataPart = set.DataPart;
                    ExportPart exportPart = set.Part;
                    Template template = set.Template;

                    // Update the StyleManager so that it is using the MapStyles which have been set against the Template
                    stylesManager.SetCurrentMapStyles(template.MapStyles);

                    // pull the spreadsheet doc out of storage for this template, need for sheet copying
                    var templateDocument = templatePackage.GetTemplateSpreadsheetDocumentByTemplateId(template.TemplateId);
                    if (templateDocument != null)
                    {
                        // keep track of all styles to add reuse
                        var mergedStyles = new Dictionary<uint, uint>();

                        OpenXmlSpreadsheet.Stylesheet sourceStylesheet = templateDocument.WorkbookPart.GetPartsOfType<WorkbookStylesPart>().First().Stylesheet;

                        // set the data on the export part..
                        // ..bindings will now be resolved
                        template.DataContext = dataPart.Data;

                        // if there's a title then override
                        // TODO: Change the 31 (and whole sheet name thing) so that we can't get silly worksheet names (duplicated etc)
                        string title = template.Title == null ? string.Empty : template.Title.ToString();
                        string worksheetName = title.ToString().Substring(0, title.Length > 31 ? 31 : title.Length);
                        if (!string.IsNullOrEmpty(worksheetName))
                        {
                            if (!string.IsNullOrEmpty(template.PresentationTemplateSheet))
                            {
                                ////TODO: Change the 31 (and whole sheet name thing) so that we can't get silly worksheet names (duplicated etc)
                                string dataSheetName = string.Concat(worksheetName, " Data");
                                exportPart.DataSheetName = dataSheetName.Substring(0, dataSheetName.Length > 31 ? 31 : dataSheetName.Length);
                                exportPart.PresentationSheetName = worksheetName;
                            }
                            else
                            {
                                exportPart.DataSheetName = worksheetName;
                            }
                        }

                        // clone a presentation worksheet into the output document, null will be returned if not needed
                        WorksheetPart presentationWSPart = TryCloneSheet(outputDocument, exportPart.PresentationSheetName, templateDocument, template.PresentationTemplateSheet);
                        if (presentationWSPart != null)
                        {
                            MergeResourcesAndStyles(presentationWSPart, templateDocument, outputDocument, mergedStyles);
                        }

                        // try clone the data worksheet into the output document - the sheet gets cloned into the 'document' (ie. the destination spreadsheet document) 
                        WorksheetPart dataWSPart = TryCloneSheet(outputDocument, exportPart.DataSheetName, templateDocument, template.DataTemplateSheet);

                        // if there's been nothing to clone, then create a new one
                        // this might be because there's no matching sheet in the template or the output doc
                        if (dataWSPart == null)
                        {
                            dataWSPart = outputDocument.CreateOrOpenSheet(exportPart.DataSheetName);
                        }

                        if (dataWSPart != null)
                        {
                            MergeResourcesAndStyles(dataWSPart, templateDocument, templateDocument, mergedStyles);
                        }
                        else
                        {
                            throw new ExportException(string.Format("No Data WorksheetPart for TemplateId <{0}> for PartId <{1}>", template.TemplateId, exportPart.PartId));
                        }

                        // *********************************************************************************************
                        // * At this point, we should have a worksheet into which things can be written (dataWSPart).
                        // *********************************************************************************************
                        ResourceStore resourceStore = null;

                        // if there's a map provided
                        if (template != null)
                        {
                            // Load resources from the TemplateCollection (legacy)
                            resourceStore = ResourceStore.Create(template.TemplateCollection, templatePackage);

                            // Create an object which can write to the worksheet
                            var mapperHelper = new ExcelSheetMapper
                            (
                                exportPart.DataSheetName,
                                dataWSPart,
                                null,
                                outputDocument,
                                stylesManager,
                                resourceStore
                            );

                            // Object that is used to manage the map creation and layout
                            mapperHelper.ProcessMap(template, dataPart, exportPart);
                        }
                        else
                        {
                            throw new InvalidOperationException("You need at least one Map");
                        }

                        // *********************************************************************************************************************
                        // * TODO: Change the ExportViewItem processing so that it uses the ExcelMapCoOrdinates (as does the ProcessExcelMap)  *
                        // *********************************************************************************************************************

                        // check if there are any mappings
                        if (exportPart.Mappings != null && exportPart.Mappings.Count > 0)
                        {
                            // work through each mapping
                            foreach (var mapping in exportPart.Mappings)
                            {
                                MappingPlaceholder dp = null;

                                // if there's a set of placeholders specified
                                if (!string.IsNullOrEmpty(mapping.PlaceholderSetId))
                                {
                                    if (!placeholderSets.ContainsKey(mapping.PlaceholderSetId))
                                    {
                                        throw new ExportException(string.Format("Unable to find PlaceholderSetId {0} for TemplateId <{1}> for PartId <{2}>", mapping.PlaceholderSetId, template.TemplateId, exportPart.PartId));
                                    }
                                    else
                                    {
                                        // get the next in the set, helps with flow of layout for optional parts
                                        dp = GeNextPlaceholderFromSet(placeholderSets[mapping.PlaceholderSetId], placeholders, outputDocument);
                                    }
                                }
                                else
                                {
                                    if (!placeholders.ContainsKey(mapping.PlaceholderId))
                                    {
                                        throw new ExportException(string.Format("Unable to find PlaceholderId {0} for TemplateId <{1}> for PartId <{2}>", mapping.PlaceholderId, template.TemplateId, exportPart.PartId));
                                    }
                                    else
                                    {
                                        dp = placeholders[mapping.PlaceholderId];
                                    }
                                }

                                // if dp cant be found, exception
                                if (dp == null)
                                {
                                    throw new ExportException(string.Format("Failed to find the mapping placeholder for TemplateId <{0}> for PartId <{1}>", template.TemplateId, exportPart.PartId));
                                }

                                // use the information in the placeholder to get the target report worksheet
                                var reportWSPart = outputDocument.WorkbookPart.Workbook.GetWorksheetPartByName(dp.SheetName);

                                // either moving a range of cells from one sheet to another
                                if (mapping is RangeMapping)
                                {
                                    var rangeMapping = (RangeMapping)mapping;

                                    // if this is a range mapping from the presentation sheet make sure there is one
                                    if (rangeMapping.UsePresentationAsSource && presentationWSPart == null)
                                    {
                                        throw new ExportException(string.Format("No presentation worksheet part during range mapping for PlaceholderId {0} for TemplateId <{1}> for PartId <{2}>", mapping.PlaceholderId, template.TemplateId, exportPart.PartId));
                                    }

                                    var sourceWSPart = rangeMapping.UsePresentationAsSource ? presentationWSPart : dataWSPart;

                                    ProcessRangeMapping(set, rangeMapping, sourceWSPart, reportWSPart);
                                }
                                else if (mapping is ParagraphMapping)
                                {
                                    // or transfering a XAML document in a cell to an Excel Drawing object text
                                    if (presentationWSPart == null)
                                    {
                                        throw new ExportException(string.Format("No presentation worksheet part found during paragraph mapping for PlaceholderId {0} for TemplateId <{1}> for PartId <{2}>", mapping.PlaceholderId, template.TemplateId, exportPart.PartId));
                                    }

                                    var paragraphMapping = (ParagraphMapping)mapping;

                                    DrawingSpreadsheet.Shape targetShape = ProcessParagraphPart(outputDocument.WorkbookPart.Workbook, exportPart, dp, paragraphMapping, presentationWSPart, reportWSPart);
                                }
                                else if (mapping is DrawingMapping)
                                {
                                    // or moving a drawing (ie a chart) from one sheet to another
                                    if (presentationWSPart == null)
                                    {
                                        throw new ExportException(string.Format("No presentation worksheet part found during drawing mapping for PlaceholderId {0} for TemplateId <{1}> for PartId <{2}>", mapping.PlaceholderId, template.TemplateId, exportPart.PartId));
                                    }

                                    var drawingMapping = (DrawingMapping)mapping;

                                    var newChartPart = ProcessDrawingPart(exportPart, dp, drawingMapping, presentationWSPart, reportWSPart);
                                    if (newChartPart != null)
                                    {
                                        // Update any formula or references with the new sheet name
                                        newChartPart.UpdateSources(template.DataTemplateSheet, exportPart.DataSheetName, new int?(dataPart.RowCount));

                                        // if there's a LegendPlaceholderId
                                        if (!string.IsNullOrEmpty(drawingMapping.LegendPlaceholderId))
                                        {
                                            // TBD....not implemented yet
                                            ProcessLegend(newChartPart, drawingMapping.LegendPlaceholderId);
                                        }
                                    }
                                }
                            }
                        }
                        else
                        {
                            // No mappings so just a do dynamic chart + shape processing
                            GenerateModelBasedElements(template, dataWSPart, resourceStore);

                            // Do source updates direct on the presentation sheet
                            if (presentationWSPart != null)
                            {
                                // then perform any dynamic charting logic if needed...
                                // For the moment, Find the first Table in the supplied template, and use this as the basis of our chart.
                                // TODO: Identify the table explicitly, then if not explicitly set, use first table...
                                Table firstTable = template.FirstDescendentOfType<Table>();
                                TableData tableData = firstTable.TableData;

                                if (tableData != null)
                                {
                                    // Process chart based on TableData
                                    string dataSheetName = set.Part.DataSheetName;
                                    ProcessDynamicChart(set.DataPart, dataSheetName, tableData, presentationWSPart);

                                    // ...and then update any formula or references with the new sheet name
                                    bool treatRowAsSeries = tableData.TreatRowAsSeries;

                                    presentationWSPart.UpdateSources(template.DataTemplateSheet,
                                                                     exportPart.DataSheetName,
                                                                     treatRowAsSeries ? null : new int?(tableData.RowData.Count));
                                }
                            }
                        }

                        // remove the pres sheet if there are mappings, and the suppress of this behaviour hasnt been overridden.
                        if (exportPart.Mappings.Count > 0 && !exportPart.SuppressPresentationRemoval)
                        {
                            outputDocument.WorkbookPart.DeleteSheet(exportPart.PresentationSheetName);
                        }

                        // hide the data sheet if thats whats wanted
                        if (exportPart.DataSheetHidden)
                        {
                            var sheet = outputDocument.WorkbookPart.Workbook.GetSheetByName(exportPart.DataSheetName);
                            if (sheet != null)
                            {
                                sheet.State = OpenXmlSpreadsheet.SheetStateValues.Hidden;
                            }
                        }
                    }
                }

                // Any code to set row height to 0 for DataTables needs to go here.

                // Hide all rows in matching ranges where there is no source data.
                HideTablesWithNoData(outputDocument.WorkbookPart.Workbook);

                // Hide superfluous rows in matching ranges where source provides less data than destination
                HideTableRowsWithNoData(outputDocument.WorkbookPart.Workbook);

                // Translate XAML RTF Into Inline text at the position of the defined name
                //ProcessXAMLRTFToInlineText(outputDocument.WorkbookPart.Workbook);

                // remove any placeholders now that processing has finished
                foreach (var id in placeholders.Keys)
                {
                    var worksheetPart = outputDocument.WorkbookPart.Workbook.GetWorksheetPartByName(placeholders[id].SheetName);
                    if (worksheetPart != null)
                    {
                        var shape = worksheetPart.GetShapeByName(id);
                        if (shape != null)
                        {
                            shape.Parent.Remove();
                        }
                    }
                }

                // if theres a pre process assembly defined call it
                if (!string.IsNullOrEmpty(metadata.PostProcessAssemblyInfo))
                {
                    var customDocProcess = this.SafeLoadDocumentCustomProcess(metadata.PostProcessAssemblyInfo);
                    customDocProcess.PostProcess(outputDocument, exportParameters, metadata, dataParts);
                }

                EndExport(outputDocument);
            }

            templatePackage.Flush();

            return stream;
        }

        /// <summary>
        /// Generates a <see cref="MemoryStream" /> containing an Excel document
        /// </summary>
        /// <param name="exportParameters">A set of <see cref="ExportParameters" /></param>
        /// <param name="metadata">Information about how to marry data to XAML derived templates</param>
        /// <param name="dataParts">The Data</param>
        /// <param name="resourcePackage">A set of re-usable resources</param>
        /// <returns>
        /// A <see cref="MemoryStream" /> containing an Excel document
        /// </returns>
        /// <exception cref="MetadataException">No sheet name. Must be supplied if no part id set</exception>
        private MemoryStream GenerateExcelInternal(ExportParameters exportParameters,
                                                   ExcelDocumentMetadata metadata,
                                                   IEnumerable<IDataPart> dataParts,
                                                   ResourcePackage resourcePackage)
        {
            // if this is set then make sure its tidied up at the end....
            bool initSheets = false;

            // base the spreadsheet based on a blank workbook, this may be overridable in the future
            var stream = new MemoryStream();

            // if a template file (Excel Workbook) is provided, use this as a basis for the output
            if (metadata.HasTemplate)
            {
                stream.Write(metadata.TemplateData, 0, metadata.TemplateData.Length);
            }
            else
            {
                // otherwise use our blank sheet
                initSheets = true;
                stream.Write(Properties.Resources.BlankWorkbook, 0, Properties.Resources.BlankWorkbook.Length);
            }

            using (SpreadsheetDocument outputDocument = SpreadsheetDocument.Open(stream, true, new OpenSettings { AutoSave = true }))
            {
                metadata.MergeResources(resourcePackage);
                var resourceStore = metadata.ResourceStore;

                // this is used during output of data to the worksheet - A StylesManager manages the styles in the document (workbook)
                var stylesManager = new ExcelStylesManager(outputDocument);
                stylesManager.Initialise(resourceStore);

                BeginExport(initSheets, outputDocument);

                // if theres a pre process assembly defined call it
                if (!string.IsNullOrEmpty(metadata.PreProcessAssemblyInfo))
                {
                    var customDocProcess = this.SafeLoadDocumentCustomProcess(metadata.PreProcessAssemblyInfo);
                    customDocProcess.PreProcess(outputDocument, exportParameters, metadata, dataParts);
                }

                // for each set
                foreach (var sheet in metadata.Sheets)
                {
                    // no partid on the sheet just process it
                    if (string.IsNullOrEmpty(sheet.PartId))
                    {
                        if (sheet.SheetName == null)
                        {
                            throw new MetadataException("No sheet name. Must be supplied if no part id set");
                        }

                        this.ProcessSheet(sheet, dataParts, outputDocument, stylesManager, resourceStore);
                    }
                    else
                    {
                        var matches = from dp in dataParts
                                      where sheet.PartId.CompareTo(dp.PartId) == 0
                                      select dp;

                        foreach (var match in matches)
                        {
                            var clonedSheet = metadata.GetSheetByInternalId(sheet.InternalId);
                            clonedSheet.DataContext = match.Data;
                            this.ProcessSheet(clonedSheet, dataParts, outputDocument, stylesManager, resourceStore);
                        }
                    }
                }

                // if theres a pre process assembly defined call it
                if (!string.IsNullOrEmpty(metadata.PostProcessAssemblyInfo))
                {
                    var customDocProcess = this.SafeLoadDocumentCustomProcess(metadata.PostProcessAssemblyInfo);
                    customDocProcess.PostProcess(outputDocument, exportParameters, metadata, dataParts);
                }

                EndExport(outputDocument);
            }

            // if a resource package has been supplied then flush it
            if (resourcePackage != null)
            {
                resourcePackage.Flush();
            }

            return stream;
        }

        /// <summary>
        /// Process the sheet
        /// </summary>
        /// <param name="sheet">The <see cref="Sheet" /></param>
        /// <param name="dataParts">A list of DataParts</param>
        /// <param name="outputDocument">The output <see cref="SpreadsheetDocument" /></param>
        /// <param name="stylesManager">A manager which manages style resources</param>
        /// <param name="resourceStore">A store containing all re-usable resources</param>
        private void ProcessSheet(Sheet sheet, IEnumerable<IDataPart> dataParts, SpreadsheetDocument outputDocument, ExcelStylesManager stylesManager, ResourceStore resourceStore)
        {
            var sheetName = BindingContainer.ConvertToString(sheet.SheetName);
            sheetName = sheetName.Substring(0, sheetName.Length > 31 ? 31 : sheetName.Length);

            var worksheetPart = outputDocument.CreateOrOpenSheet(sheetName);

            if (sheet.Content != null)
            {
                sheet.Content.DataContext = sheet.DataContext;

                // Create an object which can write to the worksheet
                // Object that is used to manage the map creation and layout
                var mapperHelper = new ExcelSheetMapper
                (
                    sheetName,
                    worksheetPart,
                    dataParts,
                    outputDocument,
                    stylesManager,
                    resourceStore
                );
                mapperHelper.ProcessMap(sheet.Content);

                // Generate charts and shapes in the target worksheet.
                GenerateModelBasedElements(sheet.Content, worksheetPart, resourceStore);
            }
        }

        #endregion Export Internal

        #endregion Private non-static methods
    }
}
