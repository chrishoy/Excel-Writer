namespace ExcelWriter
{
    using System;
    using System.Collections.Generic;
    using DocumentFormat.OpenXml.Packaging;
    using OpenXml.Excel;

    using OpenXmlSpreadsheet = DocumentFormat.OpenXml.Spreadsheet;

    /// <summary>
    /// 
    /// </summary>
    internal sealed class ExcelMapWriter
    {
        /// <summary>
        /// Write the content of the <see cref="ExcelMapCoOrdinateContainer"> to the worksheet</see>.
        /// </summary>
        /// <param name="sheetName">Name of the sheet.</param>
        /// <param name="worksheetPart">The worksheet part.</param>
        /// <param name="mapCoOrdinateContainer">The map co ordinate container.</param>
        /// <param name="stylesManager">The styles manager.</param>
        /// <param name="spreadsheetDocument">The spreadsheet document.</param>
        public static void WriteMapToExcel(string sheetName, 
                                           WorksheetPart worksheetPart, 
                                           ExcelMapCoOrdinateContainer mapCoOrdinateContainer, 
                                           ExcelStylesManager stylesManager, 
                                           SpreadsheetDocument spreadsheetDocument)
        {
            TempDiagnostics.Output(string.Format("Writing map for ExcelMap[{0}] to worksheet '{1}'", mapCoOrdinateContainer.ContainerType, sheetName));

            var dimensionConverter = new ExcelDimensionConverter("Calibri", 11.0f);

            // Manages the writing to Excel.
            var excelWriteManager = new OpenXmlExcelWriteManager(worksheetPart.Worksheet);

            //** First we need to clear the destination worksheet down (rows, columns and merged areas)
            excelWriteManager.EmptyWorksheet();

            // Build up Columns models on all nested containers.
            RowOrColumnsModel columnsModel = mapCoOrdinateContainer.BuildColumnsModel();
            int colCount = columnsModel.Count();

            // ==========================================================
            // Traverse each column assigning start and end 
            // column indexes to the maps associated with that column.
            // ==========================================================
            RowOrColumnInfo columnInfo = columnsModel.First;
            while (columnInfo != null)
            {
                double? width = columnInfo.HeightOrWidth.HasValue ? (double?)dimensionConverter.WidthToOpenXmlWidth(columnInfo.HeightOrWidth.Value) : null;
                excelWriteManager.AddColumn(width, columnInfo.Hidden);

                // Assign Excel start and end columns to maps for merge operations.
                foreach (ExcelMapCoOrdinate map in columnInfo.Maps)
                {
                    // Extend any cells that are merged across maps.
                    if (map is ExcelMapCoOrdinatePlaceholder)
                    {
                        var cell = map as ExcelMapCoOrdinatePlaceholder;
                        if (cell.MergeWith != null)
                        {
                            cell.ExcelColumnStart = cell.MergeWith.ExcelColumnStart;
                            cell.MergeWith.ExcelColumnEnd = columnInfo.ExcelIndex; //excelCol
                        }
                    }

                    if (map.ExcelColumnStart == 0)
                    {
                        map.ExcelColumnStart = columnInfo.ExcelIndex; //excelCol;
                        map.ExcelColumnEnd = columnInfo.ExcelIndex; //excelCol;
                    }
                    else
                    {
                        map.ExcelColumnEnd = columnInfo.ExcelIndex; //excelCol;
                    }
                }

                // Move on to next column in list
                columnInfo = columnInfo.Next;
            }
            TempDiagnostics.Output(string.Format("Created ColumnsModel which contains '{0}' columns", colCount));

            // Build up Rows models on all nested containers.
            RowOrColumnsModel rowsModel = mapCoOrdinateContainer.BuildRowsModel();
            int rowCount = rowsModel.Count();

            // ==========================================================
            // Traverse each row assigning start and end 
            // row indexes to the maps associated with that row.
            // ==========================================================
            RowOrColumnInfo rowInfo = rowsModel.First;
            while (rowInfo != null)
            {
                double? height = rowInfo.HeightOrWidth.HasValue ? (double?)dimensionConverter.HeightToOpenXmlHeight(rowInfo.HeightOrWidth.Value) : null;
                excelWriteManager.AddRow(height, rowInfo.Hidden);

                // Assign Excel start and end rows to maps for merge operations.
                foreach (ExcelMapCoOrdinate map in rowInfo.Maps)
                {
                    if (map.ExcelRowStart == 0)
                    {
                        map.ExcelRowStart = rowInfo.ExcelIndex; //excelRow;
                        map.ExcelRowEnd = rowInfo.ExcelIndex; //excelRow;
                    }
                    else
                    {
                        map.ExcelRowEnd = rowInfo.ExcelIndex; //excelRow;
                    }
                }

                // Move on to next Row in list
                rowInfo = rowInfo.Next;
            }
            TempDiagnostics.Output(string.Format("Created RowsModel which contains '{0}' rows", rowCount));

            // Build a layered cells dictionary for all cells, keyed by Excel row and column index
            var layeredCellsDictionary = new LayeredCellsDictionary();
            mapCoOrdinateContainer.UpdateLayeredCells(ref layeredCellsDictionary);
            TempDiagnostics.Output(string.Format("Updated Layered Cell Information for worksheet '{0}' = {1}", sheetName, layeredCellsDictionary.Count));

            // Probe the Row, Column and Cell Layered maps, embellishing them with row, column and cell formatting information.
            for (uint worksheetRow = 1; worksheetRow <= rowCount; worksheetRow++)
            {
                for (uint worksheetCol = 1; worksheetCol <= colCount; worksheetCol++)
                {
                    // We can now use the layeredCellsDictionary to build the
                    // excel workbook based on layered cell information
                    var currentCoOrdinate = new System.Drawing.Point((int)worksheetCol, (int)worksheetRow);
                    LayeredCellInfo layeredCellInfo = layeredCellsDictionary[currentCoOrdinate];

                    // Work through the layered maps to determine what needs to be written to the Excel worksheet at that Row/Column
                    ProcessLayeredCellMaps(currentCoOrdinate, layeredCellInfo);
                }
            }
            TempDiagnostics.Output(string.Format("Built Worksheet CellInfos[Cols={0},Rows={1}]", colCount, rowCount));

            //** Write the 2D array of cell information to the Excel Worksheet
            //** building a list of areas that are to be merged. 
            for (uint worksheetRow = 1; worksheetRow <= rowCount; worksheetRow++)
            {
                OpenXmlSpreadsheet.Row row = excelWriteManager.GetRow(worksheetRow);

                for (uint worksheetCol = 1; worksheetCol <= colCount; worksheetCol++)
                {
                    // Pluck out information relating to the current cell
                    var currentCoOrdinate = new System.Drawing.Point((int)worksheetCol, (int)worksheetRow);
                    ExcelCellInfo cellInfo = layeredCellsDictionary[currentCoOrdinate].CellInfo;
                    
                    // Not sure if we need this as column letters are easily (quickly) translated using existing OpenXml helpers
                    string columnLetter = excelWriteManager.GetColumnLetter(worksheetCol);

                    // All merge cells should have the same style as the source merge cell.
                    if (cellInfo.MergeFrom != null)
                    {
                        // If MergeFrom is non-null, then this cell has been already been marked as a merge cell,
                        // whose style is to be the same as the source 'MergeFrom cell.
                        // Update the source (MergeFrom) cell so it ends up containing a reference to the last (MergeTo) cell to be merged.
                        cellInfo.MergeFrom.MergeTo = cellInfo;

                        // Create/lookup styles in the target workbook and return index of created style
                        uint cellStyleIndex = stylesManager.GetOrCreateStyle(cellInfo.MergeFrom.StyleInfo);

                        // Write in to Excel (style information only)
                        var cell = new OpenXmlSpreadsheet.Cell();
                        cell.SetCellReference(columnLetter, worksheetRow);
                        cell.StyleIndex = cellStyleIndex;
                        row.Append(cell);
                        cellInfo.Cell = cell;
                    }
                    else
                    {
                        // Create/lookup styles in the target workbook and return index of created style
                        uint cellStyleIndex = stylesManager.GetOrCreateStyle(cellInfo.StyleInfo);

                        // NB! If we write a Null value then we lose cell formatting information for some reason
                        object value = cellInfo.Value == null ? string.Empty : cellInfo.Value;

                        // Write in to Excel
                        var cell = new OpenXmlSpreadsheet.Cell();
                        cell.SetCellReference(columnLetter, worksheetRow);
                        excelWriteManager.SetCellValue(cell, value);
                        cell.StyleIndex = cellStyleIndex;
                        row.Append(cell);
                        cellInfo.Cell = cell;
                    }

                    // Span the cell accross it's parent columns if specified
                    if (cellInfo.LastSpanRow == 0) cellInfo.LastSpanRow = worksheetRow;
                    if (cellInfo.LastSpanColumn == 0) cellInfo.LastSpanColumn = worksheetCol;

                    // Merge cells if required (spanning)
                    for (uint rowIdx = worksheetRow; rowIdx <= cellInfo.LastSpanRow; rowIdx++)
                    {
                        for (uint colIdx = worksheetCol; colIdx <= cellInfo.LastSpanColumn; colIdx++)
                        {
                            // Mark processed (so we don't over-write)
                            var coOrdinate = new System.Drawing.Point((int)colIdx, (int)rowIdx);
                            if (coOrdinate != currentCoOrdinate)
                            {
                                var mergeCoOrdinate = new System.Drawing.Point((int)colIdx, (int)rowIdx);
                                ExcelCellInfo mergeCellInfo = layeredCellsDictionary[mergeCoOrdinate].CellInfo;
                                if (mergeCellInfo.MergeFrom == null) mergeCellInfo.MergeFrom = cellInfo;
                            }
                        }
                    }
                }
            }

//#if DEBUG
//            if (!skipMerge)
//            { 
//#endif
            // Merge cells that have been marked to merge in the cellInfos dictionary.
            MergeCells(worksheetPart.Worksheet, (uint)rowCount, (uint)colCount, layeredCellsDictionary);
//#if DEBUG
//            } 
//#endif
            TempDiagnostics.Output(string.Format("Written CellInfos[Cols={0},Rows={1}] to Excel", colCount, rowCount));

            // Create a list of all of the the defined names in the container.
            var definedNameList = new List<ExcelDefinedNameInfo>();
            mapCoOrdinateContainer.UpdateDefinedNameList(ref definedNameList, sheetName);

            // And write into Excel workbook
            foreach (var definedNameInfo in definedNameList)
            {
                uint rangeColumnCount = definedNameInfo.EndColumnIndex - definedNameInfo.StartColumnIndex + 1;
                uint rangeRowCount = definedNameInfo.EndRowIndex - definedNameInfo.StartRowIndex + 1;

                if (rangeRowCount > 0 && rangeColumnCount > 0)
                {
                    spreadsheetDocument.WorkbookPart.Workbook.AddDefinedName(sheetName,
                                                                            definedNameInfo.DefinedName,
                                                                            (uint)definedNameInfo.StartColumnIndex,
                                                                            (uint)definedNameInfo.StartRowIndex,
                                                                            (int)rangeColumnCount,
                                                                            (int)rangeRowCount);
                }
            }

            TempDiagnostics.Output(string.Format("Defined Names added - write map for {0} complete...", mapCoOrdinateContainer.ContainerType));
        }

        /// <summary>
        /// Merges all cells that have been marked to be merged in the cellinfos dictionary.
        /// </summary>
        /// <param name="worksheet">The worksheet.</param>
        /// <param name="totalRowCount">The total number of rows in the worksheet.</param>
        /// <param name="totalColCount">The total number of columns in the worksheet.</param>
        /// <param name="cellInfos">A dictionary of cell information, keyed by row and column index.</param>
        private static void MergeCells(OpenXmlSpreadsheet.Worksheet worksheet, uint totalRowCount, uint totalColCount, LayeredCellsDictionary cellInfos)
        {
            OpenXmlSpreadsheet.SheetData sheetData = worksheet.GetFirstChild<OpenXmlSpreadsheet.SheetData>();

            // Process the collection of merged cells.
            for (uint worksheetRow = 1; worksheetRow <= totalRowCount; worksheetRow++)
            {
                for (uint worksheetCol = 1; worksheetCol <= totalColCount; worksheetCol++)
                {
                    // Get cellInfo, if it has a MergeTo then merge it.
                    var currentCoOrdinate = new System.Drawing.Point((int)worksheetCol, (int)worksheetRow);
                    ExcelCellInfo cellInfo = cellInfos[currentCoOrdinate].CellInfo;

                    if (cellInfo.MergeTo != null)
                    {
                        sheetData.MergeCells(cellInfo.Cell, cellInfo.MergeTo.Cell);
                    }
                }
            }
        }

        /// <summary>
        /// Walks down the list of layered maps, and determines information that is to be written into a single cell.
        /// </summary>
        /// <param name="coOrdinate">The co ordinate.</param>
        /// <param name="layeredCellInfo">The layered cell information.</param>
        private static void ProcessLayeredCellMaps(System.Drawing.Point coOrdinate, LayeredCellInfo layeredCellInfo)
        {
            var cellInfo = new ExcelCellInfo();

            System.Drawing.Point toCoOrdinate = coOrdinate;

            // Work up the layers (ie. Cell up to Worksheet)
            foreach (var map in layeredCellInfo.LayeredMaps)
            {
                // For cells, read the value
                if (map is ExcelMapCoOrdinateCell)
                {
                    var cell = ((ExcelMapCoOrdinateCell)map);
                    cellInfo.Value = cell.CurrentValue;
                }

                // For Placeholder derived maps (cells + placeholders),
                // determine the 'Span to' information for row/column + layer the style
                if (map is ExcelMapCoOrdinatePlaceholder)
                {
                    var placeholder = ((ExcelMapCoOrdinatePlaceholder)map);

                    // If this lower level cell is flagged to span (to end of the higher level containing element), then apply
                    if (placeholder.SpanLastColumn)
                    {
                        cellInfo.LastSpanColumn = placeholder.GetEndColumnIndex(); // higherLayerLastColumn
                        toCoOrdinate.X = (int)cellInfo.LastSpanColumn;
                    }
                    else
                    {
                        cellInfo.LastSpanColumn = (uint)placeholder.ExcelColumnEnd;
                        toCoOrdinate.X = (int)cellInfo.LastSpanColumn;
                    }

                    // If this lower level cell is flagged to span (to end of the higher level containing element), then apply
                    if (placeholder.SpanLastRow)
                    {
                        cellInfo.LastSpanRow = placeholder.GetEndRowIndex(); // higherLayerLastRow;
                        toCoOrdinate.Y = (int)cellInfo.LastSpanRow;
                    }
                    else
                    {
                        cellInfo.LastSpanRow = (uint)placeholder.ExcelRowEnd;
                        toCoOrdinate.Y = (int)cellInfo.LastSpanRow;
                    }

                    // A cell, extract the value and set the base styles
                    cellInfo.StyleInfo.ApplyCellStyles(placeholder.Styles);
                }

                // For containers, layer the style, taking into consideration the
                // position of the current co-ordinate relative to (within) the container
                if (map is ExcelMapCoOrdinateContainer)
                {
                    // A container, update the base style
                    var container = ((ExcelMapCoOrdinateContainer)map);
                    cellInfo.StyleInfo.Update(container, coOrdinate.X, coOrdinate.Y, toCoOrdinate.X, toCoOrdinate.Y);
                }
            }

            // Update the layered cell with the result
            layeredCellInfo.CellInfo = cellInfo;
        }
    }
}
