namespace ExcelWriter
{
    using System;
    using System.Collections;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text;

    using DocumentFormat.OpenXml.Packaging;

    using OpenXmlSpreadsheet = DocumentFormat.OpenXml.Spreadsheet;
    using OpenXml.Excel;

    using ExportMap = ExcelWriter;
    using ExcelModel = OpenXml.Excel.Model;

    /// <summary>
    /// 
    /// </summary>
    internal sealed class ExcelSheetMapper
    {
        #region Private Fields

        /// <summary>
        /// The sheet name
        /// </summary>
        private readonly string sheetName;
        /// <summary>
        /// The worksheet part
        /// </summary>
        private readonly WorksheetPart worksheetPart;
        /// <summary>
        /// The data parts
        /// </summary>
        private readonly IEnumerable<IDataPart> dataParts;
        /// <summary>
        /// The spreadsheet document
        /// </summary>
        private readonly SpreadsheetDocument spreadsheetDocument;
        /// <summary>
        /// The styles manager
        /// </summary>
        private readonly ExcelStylesManager stylesManager;
        /// <summary>
        /// The resource store
        /// </summary>
        private readonly ResourceStore resourceStore;

        /// <summary>
        /// The stylesheet
        /// </summary>
        private OpenXmlSpreadsheet.Stylesheet stylesheet;

        /// <summary>
        /// The worksheet co ordinates
        /// </summary>
        private ExcelMapCoOrdinateContainer worksheetCoOrdinates;

        /// <summary>
        /// The legacy process
        /// </summary>
        private bool legacyProcess;
        /// <summary>
        /// The data part
        /// </summary>
        private IDataPart dataPart;
        /// <summary>
        /// The export part
        /// </summary>
        private ExportPart exportPart;

        #endregion Private Fields

        #region Constructor

        /// <summary>
        /// Initializes a new instance of the <see cref="ExcelSheetMapper" /> class.
        /// </summary>
        /// <param name="sheetName">Name of the sheet.</param>
        /// <param name="worksheetPart">The worksheet part.</param>
        /// <param name="dataParts">The data parts.</param>
        /// <param name="spreadsheetDocument">The spreadsheet document.</param>
        /// <param name="stylesManager">The styles manager.</param>
        /// <param name="resoureStore">The resoure store.</param>
        /// <exception cref="ArgumentNullException">
        /// sheetName
        /// or
        /// worksheetPart
        /// or
        /// spreadsheetDocument
        /// or
        /// stylesManager
        /// or
        /// resoureStore
        /// </exception>
        public ExcelSheetMapper(string sheetName,
                                WorksheetPart worksheetPart,
                                IEnumerable<IDataPart> dataParts,
                                SpreadsheetDocument spreadsheetDocument,
                                ExcelStylesManager stylesManager,
                                ResourceStore resoureStore)
        {
            if (string.IsNullOrEmpty(sheetName)) throw new ArgumentNullException("sheetName");
            if (worksheetPart == null) throw new ArgumentNullException("worksheetPart");
            if (spreadsheetDocument == null) throw new ArgumentNullException("spreadsheetDocument");
            if (stylesManager == null) throw new ArgumentNullException("stylesManager");
            if (resoureStore == null) throw new ArgumentNullException("resoureStore");

            // Reset the global counter used for creating Id's
            Counter.Reset();

            this.sheetName = sheetName;
            this.worksheetPart = worksheetPart;
            this.dataParts = dataParts;
            this.resourceStore = resoureStore;
            this.spreadsheetDocument = spreadsheetDocument;
            this.stylesManager = stylesManager;
            this.stylesheet = this.spreadsheetDocument.WorkbookPart.GetPartsOfType<WorkbookStylesPart>().First().Stylesheet;
        }

        #endregion Constructor

        #region Public Methods

        /// <summary>
        /// Entry point to process an Excel Map.<br />
        /// Map is rooted at R1C1 of the worksheet.
        /// </summary>
        /// <param name="map">The map.</param>
        public void ProcessMap(BaseMap map)
        {
            this.legacyProcess = false;
            this.dataPart = null;
            this.exportPart = null;

            map.ResourceKey = map.Key;

            // Create an entity which will map elements to a worksheet.
            this.worksheetCoOrdinates = new ExcelMapCoOrdinateContainer(1, 1, map.GetContainerType()); // 1 x 1 map which represents the worksheet

            TempDiagnostics.Output(string.Format("Populating map for worksheet '{0}'", this.sheetName));

            // Recursively process the supplied map and co-ordinates.
            this.ProcessMapInternal(map, null, this.worksheetCoOrdinates);

            // Write the contents of the map into the Excel Worksheet
            ExcelMapWriter.WriteMapToExcel(sheetName, worksheetPart, worksheetCoOrdinates, stylesManager, spreadsheetDocument);
            TempDiagnostics.Output(string.Format("Excel worksheet '{0}' complete...", sheetName));
        }

        /// <summary>
        /// Legacy entry point to process an Excel Map.<br />
        /// NB! This processes <see cref="Template" /> based maps.
        /// </summary>
        /// <param name="map">A <see cref="BaseMap" /> derived element</param>
        /// <param name="data">The data.</param>
        /// <param name="exportPart">The export part.</param>
        public void ProcessMap(BaseMap map,
                               IDataPart data,
                               ExportPart exportPart)
        {
            this.legacyProcess = true;
            this.dataPart = data;
            this.exportPart = exportPart;

            map.ResourceKey = map.Key;

            TempDiagnostics.Output(string.Format("Populating map for worksheet '{0}'", this.sheetName));

            // Create an entity which will map elements to a worksheet.
            this.worksheetCoOrdinates = new ExcelMapCoOrdinateContainer(1, 1, map.GetContainerType()); // 1 x 1 map which represents the worksheet

            // Recursively process the supplied map and co-ordinates.
            this.ProcessMapInternal(map, null, this.worksheetCoOrdinates);

            // Write the contents of the map into the Excel Worksheet
            ExcelMapWriter.WriteMapToExcel(this.sheetName, this.worksheetPart, this.worksheetCoOrdinates, this.stylesManager, this.spreadsheetDocument);
            TempDiagnostics.Output(string.Format("Excel worksheet '{0}' complete...", this.sheetName));

            // Remove all model-based elements from the worksheet that were created using Chart and ShapeTemplate references
            RemoveTemplateElements(map);
        }

        #endregion Public Methods

        #region Private Helpers

        /// <summary>
        /// Recursively called procedure which
        /// </summary>
        /// <param name="map">The map being processed</param>
        /// <param name="parentMap">The parent map.</param>
        /// <param name="mapCoOrdinate">The parent map</param>
        private void ProcessMapInternal(BaseMap map, BaseMap parentMap, ExcelMapCoOrdinateContainer mapCoOrdinate)
        {
            // Set the data context of the element being traversed.
            // This may be a DataPart, CompositeDataPart TemplateId lookup data, or even a simple binding....
            this.TrySetMapDataContext(map, parentMap);

            // Is the element enabled, if so then process it.
            bool enabled = BindingContainer.ConvertToNullableBoolean(map.Enabled).GetValueOrDefault(true);

            if (enabled)
            {
                if (parentMap != null && !string.IsNullOrEmpty(parentMap.ResourceKey))
                {
                    map.ResourceKey = parentMap.ResourceKey;
                }

                // Prepare the map being processed (do not double-prepare)
                if (!map.Prepared && map.DataContext is IPreparable)
                {
                    // Not sure about the 1st parameter here... Is it even required? - it is suppoed to be the parent element.
                    ((IPreparable)(map.DataContext)).Prepare(parentMap, map);
                    map.Prepared = true;
                }

                // Now process the map depending on type...
                if (map is ExcelWriter.Template)
                {
                    this.ProcessTemplate((ExcelWriter.Template)map, exportPart, mapCoOrdinate);
                }
                else if (map is ExcelWriter.ContentControl)
                {
                    this.ProcessContentControl((ExcelWriter.ContentControl)map, mapCoOrdinate);
                }
                else if (map is ExcelWriter.StackPanel)
                {
                    this.ProcessStackPanel((ExcelWriter.StackPanel)map, mapCoOrdinate);
                }
                else if (map is ExcelWriter.Cell)
                {
                    this.ProcessCell((ExcelWriter.Cell)map, mapCoOrdinate);
                }
                else if (map is ExcelWriter.Padding)
                {
                    this.ProcessPadding((ExcelWriter.Padding)map, mapCoOrdinate);
                }
                else if (map is ExcelWriter.Property)
                {
                    this.ProcessProperty((ExcelWriter.Property)map, mapCoOrdinate);
                }
                else if (map is ExcelWriter.Table)
                {
                    this.ProcessTable((ExcelWriter.Table)map, mapCoOrdinate);
                }
                else if (map is ExcelWriter.TableData)
                {
                    this.ProcessTableData((ExcelWriter.TableData)map, mapCoOrdinate);
                }
                else if (map is ExcelWriter.Chart)
                {
                    this.ProcessChart((ExcelWriter.Chart)map, mapCoOrdinate);
                }
                else if (map is ExcelWriter.Shape)
                {
                    this.ProcessShape((ExcelWriter.Shape)map, mapCoOrdinate);
                }
                else if (map is ExcelWriter.Picture)
                {
                    this.ProcessPicture((ExcelWriter.Picture)map, mapCoOrdinate);
                }
            }
        }

        /// <summary>
        /// Processes a <see cref="ExcelWriter.ContentControl" /> which copies and uses a pre-stored <see cref="BaseMap" /> derived entity.<br />
        /// This may recurively call Process ProcessExcelMap if there are Maps within the map.
        /// </summary>
        /// <param name="contentControl">The content control.</param>
        /// <param name="containerMap">The container map.</param>
        private void ProcessContentControl(ExcelWriter.ContentControl contentControl,
                                           ExcelMapCoOrdinateContainer containerMap)
        {
            // Pull out a new instance of the map from the package
            if (contentControl.Content == null)
            {
                contentControl.Content = this.resourceStore.GetResourceByKey<BaseMap>(contentControl.ContentKey);
                contentControl.Content.ResourceKey = contentControl.Content.Key;
            }

            if (contentControl.Content != null)
            {
                this.ProcessMapInternal(contentControl.Content, contentControl, containerMap);
            }
        }

        /// <summary>
        /// Processes a <see cref="ExcelWriter.Shape" />.<br />
        /// Creates a <see cref="ExcelMapCoOrdinatePlaceholder" /> within the suplied <see cref="ExcelMapCoOrdinateContainer" /> and links<br />
        /// to the <see cref="ExcelWriter.Shape" /> so that is can be modelled and inserted in the output document.
        /// </summary>
        /// <param name="shape">The shape to be processed</param>
        /// <param name="container">The container into which it is to be processed</param>
        private void ProcessShape(ExcelWriter.Shape shape, ExcelMapCoOrdinateContainer container)
        {
            // override the style key using the selector if there is one
            string cellStyleKey = this.GetStyleKey
            (
                shape.CellStyleSelectorKey,
                shape.DataContext,
                shape.CellStyleKey
            );

            // Create a new container map located where the chart is to be written into the parent (container) map.
            var placeholderMapCoOrdinate = new ExcelMapCoOrdinatePlaceholder();
            placeholderMapCoOrdinate.DefinedName = BindingContainer.ConvertToString(shape.DefinedName);

            placeholderMapCoOrdinate.AddStyle(stylesManager.GetMapStyle(cellStyleKey));
            placeholderMapCoOrdinate.SpanLastColumn = shape.SpanLastColumn;
            placeholderMapCoOrdinate.ColumnSpan = (uint)shape.ColumnSpan;
            placeholderMapCoOrdinate.RowSpan = (uint)shape.RowSpan;
            placeholderMapCoOrdinate.SpanLastRow = shape.SpanLastRow;
            placeholderMapCoOrdinate.AssignedWidth = shape.Width;
            placeholderMapCoOrdinate.AssignedHeight = shape.Height;

            container.SetExcelMapCoOrdinate(placeholderMapCoOrdinate);

            // Store the placeholder for the chart with the chart itself for later processing
            shape.MapPlaceholder = placeholderMapCoOrdinate;
        }

        /// <summary>
        /// Processes a <see cref="ExcelWriter.Picture" />.<br />
        /// Creates a <see cref="ExcelMapCoOrdinatePlaceholder" /> within the suplied <see cref="ExcelMapCoOrdinateContainer" /> and links<br />
        /// to the <see cref="ExcelWriter.Picture" /> so that is can be modelled and inserted in the output document.
        /// </summary>
        /// <param name="picture">The picture to be processed</param>
        /// <param name="container">The container into which it is to be processed</param>
        private void ProcessPicture(ExcelWriter.Picture picture, ExcelMapCoOrdinateContainer container)
        {
            // override the style key using the selector if there is one
            string cellStyleKey = this.GetStyleKey
            (
                picture.CellStyleSelectorKey,
                picture.DataContext,
                picture.CellStyleKey
            );

            // Create a new container map located where the chart is to be written into the parent (container) map.
            var placeholderMapCoOrdinate = new ExcelMapCoOrdinatePlaceholder();
            placeholderMapCoOrdinate.DefinedName = BindingContainer.ConvertToString(picture.DefinedName);

            placeholderMapCoOrdinate.AddStyle(stylesManager.GetMapStyle(cellStyleKey));
            placeholderMapCoOrdinate.SpanLastColumn = picture.SpanLastColumn;
            placeholderMapCoOrdinate.ColumnSpan = (uint)picture.ColumnSpan;
            placeholderMapCoOrdinate.RowSpan = (uint)picture.RowSpan;
            placeholderMapCoOrdinate.SpanLastRow = picture.SpanLastRow;
            placeholderMapCoOrdinate.AssignedWidth = picture.Width;
            placeholderMapCoOrdinate.AssignedHeight = picture.Height;

            container.SetExcelMapCoOrdinate(placeholderMapCoOrdinate);

            // Store the placeholder for the chart with the chart itself for later processing
            picture.MapPlaceholder = placeholderMapCoOrdinate;
        }

        /// <summary>
        /// Processes a <see cref="ExcelWriter.Chart" />.<br />
        /// Creates a <see cref="ExcelMapCoOrdinatePlaceholder" /> within the suplied <see cref="ExcelMapCoOrdinateContainer" /> and links<br />
        /// to the <see cref="ExcelWriter.Chart" /> so that is can be modelled and inserted in the output document.
        /// </summary>
        /// <param name="chart">The chart to be processed</param>
        /// <param name="container">The container into which it is to be processed</param>
        /// <exception cref="InvalidOperationException">
        /// </exception>
        private void ProcessChart(ExcelWriter.Chart chart, ExcelMapCoOrdinateContainer container)
        {
            // override the style key using the selector if there is one
            string cellStyleKey = this.GetStyleKey
            (
                chart.CellStyleSelectorKey,
                chart.DataContext,
                chart.CellStyleKey
            );

            // Create a new container map located where the chart is to be written into the parent (container) map.
            var placeholderMapCoOrdinate = new ExcelMapCoOrdinatePlaceholder();
            container.SetExcelMapCoOrdinate(placeholderMapCoOrdinate);

            // *******************************************************************************************************************************************
            // Build a TableData object (which can also be specified in XAML and bound to the same table column entities), and set on the Table.
            // In that way, we could define a TableData in XAML as a means of populating a chart and a table simultaneously on one sheet.
            // The TableData could be constructed prior to this processing, and key referenced to populate the chart.
            // *******************************************************************************************************************************************
            string tableDataKey = BindingContainer.ConvertToString(chart.TableDataKey);
            string chartKey = string.IsNullOrEmpty(chart.Key) ? "Not Defined" : chart.Key;

            if (chart.TableData != null && !string.IsNullOrEmpty(tableDataKey))
            {
                // Exception - Can't have both TableData and TableDataKey defined
                throw new InvalidOperationException(string.Format("Both TableData and TableDataKey defined on Table Key='{0}'. Can't have both.", chartKey));
            }
            else if (!string.IsNullOrEmpty(tableDataKey))
            {
                // Try and find a DataTable with a matching key in the container... (we may need to go up the tree further!!!)
                chart.TableData = placeholderMapCoOrdinate.FirstAncendentKeyedElementOfType<TableData>(tableDataKey);
                if (chart.TableData == null)
                {
                    // None found, so create by plucking from the resources defined in the package and process the TableData
                    chart.TableData = this.resourceStore.GetResourceByKey<TableData>(tableDataKey);
                    this.ProcessMapInternal(chart.TableData, chart, container);

                    // Append the TableData to the container in which it was created so it can be key referenced by siblings and lower level entities.
                    container.AddKeyedElement(tableDataKey, chart.TableData);
                }
            }
            else if (chart.TableData == null)
            {
                // No TableData specified for this chart, so raise an exception
                throw new InvalidOperationException(string.Format("No TableData or TableDataKey defined on Chart Key='{0}'", chartKey));
            }
            else
            {
                // A TableData has been explicitly defined with the Chart.
                // Process the pre-defined TableData (populates with rows)
                this.ProcessMapInternal(chart.TableData, chart, container);
            }

            placeholderMapCoOrdinate.AddStyle(stylesManager.GetMapStyle(cellStyleKey));
            placeholderMapCoOrdinate.SpanLastColumn = chart.SpanLastColumn;
            placeholderMapCoOrdinate.ColumnSpan = (uint)chart.ColumnSpan;
            placeholderMapCoOrdinate.RowSpan = (uint)chart.RowSpan;
            placeholderMapCoOrdinate.SpanLastRow = chart.SpanLastRow;
            placeholderMapCoOrdinate.AssignedWidth = chart.Width;
            placeholderMapCoOrdinate.AssignedHeight = chart.Height;
            placeholderMapCoOrdinate.DefinedName = BindingContainer.ConvertToString(chart.DefinedName);

            // Store the placeholder for the chart with the chart itself for later processing
            chart.MapPlaceholder = placeholderMapCoOrdinate;
        }

        /// <summary>
        /// Processes a Template. This may recurively call Process ProcessExcelMap if there are Maps within the stack panel.
        /// </summary>
        /// <param name="template">The template.</param>
        /// <param name="exportPart">The export part.</param>
        /// <param name="containerMap">The container map.</param>
        private void ProcessTemplate(ExcelWriter.Template template,
                                     ExportPart exportPart,
                                     ExcelMapCoOrdinateContainer containerMap)
        {
            // Create a new container map located where the map-based element is to be written into the parent map.
            var map = new ExcelMapCoOrdinateContainer(0, 0, template.GetContainerType());
            //map.AddStyle(stylesManager.GetMapStyle(template.CellStyleKey));
            containerMap.SetExcelMapCoOrdinate(map);

            // ***********************************
            //  Process the Content Property
            // ***********************************

            // Recursive call to process the content
            if (template.Content != null)
            {
                map.MoveToNextRow();
                map.MoveToNextColumn();
                this.ProcessMapInternal(template.Content, template, map);
            }
        }

        /// <summary>
        /// Process an <see cref="ExcelWriter.Cell" />, writing the content into an excel worksheet.
        /// </summary>
        /// <param name="cell">The cell.</param>
        /// <param name="container">The container.</param>
        private void ProcessCell(ExcelWriter.Cell cell, ExcelMapCoOrdinateContainer container)
        {
            // override the style key using the selector if there is one
            string cellStyleKey = this.GetStyleKey
                (
                    cell.CellStyleSelectorKey,
                    cell.DataContext,
                    cell.CellStyleKey
                );

            // Create a new co-ordinate container located where the property is to be written into the parent map.
            var cellMapCoOrdinate = new ExcelMapCoOrdinateCell();
            cellMapCoOrdinate.AddStyle(stylesManager.GetMapStyle(cellStyleKey, cell.DataContext));
            cellMapCoOrdinate.SpanLastColumn = cell.SpanLastColumn;
            cellMapCoOrdinate.SpanLastRow = cell.SpanLastRow;
            cellMapCoOrdinate.RowSpan = (uint)cell.RowSpan;
            cellMapCoOrdinate.ColumnSpan = (uint)cell.ColumnSpan;
            cellMapCoOrdinate.DefinedName = BindingContainer.ConvertToString(cell.DefinedName);
            cellMapCoOrdinate.AssignedWidth = cell.Width;
            cellMapCoOrdinate.ColumnIsHidden = BindingContainer.ConvertToNullableBoolean(cell.ColumnIsHidden).GetValueOrDefault(false);
            cellMapCoOrdinate.AssignedHeight = cell.Height;
            cellMapCoOrdinate.RowIsHidden = BindingContainer.ConvertToNullableBoolean(cell.RowIsHidden).GetValueOrDefault(false);

            container.SetExcelMapCoOrdinate(cellMapCoOrdinate);

            //ProcessRowRelatedInfo(cellMapCoOrdinate, container, cell);
            //ProcessColumnRelatedInfo(cellMapCoOrdinate, container, cell);

            // Set value if there is one to set
            if (cell.Value != null)
            {
                cellMapCoOrdinate.CurrentValue = cell.Value;
            }
        }

        /// <summary>
        /// Process an <see cref="ExcelWriter.Padding" />, writing the content into an excel worksheet.
        /// </summary>
        /// <param name="padding">The padding.</param>
        /// <param name="container">The container.</param>
        private void ProcessPadding(ExcelWriter.Padding padding, ExcelMapCoOrdinateContainer container)
        {
            // override the style key using the selector if there is one
            string cellStyleKey = this.GetStyleKey
            (
                padding.CellStyleSelectorKey,
                padding.DataContext,
                padding.CellStyleKey
            );

            // Create a new co-ordinate container located where the property is to be written into the parent map.
            var cellMapCoOrdinate = new ExcelMapCoOrdinatePadding();
            cellMapCoOrdinate.AddStyle(stylesManager.GetMapStyle(cellStyleKey));
            cellMapCoOrdinate.SpanLastColumn = padding.SpanLastColumn;
            cellMapCoOrdinate.ColumnSpan = (uint)padding.ColumnSpan;
            cellMapCoOrdinate.RowSpan = (uint)padding.RowSpan;
            cellMapCoOrdinate.SpanLastRow = padding.SpanLastRow;

            // The cell is assigned an Excel DefinedName when written to Excel.
            cellMapCoOrdinate.DefinedName = BindingContainer.ConvertToString(padding.DefinedName);

            container.SetExcelMapCoOrdinate(cellMapCoOrdinate);

            // Set value if there is one to set
            if (padding.Value != null)
            {
                cellMapCoOrdinate.CurrentValue = padding.Value;
            }
        }

        /// <summary>
        /// Process an <see cref="Property" />, writing the content into an excel worksheet.
        /// </summary>
        /// <param name="property">The property.</param>
        /// <param name="container">The container.</param>
        private void ProcessProperty(Property property, ExcelMapCoOrdinateContainer container)
        {
            // Create a new co-ordinate container located where the property is to be written into the parent map.
            var propertyMap = new ExcelMapCoOrdinateContainer(1, 1, property.GetContainerType());
            container.SetExcelMapCoOrdinate(propertyMap);

            //TODO: Change this and depr
            bool rowIsHidden = BindingContainer.ConvertToNullableBoolean(property.RowIsHidden).GetValueOrDefault(false);

            // The cell is assigned an Excel DefinedName when written to Excel.
            propertyMap.DefinedName = BindingContainer.ConvertToString(property.DefinedName);

            // override the style key using the selector if there is one
            string headerStyleKey = this.GetStyleKey
            (
                property.HeaderStyleSelectorKey,
                property.DataContext,
                property.HeaderStyleKey
            );

            // Add the properties.
            var headerCellCoOrdinate = new ExcelMapCoOrdinateCell
            {
                CurrentValue = property.Header,
                RowIsHidden = rowIsHidden,
                AssignedHeight = property.Height,
            };
            headerCellCoOrdinate.AddStyle(stylesManager.GetMapStyle(headerStyleKey));

            propertyMap.SetExcelMapCoOrdinate(headerCellCoOrdinate);

            // Write in adjacent cell
            propertyMap.MoveToNextColumn();

            // override the style key using the selector if there is one
            string cellStyleKey = this.GetStyleKey
            (
                property.CellStyleSelectorKey,
                property.DataContext,
                property.CellStyleKey
            );

            var cellCoOrdinate = new ExcelMapCoOrdinateCell
            {
                CurrentValue = property.Value,
                RowIsHidden = rowIsHidden,
                AssignedHeight = property.Height,
            };

            cellCoOrdinate.AddStyle(stylesManager.GetMapStyle(cellStyleKey));
            propertyMap.SetExcelMapCoOrdinate(cellCoOrdinate);
        }

        /// <summary>
        /// Processes a <see cref="Table" />, writing the content to an Excel Worksheet.
        /// </summary>
        /// <param name="table">The table.</param>
        /// <param name="container">The container.</param>
        /// <exception cref="InvalidOperationException"></exception>
        private void ProcessTable(Table table, ExcelMapCoOrdinateContainer container)
        {
            // Create a new container map (table map) located where the table is to be written into the parent map.
            var tableMap = new ExcelMapCoOrdinateContainer(0, 0, table.GetContainerType());
            container.SetExcelMapCoOrdinate(tableMap);

            // *******************************************************************************************************************************************
            // Build a TableData object (which can also be specified in XAML and bound to the same table column entities), and set on the Table.
            // In that way, we could define a TableData in XAML as a means of populating a chart and a table simultaneously on one sheet.
            // The TableData could be constructed prior to this processing, or key referenced to populate the table.
            // *******************************************************************************************************************************************
            string tableDataKey = BindingContainer.ConvertToString(table.TableDataKey);
            if (table.TableData != null && !string.IsNullOrEmpty(tableDataKey))
            {
                // Exception - Can't have both TableData and TableDataKey defined
                string tableKey = string.IsNullOrEmpty(table.Key) ? "Not Defined" : table.Key;
                throw new InvalidOperationException(string.Format("Both TableData and TableDataKey defined on Table Key='{0}'. Can't have both.", tableKey));
            }
            else if (!string.IsNullOrEmpty(tableDataKey))
            {
                // We have specified a TableDataKey...
                // Try and find a DataTable with a matching key in the container... (we may need to go up the tree further!!!)
                table.TableData = tableMap.FirstAncendentKeyedElementOfType<TableData>(tableDataKey);
                if (table.TableData == null)
                {
                    // None found, so create by plucking from the resources defined in the package and process the TableData
                    table.TableData = this.resourceStore.GetResourceByKey<TableData>(tableDataKey);
                    this.ProcessMapInternal(table.TableData, table, container);

                    // Append the TableData to the container in which it was created so it can be key referenced by siblings and lower level entities.
                    container.AddKeyedElement(tableDataKey, table.TableData);
                }
            }
            else if (table.TableData == null)
            {
                // No TableData specified for this table, so create from table properties, and process
                table.CreateDefaultTableData(table.Prepared);
                this.ProcessMapInternal(table.TableData, table, container);
            }
            else
            {
                // A TableData has been explicitly defined with the Table.
                // Process the pre-defined TableData (populates with rows)
                this.ProcessMapInternal(table.TableData, table, container);
            }

            tableMap.SpanLastColumn = table.SpanLastColumn;
            tableMap.AddStyle(stylesManager.GetMapStyle(table.CellStyleKey));

            // **********************************************
            // Add header to the table map if there is one
            // **********************************************
            if (table.Header != null)
            {
                // Move to next row, first column
                tableMap.MoveToNextRow();
                tableMap.SetCurrentColumn(1);

                var exportCell = new ExcelMapCoOrdinateCell
                {
                    SpanLastColumn = true,
                    CurrentValue = table.Header,
                };
                exportCell.AddStyle(stylesManager.GetMapStyle(table.HeaderStyleKey));

                tableMap.SetExcelMapCoOrdinate(exportCell);
            }

            // **************************************************
            // Add sub-header to the table map if there is one
            // **************************************************
            if (table.SubHeader != null)
            {
                // Move to next row, first column
                tableMap.MoveToNextRow();
                tableMap.SetCurrentColumn(1);

                var exportCell = new ExcelMapCoOrdinateCell
                {
                    SpanLastColumn = true,
                    CurrentValue = table.SubHeader,
                };
                exportCell.AddStyle(stylesManager.GetMapStyle(table.SubHeaderStyleKey));
                tableMap.SetExcelMapCoOrdinate(exportCell);
            }

            // Add Space row after header and sub-header (if there are any)
            if (table.Header != null || table.SubHeader != null)
            {
                tableMap.MoveToNextRow();
                tableMap.SetCurrentColumn(1);
                var paddingCell = new ExcelMapCoOrdinateCell();
                tableMap.SetExcelMapCoOrdinate(paddingCell);
            }

            // ****************************************************************
            // Add properties to the table map if there are any.
            // Note that this will add a spacer row after the properties
            // if any are to be written into the worksheet. This
            // row will be hidden if ALL of the properties are hidden.
            // ****************************************************************
            if (table.Properties != null && table.Properties.Count > 0)
            {
                tableMap.MoveToNextRow();
                tableMap.SetCurrentColumn(1);

                // Create a container into which the properties will be inserted
                this.ProcessTable_Properties(table, tableMap);
            }

            // ****************************************************************************
            // Add the table of data to the table map (Should have same columns as table)
            // ****************************************************************************
            if (table.TableData.Columns.Count > 0)
            {
                // Get inforamtion about the table columns and the header
                TableColumnsInfo tableInfo = new TableColumnsInfo(table, this.legacyProcess);

                // Position ready to write table
                tableMap.MoveToNextRow();
                tableMap.SetCurrentColumn(1);

                // The data region within the table is assigned an Excel DefinedName.
                // If a DataRegionDefinedName property is explicitly set on the ExcelTableMap, then this is used,
                // otherwise the worksheet name is used, and multiples are appended with _1, _2 etc...
                string resolvedName = BindingContainer.ConvertToString(table.DataRegionDefinedName);
                string dataRegionDefinedName = string.IsNullOrEmpty(resolvedName)
                                            ? this.sheetName
                                            : resolvedName;

                // Write table into the parent map.
                var tableMapTable = new ExcelMapCoOrdinateContainer(0, 0, "TableMapTable");
                tableMapTable.SpanLastColumn = tableMap.SpanLastColumn;
                tableMapTable.SpanLastRow = tableMap.SpanLastRow;

                // Process the group headers - Create a defined name which maps to the group header
                this.ProcessTable_GroupHeaders(dataRegionDefinedName, tableInfo, tableMapTable);

                // Write the table region to the map (group headers, column headers & data)
                this.ProcessTable_TableData(dataRegionDefinedName, tableInfo, tableMapTable);

                // Write into main 'Export View Item Container' map
                tableMap.SetExcelMapCoOrdinate(tableMapTable);
            }

            // Add a row between table and footer.
            if (table.Footer != null || table.SubFooter != null)
            {
                tableMap.MoveToNextRow();
                var paddingCell = new ExcelMapCoOrdinateCell();
                tableMap.SetExcelMapCoOrdinate(paddingCell);
            }

            // ***********************************
            // Add the footer if there is one
            // ***********************************
            if (table.Footer != null)
            {
                // Position ready to write footer
                tableMap.MoveToNextRow();
                tableMap.SetCurrentColumn(1);

                var exportCell = new ExcelMapCoOrdinateCell
                {
                    SpanLastColumn = true,
                    CurrentValue = table.Footer,
                };
                exportCell.AddStyle(stylesManager.GetMapStyle(table.FooterStyleKey));
                tableMap.SetExcelMapCoOrdinate(exportCell);
            }

            // ***********************************
            // Add the sub-footer if there is one
            // ***********************************
            if (table.SubFooter != null)
            {
                // Position ready to write footer
                tableMap.MoveToNextRow();
                tableMap.SetCurrentColumn(1);

                var exportCell = new ExcelMapCoOrdinateCell
                {
                    SpanLastColumn = true,
                    CurrentValue = table.SubFooter,
                };
                exportCell.AddStyle(stylesManager.GetMapStyle(table.SubFooterStyleKey));
                tableMap.SetExcelMapCoOrdinate(exportCell);
            }
        }

        /// <summary>
        /// Processes a <see cref="TableData" /> element, which has no visual representation.
        /// </summary>
        /// <param name="tableData">The XAML defined element</param>
        /// <param name="mapContainer">The map container.</param>
        private void ProcessTableData(TableData tableData, ExcelMapCoOrdinateContainer mapContainer)
        {
            string tableDataKey = BindingContainer.ConvertToString(tableData.Key);
            if (!string.IsNullOrEmpty(tableDataKey))
            {
                // Append the TableData to the container in which it was created so it can be key referenced by siblings and lower level entities.
                mapContainer.AddKeyedElement(tableDataKey, tableData);
            }

            // Store reference to all rows that are to be added to the table
            IEnumerable itemsSource = tableData.ItemsSource as IEnumerable;
            if (itemsSource != null)
            {
                // Add all rows to the TableData
                foreach (object item in itemsSource)
                {
                    tableData.AddRow(item);
                }
            }
        }

        /// <summary>
        /// Processes the <see cref="TableData" /> in the <see cref="Table" /> into an <see cref="ExcelMapCoOrdinateContainer" />
        /// which will eventually be written into an Excel worksheet.
        /// </summary>
        /// <param name="dataRegionDefinedName">Name of the data region defined.</param>
        /// <param name="tableInfo">Information about the <see cref="Table" /> which contains the <see cref="TableData" /></param>
        /// <param name="container">The excel map container into which the <see cref="TableData" /> will be written</param>
        private void ProcessTable_TableData(string dataRegionDefinedName, TableColumnsInfo tableInfo, ExcelMapCoOrdinateContainer container)
        {
            // Reset before first column in the map.
            container.MoveToNextRow();
            container.SetCurrentColumn(0);

            Table table = tableInfo.Table;
            TableData tableData = table.TableData;

            double? defaultRowHeight = BindingContainer.ConvertToNullableDouble(table.DefaultRowHeight);

            // Create a TableArea map to hold the data region TODO: + padding columns if required
            var tableAreaMap = new ExcelMapCoOrdinateContainer(1, 0, "TableArea");
            tableAreaMap.SpanLastColumn = container.SpanLastColumn;
            tableAreaMap.SpanLastRow = container.SpanLastRow;

            // Create a TableData map that contains the entire table data area (including column headings).
            // This is so we can write a 'Defined Name' into the Excel document for the table data.
            var tableDataRegionMap = new ExcelMapCoOrdinateContainer(1, 0, "TableDataRegion");
            tableDataRegionMap.DefinedName = dataRegionDefinedName;
            tableDataRegionMap.SpanLastColumn = container.SpanLastColumn;
            tableDataRegionMap.SpanLastRow = container.SpanLastRow;

            // Update internal reference to DataRegion of TableData in the TableData
            // This is the first ExcelCoOrdinateContainer encountered and will be used for
            // range/formula referencing in charts..
            if (tableData.MapContainer == null) tableData.MapContainer = tableDataRegionMap;

            // Write the column headers into the map.

            // Used for column padding.
            StyleBase lastColumnHeaderStyle = null;
            bool hideColumnHeaders = table.HideColumnsHeader;

            for (int columnIdx = 0; columnIdx < tableInfo.ColumnCount; columnIdx++)
            {
                // Get column and set DataContext to Parent
                TableColumnInfo columnInfo = tableInfo.ColumnInfos[columnIdx];
                TableColumn column = columnInfo.Column;

                // Move on to the next column in the table map
                tableDataRegionMap.MoveToNextColumn();

                // Create a ColumnData map that contains the column header and column data.
                // This is so we can write a 'Defined Name' into the Excel document for the column data.
                var columnDataMap = new ExcelMapCoOrdinateContainer(1, 1, "TableDataColumnRegion");
                tableDataRegionMap.SetExcelMapCoOrdinate(columnDataMap);

                columnDataMap.SpanLastColumn = columnInfo.SpanLastColumn;
                columnDataMap.DefinedName = string.Format("{0}_{1}", dataRegionDefinedName, column.Header);

                // Update internal reference to DataRegion of columns in the TableData
                // This is the first ExcelCoOrdinateContainer encountered and will be used for
                // range/formula referencing in charts..
                if (column.DataRegion == null) column.DataRegion = columnDataMap;

                // process the cell style and style selectors
                string headerStyleKey = this.GetStyleKey
                (
                    column.HeaderStyleSelectorKey,
                    column.DataContext,
                    column.HeaderStyleKey
                );

                StyleBase headerStyle = stylesManager.GetMapStyle(headerStyleKey);
                if (columnInfo.IsLastColumn)
                {
                    // If this is the last column then store the style so it can be re-used
                    // when padding the columns out the the container boundaries
                    lastColumnHeaderStyle = headerStyle;
                }

                // Write header values into the column data map
                // and span (to bounds) if set on the table.
                var headerCellMap = new ExcelMapCoOrdinateCell()
                {
                    CurrentValue = column.Header,
                    SpanLastColumn = columnInfo.SpanLastColumn,
                    ColumnSpan = (uint)columnInfo.ColumnSpan,
                    AssignedWidth = columnInfo.Width,
                    ColumnIsHidden = columnInfo.Hidden,
                    RowIsHidden = hideColumnHeaders,
                    AssignedHeight = column.Height,
                };

                headerCellMap.AddStyle(headerStyle);
                columnDataMap.SetExcelMapCoOrdinate(headerCellMap);

                // Selector to be used when styling rows
                string rowStyleSelectorKey = table.RowStyleSelectorKey;

                // Write all row data for this column into the column data map
                foreach (TableDataRowInfo tableDataRowInfo in tableData.RowData)
                {
                    columnDataMap.MoveToNextRow();

                    // Read the named property using reflection
                    object actualValue = null;
                    try
                    {
                        actualValue = BindingContainer.GetPropValue(tableDataRowInfo.RowData, column.DisplayMember);
                    }
                    catch (System.Exception)
                    {
                        // Write out to debug window... This is a binding error.
                    }

                    // Read the cell and row styles from the column and table (considers selector keys)
                    string cellStyleKey = this.GetStyleKey(column.CellStyleSelectorKey, tableDataRowInfo.RowData, column.CellStyleKey);
                    string rowStyleKey = this.GetStyleKey(rowStyleSelectorKey, tableDataRowInfo.RowData, null);

                    if (!string.IsNullOrEmpty(column.CellTemplateMapKey))
                    {
                        // We have a reference to a CellTemplate, create a container, and pull the 
                        // a ContentControl to host the keyed map cell template...
                        var cellMapContainer = new ExcelMapCoOrdinateContainer(1, 1, "CellTemplate")
                        {
                            SpanLastColumn = columnInfo.SpanLastColumn,
                        };

                        cellMapContainer.AddStyle(stylesManager.GetMapStyle(cellStyleKey));
                        cellMapContainer.AddStyle(stylesManager.GetMapStyle(rowStyleKey));

                        // Write into the ColumnDataMap
                        columnDataMap.SetExcelMapCoOrdinate(cellMapContainer);

                        // Get a map, looked up by key, set its DataContext + ParentDataContext (Parent in this instance is Table)
                        BaseMap cellTemplateMap = this.resourceStore.GetResourceByKey<BaseMap>(column.CellTemplateMapKey);
                        cellTemplateMap.ParentDataContext = table.DataContext;
                        cellTemplateMap.DataContext = actualValue;

                        // Recursive call to process the content of the current location in the stack panel
                        this.ProcessMapInternal(cellTemplateMap, table, cellMapContainer);

                        // Add the inflated template map to the Table Items collection
                        // so that the instance can be processed later...
                        table.Items.Add(cellTemplateMap);
                    }
                    else
                    {
                        // Write the data cell into the column data map
                        var valueCellMap = new ExcelMapCoOrdinateCell
                        {
                            CurrentValue = actualValue,
                            SpanLastColumn = columnInfo.SpanLastColumn,
                            ColumnSpan = (uint)columnInfo.ColumnSpan,
                            AssignedWidth = columnInfo.Width,
                            ColumnIsHidden = columnInfo.Hidden,
                            AssignedHeight = defaultRowHeight,
                        };

                        // First, process cell for cell style selectors
                        valueCellMap.AddStyle(stylesManager.GetMapStyle(cellStyleKey, tableDataRowInfo.RowData));
                        valueCellMap.AddStyle(stylesManager.GetMapStyle(rowStyleKey));

                        columnDataMap.SetExcelMapCoOrdinate(valueCellMap);
                    }

                    // Store the index of the row for use by chart and table processing.
                    if (tableDataRowInfo.TableRowIndex == 0)
                    {
                        tableDataRowInfo.TableRowIndex = columnDataMap.CurrentRowIndex;
                    }
                }
            }

            if (tableDataRegionMap.MapRowCount > 0 && tableDataRegionMap.MapColumnCount > 0)
            {
                // Write TableDataRegion into first column of TableArea
                tableAreaMap.SetCurrentColumn(1);
                tableAreaMap.SetExcelMapCoOrdinate(tableDataRegionMap);

                if (table.PadLastColumn)
                {
                    // Add a padding column which has the same header style as the last cell in the tableDataRegionMap
                    // and which has rows that have a RowStyle which is the same as the table.
                    // We already have LastColumnHeaderStyle, so we just need to repeat for the rows.
                    tableAreaMap.MoveToNextColumn();
                    var paddingColumnMap = new ExcelMapCoOrdinateContainer(1, 1, "TablePaddingColumnRegion");
                    tableAreaMap.SetExcelMapCoOrdinate(paddingColumnMap);

                    // Padding Header
                    var paddingHeaderCellMap = new ExcelMapCoOrdinateCell()
                    {
                        SpanLastColumn = true,
                    };
                    paddingHeaderCellMap.AddStyle(lastColumnHeaderStyle);
                    // Add Padding Header cell to Column Map
                    paddingColumnMap.SetExcelMapCoOrdinate(paddingHeaderCellMap);

                    // Write all row data for this column into the column data map
                    foreach (TableDataRowInfo rowInfo in tableData.RowData)
                    {
                        paddingColumnMap.MoveToNextRow();

                        // Write the data cell into the column data map
                        //TODO: PADDING CELL
                        var paddingCellMap = new ExcelMapCoOrdinateCell
                        {
                            SpanLastColumn = true,
                            AssignedHeight = defaultRowHeight,
                        };

                        // process the cell and row style selectors
                        string rowStyleKey = this.GetStyleKey(table.RowStyleSelectorKey, rowInfo.RowData, null);
                        if (!string.IsNullOrEmpty(rowStyleKey))
                        {
                            paddingCellMap.AddStyle(stylesManager.GetMapStyle(rowStyleKey));
                        }
                        // Add Padding cell to Column Map
                        paddingColumnMap.SetExcelMapCoOrdinate(paddingCellMap);
                    }
                }

                // Write TableArea into Container.
                container.SetCurrentColumn(1);
                container.SetExcelMapCoOrdinate(tableAreaMap);

                //for (int columnIdx = 0; columnIdx < columnCount; columnIdx++)
                //{
                //    // Move on to the next column in the table map
                //    tableDataRegionMap.MoveToNextColumn();

                //    // Create a ColumnData map that contains the column header and column data.
                //    // This is so we can write a 'Defined Name' into the Excel document for the column data.
                //    var paddingColumnDataMap = new ExcelMapCoOrdinateContainer(tableDataRegionMap, 1, 1, "TableDataColumnRegion");
                //}
            }
        }

        /// <summary>
        /// Processes the group headers defined on a table into a map.
        /// </summary>
        /// <param name="dataRegionDefinedName">If specified, defines the prefix for named range when exported to Excel</param>
        /// <param name="tableInfo">The table information.</param>
        /// <param name="container">The container.</param>
        private void ProcessTable_GroupHeaders(string dataRegionDefinedName, TableColumnsInfo tableInfo, ExcelMapCoOrdinateContainer container)
        {
            if (tableInfo.HeaderRowInfos.Count() > 0)
            {
                // Create a map to hold the group headers.
                var groupHeaderMap = new ExcelMapCoOrdinateContainer(0, 0, "TableGroupHeader");
                groupHeaderMap.SpanLastColumn = container.SpanLastColumn;

                // Group Column Header Levels
                foreach (GroupHeaderRowInfo rowInfo in tableInfo.HeaderRowInfos)
                {
                    // One row per level
                    groupHeaderMap.SetCurrentColumn(1);
                    groupHeaderMap.MoveToNextRow();

                    // Create a map to hold each group header level.
                    var groupHeaderLevelMap = new ExcelMapCoOrdinateContainer(0, 0, string.Format("TableGroupHeaderLevel[{0}]", rowInfo.Level));
                    groupHeaderMap.SetExcelMapCoOrdinate(groupHeaderLevelMap);
                    groupHeaderLevelMap.SetCurrentRow(1);

                    TableColumnHeader lastColumnHeader = null;
                    ExcelMapCoOrdinatePlaceholder lastHeaderCell = null;

                    // Find the height of this level header (highest)
                    //double? headerHeight = rowInfo.Height;
                    //bool headerIsHidden = rowInfo.Hidden;

                    // Loop over all columns + 1 so we can close the last header
                    for (int colIdx = 0; colIdx < tableInfo.ColumnCount; colIdx++)
                    {
                        groupHeaderLevelMap.MoveToNextColumn();

                        // Get the header for the level
                        TableColumnInfo colInfo = tableInfo.ColumnInfos[colIdx];
                        TableColumnHeader header = colInfo.GetColumnHeader(rowInfo.Level);

                        // Write a padding cell into the map
                        var headerCell = new ExcelMapCoOrdinateCell
                        {
                            ColumnSpan = (uint)colInfo.ColumnSpan,
                            ColumnIsHidden = colInfo.Hidden,
                            RowIsHidden = rowInfo.Hidden,
                            AssignedWidth = colInfo.Width,
                            AssignedHeight = rowInfo.Height,
                            SpanLastColumn = colInfo.IsLastColumn && colInfo.SpanLastColumn,
                        };

                        // Write the header over the column
                        if (header != null)
                        {
                            headerCell.CurrentValue = header.Header;
                            headerCell.AddStyle(this.stylesManager.GetMapStyle(header.HeaderStyleKey));
                        }

                        // Write header cell into map
                        groupHeaderLevelMap.SetExcelMapCoOrdinate(headerCell);

                        if (header == lastColumnHeader)
                        {
                            // Link this header to the last so that they get merged across columns
                            if (colIdx > 0)
                            {
                                headerCell.MergeWith = lastHeaderCell;
                            }
                        }
                        else
                        {
                            // Move on to new column header, reset columns spanned
                            lastColumnHeader = header;
                            lastHeaderCell = headerCell;
                        }
                    }
                }

                groupHeaderMap.DefinedName = string.Format("{0}_{1}", dataRegionDefinedName, "GroupHeaders");

                // Add to the parent map (if any group headers were specified)
                if (groupHeaderMap.MapRowCount > 0 && groupHeaderMap.MapColumnCount > 0)
                {
                    container.MoveToNextRow();
                    container.SetCurrentColumn(1);
                    container.SetExcelMapCoOrdinate(groupHeaderMap);
                }
            }
        }

        /// <summary>
        /// Adds the properties table to the <see cref="Table" />.<br />
        /// Increments the rowIndex
        /// </summary>
        /// <param name="excelTableMap">The excel table map.</param>
        /// <param name="container">The container.</param>
        private void ProcessTable_Properties(ExcelWriter.Table excelTableMap, ExcelMapCoOrdinateContainer container)
        {
            // We need an extra container so we can put a DefinedName against the range,
            // and, at the same time, add an extra row which is not covered by the DefinedName.
            //var propMapsContainer = new ExcelMapCoOrdinateContainer(container, 0, 0, "TableMapProperties");
            var propMapsContainer = new ExcelMapCoOrdinateContainer(0, 1, "TableMapProperties");

            // Define a name on the props container
            propMapsContainer.DefinedName = string.Format("{0}_Props", this.sheetName);

            // Write all of the properties into the worksheet.
            foreach (Property propertyMap in excelTableMap.Properties)
            {
                propertyMap.DataContext = excelTableMap.DataContext;
                propMapsContainer.MoveToFirstColumn();
                propMapsContainer.MoveToNextRow();
                this.ProcessProperty(propertyMap, propMapsContainer);
            }

            if (propMapsContainer.MapRowCount > 0)
            {
                // Write the props container into the main container.
                container.SetExcelMapCoOrdinate(propMapsContainer);
            }

            // Add a padding cell on the row below the properties.
            container.MoveToNextRow();
            var paddingCell = new ExcelMapCoOrdinateCell();
            container.SetExcelMapCoOrdinate(paddingCell);

            // Hide the spacer row if all properties hidden
            // Any rows hidden?
            bool hideExtraRow = AnyPropertyRowsHidden(excelTableMap.Properties);
            if (hideExtraRow)
            {
                paddingCell.RowIsHidden = true;
            }
        }

        /// <summary>
        /// Processes a stack panel. This may recursively call ProcessMap if there are Maps within the stack panel,
        /// or the <see cref="StackPanel" /> has an ItemsSource and ItemTemplateMapKey.
        /// </summary>
        /// <param name="stackPanel">The stack panel.</param>
        /// <param name="containerMap">The container map.</param>
        /// <exception cref="InvalidOperationException">
        /// </exception>
        private void ProcessStackPanel(ExcelWriter.StackPanel stackPanel,
                                       ExcelMapCoOrdinateContainer containerMap)
        {
            // Check that we don't have both an Items collection and an ItemsSource
            if (stackPanel.ItemsSource != null && stackPanel.Items != null && stackPanel.Items.Count > 0)
            {
                throw new InvalidOperationException(string.Format("Error: cannot have ItemsSource and Items set in {0}", stackPanel));
            }

            if (stackPanel.ItemsSource != null && string.IsNullOrEmpty(stackPanel.ItemTemplateMapKey))
            {
                StringBuilder sb = new StringBuilder();
                sb.AppendLine("Error: cannot have ItemsSource with no ItemTemplateMapKey specified.");
                sb.AppendLine("An ItemTemplateMapKey uses an instance of a Map defined in resources identified by Key:");
                sb.AppendLine(string.Format("Map=<{0}>, ItemTemplateMapKey='{1}'", stackPanel, stackPanel.ItemTemplateMapKey));
                sb.AppendLine(string.Format("ItemsSource=<{0}>", stackPanel.ItemsSource));
                throw new InvalidOperationException(sb.ToString());
            }

            // Set a StackPanel map within the supplied mapCoOrdinate at the current location
            // The stackPanelMapCoOrdinate is what all information is written into.
            var stackPanelMap = new ExcelMapCoOrdinateContainer(0, 0, stackPanel.GetContainerType());
            stackPanelMap.SpanLastColumn = stackPanel.SpanLastColumn;
            stackPanelMap.SpanLastRow = stackPanel.SpanLastRow;
            stackPanelMap.AddStyle(stylesManager.GetMapStyle(stackPanel.CellStyleKey));
            stackPanelMap.DefinedName = BindingContainer.ConvertToString(stackPanel.DefinedName);
            containerMap.SetExcelMapCoOrdinate(stackPanelMap);

            // if we have 
            // 1. a map 
            // or
            // 2. a map id and a package (so we can inflate a new instance if needed, a must really for dynamics)
            if (stackPanel.ItemsSource != null && !string.IsNullOrEmpty(stackPanel.ItemTemplateMapKey))
            {
                // Itterate over the ItemSource
                var itemsSource = stackPanel.ItemsSource as System.Collections.IEnumerable;

                if (itemsSource != null)
                {
                    // Create a ContentControl for each item that has the required ContentKey...
                    // This will pull in the specified map (identitied by Key) for each item.
                    foreach (object dataItem in itemsSource)
                    {
                        stackPanel.Items.Add(new ExcelWriter.ContentControl
                        {
                            ContentKey = stackPanel.ItemTemplateMapKey,
                            DataContext = dataItem,
                        });
                    }
                }
            }

            // StackPanel should either have an explicit map (items) or have been populated abovedoes not contain a templated ItemSource
            foreach (BaseMap content in stackPanel.Items)
            {
                if (content.IsVisual)
                {
                    // As this is a stack-panel, each visual element within the panel will increase the 
                    // number of row or column elements in the map depending on the orientation
                    if (stackPanel.Orientation == System.Windows.Controls.Orientation.Horizontal)
                    {
                        // Only 1 row
                        stackPanelMap.SetCurrentRow(1);
                        stackPanelMap.MoveToNextColumn();
                    }
                    else
                    {
                        // Only 1 column
                        stackPanelMap.SetCurrentColumn(1);
                        stackPanelMap.MoveToNextRow();
                    }
                }

                // Recursive call to process the content of the current location in the stack panel
                this.ProcessMapInternal(content, stackPanel, stackPanelMap);
            }
        }

        /// <summary>
        /// Determine the number of column groups that are specified in the column headers
        /// These are additional grouped columns above the column headers, one for each level.
        /// </summary>
        /// <param name="columnHeaders">The column headers.</param>
        /// <returns></returns>
        private static IOrderedEnumerable<IGrouping<int, TableColumnHeader>> GetColumnGroupHeaders(TableColumnHeaderCollection columnHeaders)
        {
            var columnHeaderGroups = columnHeaders.GroupBy(x => x.Level).OrderBy(x => x.Key);
            return columnHeaderGroups;
        }

        /// <summary>
        /// Gets the style key.
        /// </summary>
        /// <param name="cellStyleSelectorKey">The cell style selector key.</param>
        /// <param name="item">The item.</param>
        /// <param name="defaultCellStyleKey">The default cell style key.</param>
        /// <returns></returns>
        private string GetStyleKey(string cellStyleSelectorKey, object item, string defaultCellStyleKey)
        {
            // default to the key
            string result = defaultCellStyleKey;

            // but if a CellStyleSelectorKey has been set then use it to get the style key
            if (!string.IsNullOrEmpty(cellStyleSelectorKey))
            {
                // get the from the store using the key
                var cellStyleSelector = this.resourceStore.GetResourceByKey<CellStyleSelector>(cellStyleSelectorKey);
                if (cellStyleSelector != null)
                {
                    try
                    {
                        // call the select method to get the result
                        result = cellStyleSelector.SelectCellStyleKey(item);
                    }
                    catch (System.Exception ex)
                    {
                        // Write out to debug window that the cellStyleSelector has thrown an exception
                        System.Diagnostics.Debug.Write(string.Format("Selector '{0}' threw an exception. '{1}'", cellStyleSelectorKey, ex.Message));
                    }
                }
                else
                {
                    // Write out to debug window that a key has been specified that doesn't exist
                    System.Diagnostics.Debug.Write(string.Format("CellStyleSelector with key '{0}' does not exist in the xaml defined CellStyleSelectorCollection.", cellStyleSelectorKey));
                }
            }
            return result;
        }

        /// <summary>
        /// Attempts to set the data context of the supplied <see cref="BaseMap" />
        /// </summary>
        /// <param name="map">The <see cref="BaseMap" /> which is to have its data context set.</param>
        /// <param name="parentMap">The parent <see cref="BaseMap" /> from where data context will flow down.</param>
        private void TrySetMapDataContext(BaseMap map, BaseMap parentMap)
        {
            if (this.legacyProcess)
            {
                // If the evaluation of a binding on the data context results in data, then set.
                // If no data context, and there is a TemplateId associates with the map, then look up the 
                // required DataContext via TemplateId - PartId mappings defined in metadata.
                // Ie. Priority is given to explicit DataContext, not TemplateId set DataContect
                if (map.ParentDataContext == null && parentMap != null)
                {
                    map.ParentDataContext = parentMap.DataContext;
                }

                if (map.DataContext == null)
                {
                    object actualData = map.ParentDataContext;

                    // if the map has a template id - Look up the DataContect using TemplateId/PartId mapping in metadata package
                    if (!string.IsNullOrEmpty(map.TemplateId))
                    {
                        // then if the part is composite, attempt to set the datacontext from one if its data parts
                        if (map.ParentDataContext is ICompositeDataPart && exportPart.CompositeTemplateMappings != null)
                        {
                            string childPartId = GetMatchingChildPartId(map, exportPart);

                            // a child part id is found
                            if (!string.IsNullOrEmpty(childPartId))
                            {
                                var cdp = (ICompositeDataPart)map.ParentDataContext;

                                var childDataPart = GetChildDataPart(childPartId, cdp);
                                actualData = childDataPart;
                            }
                        }
                    }

                    // Finally set the data context for the map
                    map.DataContext = actualData;
                }
            }
            else
            {
                // If the evaluation of a binding on the data context results in data, then set.
                // If no data context, and there is a TemplateId associates with the map, then look up the 
                // required DataContext via TemplateId - PartId mappings defined in metadata.
                // Ie. Priority is given to explicit DataContext, not TemplateId set DataContect
                map.ParentDataContext = parentMap != null ? parentMap.DataContext : null;

                // 
                if (!string.IsNullOrEmpty(map.PartId) && this.dataParts != null)
                {
                    var matches = from dp in this.dataParts
                                  where map.PartId.CompareTo(dp.PartId) == 0
                                  select dp;

                    // stack panels might have a partId set and there might be more than 1 
                    if (map is StackPanel)  // extend to an interface if other items need this behaviour
                    {
                        var stackPanel = (StackPanel)map;
                        if (!string.IsNullOrEmpty(stackPanel.ItemTemplateMapKey) && matches.Count() > 1)
                        {
                            stackPanel.ItemsSource = matches;
                        }
                        else 
                        {
                            var match = matches.FirstOrDefault(); 
                            if (match != null)
                            {
                                map.DataContext = match.Data;
                            }
                        }
                    }
                    else
                    {
                        var match = matches.FirstOrDefault(); // address multiple matching dps
                        if (match != null)
                        {
                            map.DataContext = match.Data;
                        }
                    }
                }

                // If DataContext still note set, then flow down from parent
                if (map.DataContext == null)
                {
                    map.DataContext = map.ParentDataContext;
                }
            }

            // If the map is a positionable map, it will have a Placement, which can be bound.
            if (map is PositionableMap)
            {
                DataContextBase placement = (map as PositionableMap).Placement;
                placement.ParentDataContext = map.DataContext;

                // No binding set on DataContext, so flow from parent
                if (BindingContainer.GetSourceBindingOrValue(placement.DataContext) == null)
                {
                    placement.DataContext = map.DataContext;
                }
            }
        }

        /// <summary>
        /// Gets the child data part.
        /// </summary>
        /// <param name="childPartId">The child part identifier.</param>
        /// <param name="cdp">The CDP.</param>
        /// <returns></returns>
        private static IDataPart GetChildDataPart(string childPartId, ICompositeDataPart cdp)
        {
            // so get the actual part
            var childDataPart = (from c in cdp.DataParts
                                 where c.PartId == childPartId
                                 select c).FirstOrDefault();
            return childDataPart;
        }

        /// <summary>
        /// Gets the matching child part identifier.
        /// </summary>
        /// <param name="map">The map.</param>
        /// <param name="exportPart">The export part.</param>
        /// <returns></returns>
        private static string GetMatchingChildPartId(BaseMap map, ExportPart exportPart)
        {
            // look in the comp's table mappings for the part id that maps to this template id
            string childPartId = (from c in exportPart.CompositeTemplateMappings
                                  where c.TemplateId == map.TemplateId
                                  select c.PartId).FirstOrDefault();
            return childPartId;
        }

        /// <summary>
        /// Will attempt to return the worksheet part for the specified resource and template sheet name
        /// If not found will return the provided designer worksheet part
        /// </summary>
        /// <param name="resourceKey">The resource key.</param>
        /// <param name="templateSheetName">Name of the template sheet.</param>
        /// <param name="designerWorksheetPart">The designer worksheet part.</param>
        /// <returns></returns>
        private WorksheetPart GetDesignerWorksheet(string resourceKey, string templateSheetName, WorksheetPart designerWorksheetPart)
        {
            if (string.IsNullOrEmpty(templateSheetName))
            {
                return designerWorksheetPart;
            }
            if (this.resourceStore.ResourceDataDictionary.ContainsKey(resourceKey))
            {
                var designerSheetName = this.resourceStore.ResourceDataDictionary[resourceKey].DesignerFileName;

                if (!string.IsNullOrEmpty(designerSheetName))
                {
                    var designerSpreadsheetDocument = this.resourceStore.GetDesignerSpreadsheetDocumentByKey(resourceKey);
                    if (designerSpreadsheetDocument != null)
                    {
                        designerWorksheetPart = designerSpreadsheetDocument.GetWorksheetPart(templateSheetName);
                    }
                }
            }
            return designerWorksheetPart;
        }

        /// <summary>
        /// Remove all model-based elements from the worksheet that were created using Chart and Shape Template references
        /// </summary>
        /// <param name="map">The root of the map which was used to create a worksheet</param>
        private void RemoveTemplateElements(BaseMap map)
        {
            // Where am I removing the chart/shape from?
            string worksheetName = this.worksheetPart.GetSheetName();

            // Chart Templates
            IEnumerable<ChartTemplate> chartTemplates;
            if (map is Template)
            {
                // If supplied map in a legacy Template, then read from TemplateCollection, 
                // which is where all of the maps were historically defined.
                chartTemplates = (map as Template).TemplateCollection.GetElementsOfType<ChartTemplate>();
            }
            else
            {
                // Otherwise simply trawl all descendents.
                chartTemplates = map.AllDescendentsOfType<ChartTemplate>();
            }

            // Remove the chart from the output worksheet (if it was ever there.... The template will be the first one!!)
            foreach (ChartTemplate chartTemplate in chartTemplates)
            {
                string chartName = BindingContainer.ConvertToString(chartTemplate.TemplateChartName);
                OpenXml.Excel.Model.ChartModel chartModel = OpenXml.Excel.Model.ChartModel.GetChartModel(this.worksheetPart.Worksheet, chartName);

                if (chartModel != null)
                {
                    chartModel.RemoveChart();
                }
            }

            // Shape Templates
            IEnumerable<ShapeTemplate> shapeTemplates;
            if (map is Template)
            {
                shapeTemplates = (map as Template).TemplateCollection.GetElementsOfType<ShapeTemplate>();
            }
            else
            {
                shapeTemplates = map.AllDescendentsOfType<ShapeTemplate>();
            }

            foreach (ShapeTemplate shapeTemplate in shapeTemplates)
            {
                string shapeName = BindingContainer.ConvertToString(shapeTemplate.TemplateShapeName);
                OpenXml.Excel.Model.ShapeModel shapeModel = OpenXml.Excel.Model.ShapeModel.GetShapeModel(this.worksheetPart.Worksheet, shapeName);
                if (shapeModel != null)
                {
                    shapeModel.RemoveShape();
                }
            }

            // Picture Templates
            IEnumerable<PictureTemplate> pictureTemplates;
            if (map is Template)
            {
                pictureTemplates = (map as Template).TemplateCollection.GetElementsOfType<PictureTemplate>();
            }
            else
            {
                pictureTemplates = map.AllDescendentsOfType<PictureTemplate>();
            }

            foreach (PictureTemplate pictureTemplate in pictureTemplates)
            {
                string pictureName = BindingContainer.ConvertToString(pictureTemplate.TemplatePictureName);
                OpenXml.Excel.Model.PictureModel pictureModel = OpenXml.Excel.Model.PictureModel.GetPictureModel(this.worksheetPart.Worksheet, pictureName);
                if (pictureModel != null)
                {
                    pictureModel.RemovePicture();
                }
            }

        }

        /// <summary>
        /// Are any property rows in the collection hidden?
        /// </summary>
        /// <param name="properties">A collection of <see cref="Property">Properties</see></param>
        /// <returns>
        /// True if any marked RowIsHidden=True
        /// </returns>
        private static bool AnyPropertyRowsHidden(PropertyCollection properties)
        {
            if (properties == null)
            {
                return false;
            }

            return properties.Any(p => BindingContainer.ConvertToNullableBoolean(p.RowIsHidden).GetValueOrDefault(false));
        }

        #endregion Private Helpers
    }
}
