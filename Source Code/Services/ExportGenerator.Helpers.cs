namespace ExcelWriter
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Linq;
    using System.Text;
    using System.Windows.Documents;
    using System.Windows.Media;
    using DocumentFormat.OpenXml;
    using DocumentFormat.OpenXml.Drawing;
    using DocumentFormat.OpenXml.Packaging;
    using OpenXml.Excel;
    using DrawingSpreadsheet = DocumentFormat.OpenXml.Drawing.Spreadsheet;
    using Spreadsheet = DocumentFormat.OpenXml.Spreadsheet;

    /// <summary>
    /// 
    /// </summary>
    public sealed partial class ExportGenerator
    {
        /// <summary>
        /// Adds a set of parameters as debug data parts.
        /// </summary>
        /// <param name="dataParts">The data parts.</param>
        /// <param name="metadata">The metadata.</param>
        /// <param name="exportParameters">The export parameters.</param>
        /// <returns></returns>
        private static IEnumerable<IDataPart> AddDebugPart(IEnumerable<IDataPart> dataParts, Book metadata, ExportParameters exportParameters)
        {
            if (exportParameters != null && exportParameters.IncludeDebug)
            {
                var exportDP = new ExportParametersDataPart(exportParameters);
                dataParts = dataParts.Concat(new[] { exportDP });

                metadata.Parts.Add(new ExportPart
                {
                    PartId = exportDP.PartId,
                    DataSheetHidden = true,
                    DataSheetName = "_Parameters",
                    TemplateId = "Common.DebugTemplate"
                });
            }
            return dataParts;
        }

        ///// <summary>
        ///// Saves the memory stream to my documents with the provided title
        ///// </summary>
        ///// <param name="title">The title.</param>
        ///// <param name="stream">The stream.</param>
        ///// <returns>
        ///// The full path of the saved file
        ///// </returns>
        //public static string Save(string title, MemoryStream stream)
        //{
        //    string filePath = FileHelper.GetMyDocumentFilePath(title);
        //    using (FileStream fileStream = File.OpenWrite(filePath))
        //    {
        //        stream.WriteTo(fileStream);
        //    }
        //    return filePath;
        //}

        ///// <summary>
        ///// Saves the stream to the my documents with the title
        ///// Optionally opends
        ///// </summary>
        ///// <param name="open">if set to <c>true</c> [open].</param>
        ///// <param name="title">The title.</param>
        ///// <param name="stream">The stream.</param>
        ///// <returns>
        ///// The full path of the saved file
        ///// </returns>
        //public static string SaveAndOpen(bool open, string title, MemoryStream stream)
        //{
        //    // call save and open
        //    string outputFile = Save(title, stream);
        //    if (open)
        //    {
        //        ThreadPool.QueueUserWorkItem(
        //            (obj) =>
        //            {
        //                Process.Start(outputFile);
        //            });
        //    }
        //    return outputFile;
        //}

        #region Build of export sets

        /// <summary>
        /// Creates a list of export sets matching parts based on the provided data.
        /// For each data part try and find a matching export part.
        /// If any mandatory export parts havent been provided with data an exception will be thrown.
        /// </summary>
        /// <param name="dataParts">The data parts.</param>
        /// <param name="exportMetadata">The export metadata.</param>
        /// <param name="templatePackage">The template package.</param>
        /// <returns></returns>
        /// <exception cref="MetadataException"></exception>
        private List<ExportTripleSet> BuildSets(
            IEnumerable<IDataPart> dataParts, 
            Book exportMetadata, 
            ExcelTemplatePackage templatePackage)
        {
            // pair up the data with the export part and keep them for use later
            var sets = new List<ExportTripleSet>();

            // for each data part get the matching export part
            foreach (var dataPart in dataParts)
            {
                if (dataPart == null)
                {
                    continue;
                }

                // return the ExportPart(s) that matches with the provided data part
                foreach (var exportPart in SafePartMatch(dataPart, exportMetadata.Parts))
                {
                    // return a new instance of the template
                    var template = templatePackage.GetTemplateByTemplateId(exportPart.TemplateId);

                    // and add the 3 items to the set
                    sets.Add(new ExportTripleSet(dataPart, exportPart, template));
                }
            }

            // get a list of mandatory parts
            var mandatoryExportParts = from p in exportMetadata.Parts
                                       where p.IsMandatory
                                       select p;

            StringBuilder errors = new StringBuilder();

            // and for each of these check that they've been added to the set
            foreach (var mandatory in mandatoryExportParts)
            {
                bool match = (from s in sets
                              where s.Part == mandatory
                              select s).Any();
                // if we cant find one them add to our list of errors
                if (!match)
                {
                    errors.Append("No DataPart supplied for mandatory export part <");
                    errors.Append(mandatory.PartId);
                    errors.Append(">");
                    errors.Append(Environment.NewLine);
                }
            }

            // then throw any missing mandatorys as a metadata exception
            if (errors.Length > 0)
            {
                throw new MetadataException(errors.ToString());
            }

            return sets;
        }

        /// <summary>
        /// Tries to find an IPart from the list that matches with the provided IDataPart.
        /// ExportException if no matches or more than one match.
        /// Also calls IPart.Valid method and an ExportException will be thrown if the match proves to be invalid
        /// </summary>
        /// <param name="dataPart">The IDataPart to try and find an IPart match on</param>
        /// <param name="parts">The parts.</param>
        /// <returns>
        /// The IPart match
        /// </returns>
        private static IEnumerable<ExportPart> SafePartMatch(IDataPart dataPart, IEnumerable<ExportPart> parts)
        {
            // match on PartId
            foreach (var exportPart in from p in parts
                                       where p.PartId.CompareTo(dataPart.PartId) == 0
                                       select p)
            {
                // make sure its valid
                string exportPartErrorMsg = null;
                if (exportPart.Valid(out exportPartErrorMsg))
                {
                    yield return exportPart;
                }
            }
        }

        #endregion

        #region Helpers to prepare before and tidy after the process

        /// <summary>
        /// Sets up the document ready for export
        /// </summary>
        /// <param name="initSheets">if set to <c>true</c> [initialize sheets].</param>
        /// <param name="document">The document.</param>
        private static void BeginExport(bool initSheets, SpreadsheetDocument document)
        {
            // good practice to set this to true, otherwise any charts that have been populated during this process wont redraw
            document.WorkbookPart.Workbook.CalculationProperties = new DocumentFormat.OpenXml.Spreadsheet.CalculationProperties
            {
                ForceFullCalculation = true,
                FullCalculationOnLoad = true
            };

            //Spreadsheet.Stylesheet stylesheet = document.WorkbookPart.GetPartsOfType<WorkbookStylesPart>().First().Stylesheet;
            //stylesheet.InitializeIndexedColors();

            if (initSheets)
            {
                // delete all sheets other that data and report sheet
                foreach (var sheet in document.WorkbookPart.Workbook.Descendants<Spreadsheet.Sheet>().ToList())
                {
                    document.WorkbookPart.DeleteSheet(sheet);
                }
            }
        }

        /// <summary>
        /// Tidy up routine called at the end of the export
        /// </summary>
        /// <param name="document">The document.</param>
        private static void EndExport(SpreadsheetDocument document)
        {
            // set active tab to null,
            // so it defaults to the 1st sheet
            var view = document.WorkbookPart.Workbook.BookViews.GetFirstChild<Spreadsheet.WorkbookView>();
            if (view != null)
            {
                view.ActiveTab = null;
            }

            // ensure we don't have tab selected
            foreach (var wsp in document.WorkbookPart.WorksheetParts)
            {
                if (wsp.Worksheet.SheetViews != null)
                {
                    var sheetView = wsp.Worksheet.SheetViews.GetFirstChild<Spreadsheet.SheetView>();
                    if (sheetView != null && sheetView.TabSelected != null)
                    {
                        sheetView.TabSelected = BooleanValue.FromBoolean(false);
                    }
                }
            }
        }

        #endregion

        /// <summary>
        /// Xamls the section document reader.
        /// </summary>
        /// <param name="XAMLString">The xaml string.</param>
        /// <returns></returns>
        public static Section XamlSectionDocumentReader(string XAMLString)
        {
            System.Xml.XmlDocument xmlDoc = new System.Xml.XmlDocument();
            xmlDoc.LoadXml(XAMLString);
            Stream XAMLStream = new MemoryStream(System.Text.ASCIIEncoding.ASCII.GetBytes(xmlDoc.OuterXml));
            return (Section)System.Windows.Markup.XamlReader.Load(XAMLStream);
        }

        #region The RTF XAML paragraph format has a very different structure to the drawing.spreadheet.paragraph format this section contains methods that convert between

        /// <summary>
        /// Converts the paragraph.
        /// </summary>
        /// <param name="worksheetPart">The worksheet part.</param>
        /// <param name="RTFSection">The RTF section.</param>
        /// <param name="matchShape">The match shape.</param>
        /// <returns></returns>
        public static DrawingSpreadsheet.Shape ConvertParagraph(WorksheetPart worksheetPart, Section RTFSection, DrawingSpreadsheet.Shape matchShape)
        {
            DrawingSpreadsheet.Shape NewShape = new DrawingSpreadsheet.Shape();
            DocumentFormat.OpenXml.Drawing.Run ShapeRun = null;

            // get the graphic frame from the source anchor
            var sourceAnchor = worksheetPart.DrawingsPart.WorksheetDrawing.GetFirstChild<DrawingSpreadsheet.TwoCellAnchor>();
            var sourceShape = sourceAnchor.Descendants<DrawingSpreadsheet.Shape>().FirstOrDefault();

            // add it to the target anchor (ie. the one with the shape removed)
            DocumentFormat.OpenXml.Drawing.Spreadsheet.Shape targetShape = (DocumentFormat.OpenXml.Drawing.Spreadsheet.Shape)NewShape.CloneNode(true);

            targetShape.Append(matchShape.NonVisualShapeProperties.CloneNode(true));
            targetShape.Append(matchShape.ShapeProperties.CloneNode(true));
            targetShape.Append(matchShape.ShapeStyle.CloneNode(true));

            targetShape.Append(matchShape.TextBody.CloneNode(true));

            //Remove the text associated with the shape

            foreach (DocumentFormat.OpenXml.Drawing.Paragraph Paragraph in targetShape.TextBody.Descendants<DocumentFormat.OpenXml.Drawing.Paragraph>())
            {
                foreach (DocumentFormat.OpenXml.Drawing.Run RunToRemove in Paragraph.Descendants<DocumentFormat.OpenXml.Drawing.Run>())
                {
                    RunToRemove.Remove();
                }
            }

            try
            {
                foreach (System.Windows.Documents.Paragraph p in RTFSection.Blocks)
                {
                    InlineCollection ParagraphInLines = p.Inlines;


                    DocumentFormat.OpenXml.Drawing.Paragraph ShapeParagraph = new DocumentFormat.OpenXml.Drawing.Paragraph();

                    foreach (var InLine in ParagraphInLines)
                    {
                        if (InLine.GetType() == typeof(System.Windows.Documents.Span))
                        {
                            Span s = (Span)InLine;

                            Brush SourceTextColour = s.Foreground;

                            foreach (System.Windows.Documents.Run r in s.Inlines)
                            {
                                if (r.Text != "\n")
                                {
                                    ShapeRun = new DocumentFormat.OpenXml.Drawing.Run();

                                    DocumentFormat.OpenXml.Drawing.RunProperties ShapeRunProperties = new DocumentFormat.OpenXml.Drawing.RunProperties();
                                    DocumentFormat.OpenXml.Drawing.Text ShapeText = new DocumentFormat.OpenXml.Drawing.Text(r.Text);

                                    //the font family will be inherited from the sheet of the target shape
                                    ShapeRunProperties.FontSize = new Int32Value(System.Convert.ToInt32(s.FontSize * 100));

                                    SolidFill textFill = new SolidFill();

                                    SystemColor textColour = new SystemColor();

                                    textColour.Val = new EnumValue<SystemColorValues>(SystemColorValues.WindowText);

                                    Int64Value ColourMask = 0xFF000000;
                                    Int64Value SourceColour = System.Convert.ToInt64(SourceTextColour.ToString().Replace("#", "0x"), 16);

                                    SourceColour = SourceColour % ColourMask;

                                    textColour.LastColor = new HexBinaryValue(string.Format("{0,10:X}", SourceColour));

                                    textFill.SystemColor = textColour;

                                    ShapeRunProperties.Append(textFill);

                                    if (r.FontWeight == System.Windows.FontWeights.Bold)
                                    {
                                        ShapeRunProperties.Bold = true;
                                    }

                                    if (r.FontStyle == System.Windows.FontStyles.Italic)
                                    {
                                        ShapeRunProperties.Italic = true;
                                    }

                                    if (r.TextDecorations == System.Windows.TextDecorations.Underline)
                                    {
                                        ShapeRunProperties.Underline = new EnumValue<TextUnderlineValues>(TextUnderlineValues.Single);
                                    }

                                    ShapeRun.Text = ShapeText;
                                    ShapeRun.RunProperties = ShapeRunProperties;

                                    ShapeParagraph.Append(ShapeRun);
                                }
                            }

                        }
                        else
                        {
                            //do something else
                        }
                    }

                    targetShape.TextBody.Append(ShapeParagraph);
                }
            }
            catch (System.Exception ex)
            {
                throw ex;
            }

            return targetShape;

        }

        #endregion
    }
}
