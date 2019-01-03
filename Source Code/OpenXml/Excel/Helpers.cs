namespace ExcelWriter.OpenXml.Excel
{
    using System.Collections.Generic;
    using System.Linq;
    using System.Windows.Documents;
    using DocumentFormat.OpenXml;
    using DocumentFormat.OpenXml.Spreadsheet;
    using System.Windows;
    using DocumentFormat.OpenXml.Packaging;

    /// <summary>
    /// 
    /// </summary>
    public sealed class Helpers
    {               
        #region Style merging

        /// <summary>
        /// Checks the shared string id
        /// </summary>
        /// <param name="stringId">The string identifier.</param>
        /// <param name="sourceDocument">The source document.</param>
        /// <param name="targetDocument">The target document.</param>
        /// <returns></returns>
        public static int GetOrCreateSharedStringIndex(int stringId, SpreadsheetDocument sourceDocument, SpreadsheetDocument targetDocument)
        {
            if (sourceDocument == null ||
                sourceDocument.WorkbookPart == null || 
                targetDocument == null ||
                targetDocument.WorkbookPart == null)
            { 
                return stringId; 
            }

            // if there's no shared string table in the source then dont bother trying
            if (sourceDocument.WorkbookPart.SharedStringTablePart == null) 
            {
                return stringId;
            }

            var sourceItems = sourceDocument.WorkbookPart.SharedStringTablePart.SharedStringTable.Descendants<SharedStringItem>();

            if (targetDocument.WorkbookPart.SharedStringTablePart == null)
            {
                targetDocument.WorkbookPart.AddNewPart<SharedStringTablePart>();
                targetDocument.WorkbookPart.SharedStringTablePart.SharedStringTable = new SharedStringTable();
            }                        
            var targetItems = targetDocument.WorkbookPart.SharedStringTablePart.SharedStringTable.Descendants<SharedStringItem>();
            
            var sourceItem = sourceItems.ElementAtOrDefault(stringId);
            var targetItem = targetItems.ElementAtOrDefault(stringId);

            if (sourceItem != null && sourceItem.Text != null)
            {
                // the source and target are the same string so just return what was sent in            
                if (targetItem != null && targetItem.Text != null && sourceItem.Text.Text == targetItem.Text.Text)
                {
                    return stringId;
                }

                // not the same so check if the same string isnt elsewhere
                var match = (from ss in targetItems
                             where ss.Text != null
                             && ss.Text.Text == sourceItem.Text.Text
                             select ss).FirstOrDefault();
                if (match != null) 
                {
                    return targetItems.ToList().IndexOf(match);
                }

                // still not found create a new one and add to the end
                SharedStringItem newSharedString = (SharedStringItem)sourceItem.CloneNode(true);
                targetDocument.WorkbookPart.SharedStringTablePart.SharedStringTable.Append(newSharedString);
                return targetItems.ToList().IndexOf(newSharedString);
            }
            return stringId;
        }

        /// <summary>
        /// Creates a differential format in the target based on the source using the sourceIndex
        /// </summary>
        /// <param name="sourceIndex">Index of the source.</param>
        /// <param name="source">The source.</param>
        /// <param name="target">The target.</param>
        /// <returns></returns>
        /// <exception cref="OpenXmlException"></exception>
        public static uint CreateDifferentialFormatIndex(uint sourceIndex, Stylesheet source, Stylesheet target)
        {
            var df = SafeIndex<DifferentialFormats, DifferentialFormat>(source.DifferentialFormats, (int)sourceIndex);
            if (df == null)
            {
                // this really ought to be an exception
                // return 0;
                throw new OpenXmlException(string.Format("Unable to find CellFormat with index <{0}>", sourceIndex));
            }

            var cloned = (DifferentialFormat)df.CloneNode(true);
            return target.AddDifferentialFormat(cloned);
        }

        /// <summary>
        /// Check to see if the source index matches with the target
        /// Return if matched and creating if necessary
        /// </summary>
        /// <param name="sourceId">The id of the source index</param>
        /// <param name="sourceStyleSheet">The source style sheet.</param>
        /// <param name="targetStyleSheet">The target style sheet.</param>
        /// <param name="mergedStyles">The merged styles.</param>
        /// <returns>
        /// The return id, might be the same as the source id
        /// </returns>
        public static uint GetOrCreateCellStyleIndex(uint sourceId, Stylesheet sourceStyleSheet, Stylesheet targetStyleSheet, Dictionary<uint, uint> mergedStyles)
        {
            if (mergedStyles.ContainsKey(sourceId))
            {
                return mergedStyles[sourceId];
            }

            uint newIndex = Helpers.GetOrCreateCellStyleIndexInner(sourceId, sourceStyleSheet, targetStyleSheet);
            mergedStyles.Add(sourceId, newIndex);

            return newIndex;
        }

        /// <summary>
        /// Gets the or create cell style index inner.
        /// </summary>
        /// <param name="index">The index.</param>
        /// <param name="source">The source.</param>
        /// <param name="target">The target.</param>
        /// <returns></returns>
        private static uint GetOrCreateCellStyleIndexInner(uint index, Stylesheet source, Stylesheet target)
        {
            var cellFormat = SafeIndex<CellFormats, CellFormat>(source.CellFormats, (int)index);
            if (cellFormat == null)
            {
                return 0;
                // this really ought to be an exception, it should never happen if i understand this correctly
                //throw new OpenXmlException(string.Format("Unable to find CellFormat with index <{0}>", index));
            }

            var cloned = (CellFormat)cellFormat.CloneNode(true);

            // get or create border if there is one
            if (cloned.BorderId != null && cloned.BorderId.HasValue)
            {
                cloned.BorderId.Value = GetOrCreateBorder(cloned.BorderId.Value, source, target);
            }

            // get or create fill if there is one
            if (cloned.FillId != null && cloned.FillId.HasValue)
            {
                cloned.FillId.Value = GetOrCreateFill(cloned.FillId.Value, source, target);
            }

            // get or create font if there is one
            if (cloned.FontId != null && cloned.FontId.HasValue)
            {
                cloned.FontId.Value = GetOrCreateFont(cloned.FontId.Value, source, target);
            }

            // get or create number format if there is one
            if (cloned.NumberFormatId != null && cloned.NumberFormatId.HasValue)
            {
                uint result = 0;
                if (TryGetOrCreateNumberingFormat(cloned.NumberFormatId.Value, source, target, out result))
                {
                    cloned.NumberFormatId.Value = result;
                }
            }

            // ?? get or create format if there is one
            //if (cloned.FormatId.HasValue){ } 

            return target.AddCellFormat(cloned);
        }

        /// <summary>
        /// Safes the index.
        /// </summary>
        /// <typeparam name="TParent">The type of the parent.</typeparam>
        /// <typeparam name="TChild">The type of the child.</typeparam>
        /// <param name="parent">The parent.</param>
        /// <param name="index">The index.</param>
        /// <returns></returns>
        private static TChild SafeIndex<TParent, TChild>(TParent parent, int index) where TParent : OpenXmlCompositeElement where TChild : OpenXmlElement
        {
            if (parent != null)
            {
                var elements = parent.Elements<TChild>().ToList();
                if (elements.Count > index)
                {
                    return elements[index];
                }
            }
            return null;
        }

        /// <summary>
        /// Gets the or create border.
        /// </summary>
        /// <param name="index">The index.</param>
        /// <param name="source">The source.</param>
        /// <param name="target">The target.</param>
        /// <returns></returns>
        private static uint GetOrCreateBorder(uint index, Stylesheet source, Stylesheet target)
        {
            var child = SafeIndex<Borders, Border>(source.Borders, (int)index);
            if (child == null)
            {
                return 0;
            }

            uint id = 0;

            if (target.TryMatchBorder(child, out id))
            {
                return id;
            }
            return target.AddBorder(child);
        }

        /// <summary>
        /// Gets the or create fill.
        /// </summary>
        /// <param name="index">The index.</param>
        /// <param name="source">The source.</param>
        /// <param name="target">The target.</param>
        /// <returns></returns>
        private static uint GetOrCreateFill(uint index, Stylesheet source, Stylesheet target)
        {
            var child = SafeIndex<Fills, Fill>(source.Fills, (int)index);
            if (child == null)
            {
                return 0;
            }

            uint id = 0;
            if (target.TryMatchFill(child, out id))
            {
                return id;
            }
            return target.AddFill(child);
        }

        /// <summary>
        /// Gets the or create font.
        /// </summary>
        /// <param name="index">The index.</param>
        /// <param name="source">The source.</param>
        /// <param name="target">The target.</param>
        /// <returns></returns>
        private static uint GetOrCreateFont(uint index, Stylesheet source, Stylesheet target)
        {
            var child = SafeIndex<Fonts, Font>(source.Fonts, (int)index);
            if (child == null)
            {
                return 0;
            }

            uint id = 0;
            if (target.TryMatchFont(child, out id))
            {
                return id;
            }
            return target.AddFont(child);
        }

        /// <summary>
        /// Tries the get or create numbering format.
        /// </summary>
        /// <param name="index">The index.</param>
        /// <param name="source">The source.</param>
        /// <param name="target">The target.</param>
        /// <param name="result">The result.</param>
        /// <returns></returns>
        private static bool TryGetOrCreateNumberingFormat(uint index, Stylesheet source, Stylesheet target, out uint result)
        {
            if (source.NumberingFormats == null)
            {
                source.NumberingFormats = new NumberingFormats();
            }

            var match = (from nf in source.NumberingFormats.Descendants<NumberingFormat>()
                         where nf.NumberFormatId.HasValue && nf.NumberFormatId.Value == index
                         select nf).FirstOrDefault();

            if (match == null)
            {
                result = 0;
                return false;
            }

            uint id = 0;

            //// to fix naughty use of reserved number formats
            //if (index < 164) 
            //{
            //    result = target.AddNumberingFormat(match);
            //    return true;
            //}

            if (target.TryMatchNumberingFormat(match, out id))
            {
                result = id;
            }
            else
            {
                result = target.AddNumberingFormat(match);
            }
            return true;
        }

        #endregion

        /// <summary>
        /// Recursive function to dive down an XML tree and look for any NumberReference. When it finds one it will replace the original
        /// source worksheet (and/or file reference) to the new worksheet name.
        /// This assumes all data sources are going to be updated to the new worksheet name. Will break any data source references
        /// that point to external files.
        /// </summary>
        /// <param name="element">Element to look for NumberReferences on.</param>
        /// <param name="oldSourceName">Old name of the source.</param>
        /// <param name="newSourceName">New name of the source.</param>
        /// <param name="tableRowCount">The table row count.</param>
        public static void UpdateDataSourcesForChildren(OpenXmlCompositeElement element, string oldSourceName, string newSourceName, int? tableRowCount)
        {
            foreach (var child in element.ChildElements)
            {
                if (child is DocumentFormat.OpenXml.Drawing.Charts.NumberReference)
                {
                    DocumentFormat.OpenXml.Drawing.Charts.NumberReference reference = (DocumentFormat.OpenXml.Drawing.Charts.NumberReference)child;
                    reference.Formula.Text = UpdateFormula(reference.Formula.Text, oldSourceName, newSourceName, tableRowCount);
                }
                else if (child is DocumentFormat.OpenXml.Drawing.Charts.StringReference)
                {
                    DocumentFormat.OpenXml.Drawing.Charts.StringReference reference = (DocumentFormat.OpenXml.Drawing.Charts.StringReference)child;
                    reference.Formula.Text = UpdateFormula(reference.Formula.Text, oldSourceName, newSourceName, tableRowCount);
                }
                else if (child is DocumentFormat.OpenXml.Drawing.Charts.MultiLevelStringReference)
                {
                    DocumentFormat.OpenXml.Drawing.Charts.MultiLevelStringReference reference = (DocumentFormat.OpenXml.Drawing.Charts.MultiLevelStringReference)child;
                    reference.Formula.Text = UpdateFormula(reference.Formula.Text, oldSourceName, newSourceName, tableRowCount);
                }
                else if (child != null && typeof(OpenXmlCompositeElement).IsAssignableFrom(child.GetType()) && ((OpenXmlCompositeElement)child).HasChildren)
                {
                    UpdateDataSourcesForChildren((OpenXmlCompositeElement)child, oldSourceName, newSourceName, tableRowCount);
                }
            }
        }

        /// <summary>
        /// Updates the formula string.
        /// First of all replacing old with new source names
        /// And if necessary changing any ranges with the supplied row count
        /// </summary>
        /// <param name="formulaText">The formula text.</param>
        /// <param name="oldSourceName">Old name of the source.</param>
        /// <param name="newSourceName">New name of the source.</param>
        /// <param name="rowCount">The row count.</param>
        /// <returns></returns>
        public static string UpdateFormula(string formulaText, string oldSourceName, string newSourceName, int? rowCount)
        {
            if (!newSourceName.StartsWith("'"))
            {
                newSourceName = string.Concat("'", newSourceName);
            }
            if (!newSourceName.EndsWith("'"))
            {
                newSourceName = string.Concat(newSourceName, "'");
            }

            // perform a simple string replacement of old with new name
            formulaText = formulaText.Replace(oldSourceName, newSourceName);

            // if a row count has been provided, then make sure ranges are updated to reflect this
            // eg change =Sheet!$A$3:$C$10 to =Sheet!$A$3:$C$20
            if (rowCount.HasValue)
            {
                // split using the : to get right and left hand side
                var split = formulaText.Split(':');

                // if more than part found
                if (split.Length > 1)
                {

                    int startCount = 0;

                    // get the left hand
                    string lhs = split[0];

                    // and split again....
                    var lhsSplit = lhs.Split('$');
                    if (lhsSplit.Length > 2)
                    {
                        //  to get the starting row number
                        string startStringCount = lhsSplit[2];
                        if (int.TryParse(startStringCount, out startCount))
                        {
                            // now split the right hand side
                            string rhs = split[1];
                            var rhsSplit = rhs.Split('$');

                            // and increment the start point by the number of rows expected
                            if (rhsSplit.Length > 2)
                            {
                                rhsSplit[2] = (startCount + (rowCount.Value - 1)).ToString();
                            }

                            // rebuild the range string
                            split[1] = string.Join("$", rhsSplit);
                        }
                    }
                    formulaText = string.Join(":", split);
                }
            }

            return formulaText;
        }

        /// <summary>
        /// Converts the paragraph.
        /// </summary>
        /// <param name="pgh">The PGH.</param>
        /// <param name="ils">The ils.</param>
        /// <param name="FontName">Name of the font.</param>
        /// <param name="Size">The size.</param>
        /// <returns></returns>
        public static InlineString ConvertParagraph(System.Windows.Documents.Paragraph pgh, InlineString ils, string FontName, decimal Size)
        {

            if (ils == null)
            {
                ils = new InlineString();
            }

            foreach (Inline ilSpan in pgh.Inlines)
            {
                if (ilSpan.GetType() == typeof(Span))
                {
                    Span sp = (Span)ilSpan;

                    foreach (object obj in sp.Inlines)
                    {
                        if (obj.GetType() == typeof(System.Windows.Documents.Run))
                        {
                            System.Windows.Documents.Run rn = (System.Windows.Documents.Run)obj;

                            List<OpenXmlLeafElement> DecorationList = new List<OpenXmlLeafElement>();

                            FontName fn = new FontName { Val = FontName };
                            FontSize fs = new FontSize { Val = (double) Size };
                            //Font.FontFamilyNumbering = new FontFamilyNumbering { Val = 2 };

                            if (rn.FontStyle == FontStyles.Italic)
                            {
                                DecorationList.Add(new DocumentFormat.OpenXml.Spreadsheet.Italic());
                            }

                            if (sp.TextDecorations.Contains(TextDecorations.Underline.First()))
                            {
                                DecorationList.Add(new DocumentFormat.OpenXml.Spreadsheet.Underline());
                            }

                            if (rn.FontWeight == FontWeights.Bold)
                            {
                                DecorationList.Add(new DocumentFormat.OpenXml.Spreadsheet.Bold());
                            }

                            ils.AppendChild(NewRunText(rn.Text, DecorationList, fn,fs));
                        }
                    }
                }
            }

            return ils;
        }

        /// <summary>
        /// Converts the paragraph list.
        /// </summary>
        /// <param name="pghList">The PGH list.</param>
        /// <param name="FontName">Name of the font.</param>
        /// <param name="Size">The size.</param>
        /// <returns></returns>
        public static List<InlineString> ConvertParagraphList(BlockCollection pghList, string FontName, decimal Size)
        {
            List<InlineString> NewList = new List<InlineString>();

            foreach (Paragraph pgh in pghList)
            {
                InlineString ils = new InlineString();

                NewList.Add(ConvertParagraph(pgh, ils, FontName, Size));
            }

            return NewList;
        }

        /// <summary>
        /// News the run text.
        /// </summary>
        /// <param name="NewText">The new text.</param>
        /// <param name="DecorationList">The decoration list.</param>
        /// <param name="FontName">Name of the font.</param>
        /// <param name="FontSize">Size of the font.</param>
        /// <returns></returns>
        private static DocumentFormat.OpenXml.Spreadsheet.Run NewRunText(string NewText, List<OpenXmlLeafElement> DecorationList, FontName FontName, FontSize FontSize)
        {
            DocumentFormat.OpenXml.Spreadsheet.Run run = new DocumentFormat.OpenXml.Spreadsheet.Run();

            Text txt = new Text(NewText);

            txt.SetAttribute(new OpenXmlAttribute("xml:space", null, "preserve"));

            RunProperties pPr = new RunProperties();
            run.Append(pPr);

            if (DecorationList != null)
            {
                foreach (OpenXmlLeafElement Decoration in DecorationList)
                {
                    pPr.Append(Decoration);
                }

                pPr.Append(FontName);
                pPr.Append(FontSize);
            }

            run.Append(txt);

            return run;
        }

        /// <summary>
        /// Return an index number for a set of column letters
        /// </summary>
        /// <param name="ColLetters">The col letters.</param>
        /// <returns></returns>
        public static uint ColLetterNumber(char[] ColLetters)
        {
            uint Number = 0;
            int index = 0;

            // To properly process units in a forward sequence the char arrays need to be reversed.

            char[] Chars = ColLetters.Reverse().ToArray();

            foreach (char c in Chars) // starting with the normal rightmost (as the string is reversed)
            {
                byte b = (byte)c;

                int i = (int)b;

                Number += (uint)((i - 64) * System.Math.Pow(26, index));

                index++;
            }

            return Number;
        }

    }
}
