namespace ExcelWriter.OpenXml.Excel
{
    using System.Linq;
    using DocumentFormat.OpenXml.Spreadsheet;

    public static class WorksheetExtensions
    {

        /// <summary>
        /// Sets the page setup orientation of a worksheet
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="orientation"></param>
        public static void SetOrientation(this Worksheet worksheet, OrientationValues orientation)
        {
            PageSetup pageSetup = worksheet.Descendants<PageSetup>().FirstOrDefault();
            if (pageSetup == null)
            {
                worksheet.AppendChild(
                        new PageSetup
                        {
                            Orientation = orientation
                        });                
            }
        }

        /// <summary>
        /// Using the worksheet hide the rows in the required limits.
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="RowStart"></param>
        /// <param name="RowEnd"></param>
        public static void HideRows(this Worksheet worksheet, uint RowStart, uint RowEnd)
        {
            foreach (Row row in worksheet.Descendants<Row>().ToList())
            {
                if ((row.RowIndex >= RowStart) & (row.RowIndex <= RowEnd))
                {
                    row.Hidden = true;
                }
            }
        }

        /// <summary>
        /// Using the worksheet show the rows in the required limits.
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="RowStart"></param>
        /// <param name="RowEnd"></param>
        public static void ShowRows(this Worksheet worksheet, uint RowStart, uint RowEnd)
        {
            foreach (Row row in worksheet.Descendants<Row>().ToList())
            {
                if ((row.RowIndex >= RowStart) & (row.RowIndex <= RowEnd))
                {
                    row.Hidden = false;
                }
            }
        }

        ///// <summary>
        ///// Iterate through the list of inline strings and add each to the consecutive rows from row index 
        ///// </summary>
        ///// <param name="worksheet"></param>
        ///// <param name="SheetData"></param>
        ///// <param name="ss"></param>
        ///// <param name="ilsList"></param>
        ///// <param name="ColIndex"></param>
        ///// <param name="RowIndex"></param>
        ///// <param name="MergedSpan"></param>
        //public static void ApplyInlineStringList(this Worksheet worksheet, SheetData SheetData,Stylesheet ss, List<InlineString> ilsList, uint ColIndex, uint RowIndex,uint MergedSpan)
        //{
        //    uint Row = RowIndex;

        //    foreach (InlineString ils in ilsList)
        //    {
        //        Cell cell = SheetData.GetCell(ColIndex, Row);
        //        Cell endcell = SheetData.GetCell(ColIndex + MergedSpan, Row);

        //        cell.RemoveAllChildren();

        //        if (cell.InlineString == null)
        //        {
        //            cell.DataType = CellValues.InlineString;
        //            cell.AppendChild(ils);
        //        }

        //        SheetData.MergeCells(cell, endcell);

        //        CellFormat cellFormat3 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyAlignment = true };

        //        //The Inline texts will always be aligned to top and will always wrap.
        //        Alignment alignment = new Alignment() { Vertical = VerticalAlignmentValues.Top, WrapText = true };

        //        cellFormat3.Append(alignment);

        //        ss.CellFormats.Count = (uint)ss.CellFormats.Count();
        //        uint cellFormatId = ss.CellFormats.Count - 1;

        //        cell.StyleIndex = cellFormatId;

        //        Row++;
        //    }
        //}

        ///// <summary>
        ///// The row height must be calculated from the amount of text, the width of the row, the font and font size and any text effects or decorations. This 
        ///// height will be applied to the inlinetext cell.
        ///// </summary>
        ///// <param name="row"></param>
        ///// <param name="FontFamily"></param>
        ///// <param name="emsize"></param>
        ///// <param name="LineWrapWidth"></param>
        ///// <param name="WidthMultiplier"></param>
        ///// <param name="HeightMultiplier"></param>
        ///// <returns></returns>
        //public static decimal RowHeight(Row row, string FontFamily, decimal emsize, decimal LineWrapWidth, decimal WidthMultiplier, decimal HeightMultiplier)
        //{
        //    decimal Height = 0;
        //    System.Drawing.Size RunLineSize;
        //    decimal TextLineWidth = 0;
        //    decimal TextLineHeight = 0;

        //    foreach (Cell c in row.Elements<Cell>())
        //    {
        //        if (c.InlineString != null)
        //        {
        //            foreach (DocumentFormat.OpenXml.Spreadsheet.Run r in c.InlineString.Elements<DocumentFormat.OpenXml.Spreadsheet.Run>())
        //            {
        //                System.Drawing.Font f = new System.Drawing.Font(FontFamily,(float) emsize);

        //                System.Drawing.FontStyle fs = new System.Drawing.FontStyle();

        //                if (r.RunProperties != null)
        //                {
        //                    foreach (DocumentFormat.OpenXml.Spreadsheet.Bold bold in r.RunProperties.Elements<DocumentFormat.OpenXml.Spreadsheet.Bold>())
        //                    {
        //                        fs = fs | System.Drawing.FontStyle.Bold;
        //                    }

        //                    foreach (DocumentFormat.OpenXml.Spreadsheet.Italic italic in r.RunProperties.Elements<DocumentFormat.OpenXml.Spreadsheet.Italic>())
        //                    {
        //                        fs = fs | System.Drawing.FontStyle.Italic;
        //                    }

        //                    foreach (DocumentFormat.OpenXml.Spreadsheet.Underline ul in r.RunProperties.Elements<DocumentFormat.OpenXml.Spreadsheet.Underline>())
        //                    {
        //                        fs = fs | System.Drawing.FontStyle.Underline;
        //                    }
        //                }

        //                if (f == null)
        //                {
        //                    f = new System.Drawing.Font(FontFamily,(float) emsize, fs);
        //                }

        //                Text txt = r.Text;

        //                RunLineSize = TextRenderer.MeasureText(txt.Text, f);

        //                TextLineWidth += RunLineSize.Width;
        //                if (RunLineSize.Height > TextLineHeight)
        //                {
        //                    TextLineHeight = RunLineSize.Height;
        //                }
        //            }
        //        }
        //    }

        //    decimal NumberWrappedLines = decimal.Truncate((TextLineWidth * WidthMultiplier) / LineWrapWidth) + 1;

        //    Height = NumberWrappedLines * TextLineHeight * HeightMultiplier;

        //    return Height;
        //}
    }

}   
   
