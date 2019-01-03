namespace ExcelWriter.OpenXml.Excel
{
    using System.Collections.Generic;
    using System.Linq;
    using DocumentFormat.OpenXml.ExtendedProperties;
    using DocumentFormat.OpenXml.Packaging;
    using DocumentFormat.OpenXml.Spreadsheet;
    using DocumentFormat.OpenXml.VariantTypes;

    using ss = DocumentFormat.OpenXml.Spreadsheet;

    /// <summary>
    /// Extension methods for <see cref="SpreadsheetDocument"/>
    /// </summary>
    public static class SpreadsheetDocumentExtensions
    {
        #region Public Static Methods

        /// <summary>
        /// Either opens the <see cref="WorksheetPart" /> which has a specified sheet name, or creates that sheet
        /// within the supplied <see cref="SpreadsheetDocument" />.
        /// </summary>
        /// <param name="document">A <see cref="SpreadsheetDocument" /></param>
        /// <param name="sheetName">The worksheet name</param>
        /// <returns>
        /// The <see cref="WorksheetPart" />
        /// </returns>
        public static WorksheetPart CreateOrOpenSheet(this SpreadsheetDocument document, string sheetName)
        {
            WorkbookPart workbookPart = document.WorkbookPart;
            Workbook workbook = workbookPart.Workbook;
            var sheets = workbook.GetFirstChild<Sheets>();

            Sheet sheet = sheets.OfType<Sheet>().FirstOrDefault(x => string.Equals(x.Name, sheetName));
            
            if (sheet == null)
            {
                var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                var worksheet = new Worksheet();
                var sheetData = new SheetData();

                string worksheetPathId = workbookPart.GetIdOfPart(worksheetPart);
                uint sheetId = sheets.GetNextSheetId();

                worksheet.Append(sheetData);
                worksheet.Save(worksheetPart);

                sheet = new Sheet
                {
                    Id = worksheetPathId,
                    Name = sheetName,
                    SheetId = sheetId
                };
                sheets.Append(sheet);
                workbook.Save();
            }

            return (WorksheetPart)workbookPart.GetPartById(sheet.Id.Value);
        }

        /// <summary>
        /// Gets the worksheet part.
        /// </summary>
        /// <param name="document">The document.</param>
        /// <param name="sheetName">Name of the sheet.</param>
        /// <returns></returns>
        public static WorksheetPart GetWorksheetPart(this SpreadsheetDocument document, string sheetName)
        {
            WorkbookPart workbookPart = document.WorkbookPart;
            Workbook workbook = workbookPart.Workbook;
            var sheets = workbook.GetFirstChild<Sheets>();

            Sheet sheet = sheets.OfType<Sheet>().FirstOrDefault(x => string.Equals(x.Name, sheetName));
            if (sheet != null)
            {
                return (WorksheetPart)workbookPart.GetPartById(sheet.Id.Value);
            } 
            return null;
        }

        /// <summary>
        /// Gets the sheet data.
        /// </summary>
        /// <param name="document">The document.</param>
        /// <param name="sheetName">Name of the sheet.</param>
        /// <returns></returns>
        public static SheetData GetSheetData(this SpreadsheetDocument document, string sheetName)
        {
            WorkbookPart workbookPart = document.WorkbookPart;

            Sheet sheet = workbookPart.Workbook.Sheets.
                Cast<Sheet>().
                Single(x => string.Equals(x.Name, sheetName));
            WorksheetPart worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id);
            SheetData sheetData = worksheetPart.Worksheet.OfType<SheetData>().First();

            return sheetData;
        }

        /// <summary>
        /// Gets the stylesheet.
        /// </summary>
        /// <param name="document">The document.</param>
        /// <returns></returns>
        public static Stylesheet GetStylesheet(this SpreadsheetDocument document)
        {
            WorkbookStylesPart workbookStylesPart =
                document.WorkbookPart.GetPartsOfType<WorkbookStylesPart>().First();
            return workbookStylesPart.Stylesheet;
        }

        /// <summary>
        /// Initializes the blank document.
        /// </summary>
        /// <param name="document">The document.</param>
        public static void InitializeBlankDocument(this SpreadsheetDocument document)
        {
            // Add this so we can use named ranges.
            ExtendedFilePropertiesPart extendedFilePropertiesPart = document.AddNewPart<ExtendedFilePropertiesPart>("rId3");
            InitializeExtendedFileProperties(extendedFilePropertiesPart);

            WorkbookPart workbookPart = document.AddWorkbookPart();
            
            Workbook workbook = new Workbook();
            Sheets sheets = new Sheets();
            DefinedNames definedNames = new DefinedNames();

            WorkbookStylesPart workbookStylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
            workbookStylesPart.Stylesheet = new Stylesheet();

            InitializeStylesheet(workbookStylesPart.Stylesheet);

            workbook.Append(sheets);
            workbook.Append(definedNames);
            workbookPart.Workbook = workbook;
        }

        /// <summary>
        /// Sets the orientation.
        /// </summary>
        /// <param name="document">The document.</param>
        /// <param name="orientation">The orientation.</param>
        public static void SetOrientation(this SpreadsheetDocument document, OrientationValues orientation)
        {
            WorkbookPart workbookPart = document.WorkbookPart;
            IEnumerable<string> worksheetIds = workbookPart.Workbook.Descendants<Sheet>().Select(w => w.Id.Value);

            foreach (string worksheetId in worksheetIds)
            {
                WorksheetPart worksheetPart = ((WorksheetPart)workbookPart.GetPartById(worksheetId));
                DocumentFormat.OpenXml.Spreadsheet.PageSetup pageSetup =
                    worksheetPart.Worksheet.Descendants<DocumentFormat.OpenXml.Spreadsheet.PageSetup>().FirstOrDefault();
                if (pageSetup == null)
                {
                    pageSetup = worksheetPart.Worksheet.AppendChild<DocumentFormat.OpenXml.Spreadsheet.PageSetup>(
                        new DocumentFormat.OpenXml.Spreadsheet.PageSetup()
                        {
                            Orientation = orientation
                        });
                }
            }
        }

        /// <summary>
        /// Saves the and close.
        /// </summary>
        /// <param name="document">The document.</param>
        public static void SaveAndClose(this SpreadsheetDocument document)
        {
            foreach (WorksheetPart worksheetPart in document.WorkbookPart.GetPartsOfType<WorksheetPart>())
            {
                worksheetPart.Worksheet.Save();
            }

            document.WorkbookPart.Workbook.Save();
            document.Close();
        }

        #endregion

        #region Private Static Methods

        /// <summary>
        /// Generates the content of the extended file properties part. Add this so we can use named ranges.
        /// </summary>
        /// <param name="extendedFilePropertiesPart">The extended file properties part.</param>
        private static void InitializeExtendedFileProperties(ExtendedFilePropertiesPart extendedFilePropertiesPart)
        {
            DocumentFormat.OpenXml.ExtendedProperties.Properties properties1 = new DocumentFormat.OpenXml.ExtendedProperties.Properties();
            properties1.AddNamespaceDeclaration("vt", "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes");
            DocumentSecurity documentSecurity1 = new DocumentSecurity();
            documentSecurity1.Text = "0";
            ScaleCrop scaleCrop1 = new ScaleCrop();
            scaleCrop1.Text = "false";

            HeadingPairs headingPairs1 = new HeadingPairs();

            VTVector vTVector1 = new VTVector() { BaseType = VectorBaseValues.Variant, Size = 4U };

            Variant variant1 = new Variant();
            VTLPSTR vTLPSTR1 = new VTLPSTR();
            vTLPSTR1.Text = "Worksheets";

            variant1.Append(vTLPSTR1);

            Variant variant2 = new Variant();
            VTInt32 vTInt321 = new VTInt32();
            vTInt321.Text = "2";

            variant2.Append(vTInt321);

            Variant variant3 = new Variant();
            VTLPSTR vTLPSTR2 = new VTLPSTR();
            vTLPSTR2.Text = "Named Ranges";

            variant3.Append(vTLPSTR2);

            Variant variant4 = new Variant();
            VTInt32 vTInt322 = new VTInt32();
            vTInt322.Text = "2";

            variant4.Append(vTInt322);

            vTVector1.Append(variant1);
            vTVector1.Append(variant2);
            vTVector1.Append(variant3);
            vTVector1.Append(variant4);

            headingPairs1.Append(vTVector1);

            TitlesOfParts titlesOfParts1 = new TitlesOfParts();

            VTVector vTVector2 = new VTVector() { BaseType = VectorBaseValues.Lpstr, Size = 4U };
            VTLPSTR vTLPSTR3 = new VTLPSTR();
            vTLPSTR3.Text = "Model Example Table 1";
            VTLPSTR vTLPSTR4 = new VTLPSTR();
            vTLPSTR4.Text = "Model Example Table 2";
            VTLPSTR vTLPSTR5 = new VTLPSTR();
            vTLPSTR5.Text = "Legs";
            VTLPSTR vTLPSTR6 = new VTLPSTR();
            vTLPSTR6.Text = "Name";

            vTVector2.Append(vTLPSTR3);
            vTVector2.Append(vTLPSTR4);
            vTVector2.Append(vTLPSTR5);
            vTVector2.Append(vTLPSTR6);

            titlesOfParts1.Append(vTVector2);
            LinksUpToDate linksUpToDate1 = new LinksUpToDate();
            linksUpToDate1.Text = "false";
            SharedDocument sharedDocument1 = new SharedDocument();
            sharedDocument1.Text = "false";
            HyperlinksChanged hyperlinksChanged1 = new HyperlinksChanged();
            hyperlinksChanged1.Text = "false";
            ApplicationVersion applicationVersion1 = new ApplicationVersion();
            applicationVersion1.Text = "12.0000";

            properties1.Append(documentSecurity1);
            properties1.Append(scaleCrop1);
            properties1.Append(headingPairs1);
            properties1.Append(titlesOfParts1);
            properties1.Append(linksUpToDate1);
            properties1.Append(sharedDocument1);
            properties1.Append(hyperlinksChanged1);
            properties1.Append(applicationVersion1);

            extendedFilePropertiesPart.Properties = properties1;
        }

        /// <summary>
        /// Initializes a blank stylesheet with default styles.
        /// </summary>
        /// <param name="stylesheet">The stylesheet.</param>
        private static void InitializeStylesheet(Stylesheet stylesheet)
        {
            Fonts fonts1 = new Fonts() { Count = 20U };

            Font font1 = new Font();
            FontSize fontSize1 = new FontSize() { Val = 11D };
            Color color1 = new Color() { Theme = 1U };
            FontName fontName1 = new FontName() { Val = "Calibri" };
            FontFamilyNumbering fontFamilyNumbering1 = new FontFamilyNumbering() { Val = 2 };
            FontScheme fontScheme1 = new FontScheme() { Val = FontSchemeValues.Minor };

            font1.Append(fontSize1);
            font1.Append(color1);
            font1.Append(fontName1);
            font1.Append(fontFamilyNumbering1);
            font1.Append(fontScheme1);

            Font font2 = new Font();
            FontSize fontSize2 = new FontSize() { Val = 11D };
            Color color2 = new Color() { Indexed = 8U };
            FontName fontName2 = new FontName() { Val = "Calibri" };
            FontFamilyNumbering fontFamilyNumbering2 = new FontFamilyNumbering() { Val = 2 };

            font2.Append(fontSize2);
            font2.Append(color2);
            font2.Append(fontName2);
            font2.Append(fontFamilyNumbering2);

            Font font3 = new Font();
            FontSize fontSize3 = new FontSize() { Val = 11D };
            Color color3 = new Color() { Indexed = 10U };
            FontName fontName3 = new FontName() { Val = "Calibri" };
            FontFamilyNumbering fontFamilyNumbering3 = new FontFamilyNumbering() { Val = 2 };

            font3.Append(fontSize3);
            font3.Append(color3);
            font3.Append(fontName3);
            font3.Append(fontFamilyNumbering3);

            Font font4 = new Font();
            FontSize fontSize4 = new FontSize() { Val = 11D };
            Color color4 = new Color() { Theme = 1U };
            FontName fontName4 = new FontName() { Val = "Calibri" };
            FontFamilyNumbering fontFamilyNumbering4 = new FontFamilyNumbering() { Val = 2 };
            FontScheme fontScheme2 = new FontScheme() { Val = FontSchemeValues.Minor };

            font4.Append(fontSize4);
            font4.Append(color4);
            font4.Append(fontName4);
            font4.Append(fontFamilyNumbering4);
            font4.Append(fontScheme2);

            Font font5 = new Font();
            FontSize fontSize5 = new FontSize() { Val = 11D };
            Color color5 = new Color() { Theme = 0U };
            FontName fontName5 = new FontName() { Val = "Calibri" };
            FontFamilyNumbering fontFamilyNumbering5 = new FontFamilyNumbering() { Val = 2 };
            FontScheme fontScheme3 = new FontScheme() { Val = FontSchemeValues.Minor };

            font5.Append(fontSize5);
            font5.Append(color5);
            font5.Append(fontName5);
            font5.Append(fontFamilyNumbering5);
            font5.Append(fontScheme3);

            Font font6 = new Font();
            FontSize fontSize6 = new FontSize() { Val = 11D };
            Color color6 = new Color() { Rgb = "FF9C0006" };
            FontName fontName6 = new FontName() { Val = "Calibri" };
            FontFamilyNumbering fontFamilyNumbering6 = new FontFamilyNumbering() { Val = 2 };
            FontScheme fontScheme4 = new FontScheme() { Val = FontSchemeValues.Minor };

            font6.Append(fontSize6);
            font6.Append(color6);
            font6.Append(fontName6);
            font6.Append(fontFamilyNumbering6);
            font6.Append(fontScheme4);

            Font font7 = new Font();
            Bold bold1 = new Bold();
            FontSize fontSize7 = new FontSize() { Val = 11D };
            Color color7 = new Color() { Rgb = "FFFA7D00" };
            FontName fontName7 = new FontName() { Val = "Calibri" };
            FontFamilyNumbering fontFamilyNumbering7 = new FontFamilyNumbering() { Val = 2 };
            FontScheme fontScheme5 = new FontScheme() { Val = FontSchemeValues.Minor };

            font7.Append(bold1);
            font7.Append(fontSize7);
            font7.Append(color7);
            font7.Append(fontName7);
            font7.Append(fontFamilyNumbering7);
            font7.Append(fontScheme5);

            Font font8 = new Font();
            Bold bold2 = new Bold();
            FontSize fontSize8 = new FontSize() { Val = 11D };
            Color color8 = new Color() { Theme = 0U };
            FontName fontName8 = new FontName() { Val = "Calibri" };
            FontFamilyNumbering fontFamilyNumbering8 = new FontFamilyNumbering() { Val = 2 };
            FontScheme fontScheme6 = new FontScheme() { Val = FontSchemeValues.Minor };

            font8.Append(bold2);
            font8.Append(fontSize8);
            font8.Append(color8);
            font8.Append(fontName8);
            font8.Append(fontFamilyNumbering8);
            font8.Append(fontScheme6);

            Font font9 = new Font();
            Italic italic1 = new Italic();
            FontSize fontSize9 = new FontSize() { Val = 11D };
            Color color9 = new Color() { Rgb = "FF7F7F7F" };
            FontName fontName9 = new FontName() { Val = "Calibri" };
            FontFamilyNumbering fontFamilyNumbering9 = new FontFamilyNumbering() { Val = 2 };
            FontScheme fontScheme7 = new FontScheme() { Val = FontSchemeValues.Minor };

            font9.Append(italic1);
            font9.Append(fontSize9);
            font9.Append(color9);
            font9.Append(fontName9);
            font9.Append(fontFamilyNumbering9);
            font9.Append(fontScheme7);

            Font font10 = new Font();
            FontSize fontSize10 = new FontSize() { Val = 11D };
            Color color10 = new Color() { Rgb = "FF006100" };
            FontName fontName10 = new FontName() { Val = "Calibri" };
            FontFamilyNumbering fontFamilyNumbering10 = new FontFamilyNumbering() { Val = 2 };
            FontScheme fontScheme8 = new FontScheme() { Val = FontSchemeValues.Minor };

            font10.Append(fontSize10);
            font10.Append(color10);
            font10.Append(fontName10);
            font10.Append(fontFamilyNumbering10);
            font10.Append(fontScheme8);

            Font font11 = new Font();
            Bold bold3 = new Bold();
            FontSize fontSize11 = new FontSize() { Val = 15D };
            Color color11 = new Color() { Theme = 3U };
            FontName fontName11 = new FontName() { Val = "Calibri" };
            FontFamilyNumbering fontFamilyNumbering11 = new FontFamilyNumbering() { Val = 2 };
            FontScheme fontScheme9 = new FontScheme() { Val = FontSchemeValues.Minor };

            font11.Append(bold3);
            font11.Append(fontSize11);
            font11.Append(color11);
            font11.Append(fontName11);
            font11.Append(fontFamilyNumbering11);
            font11.Append(fontScheme9);

            var font12 = new Font();
            var bold4 = new Bold();
            var fontSize12 = new FontSize { Val = 13D };
            var color12 = new Color { Theme = 3U };
            var fontName12 = new FontName { Val = "Calibri" };
            var fontFamilyNumbering12 = new FontFamilyNumbering() { Val = 2 };
            var fontScheme10 = new FontScheme { Val = FontSchemeValues.Minor };

            font12.Append(bold4);
            font12.Append(fontSize12);
            font12.Append(color12);
            font12.Append(fontName12);
            font12.Append(fontFamilyNumbering12);
            font12.Append(fontScheme10);

            Font font13 = new Font();
            Bold bold5 = new Bold();
            FontSize fontSize13 = new FontSize() { Val = 11D };
            Color color13 = new Color() { Theme = 3U };
            FontName fontName13 = new FontName() { Val = "Calibri" };
            FontFamilyNumbering fontFamilyNumbering13 = new FontFamilyNumbering() { Val = 2 };
            FontScheme fontScheme11 = new FontScheme() { Val = FontSchemeValues.Minor };

            font13.Append(bold5);
            font13.Append(fontSize13);
            font13.Append(color13);
            font13.Append(fontName13);
            font13.Append(fontFamilyNumbering13);
            font13.Append(fontScheme11);

            Font font14 = new Font();
            FontSize fontSize14 = new FontSize() { Val = 11D };
            Color color14 = new Color() { Rgb = "FF3F3F76" };
            FontName fontName14 = new FontName() { Val = "Calibri" };
            FontFamilyNumbering fontFamilyNumbering14 = new FontFamilyNumbering() { Val = 2 };
            FontScheme fontScheme12 = new FontScheme() { Val = FontSchemeValues.Minor };

            font14.Append(fontSize14);
            font14.Append(color14);
            font14.Append(fontName14);
            font14.Append(fontFamilyNumbering14);
            font14.Append(fontScheme12);

            Font font15 = new Font();
            FontSize fontSize15 = new FontSize() { Val = 11D };
            Color color15 = new Color() { Rgb = "FFFA7D00" };
            FontName fontName15 = new FontName() { Val = "Calibri" };
            FontFamilyNumbering fontFamilyNumbering15 = new FontFamilyNumbering() { Val = 2 };
            FontScheme fontScheme13 = new FontScheme() { Val = FontSchemeValues.Minor };

            font15.Append(fontSize15);
            font15.Append(color15);
            font15.Append(fontName15);
            font15.Append(fontFamilyNumbering15);
            font15.Append(fontScheme13);

            Font font16 = new Font();
            FontSize fontSize16 = new FontSize() { Val = 11D };
            Color color16 = new Color() { Rgb = "FF9C6500" };
            FontName fontName16 = new FontName() { Val = "Calibri" };
            FontFamilyNumbering fontFamilyNumbering16 = new FontFamilyNumbering() { Val = 2 };
            FontScheme fontScheme14 = new FontScheme() { Val = FontSchemeValues.Minor };

            font16.Append(fontSize16);
            font16.Append(color16);
            font16.Append(fontName16);
            font16.Append(fontFamilyNumbering16);
            font16.Append(fontScheme14);

            Font font17 = new Font();
            Bold bold6 = new Bold();
            FontSize fontSize17 = new FontSize() { Val = 11D };
            Color color17 = new Color() { Rgb = "FF3F3F3F" };
            FontName fontName17 = new FontName() { Val = "Calibri" };
            FontFamilyNumbering fontFamilyNumbering17 = new FontFamilyNumbering() { Val = 2 };
            FontScheme fontScheme15 = new FontScheme() { Val = FontSchemeValues.Minor };

            font17.Append(bold6);
            font17.Append(fontSize17);
            font17.Append(color17);
            font17.Append(fontName17);
            font17.Append(fontFamilyNumbering17);
            font17.Append(fontScheme15);

            Font font18 = new Font();
            Bold bold7 = new Bold();
            FontSize fontSize18 = new FontSize() { Val = 18D };
            Color color18 = new Color() { Theme = 3U };
            FontName fontName18 = new FontName() { Val = "Cambria" };
            FontFamilyNumbering fontFamilyNumbering18 = new FontFamilyNumbering() { Val = 2 };
            FontScheme fontScheme16 = new FontScheme() { Val = FontSchemeValues.Major };

            font18.Append(bold7);
            font18.Append(fontSize18);
            font18.Append(color18);
            font18.Append(fontName18);
            font18.Append(fontFamilyNumbering18);
            font18.Append(fontScheme16);

            Font font19 = new Font();
            Bold bold8 = new Bold();
            FontSize fontSize19 = new FontSize() { Val = 11D };
            Color color19 = new Color() { Theme = 1U };
            FontName fontName19 = new FontName() { Val = "Calibri" };
            FontFamilyNumbering fontFamilyNumbering19 = new FontFamilyNumbering() { Val = 2 };
            FontScheme fontScheme17 = new FontScheme() { Val = FontSchemeValues.Minor };

            font19.Append(bold8);
            font19.Append(fontSize19);
            font19.Append(color19);
            font19.Append(fontName19);
            font19.Append(fontFamilyNumbering19);
            font19.Append(fontScheme17);

            Font font20 = new Font();
            FontSize fontSize20 = new FontSize() { Val = 11D };
            Color color20 = new Color() { Rgb = "FFFF0000" };
            FontName fontName20 = new FontName() { Val = "Calibri" };
            FontFamilyNumbering fontFamilyNumbering20 = new FontFamilyNumbering() { Val = 2 };
            FontScheme fontScheme18 = new FontScheme() { Val = FontSchemeValues.Minor };

            font20.Append(fontSize20);
            font20.Append(color20);
            font20.Append(fontName20);
            font20.Append(fontFamilyNumbering20);
            font20.Append(fontScheme18);

            fonts1.Append(font1);
            fonts1.Append(font2);
            fonts1.Append(font3);
            fonts1.Append(font4);
            fonts1.Append(font5);
            fonts1.Append(font6);
            fonts1.Append(font7);
            fonts1.Append(font8);
            fonts1.Append(font9);
            fonts1.Append(font10);
            fonts1.Append(font11);
            fonts1.Append(font12);
            fonts1.Append(font13);
            fonts1.Append(font14);
            fonts1.Append(font15);
            fonts1.Append(font16);
            fonts1.Append(font17);
            fonts1.Append(font18);
            fonts1.Append(font19);
            fonts1.Append(font20);

            Fills fills1 = new Fills() { Count = 34U };

            Fill fill1 = new Fill();
            PatternFill patternFill1 = new PatternFill() { PatternType = PatternValues.None };

            fill1.Append(patternFill1);

            Fill fill2 = new Fill();
            PatternFill patternFill2 = new PatternFill() { PatternType = PatternValues.Gray125 };

            fill2.Append(patternFill2);

            Fill fill3 = new Fill();

            PatternFill patternFill3 = new PatternFill() { PatternType = PatternValues.Solid };
            ForegroundColor foregroundColor1 = new ForegroundColor() { Indexed = 10U };
            BackgroundColor backgroundColor1 = new BackgroundColor() { Indexed = 64U };

            patternFill3.Append(foregroundColor1);
            patternFill3.Append(backgroundColor1);

            fill3.Append(patternFill3);

            Fill fill4 = new Fill();

            PatternFill patternFill4 = new PatternFill() { PatternType = PatternValues.Solid };
            ForegroundColor foregroundColor2 = new ForegroundColor() { Theme = 4U, Tint = 0.79998168889431442D };
            BackgroundColor backgroundColor2 = new BackgroundColor() { Indexed = 65U };

            patternFill4.Append(foregroundColor2);
            patternFill4.Append(backgroundColor2);

            fill4.Append(patternFill4);

            Fill fill5 = new Fill();

            PatternFill patternFill5 = new PatternFill() { PatternType = PatternValues.Solid };
            ForegroundColor foregroundColor3 = new ForegroundColor() { Theme = 5U, Tint = 0.79998168889431442D };
            BackgroundColor backgroundColor3 = new BackgroundColor() { Indexed = 65U };

            patternFill5.Append(foregroundColor3);
            patternFill5.Append(backgroundColor3);

            fill5.Append(patternFill5);

            Fill fill6 = new Fill();

            PatternFill patternFill6 = new PatternFill() { PatternType = PatternValues.Solid };
            ForegroundColor foregroundColor4 = new ForegroundColor() { Theme = 6U, Tint = 0.79998168889431442D };
            BackgroundColor backgroundColor4 = new BackgroundColor() { Indexed = 65U };

            patternFill6.Append(foregroundColor4);
            patternFill6.Append(backgroundColor4);

            fill6.Append(patternFill6);

            Fill fill7 = new Fill();

            PatternFill patternFill7 = new PatternFill() { PatternType = PatternValues.Solid };
            ForegroundColor foregroundColor5 = new ForegroundColor() { Theme = 7U, Tint = 0.79998168889431442D };
            BackgroundColor backgroundColor5 = new BackgroundColor() { Indexed = 65U };

            patternFill7.Append(foregroundColor5);
            patternFill7.Append(backgroundColor5);

            fill7.Append(patternFill7);

            Fill fill8 = new Fill();

            PatternFill patternFill8 = new PatternFill() { PatternType = PatternValues.Solid };
            ForegroundColor foregroundColor6 = new ForegroundColor() { Theme = 8U, Tint = 0.79998168889431442D };
            BackgroundColor backgroundColor6 = new BackgroundColor() { Indexed = 65U };

            patternFill8.Append(foregroundColor6);
            patternFill8.Append(backgroundColor6);

            fill8.Append(patternFill8);

            Fill fill9 = new Fill();

            PatternFill patternFill9 = new PatternFill() { PatternType = PatternValues.Solid };
            ForegroundColor foregroundColor7 = new ForegroundColor() { Theme = 9U, Tint = 0.79998168889431442D };
            BackgroundColor backgroundColor7 = new BackgroundColor() { Indexed = 65U };

            patternFill9.Append(foregroundColor7);
            patternFill9.Append(backgroundColor7);

            fill9.Append(patternFill9);

            Fill fill10 = new Fill();

            PatternFill patternFill10 = new PatternFill() { PatternType = PatternValues.Solid };
            ForegroundColor foregroundColor8 = new ForegroundColor() { Theme = 4U, Tint = 0.59999389629810485D };
            BackgroundColor backgroundColor8 = new BackgroundColor() { Indexed = 65U };

            patternFill10.Append(foregroundColor8);
            patternFill10.Append(backgroundColor8);

            fill10.Append(patternFill10);

            Fill fill11 = new Fill();

            PatternFill patternFill11 = new PatternFill() { PatternType = PatternValues.Solid };
            ForegroundColor foregroundColor9 = new ForegroundColor() { Theme = 5U, Tint = 0.59999389629810485D };
            BackgroundColor backgroundColor9 = new BackgroundColor() { Indexed = 65U };

            patternFill11.Append(foregroundColor9);
            patternFill11.Append(backgroundColor9);

            fill11.Append(patternFill11);

            Fill fill12 = new Fill();

            PatternFill patternFill12 = new PatternFill() { PatternType = PatternValues.Solid };
            ForegroundColor foregroundColor10 = new ForegroundColor() { Theme = 6U, Tint = 0.59999389629810485D };
            BackgroundColor backgroundColor10 = new BackgroundColor() { Indexed = 65U };

            patternFill12.Append(foregroundColor10);
            patternFill12.Append(backgroundColor10);

            fill12.Append(patternFill12);

            Fill fill13 = new Fill();

            PatternFill patternFill13 = new PatternFill() { PatternType = PatternValues.Solid };
            ForegroundColor foregroundColor11 = new ForegroundColor() { Theme = 7U, Tint = 0.59999389629810485D };
            BackgroundColor backgroundColor11 = new BackgroundColor() { Indexed = 65U };

            patternFill13.Append(foregroundColor11);
            patternFill13.Append(backgroundColor11);

            fill13.Append(patternFill13);

            Fill fill14 = new Fill();

            PatternFill patternFill14 = new PatternFill() { PatternType = PatternValues.Solid };
            ForegroundColor foregroundColor12 = new ForegroundColor() { Theme = 8U, Tint = 0.59999389629810485D };
            BackgroundColor backgroundColor12 = new BackgroundColor() { Indexed = 65U };

            patternFill14.Append(foregroundColor12);
            patternFill14.Append(backgroundColor12);

            fill14.Append(patternFill14);

            Fill fill15 = new Fill();

            PatternFill patternFill15 = new PatternFill() { PatternType = PatternValues.Solid };
            ForegroundColor foregroundColor13 = new ForegroundColor() { Theme = 9U, Tint = 0.59999389629810485D };
            BackgroundColor backgroundColor13 = new BackgroundColor() { Indexed = 65U };

            patternFill15.Append(foregroundColor13);
            patternFill15.Append(backgroundColor13);

            fill15.Append(patternFill15);

            Fill fill16 = new Fill();

            PatternFill patternFill16 = new PatternFill() { PatternType = PatternValues.Solid };
            ForegroundColor foregroundColor14 = new ForegroundColor() { Theme = 4U, Tint = 0.39997558519241921D };
            BackgroundColor backgroundColor14 = new BackgroundColor() { Indexed = 65U };

            patternFill16.Append(foregroundColor14);
            patternFill16.Append(backgroundColor14);

            fill16.Append(patternFill16);

            Fill fill17 = new Fill();

            PatternFill patternFill17 = new PatternFill() { PatternType = PatternValues.Solid };
            ForegroundColor foregroundColor15 = new ForegroundColor() { Theme = 5U, Tint = 0.39997558519241921D };
            BackgroundColor backgroundColor15 = new BackgroundColor() { Indexed = 65U };

            patternFill17.Append(foregroundColor15);
            patternFill17.Append(backgroundColor15);

            fill17.Append(patternFill17);

            Fill fill18 = new Fill();

            PatternFill patternFill18 = new PatternFill() { PatternType = PatternValues.Solid };
            ForegroundColor foregroundColor16 = new ForegroundColor() { Theme = 6U, Tint = 0.39997558519241921D };
            BackgroundColor backgroundColor16 = new BackgroundColor() { Indexed = 65U };

            patternFill18.Append(foregroundColor16);
            patternFill18.Append(backgroundColor16);

            fill18.Append(patternFill18);

            Fill fill19 = new Fill();

            PatternFill patternFill19 = new PatternFill() { PatternType = PatternValues.Solid };
            ForegroundColor foregroundColor17 = new ForegroundColor() { Theme = 7U, Tint = 0.39997558519241921D };
            BackgroundColor backgroundColor17 = new BackgroundColor() { Indexed = 65U };

            patternFill19.Append(foregroundColor17);
            patternFill19.Append(backgroundColor17);

            fill19.Append(patternFill19);

            Fill fill20 = new Fill();

            PatternFill patternFill20 = new PatternFill() { PatternType = PatternValues.Solid };
            ForegroundColor foregroundColor18 = new ForegroundColor() { Theme = 8U, Tint = 0.39997558519241921D };
            BackgroundColor backgroundColor18 = new BackgroundColor() { Indexed = 65U };

            patternFill20.Append(foregroundColor18);
            patternFill20.Append(backgroundColor18);

            fill20.Append(patternFill20);

            Fill fill21 = new Fill();

            PatternFill patternFill21 = new PatternFill() { PatternType = PatternValues.Solid };
            ForegroundColor foregroundColor19 = new ForegroundColor() { Theme = 9U, Tint = 0.39997558519241921D };
            BackgroundColor backgroundColor19 = new BackgroundColor() { Indexed = 65U };

            patternFill21.Append(foregroundColor19);
            patternFill21.Append(backgroundColor19);

            fill21.Append(patternFill21);

            Fill fill22 = new Fill();

            PatternFill patternFill22 = new PatternFill() { PatternType = PatternValues.Solid };
            ForegroundColor foregroundColor20 = new ForegroundColor() { Theme = 4U };

            patternFill22.Append(foregroundColor20);

            fill22.Append(patternFill22);

            Fill fill23 = new Fill();

            PatternFill patternFill23 = new PatternFill() { PatternType = PatternValues.Solid };
            ForegroundColor foregroundColor21 = new ForegroundColor() { Theme = 5U };

            patternFill23.Append(foregroundColor21);

            fill23.Append(patternFill23);

            Fill fill24 = new Fill();

            PatternFill patternFill24 = new PatternFill() { PatternType = PatternValues.Solid };
            ForegroundColor foregroundColor22 = new ForegroundColor() { Theme = 6U };

            patternFill24.Append(foregroundColor22);

            fill24.Append(patternFill24);

            Fill fill25 = new Fill();

            PatternFill patternFill25 = new PatternFill() { PatternType = PatternValues.Solid };
            ForegroundColor foregroundColor23 = new ForegroundColor() { Theme = 7U };

            patternFill25.Append(foregroundColor23);

            fill25.Append(patternFill25);

            Fill fill26 = new Fill();

            PatternFill patternFill26 = new PatternFill() { PatternType = PatternValues.Solid };
            ForegroundColor foregroundColor24 = new ForegroundColor() { Theme = 8U };

            patternFill26.Append(foregroundColor24);

            fill26.Append(patternFill26);

            Fill fill27 = new Fill();

            PatternFill patternFill27 = new PatternFill() { PatternType = PatternValues.Solid };
            ForegroundColor foregroundColor25 = new ForegroundColor() { Theme = 9U };

            patternFill27.Append(foregroundColor25);

            fill27.Append(patternFill27);

            Fill fill28 = new Fill();

            PatternFill patternFill28 = new PatternFill() { PatternType = PatternValues.Solid };
            ForegroundColor foregroundColor26 = new ForegroundColor() { Rgb = "FFFFC7CE" };

            patternFill28.Append(foregroundColor26);

            fill28.Append(patternFill28);

            Fill fill29 = new Fill();

            PatternFill patternFill29 = new PatternFill() { PatternType = PatternValues.Solid };
            ForegroundColor foregroundColor27 = new ForegroundColor() { Rgb = "FFF2F2F2" };

            patternFill29.Append(foregroundColor27);

            fill29.Append(patternFill29);

            Fill fill30 = new Fill();

            PatternFill patternFill30 = new PatternFill() { PatternType = PatternValues.Solid };
            ForegroundColor foregroundColor28 = new ForegroundColor() { Rgb = "FFA5A5A5" };

            patternFill30.Append(foregroundColor28);

            fill30.Append(patternFill30);

            Fill fill31 = new Fill();

            PatternFill patternFill31 = new PatternFill() { PatternType = PatternValues.Solid };
            ForegroundColor foregroundColor29 = new ForegroundColor() { Rgb = "FFC6EFCE" };

            patternFill31.Append(foregroundColor29);

            fill31.Append(patternFill31);

            Fill fill32 = new Fill();

            PatternFill patternFill32 = new PatternFill() { PatternType = PatternValues.Solid };
            ForegroundColor foregroundColor30 = new ForegroundColor() { Rgb = "FFFFCC99" };

            patternFill32.Append(foregroundColor30);

            fill32.Append(patternFill32);

            Fill fill33 = new Fill();

            PatternFill patternFill33 = new PatternFill() { PatternType = PatternValues.Solid };
            ForegroundColor foregroundColor31 = new ForegroundColor() { Rgb = "FFFFEB9C" };

            patternFill33.Append(foregroundColor31);

            fill33.Append(patternFill33);

            Fill fill34 = new Fill();

            PatternFill patternFill34 = new PatternFill() { PatternType = PatternValues.Solid };
            ForegroundColor foregroundColor32 = new ForegroundColor() { Rgb = "FFFFFFCC" };

            patternFill34.Append(foregroundColor32);

            fill34.Append(patternFill34);

            fills1.Append(fill1);
            fills1.Append(fill2);
            fills1.Append(fill3);
            fills1.Append(fill4);
            fills1.Append(fill5);
            fills1.Append(fill6);
            fills1.Append(fill7);
            fills1.Append(fill8);
            fills1.Append(fill9);
            fills1.Append(fill10);
            fills1.Append(fill11);
            fills1.Append(fill12);
            fills1.Append(fill13);
            fills1.Append(fill14);
            fills1.Append(fill15);
            fills1.Append(fill16);
            fills1.Append(fill17);
            fills1.Append(fill18);
            fills1.Append(fill19);
            fills1.Append(fill20);
            fills1.Append(fill21);
            fills1.Append(fill22);
            fills1.Append(fill23);
            fills1.Append(fill24);
            fills1.Append(fill25);
            fills1.Append(fill26);
            fills1.Append(fill27);
            fills1.Append(fill28);
            fills1.Append(fill29);
            fills1.Append(fill30);
            fills1.Append(fill31);
            fills1.Append(fill32);
            fills1.Append(fill33);
            fills1.Append(fill34);

            Borders borders1 = new Borders() { Count = 10U };

            DocumentFormat.OpenXml.Spreadsheet.Border border1 = new DocumentFormat.OpenXml.Spreadsheet.Border();
            LeftBorder leftBorder1 = new LeftBorder();
            RightBorder rightBorder1 = new RightBorder();
            TopBorder topBorder1 = new TopBorder();
            BottomBorder bottomBorder1 = new BottomBorder();
            DiagonalBorder diagonalBorder1 = new DiagonalBorder();

            border1.Append(leftBorder1);
            border1.Append(rightBorder1);
            border1.Append(topBorder1);
            border1.Append(bottomBorder1);
            border1.Append(diagonalBorder1);

            DocumentFormat.OpenXml.Spreadsheet.Border border2 = new DocumentFormat.OpenXml.Spreadsheet.Border();

            LeftBorder leftBorder2 = new LeftBorder() { Style = BorderStyleValues.Thin };
            Color color21 = new Color() { Rgb = "FF7F7F7F" };

            leftBorder2.Append(color21);

            RightBorder rightBorder2 = new RightBorder() { Style = BorderStyleValues.Thin };
            Color color22 = new Color() { Rgb = "FF7F7F7F" };

            rightBorder2.Append(color22);

            TopBorder topBorder2 = new TopBorder() { Style = BorderStyleValues.Thin };
            Color color23 = new Color() { Rgb = "FF7F7F7F" };

            topBorder2.Append(color23);

            BottomBorder bottomBorder2 = new BottomBorder() { Style = BorderStyleValues.Thin };
            Color color24 = new Color() { Rgb = "FF7F7F7F" };

            bottomBorder2.Append(color24);
            DiagonalBorder diagonalBorder2 = new DiagonalBorder();

            border2.Append(leftBorder2);
            border2.Append(rightBorder2);
            border2.Append(topBorder2);
            border2.Append(bottomBorder2);
            border2.Append(diagonalBorder2);

            DocumentFormat.OpenXml.Spreadsheet.Border border3 = new DocumentFormat.OpenXml.Spreadsheet.Border();

            LeftBorder leftBorder3 = new LeftBorder() { Style = BorderStyleValues.Double };
            Color color25 = new Color() { Rgb = "FF3F3F3F" };

            leftBorder3.Append(color25);

            RightBorder rightBorder3 = new RightBorder() { Style = BorderStyleValues.Double };
            Color color26 = new Color() { Rgb = "FF3F3F3F" };

            rightBorder3.Append(color26);

            TopBorder topBorder3 = new TopBorder() { Style = BorderStyleValues.Double };
            Color color27 = new Color() { Rgb = "FF3F3F3F" };

            topBorder3.Append(color27);

            BottomBorder bottomBorder3 = new BottomBorder() { Style = BorderStyleValues.Double };
            Color color28 = new Color() { Rgb = "FF3F3F3F" };

            bottomBorder3.Append(color28);
            DiagonalBorder diagonalBorder3 = new DiagonalBorder();

            border3.Append(leftBorder3);
            border3.Append(rightBorder3);
            border3.Append(topBorder3);
            border3.Append(bottomBorder3);
            border3.Append(diagonalBorder3);

            DocumentFormat.OpenXml.Spreadsheet.Border border4 = new DocumentFormat.OpenXml.Spreadsheet.Border();
            LeftBorder leftBorder4 = new LeftBorder();
            RightBorder rightBorder4 = new RightBorder();
            TopBorder topBorder4 = new TopBorder();

            BottomBorder bottomBorder4 = new BottomBorder() { Style = BorderStyleValues.Thick };
            Color color29 = new Color() { Theme = 4U };

            bottomBorder4.Append(color29);
            DiagonalBorder diagonalBorder4 = new DiagonalBorder();

            border4.Append(leftBorder4);
            border4.Append(rightBorder4);
            border4.Append(topBorder4);
            border4.Append(bottomBorder4);
            border4.Append(diagonalBorder4);

            DocumentFormat.OpenXml.Spreadsheet.Border border5 = new DocumentFormat.OpenXml.Spreadsheet.Border();
            LeftBorder leftBorder5 = new LeftBorder();
            RightBorder rightBorder5 = new RightBorder();
            TopBorder topBorder5 = new TopBorder();

            BottomBorder bottomBorder5 = new BottomBorder() { Style = BorderStyleValues.Thick };
            Color color30 = new Color() { Theme = 4U, Tint = 0.499984740745262D };

            bottomBorder5.Append(color30);
            DiagonalBorder diagonalBorder5 = new DiagonalBorder();

            border5.Append(leftBorder5);
            border5.Append(rightBorder5);
            border5.Append(topBorder5);
            border5.Append(bottomBorder5);
            border5.Append(diagonalBorder5);

            DocumentFormat.OpenXml.Spreadsheet.Border border6 = new DocumentFormat.OpenXml.Spreadsheet.Border();
            LeftBorder leftBorder6 = new LeftBorder();
            RightBorder rightBorder6 = new RightBorder();
            TopBorder topBorder6 = new TopBorder();

            BottomBorder bottomBorder6 = new BottomBorder() { Style = BorderStyleValues.Medium };
            Color color31 = new Color() { Theme = 4U, Tint = 0.39997558519241921D };

            bottomBorder6.Append(color31);
            DiagonalBorder diagonalBorder6 = new DiagonalBorder();

            border6.Append(leftBorder6);
            border6.Append(rightBorder6);
            border6.Append(topBorder6);
            border6.Append(bottomBorder6);
            border6.Append(diagonalBorder6);

            DocumentFormat.OpenXml.Spreadsheet.Border border7 = new DocumentFormat.OpenXml.Spreadsheet.Border();
            LeftBorder leftBorder7 = new LeftBorder();
            RightBorder rightBorder7 = new RightBorder();
            TopBorder topBorder7 = new TopBorder();

            BottomBorder bottomBorder7 = new BottomBorder() { Style = BorderStyleValues.Double };
            Color color32 = new Color() { Rgb = "FFFF8001" };

            bottomBorder7.Append(color32);
            DiagonalBorder diagonalBorder7 = new DiagonalBorder();

            border7.Append(leftBorder7);
            border7.Append(rightBorder7);
            border7.Append(topBorder7);
            border7.Append(bottomBorder7);
            border7.Append(diagonalBorder7);

            DocumentFormat.OpenXml.Spreadsheet.Border border8 = new DocumentFormat.OpenXml.Spreadsheet.Border();

            LeftBorder leftBorder8 = new LeftBorder() { Style = BorderStyleValues.Thin };
            Color color33 = new Color() { Rgb = "FFB2B2B2" };

            leftBorder8.Append(color33);

            RightBorder rightBorder8 = new RightBorder() { Style = BorderStyleValues.Thin };
            Color color34 = new Color() { Rgb = "FFB2B2B2" };

            rightBorder8.Append(color34);

            TopBorder topBorder8 = new TopBorder() { Style = BorderStyleValues.Thin };
            Color color35 = new Color() { Rgb = "FFB2B2B2" };

            topBorder8.Append(color35);

            BottomBorder bottomBorder8 = new BottomBorder() { Style = BorderStyleValues.Thin };
            Color color36 = new Color() { Rgb = "FFB2B2B2" };

            bottomBorder8.Append(color36);
            DiagonalBorder diagonalBorder8 = new DiagonalBorder();

            border8.Append(leftBorder8);
            border8.Append(rightBorder8);
            border8.Append(topBorder8);
            border8.Append(bottomBorder8);
            border8.Append(diagonalBorder8);

            DocumentFormat.OpenXml.Spreadsheet.Border border9 = new DocumentFormat.OpenXml.Spreadsheet.Border();

            LeftBorder leftBorder9 = new LeftBorder() { Style = BorderStyleValues.Thin };
            Color color37 = new Color() { Rgb = "FF3F3F3F" };

            leftBorder9.Append(color37);

            RightBorder rightBorder9 = new RightBorder() { Style = BorderStyleValues.Thin };
            Color color38 = new Color() { Rgb = "FF3F3F3F" };

            rightBorder9.Append(color38);

            TopBorder topBorder9 = new TopBorder() { Style = BorderStyleValues.Thin };
            Color color39 = new Color() { Rgb = "FF3F3F3F" };

            topBorder9.Append(color39);

            BottomBorder bottomBorder9 = new BottomBorder() { Style = BorderStyleValues.Thin };
            Color color40 = new Color() { Rgb = "FF3F3F3F" };

            bottomBorder9.Append(color40);
            DiagonalBorder diagonalBorder9 = new DiagonalBorder();

            border9.Append(leftBorder9);
            border9.Append(rightBorder9);
            border9.Append(topBorder9);
            border9.Append(bottomBorder9);
            border9.Append(diagonalBorder9);

            DocumentFormat.OpenXml.Spreadsheet.Border border10 = new DocumentFormat.OpenXml.Spreadsheet.Border();
            LeftBorder leftBorder10 = new LeftBorder();
            RightBorder rightBorder10 = new RightBorder();

            TopBorder topBorder10 = new TopBorder() { Style = BorderStyleValues.Thin };
            Color color41 = new Color() { Theme = 4U };

            topBorder10.Append(color41);

            BottomBorder bottomBorder10 = new BottomBorder() { Style = BorderStyleValues.Double };
            Color color42 = new Color() { Theme = 4U };

            bottomBorder10.Append(color42);
            DiagonalBorder diagonalBorder10 = new DiagonalBorder();

            border10.Append(leftBorder10);
            border10.Append(rightBorder10);
            border10.Append(topBorder10);
            border10.Append(bottomBorder10);
            border10.Append(diagonalBorder10);

            borders1.Append(border1);
            borders1.Append(border2);
            borders1.Append(border3);
            borders1.Append(border4);
            borders1.Append(border5);
            borders1.Append(border6);
            borders1.Append(border7);
            borders1.Append(border8);
            borders1.Append(border9);
            borders1.Append(border10);

            CellStyleFormats cellStyleFormats1 = new CellStyleFormats() { Count = 42U };
            CellFormat cellFormat1 = new CellFormat() { NumberFormatId = 0U, FontId = 0U, FillId = 0U, BorderId = 0U };
            CellFormat cellFormat2 = new CellFormat() { NumberFormatId = 0U, FontId = 3U, FillId = 3U, BorderId = 0U, ApplyNumberFormat = false, ApplyBorder = false, ApplyAlignment = false, ApplyProtection = false };
            CellFormat cellFormat3 = new CellFormat() { NumberFormatId = 0U, FontId = 3U, FillId = 4U, BorderId = 0U, ApplyNumberFormat = false, ApplyBorder = false, ApplyAlignment = false, ApplyProtection = false };
            CellFormat cellFormat4 = new CellFormat() { NumberFormatId = 0U, FontId = 3U, FillId = 5U, BorderId = 0U, ApplyNumberFormat = false, ApplyBorder = false, ApplyAlignment = false, ApplyProtection = false };
            CellFormat cellFormat5 = new CellFormat() { NumberFormatId = 0U, FontId = 3U, FillId = 6U, BorderId = 0U, ApplyNumberFormat = false, ApplyBorder = false, ApplyAlignment = false, ApplyProtection = false };
            CellFormat cellFormat6 = new CellFormat() { NumberFormatId = 0U, FontId = 3U, FillId = 7U, BorderId = 0U, ApplyNumberFormat = false, ApplyBorder = false, ApplyAlignment = false, ApplyProtection = false };
            CellFormat cellFormat7 = new CellFormat() { NumberFormatId = 0U, FontId = 3U, FillId = 8U, BorderId = 0U, ApplyNumberFormat = false, ApplyBorder = false, ApplyAlignment = false, ApplyProtection = false };
            CellFormat cellFormat8 = new CellFormat() { NumberFormatId = 0U, FontId = 3U, FillId = 9U, BorderId = 0U, ApplyNumberFormat = false, ApplyBorder = false, ApplyAlignment = false, ApplyProtection = false };
            CellFormat cellFormat9 = new CellFormat() { NumberFormatId = 0U, FontId = 3U, FillId = 10U, BorderId = 0U, ApplyNumberFormat = false, ApplyBorder = false, ApplyAlignment = false, ApplyProtection = false };
            CellFormat cellFormat10 = new CellFormat() { NumberFormatId = 0U, FontId = 3U, FillId = 11U, BorderId = 0U, ApplyNumberFormat = false, ApplyBorder = false, ApplyAlignment = false, ApplyProtection = false };
            CellFormat cellFormat11 = new CellFormat() { NumberFormatId = 0U, FontId = 3U, FillId = 12U, BorderId = 0U, ApplyNumberFormat = false, ApplyBorder = false, ApplyAlignment = false, ApplyProtection = false };
            CellFormat cellFormat12 = new CellFormat() { NumberFormatId = 0U, FontId = 3U, FillId = 13U, BorderId = 0U, ApplyNumberFormat = false, ApplyBorder = false, ApplyAlignment = false, ApplyProtection = false };
            CellFormat cellFormat13 = new CellFormat() { NumberFormatId = 0U, FontId = 3U, FillId = 14U, BorderId = 0U, ApplyNumberFormat = false, ApplyBorder = false, ApplyAlignment = false, ApplyProtection = false };
            CellFormat cellFormat14 = new CellFormat() { NumberFormatId = 0U, FontId = 4U, FillId = 15U, BorderId = 0U, ApplyNumberFormat = false, ApplyBorder = false, ApplyAlignment = false, ApplyProtection = false };
            CellFormat cellFormat15 = new CellFormat() { NumberFormatId = 0U, FontId = 4U, FillId = 16U, BorderId = 0U, ApplyNumberFormat = false, ApplyBorder = false, ApplyAlignment = false, ApplyProtection = false };
            CellFormat cellFormat16 = new CellFormat() { NumberFormatId = 0U, FontId = 4U, FillId = 17U, BorderId = 0U, ApplyNumberFormat = false, ApplyBorder = false, ApplyAlignment = false, ApplyProtection = false };
            CellFormat cellFormat17 = new CellFormat() { NumberFormatId = 0U, FontId = 4U, FillId = 18U, BorderId = 0U, ApplyNumberFormat = false, ApplyBorder = false, ApplyAlignment = false, ApplyProtection = false };
            CellFormat cellFormat18 = new CellFormat() { NumberFormatId = 0U, FontId = 4U, FillId = 19U, BorderId = 0U, ApplyNumberFormat = false, ApplyBorder = false, ApplyAlignment = false, ApplyProtection = false };
            CellFormat cellFormat19 = new CellFormat() { NumberFormatId = 0U, FontId = 4U, FillId = 20U, BorderId = 0U, ApplyNumberFormat = false, ApplyBorder = false, ApplyAlignment = false, ApplyProtection = false };
            CellFormat cellFormat20 = new CellFormat() { NumberFormatId = 0U, FontId = 4U, FillId = 21U, BorderId = 0U, ApplyNumberFormat = false, ApplyBorder = false, ApplyAlignment = false, ApplyProtection = false };
            CellFormat cellFormat21 = new CellFormat() { NumberFormatId = 0U, FontId = 4U, FillId = 22U, BorderId = 0U, ApplyNumberFormat = false, ApplyBorder = false, ApplyAlignment = false, ApplyProtection = false };
            CellFormat cellFormat22 = new CellFormat() { NumberFormatId = 0U, FontId = 4U, FillId = 23U, BorderId = 0U, ApplyNumberFormat = false, ApplyBorder = false, ApplyAlignment = false, ApplyProtection = false };
            CellFormat cellFormat23 = new CellFormat() { NumberFormatId = 0U, FontId = 4U, FillId = 24U, BorderId = 0U, ApplyNumberFormat = false, ApplyBorder = false, ApplyAlignment = false, ApplyProtection = false };
            CellFormat cellFormat24 = new CellFormat() { NumberFormatId = 0U, FontId = 4U, FillId = 25U, BorderId = 0U, ApplyNumberFormat = false, ApplyBorder = false, ApplyAlignment = false, ApplyProtection = false };
            CellFormat cellFormat25 = new CellFormat() { NumberFormatId = 0U, FontId = 4U, FillId = 26U, BorderId = 0U, ApplyNumberFormat = false, ApplyBorder = false, ApplyAlignment = false, ApplyProtection = false };
            CellFormat cellFormat26 = new CellFormat() { NumberFormatId = 0U, FontId = 5U, FillId = 27U, BorderId = 0U, ApplyNumberFormat = false, ApplyBorder = false, ApplyAlignment = false, ApplyProtection = false };
            CellFormat cellFormat27 = new CellFormat() { NumberFormatId = 0U, FontId = 6U, FillId = 28U, BorderId = 1U, ApplyNumberFormat = false, ApplyAlignment = false, ApplyProtection = false };
            CellFormat cellFormat28 = new CellFormat() { NumberFormatId = 0U, FontId = 7U, FillId = 29U, BorderId = 2U, ApplyNumberFormat = false, ApplyAlignment = false, ApplyProtection = false };
            CellFormat cellFormat29 = new CellFormat() { NumberFormatId = 0U, FontId = 8U, FillId = 0U, BorderId = 0U, ApplyNumberFormat = false, ApplyFill = false, ApplyBorder = false, ApplyAlignment = false, ApplyProtection = false };
            CellFormat cellFormat30 = new CellFormat() { NumberFormatId = 0U, FontId = 9U, FillId = 30U, BorderId = 0U, ApplyNumberFormat = false, ApplyBorder = false, ApplyAlignment = false, ApplyProtection = false };
            CellFormat cellFormat31 = new CellFormat() { NumberFormatId = 0U, FontId = 10U, FillId = 0U, BorderId = 3U, ApplyNumberFormat = false, ApplyFill = false, ApplyAlignment = false, ApplyProtection = false };
            CellFormat cellFormat32 = new CellFormat() { NumberFormatId = 0U, FontId = 11U, FillId = 0U, BorderId = 4U, ApplyNumberFormat = false, ApplyFill = false, ApplyAlignment = false, ApplyProtection = false };
            CellFormat cellFormat33 = new CellFormat() { NumberFormatId = 0U, FontId = 12U, FillId = 0U, BorderId = 5U, ApplyNumberFormat = false, ApplyFill = false, ApplyAlignment = false, ApplyProtection = false };
            CellFormat cellFormat34 = new CellFormat() { NumberFormatId = 0U, FontId = 12U, FillId = 0U, BorderId = 0U, ApplyNumberFormat = false, ApplyFill = false, ApplyBorder = false, ApplyAlignment = false, ApplyProtection = false };
            CellFormat cellFormat35 = new CellFormat() { NumberFormatId = 0U, FontId = 13U, FillId = 31U, BorderId = 1U, ApplyNumberFormat = false, ApplyAlignment = false, ApplyProtection = false };
            CellFormat cellFormat36 = new CellFormat() { NumberFormatId = 0U, FontId = 14U, FillId = 0U, BorderId = 6U, ApplyNumberFormat = false, ApplyFill = false, ApplyAlignment = false, ApplyProtection = false };
            CellFormat cellFormat37 = new CellFormat() { NumberFormatId = 0U, FontId = 15U, FillId = 32U, BorderId = 0U, ApplyNumberFormat = false, ApplyBorder = false, ApplyAlignment = false, ApplyProtection = false };
            CellFormat cellFormat38 = new CellFormat() { NumberFormatId = 0U, FontId = 1U, FillId = 33U, BorderId = 7U, ApplyNumberFormat = false, ApplyFont = false, ApplyAlignment = false, ApplyProtection = false };
            CellFormat cellFormat39 = new CellFormat() { NumberFormatId = 0U, FontId = 16U, FillId = 28U, BorderId = 8U, ApplyNumberFormat = false, ApplyAlignment = false, ApplyProtection = false };
            CellFormat cellFormat40 = new CellFormat() { NumberFormatId = 0U, FontId = 17U, FillId = 0U, BorderId = 0U, ApplyNumberFormat = false, ApplyFill = false, ApplyBorder = false, ApplyAlignment = false, ApplyProtection = false };
            CellFormat cellFormat41 = new CellFormat() { NumberFormatId = 0U, FontId = 18U, FillId = 0U, BorderId = 9U, ApplyNumberFormat = false, ApplyFill = false, ApplyAlignment = false, ApplyProtection = false };
            CellFormat cellFormat42 = new CellFormat() { NumberFormatId = 0U, FontId = 19U, FillId = 0U, BorderId = 0U, ApplyNumberFormat = false, ApplyFill = false, ApplyBorder = false, ApplyAlignment = false, ApplyProtection = false };

            cellStyleFormats1.Append(cellFormat1);
            cellStyleFormats1.Append(cellFormat2);
            cellStyleFormats1.Append(cellFormat3);
            cellStyleFormats1.Append(cellFormat4);
            cellStyleFormats1.Append(cellFormat5);
            cellStyleFormats1.Append(cellFormat6);
            cellStyleFormats1.Append(cellFormat7);
            cellStyleFormats1.Append(cellFormat8);
            cellStyleFormats1.Append(cellFormat9);
            cellStyleFormats1.Append(cellFormat10);
            cellStyleFormats1.Append(cellFormat11);
            cellStyleFormats1.Append(cellFormat12);
            cellStyleFormats1.Append(cellFormat13);
            cellStyleFormats1.Append(cellFormat14);
            cellStyleFormats1.Append(cellFormat15);
            cellStyleFormats1.Append(cellFormat16);
            cellStyleFormats1.Append(cellFormat17);
            cellStyleFormats1.Append(cellFormat18);
            cellStyleFormats1.Append(cellFormat19);
            cellStyleFormats1.Append(cellFormat20);
            cellStyleFormats1.Append(cellFormat21);
            cellStyleFormats1.Append(cellFormat22);
            cellStyleFormats1.Append(cellFormat23);
            cellStyleFormats1.Append(cellFormat24);
            cellStyleFormats1.Append(cellFormat25);
            cellStyleFormats1.Append(cellFormat26);
            cellStyleFormats1.Append(cellFormat27);
            cellStyleFormats1.Append(cellFormat28);
            cellStyleFormats1.Append(cellFormat29);
            cellStyleFormats1.Append(cellFormat30);
            cellStyleFormats1.Append(cellFormat31);
            cellStyleFormats1.Append(cellFormat32);
            cellStyleFormats1.Append(cellFormat33);
            cellStyleFormats1.Append(cellFormat34);
            cellStyleFormats1.Append(cellFormat35);
            cellStyleFormats1.Append(cellFormat36);
            cellStyleFormats1.Append(cellFormat37);
            cellStyleFormats1.Append(cellFormat38);
            cellStyleFormats1.Append(cellFormat39);
            cellStyleFormats1.Append(cellFormat40);
            cellStyleFormats1.Append(cellFormat41);
            cellStyleFormats1.Append(cellFormat42);

            var cellFormats1 = new CellFormats { Count = 3U };
            var cellFormat43 = new CellFormat { NumberFormatId = 0U, FontId = 0U, FillId = 0U, BorderId = 0U, FormatId = 0U };
            var cellFormat44 = new CellFormat { NumberFormatId = 0U, FontId = 0U, FillId = 2U, BorderId = 0U, FormatId = 0U, ApplyFill = true };
            var cellFormat45 = new CellFormat { NumberFormatId = 0U, FontId = 2U, FillId = 0U, BorderId = 0U, FormatId = 0U, ApplyFont = true };

            cellFormats1.Append(cellFormat43);
            cellFormats1.Append(cellFormat44);
            cellFormats1.Append(cellFormat45);

            var cellStyles1 = new ss.CellStyles { Count = 42U };
            var cellStyle1 = new ss.CellStyle { Name = "20% - Accent1", FormatId = 1U, BuiltinId = 30U, CustomBuiltin = true };
            var cellStyle2 = new ss.CellStyle { Name = "20% - Accent2", FormatId = 2U, BuiltinId = 34U, CustomBuiltin = true };
            var cellStyle3 = new ss.CellStyle { Name = "20% - Accent3", FormatId = 3U, BuiltinId = 38U, CustomBuiltin = true };
            var cellStyle4 = new ss.CellStyle { Name = "20% - Accent4", FormatId = 4U, BuiltinId = 42U, CustomBuiltin = true };
            var cellStyle5 = new ss.CellStyle { Name = "20% - Accent5", FormatId = 5U, BuiltinId = 46U, CustomBuiltin = true };
            var cellStyle6 = new ss.CellStyle { Name = "20% - Accent6", FormatId = 6U, BuiltinId = 50U, CustomBuiltin = true };
            var cellStyle7 = new ss.CellStyle { Name = "40% - Accent1", FormatId = 7U, BuiltinId = 31U, CustomBuiltin = true };
            var cellStyle8 = new ss.CellStyle { Name = "40% - Accent2", FormatId = 8U, BuiltinId = 35U, CustomBuiltin = true };
            var cellStyle9 = new ss.CellStyle { Name = "40% - Accent3", FormatId = 9U, BuiltinId = 39U, CustomBuiltin = true };
            var cellStyle10 = new ss.CellStyle { Name = "40% - Accent4", FormatId = 10U, BuiltinId = 43U, CustomBuiltin = true };
            var cellStyle11 = new ss.CellStyle { Name = "40% - Accent5", FormatId = 11U, BuiltinId = 47U, CustomBuiltin = true };
            var cellStyle12 = new ss.CellStyle { Name = "40% - Accent6", FormatId = 12U, BuiltinId = 51U, CustomBuiltin = true };
            var cellStyle13 = new ss.CellStyle { Name = "60% - Accent1", FormatId = 13U, BuiltinId = 32U, CustomBuiltin = true };
            var cellStyle14 = new ss.CellStyle { Name = "60% - Accent2", FormatId = 14U, BuiltinId = 36U, CustomBuiltin = true };
            var cellStyle15 = new ss.CellStyle { Name = "60% - Accent3", FormatId = 15U, BuiltinId = 40U, CustomBuiltin = true };
            var cellStyle16 = new ss.CellStyle { Name = "60% - Accent4", FormatId = 16U, BuiltinId = 44U, CustomBuiltin = true };
            var cellStyle17 = new ss.CellStyle { Name = "60% - Accent5", FormatId = 17U, BuiltinId = 48U, CustomBuiltin = true };
            var cellStyle18 = new ss.CellStyle { Name = "60% - Accent6", FormatId = 18U, BuiltinId = 52U, CustomBuiltin = true };
            var cellStyle19 = new ss.CellStyle { Name = "Accent1", FormatId = 19U, BuiltinId = 29U, CustomBuiltin = true };
            var cellStyle20 = new ss.CellStyle { Name = "Accent2", FormatId = 20U, BuiltinId = 33U, CustomBuiltin = true };
            var cellStyle21 = new ss.CellStyle { Name = "Accent3", FormatId = 21U, BuiltinId = 37U, CustomBuiltin = true };
            var cellStyle22 = new ss.CellStyle { Name = "Accent4", FormatId = 22U, BuiltinId = 41U, CustomBuiltin = true };
            var cellStyle23 = new ss.CellStyle { Name = "Accent5", FormatId = 23U, BuiltinId = 45U, CustomBuiltin = true };
            var cellStyle24 = new ss.CellStyle { Name = "Accent6", FormatId = 24U, BuiltinId = 49U, CustomBuiltin = true };
            var cellStyle25 = new ss.CellStyle { Name = "Bad", FormatId = 25U, BuiltinId = 27U, CustomBuiltin = true };
            var cellStyle26 = new ss.CellStyle { Name = "Calculation", FormatId = 26U, BuiltinId = 22U, CustomBuiltin = true };
            var cellStyle27 = new ss.CellStyle { Name = "Check Cell", FormatId = 27U, BuiltinId = 23U, CustomBuiltin = true };
            var cellStyle28 = new ss.CellStyle { Name = "Explanatory Text", FormatId = 28U, BuiltinId = 53U, CustomBuiltin = true };
            var cellStyle29 = new ss.CellStyle { Name = "Good", FormatId = 29U, BuiltinId = 26U, CustomBuiltin = true };
            var cellStyle30 = new ss.CellStyle { Name = "Heading 1", FormatId = 30U, BuiltinId = 16U, CustomBuiltin = true };
            var cellStyle31 = new ss.CellStyle { Name = "Heading 2", FormatId = 31U, BuiltinId = 17U, CustomBuiltin = true };
            var cellStyle32 = new ss.CellStyle { Name = "Heading 3", FormatId = 32U, BuiltinId = 18U, CustomBuiltin = true };
            var cellStyle33 = new ss.CellStyle { Name = "Heading 4", FormatId = 33U, BuiltinId = 19U, CustomBuiltin = true };
            var cellStyle34 = new ss.CellStyle { Name = "Input", FormatId = 34U, BuiltinId = 20U, CustomBuiltin = true };
            var cellStyle35 = new ss.CellStyle { Name = "Linked Cell", FormatId = 35U, BuiltinId = 24U, CustomBuiltin = true };
            var cellStyle36 = new ss.CellStyle { Name = "Neutral", FormatId = 36U, BuiltinId = 28U, CustomBuiltin = true };
            var cellStyle37 = new ss.CellStyle { Name = "Normal", FormatId = 0U, BuiltinId = 0U };
            var cellStyle38 = new ss.CellStyle { Name = "Note", FormatId = 37U, BuiltinId = 10U, CustomBuiltin = true };
            var cellStyle39 = new ss.CellStyle { Name = "Output", FormatId = 38U, BuiltinId = 21U, CustomBuiltin = true };
            var cellStyle40 = new ss.CellStyle { Name = "Title", FormatId = 39U, BuiltinId = 15U, CustomBuiltin = true };
            var cellStyle41 = new ss.CellStyle { Name = "Total", FormatId = 40U, BuiltinId = 25U, CustomBuiltin = true };
            var cellStyle42 = new ss.CellStyle { Name = "Warning Text", FormatId = 41U, BuiltinId = 11U, CustomBuiltin = true };

            cellStyles1.Append(cellStyle1);
            cellStyles1.Append(cellStyle2);
            cellStyles1.Append(cellStyle3);
            cellStyles1.Append(cellStyle4);
            cellStyles1.Append(cellStyle5);
            cellStyles1.Append(cellStyle6);
            cellStyles1.Append(cellStyle7);
            cellStyles1.Append(cellStyle8);
            cellStyles1.Append(cellStyle9);
            cellStyles1.Append(cellStyle10);
            cellStyles1.Append(cellStyle11);
            cellStyles1.Append(cellStyle12);
            cellStyles1.Append(cellStyle13);
            cellStyles1.Append(cellStyle14);
            cellStyles1.Append(cellStyle15);
            cellStyles1.Append(cellStyle16);
            cellStyles1.Append(cellStyle17);
            cellStyles1.Append(cellStyle18);
            cellStyles1.Append(cellStyle19);
            cellStyles1.Append(cellStyle20);
            cellStyles1.Append(cellStyle21);
            cellStyles1.Append(cellStyle22);
            cellStyles1.Append(cellStyle23);
            cellStyles1.Append(cellStyle24);
            cellStyles1.Append(cellStyle25);
            cellStyles1.Append(cellStyle26);
            cellStyles1.Append(cellStyle27);
            cellStyles1.Append(cellStyle28);
            cellStyles1.Append(cellStyle29);
            cellStyles1.Append(cellStyle30);
            cellStyles1.Append(cellStyle31);
            cellStyles1.Append(cellStyle32);
            cellStyles1.Append(cellStyle33);
            cellStyles1.Append(cellStyle34);
            cellStyles1.Append(cellStyle35);
            cellStyles1.Append(cellStyle36);
            cellStyles1.Append(cellStyle37);
            cellStyles1.Append(cellStyle38);
            cellStyles1.Append(cellStyle39);
            cellStyles1.Append(cellStyle40);
            cellStyles1.Append(cellStyle41);
            cellStyles1.Append(cellStyle42);
            DifferentialFormats differentialFormats1 = new DifferentialFormats() { Count = 0U };
            TableStyles tableStyles1 = new TableStyles() { Count = 0U, DefaultTableStyle = "TableStyleMedium9", DefaultPivotStyle = "PivotStyleLight16" };

            // These are the default colours in the Excel 2003 colour pallette. These are here for backwards compatability.
            Colors colors1 = new Colors();
            IndexedColors indexedColors1 = new IndexedColors();
            // Default Excel 2003 Colours
            //RgbColor rgbColor1 = new RgbColor() { Rgb = "00000000" };
            //RgbColor rgbColor2 = new RgbColor() { Rgb = "00FFFFFF" };
            //RgbColor rgbColor3 = new RgbColor() { Rgb = "00FF0000" };
            //RgbColor rgbColor4 = new RgbColor() { Rgb = "0000FF00" };
            //RgbColor rgbColor5 = new RgbColor() { Rgb = "000000FF" };
            //RgbColor rgbColor6 = new RgbColor() { Rgb = "00FFFF00" };
            //RgbColor rgbColor7 = new RgbColor() { Rgb = "00FF00FF" };
            //RgbColor rgbColor8 = new RgbColor() { Rgb = "0000FFFF" };
            //RgbColor rgbColor9 = new RgbColor() { Rgb = "00000000" };
            //RgbColor rgbColor10 = new RgbColor() { Rgb = "00FFFFFF" };
            //RgbColor rgbColor11 = new RgbColor() { Rgb = "00FF0000" };
            //RgbColor rgbColor12 = new RgbColor() { Rgb = "0000FF00" };
            //RgbColor rgbColor13 = new RgbColor() { Rgb = "000000FF" };
            //RgbColor rgbColor14 = new RgbColor() { Rgb = "00FFFF00" };
            //RgbColor rgbColor15 = new RgbColor() { Rgb = "00FF00FF" };
            //RgbColor rgbColor16 = new RgbColor() { Rgb = "0000FFFF" };
            //RgbColor rgbColor17 = new RgbColor() { Rgb = "00800000" };
            //RgbColor rgbColor18 = new RgbColor() { Rgb = "00008000" };
            //RgbColor rgbColor19 = new RgbColor() { Rgb = "00000080" };
            //RgbColor rgbColor20 = new RgbColor() { Rgb = "00808000" };
            //RgbColor rgbColor21 = new RgbColor() { Rgb = "00800080" };
            //RgbColor rgbColor22 = new RgbColor() { Rgb = "00008080" };
            //RgbColor rgbColor23 = new RgbColor() { Rgb = "00C0C0C0" };
            //RgbColor rgbColor24 = new RgbColor() { Rgb = "00808080" };
            //RgbColor rgbColor25 = new RgbColor() { Rgb = "009999FF" };
            //RgbColor rgbColor26 = new RgbColor() { Rgb = "00993366" };
            //RgbColor rgbColor27 = new RgbColor() { Rgb = "00FFFFCC" };
            //RgbColor rgbColor28 = new RgbColor() { Rgb = "00CCFFFF" };
            //RgbColor rgbColor29 = new RgbColor() { Rgb = "00660066" };
            //RgbColor rgbColor30 = new RgbColor() { Rgb = "00FF8080" };
            //RgbColor rgbColor31 = new RgbColor() { Rgb = "000066CC" };
            //RgbColor rgbColor32 = new RgbColor() { Rgb = "00CCCCFF" };
            //RgbColor rgbColor33 = new RgbColor() { Rgb = "00000080" };
            //RgbColor rgbColor34 = new RgbColor() { Rgb = "00FF00FF" };
            //RgbColor rgbColor35 = new RgbColor() { Rgb = "00FFFF00" };
            //RgbColor rgbColor36 = new RgbColor() { Rgb = "0000FFFF" };
            //RgbColor rgbColor37 = new RgbColor() { Rgb = "00800080" };
            //RgbColor rgbColor38 = new RgbColor() { Rgb = "00800000" };
            //RgbColor rgbColor39 = new RgbColor() { Rgb = "00008080" };
            //RgbColor rgbColor40 = new RgbColor() { Rgb = "000000FF" };
            //RgbColor rgbColor41 = new RgbColor() { Rgb = "0000CCFF" };
            //RgbColor rgbColor42 = new RgbColor() { Rgb = "00CCFFFF" };
            //RgbColor rgbColor43 = new RgbColor() { Rgb = "00CCFFCC" };
            //RgbColor rgbColor44 = new RgbColor() { Rgb = "00FFFF99" };
            //RgbColor rgbColor45 = new RgbColor() { Rgb = "0099CCFF" };
            //RgbColor rgbColor46 = new RgbColor() { Rgb = "00FF99CC" };
            //RgbColor rgbColor47 = new RgbColor() { Rgb = "00CC99FF" };
            //RgbColor rgbColor48 = new RgbColor() { Rgb = "00FFCC99" };
            //RgbColor rgbColor49 = new RgbColor() { Rgb = "003366FF" };
            //RgbColor rgbColor50 = new RgbColor() { Rgb = "0033CCCC" };
            //RgbColor rgbColor51 = new RgbColor() { Rgb = "0099CC00" };
            //RgbColor rgbColor52 = new RgbColor() { Rgb = "00FFCC00" };
            //RgbColor rgbColor53 = new RgbColor() { Rgb = "00FF9900" };
            //RgbColor rgbColor54 = new RgbColor() { Rgb = "00FF6600" };
            //RgbColor rgbColor55 = new RgbColor() { Rgb = "00666699" };
            //RgbColor rgbColor56 = new RgbColor() { Rgb = "00969696" };
            //RgbColor rgbColor57 = new RgbColor() { Rgb = "00003366" };
            //RgbColor rgbColor58 = new RgbColor() { Rgb = "00339966" };
            //RgbColor rgbColor59 = new RgbColor() { Rgb = "00003300" };
            //RgbColor rgbColor60 = new RgbColor() { Rgb = "00333300" };
            //RgbColor rgbColor61 = new RgbColor() { Rgb = "00993300" };
            //RgbColor rgbColor62 = new RgbColor() { Rgb = "00993366" };
            //RgbColor rgbColor63 = new RgbColor() { Rgb = "00333399" };
            //RgbColor rgbColor64 = new RgbColor() { Rgb = "00333333" };

            // GAM Excel 2003 Colours
            RgbColor rgbColor1 = new RgbColor() { Rgb = "00000000" };
            RgbColor rgbColor2 = new RgbColor() { Rgb = "00FFFFFF" };
            RgbColor rgbColor3 = new RgbColor() { Rgb = "00FF0000" };
            RgbColor rgbColor4 = new RgbColor() { Rgb = "0000FF00" };
            RgbColor rgbColor5 = new RgbColor() { Rgb = "000000FF" };
            RgbColor rgbColor6 = new RgbColor() { Rgb = "00FFFF00" };
            RgbColor rgbColor7 = new RgbColor() { Rgb = "00FF00FF" };
            RgbColor rgbColor8 = new RgbColor() { Rgb = "0000FFFF" };
            RgbColor rgbColor9 = new RgbColor() { Rgb = "00FFFFFF" };
            RgbColor rgbColor10 = new RgbColor() { Rgb = "00000000" };
            RgbColor rgbColor11 = new RgbColor() { Rgb = "00006666" };
            RgbColor rgbColor12 = new RgbColor() { Rgb = "009EBEE8" };
            RgbColor rgbColor13 = new RgbColor() { Rgb = "00BBB6B4" };
            RgbColor rgbColor14 = new RgbColor() { Rgb = "00D8BC8F" };
            RgbColor rgbColor15 = new RgbColor() { Rgb = "00AAB490" };
            RgbColor rgbColor16 = new RgbColor() { Rgb = "00D797C3" };
            RgbColor rgbColor17 = new RgbColor() { Rgb = "00E3A197" };
            RgbColor rgbColor18 = new RgbColor() { Rgb = "00D7CD91" };
            RgbColor rgbColor19 = new RgbColor() { Rgb = "00AF9BBF" };
            RgbColor rgbColor20 = new RgbColor() { Rgb = "0066CAB9" };
            RgbColor rgbColor21 = new RgbColor() { Rgb = "00006666" };
            RgbColor rgbColor22 = new RgbColor() { Rgb = "0066A3A3" };
            RgbColor rgbColor23 = new RgbColor() { Rgb = "00B2D1D1" };
            RgbColor rgbColor24 = new RgbColor() { Rgb = "00F2F7F7" };
            RgbColor rgbColor25 = new RgbColor() { Rgb = "009EBEE8" };
            RgbColor rgbColor26 = new RgbColor() { Rgb = "00BBB6B4" };
            RgbColor rgbColor27 = new RgbColor() { Rgb = "00D8BC8F" };
            RgbColor rgbColor28 = new RgbColor() { Rgb = "00AAB490" };
            RgbColor rgbColor29 = new RgbColor() { Rgb = "00D797C3" };
            RgbColor rgbColor30 = new RgbColor() { Rgb = "00E3A197" };
            RgbColor rgbColor31 = new RgbColor() { Rgb = "00D7CD91" };
            RgbColor rgbColor32 = new RgbColor() { Rgb = "00AF9BBF" };
            RgbColor rgbColor33 = new RgbColor() { Rgb = "0066CAB9" };
            RgbColor rgbColor34 = new RgbColor() { Rgb = "00006666" };
            RgbColor rgbColor35 = new RgbColor() { Rgb = "0066A3A3" };
            RgbColor rgbColor36 = new RgbColor() { Rgb = "00B2D1D1" };
            RgbColor rgbColor37 = new RgbColor() { Rgb = "00F2F7F7" };
            RgbColor rgbColor38 = new RgbColor() { Rgb = "00800000" };
            RgbColor rgbColor39 = new RgbColor() { Rgb = "00008080" };
            RgbColor rgbColor40 = new RgbColor() { Rgb = "000000FF" };
            RgbColor rgbColor41 = new RgbColor() { Rgb = "0000CCFF" };
            RgbColor rgbColor42 = new RgbColor() { Rgb = "00CCFFFF" };
            RgbColor rgbColor43 = new RgbColor() { Rgb = "00CCFFCC" };
            RgbColor rgbColor44 = new RgbColor() { Rgb = "00FFFF99" };
            RgbColor rgbColor45 = new RgbColor() { Rgb = "0099CCFF" };
            RgbColor rgbColor46 = new RgbColor() { Rgb = "00FF99CC" };
            RgbColor rgbColor47 = new RgbColor() { Rgb = "00CC99FF" };
            RgbColor rgbColor48 = new RgbColor() { Rgb = "00FFCC99" };
            RgbColor rgbColor49 = new RgbColor() { Rgb = "003366FF" };
            RgbColor rgbColor50 = new RgbColor() { Rgb = "0033CCCC" };
            RgbColor rgbColor51 = new RgbColor() { Rgb = "0099CC00" };
            RgbColor rgbColor52 = new RgbColor() { Rgb = "00FFCC00" };
            RgbColor rgbColor53 = new RgbColor() { Rgb = "00FF9900" };
            RgbColor rgbColor54 = new RgbColor() { Rgb = "00FF6600" };
            RgbColor rgbColor55 = new RgbColor() { Rgb = "00666699" };
            RgbColor rgbColor56 = new RgbColor() { Rgb = "00969696" };
            RgbColor rgbColor57 = new RgbColor() { Rgb = "00003366" };
            RgbColor rgbColor58 = new RgbColor() { Rgb = "00339966" };
            RgbColor rgbColor59 = new RgbColor() { Rgb = "00003300" };
            RgbColor rgbColor60 = new RgbColor() { Rgb = "00333300" };
            RgbColor rgbColor61 = new RgbColor() { Rgb = "00993300" };
            RgbColor rgbColor62 = new RgbColor() { Rgb = "00993366" };
            RgbColor rgbColor63 = new RgbColor() { Rgb = "00333399" };
            RgbColor rgbColor64 = new RgbColor() { Rgb = "00333333" };
            indexedColors1.Append(rgbColor1);
            indexedColors1.Append(rgbColor2);
            indexedColors1.Append(rgbColor3);
            indexedColors1.Append(rgbColor4);
            indexedColors1.Append(rgbColor5);
            indexedColors1.Append(rgbColor6);
            indexedColors1.Append(rgbColor7);
            indexedColors1.Append(rgbColor8);
            indexedColors1.Append(rgbColor9);
            indexedColors1.Append(rgbColor10);
            indexedColors1.Append(rgbColor11);
            indexedColors1.Append(rgbColor12);
            indexedColors1.Append(rgbColor13);
            indexedColors1.Append(rgbColor14);
            indexedColors1.Append(rgbColor15);
            indexedColors1.Append(rgbColor16);
            indexedColors1.Append(rgbColor17);
            indexedColors1.Append(rgbColor18);
            indexedColors1.Append(rgbColor19);
            indexedColors1.Append(rgbColor20);
            indexedColors1.Append(rgbColor21);
            indexedColors1.Append(rgbColor22);
            indexedColors1.Append(rgbColor23);
            indexedColors1.Append(rgbColor24);
            indexedColors1.Append(rgbColor25);
            indexedColors1.Append(rgbColor26);
            indexedColors1.Append(rgbColor27);
            indexedColors1.Append(rgbColor28);
            indexedColors1.Append(rgbColor29);
            indexedColors1.Append(rgbColor30);
            indexedColors1.Append(rgbColor31);
            indexedColors1.Append(rgbColor32);
            indexedColors1.Append(rgbColor33);
            indexedColors1.Append(rgbColor34);
            indexedColors1.Append(rgbColor35);
            indexedColors1.Append(rgbColor36);
            indexedColors1.Append(rgbColor37);
            indexedColors1.Append(rgbColor38);
            indexedColors1.Append(rgbColor39);
            indexedColors1.Append(rgbColor40);
            indexedColors1.Append(rgbColor41);
            indexedColors1.Append(rgbColor42);
            indexedColors1.Append(rgbColor43);
            indexedColors1.Append(rgbColor44);
            indexedColors1.Append(rgbColor45);
            indexedColors1.Append(rgbColor46);
            indexedColors1.Append(rgbColor47);
            indexedColors1.Append(rgbColor48);
            indexedColors1.Append(rgbColor49);
            indexedColors1.Append(rgbColor50);
            indexedColors1.Append(rgbColor51);
            indexedColors1.Append(rgbColor52);
            indexedColors1.Append(rgbColor53);
            indexedColors1.Append(rgbColor54);
            indexedColors1.Append(rgbColor55);
            indexedColors1.Append(rgbColor56);
            indexedColors1.Append(rgbColor57);
            indexedColors1.Append(rgbColor58);
            indexedColors1.Append(rgbColor59);
            indexedColors1.Append(rgbColor60);
            indexedColors1.Append(rgbColor61);
            indexedColors1.Append(rgbColor62);
            indexedColors1.Append(rgbColor63);
            indexedColors1.Append(rgbColor64);
            colors1.Append(indexedColors1);

            stylesheet.NumberingFormats = new NumberingFormats();
            stylesheet.Append(fonts1);
            stylesheet.Append(fills1);
            stylesheet.Append(borders1);
            stylesheet.Append(cellStyleFormats1);
            stylesheet.Append(cellFormats1);
            stylesheet.Append(cellStyles1);
            stylesheet.Append(differentialFormats1);
            stylesheet.Append(tableStyles1);
            stylesheet.Append(colors1);
        }

        #endregion
    }
}
