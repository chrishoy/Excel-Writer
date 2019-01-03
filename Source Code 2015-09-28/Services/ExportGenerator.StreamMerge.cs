namespace ExcelWriter
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using DocumentFormat.OpenXml;
    using DocumentFormat.OpenXml.Packaging;
    using DocumentFormat.OpenXml.Spreadsheet;
    using OpenXml.Excel;
    using Spreadsheet = DocumentFormat.OpenXml.Spreadsheet;

    public sealed partial class ExportGenerator
    {
        private enum InsertAppend
        {
            Insert,
            Append
        }

        /// <summary>
        /// Appends the sheet to excel stream.
        /// </summary>
        /// <param name="inputStream">The input stream.</param>
        /// <param name="workbookPath">The workbook path.</param>
        /// <param name="worksheetName">Name of the worksheet.</param>
        public static void InsertSheetToExcelStream(MemoryStream inputStream, string workbookPath, string worksheetName)
        {
            if (inputStream == null)
            {
                throw new ArgumentNullException("inputStream");
            }

            if (string.IsNullOrEmpty(workbookPath))
            {
                throw new ArgumentNullException("workbookPath");
            }

            if (string.IsNullOrEmpty(worksheetName))
            {
                throw new ArgumentNullException("worksheetName");
            }

            InsertAppendSheetToExcelStream(InsertAppend.Insert, inputStream, workbookPath, worksheetName);
        }

        /// <summary>
        /// Appends the sheet to excel stream.
        /// </summary>
        /// <param name="inputStream">The input stream.</param>
        /// <param name="workbookPath">The workbook path.</param>
        /// <param name="worksheetName">Name of the worksheet.</param>
        public static void AppendSheetToExcelStream(MemoryStream inputStream, string workbookPath, string worksheetName)
        {
            if (inputStream == null) 
            {
                throw new ArgumentNullException("inputStream");
            }

            if (string.IsNullOrEmpty(workbookPath))
            {
                throw new ArgumentNullException("workbookPath");
            }

            if (string.IsNullOrEmpty(worksheetName))
            {
                throw new ArgumentNullException("worksheetName");
            }

            InsertAppendSheetToExcelStream(InsertAppend.Append, inputStream, workbookPath, worksheetName);

            //var fi = new FileInfo(workbookPath);
            //if(!fi.Exists)
            //{
            //    throw new ExportException(string.Format("Supplied workbook path does not exist <{0}>", workbookPath));
            //}

            //// start the append process by opening the inputstream as a autosave spreadsheet document
            //using (SpreadsheetDocument target = SpreadsheetDocument.Open(inputStream, true, new OpenSettings { AutoSave = true }))
            //{
            //    // next open the workbook containing the sheet to append
            //    using (var source = SpreadsheetDocument.Open(workbookPath, false)) 
            //    {
            //        // get sheet from the source
            //        var wsp = source.GetWorksheetPart(worksheetName);
            //        if(wsp == null)
            //        {
            //            throw new ExportException(string.Format("Unable to clone worksheet <{0}> from workbook path <{1}>", worksheetName, workbookPath));                        
            //        }

            //        // and insert a clone into the target
            //        var clone = target.WorkbookPart.InsertClonedWorksheetPart(wsp, SheetStateValues.Visible, worksheetName);

            //        // call the resource and style merger, this merges the index driven styles
            //        // between source and target stream
            //        ExportGenerator.MergeResourcesAndStyles(clone, source, target);
            //    }
            //}
        }

        /// <summary>
        /// Appends the sheet to excel stream.
        /// </summary>
        /// <param name="inputStream">The input stream.</param>
        /// <param name="workbookPath">The workbook path.</param>
        /// <param name="worksheetName">Name of the worksheet.</param>
        private static void InsertAppendSheetToExcelStream(InsertAppend insertAppend, MemoryStream inputStream, string workbookPath, string worksheetName)
        {
            if (inputStream == null)
            {
                throw new ArgumentNullException("inputStream");
            }

            if (string.IsNullOrEmpty(workbookPath))
            {
                throw new ArgumentNullException("workbookPath");
            }

            if (string.IsNullOrEmpty(worksheetName))
            {
                throw new ArgumentNullException("worksheetName");
            }

            var fi = new FileInfo(workbookPath);
            if (!fi.Exists)
            {
                throw new ExportException(string.Format("Supplied workbook path does not exist <{0}>", workbookPath));
            }

            // start the append process by opening the inputstream as a autosave spreadsheet document
            using (SpreadsheetDocument target = SpreadsheetDocument.Open(inputStream, true, new OpenSettings { AutoSave = true }))
            {
                // next open the workbook containing the sheet to append
                using (var source = SpreadsheetDocument.Open(workbookPath, false))
                {
                    // get sheet from the source
                    var wsp = source.GetWorksheetPart(worksheetName);
                    if (wsp == null)
                    {
                        throw new ExportException(string.Format("Unable to clone worksheet <{0}> from workbook path <{1}>", worksheetName, workbookPath));
                    }

                    // and insert a clone into the target
                    if (insertAppend == InsertAppend.Append)
                    {
                        var clone = target.WorkbookPart.InsertClonedWorksheetPart(wsp, SheetStateValues.Visible, worksheetName);

                        // call the resource and style merger, this merges the index driven styles
                        // between source and target stream
                        ExportGenerator.MergeResourcesAndStyles(clone, source, target);
                    }
                    else
                    {
                        var clone = target.WorkbookPart.InsertClonedWorksheetPartAtFirst(wsp, SheetStateValues.Visible, worksheetName);

                        // call the resource and style merger, this merges the index driven styles
                        // between source and target stream
                        ExportGenerator.MergeResourcesAndStyles(clone, source, target);
                    }
                }
            }
        }

        /// <summary>
        /// Takes many streams of excel and merges into one
        /// </summary>
        /// <param name="inputStreams">The input streams.</param>
        /// <returns>The merged output stream</returns>
        public static MemoryStream MergeExcelStreams(List<MemoryStream> inputStreams)
        {
            // no streams return null
            if (inputStreams == null || inputStreams.Count == 0)
            {
                return null;
            }

            // only 1 just return it, dont bother to merge
            if (inputStreams.Count == 1)
            {
                return inputStreams[0];
            }

            // start the merge process, using the 1st stream of the starting point
            using (SpreadsheetDocument target = SpreadsheetDocument.Open(inputStreams[0], true, new OpenSettings { AutoSave = true }))
            {
                // use this dictionary as a map to keep track of styles created during the merge process
                // for example style 1 in source may be style 2 in the target
                var mergedStyles = new Dictionary<uint, uint>();

                // make sure defined names is there, its always used so may as well do now
                if (target.WorkbookPart.Workbook.DefinedNames == null)
                {
                    target.WorkbookPart.Workbook.DefinedNames = new Spreadsheet.DefinedNames();
                }

                bool first = true;
                foreach (var stream in inputStreams)
                {
                    // skip the 1st, its the start point
                    if (first)
                    {
                        first = false;
                        continue;
                    }

                    // start moving through the rest of the stream, creating a spreadsheet doc for each
                    using (SpreadsheetDocument source = SpreadsheetDocument.Open(stream, true, new OpenSettings { AutoSave = true }))
                    {
                        // use this to keep track of sheet id changes during copy of sheets from source to target
                        var sheetIdMap = new Dictionary<int, int>();

                        // work through each sheet in the workbook
                        foreach (var sheet in source.WorkbookPart.Workbook.Descendants<Spreadsheet.Sheet>())
                        {
                            // very hidden(?!) sheets are not cloned, they seem to contain things like macros
                            if (sheet.State != null && sheet.State.HasValue && sheet.State.Value == SheetStateValues.VeryHidden)
                            {
                                continue;
                            }

                            // get sheet from the source
                            var wsp = (WorksheetPart)source.WorkbookPart.GetPartById(sheet.Id);
                            // and insert a clone into the target
                            var clone = target.WorkbookPart.InsertClonedWorksheetPart(wsp, sheet.State, sheet.Name.HasValue ? sheet.Name.Value : null);

                            int sheetIndex = wsp.GetSheetIndex();
                            int cloneSheetIndex = clone.GetSheetIndex();

                            // add the original and cloned sheet index to the map
                            if (cloneSheetIndex != -1 && sheetIndex != -1 && !sheetIdMap.ContainsKey(sheetIndex))
                            {
                                sheetIdMap.Add(sheetIndex, cloneSheetIndex);
                            }

                            // clear style for the next sheet, might be worth trying this foreach stream
                            mergedStyles.Clear();

                            // call the resource and style merger, this merges the index driven styles
                            // between source and target stream
                            ExportGenerator.MergeResourcesAndStyles(clone, source, target, mergedStyles);
                        }

                        foreach (var dn in source.WorkbookPart.Workbook.Descendants<Spreadsheet.DefinedName>())
                        {
                            var clonedDefinedName = (Spreadsheet.DefinedName)dn.CloneNode(true);

                            if (clonedDefinedName == null)
                            {
                                continue;
                            }

                            if (clonedDefinedName.LocalSheetId != null)
                            {
                                int clonedSheetIndex = (int)clonedDefinedName.LocalSheetId.Value;
                                if (sheetIdMap.ContainsKey(clonedSheetIndex))
                                {
                                    clonedDefinedName.LocalSheetId = new UInt32Value((uint)sheetIdMap[clonedSheetIndex]);
                                }
                            }

                            target.WorkbookPart.Workbook.DefinedNames.AppendChild<Spreadsheet.DefinedName>(clonedDefinedName);
                        }
                    }
                }
            }
            return inputStreams[0];
        }     
    }
}
