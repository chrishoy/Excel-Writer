namespace ExcelWriter.OpenXml.Excel
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Linq;
    using DocumentFormat.OpenXml;
    using DocumentFormat.OpenXml.Packaging;
    using DocumentFormat.OpenXml.Spreadsheet;

    /// <summary>
    /// Extension methods for <see cref="WorkbookPart" />
    /// </summary>
    public static class WorkbookPartExtensions
    {
        /// <summary>
        /// Inserts the cloned worksheet part.
        /// </summary>
        /// <param name="workbookPart">The workbook part.</param>
        /// <param name="sourceSheetName">Name of the source sheet.</param>
        /// <param name="targetSheetName">Name of the target sheet.</param>
        /// <returns></returns>
        /// <exception cref="ArgumentNullException">
        /// workbookPart
        /// or
        /// sourceSheetName
        /// or
        /// targetSheetName
        /// </exception>
        public static WorksheetPart InsertClonedWorksheetPart(this WorkbookPart workbookPart, string sourceSheetName, string targetSheetName)
        {
            if (workbookPart == null)
            {
                throw new ArgumentNullException("workbookPart");
            }
            if (string.IsNullOrEmpty(sourceSheetName))
            {
                throw new ArgumentNullException("sourceSheetName");
            }
            if (string.IsNullOrEmpty(targetSheetName))
            {
                throw new ArgumentNullException("targetSheetName");
            }

            // find the source sheet
            var sourceWorksheetPart = workbookPart.Workbook.GetWorksheetPartByName(sourceSheetName);

            // return null now if there isnt a match for the source sheet name
            if (sourceWorksheetPart == null)
            {
                return null;
            }

            var state = workbookPart.Workbook.GetWorksheetStateName(sourceSheetName);

            // take advantage of AddPart for deep cloning
            return workbookPart.InsertClonedWorksheetPart(sourceWorksheetPart, state, targetSheetName);            
        }

        /// <summary>
        /// 
        /// </summary>
        private enum InsertAppend
        {
            /// <summary>
            /// The insert
            /// </summary>
            Insert,
            /// <summary>
            /// The append
            /// </summary>
            Append
        }

        /// <summary>
        /// Inserts a cloned worksheet part based on the worksheet with the provided source sheet name as the first tab
        /// </summary>
        /// <param name="targetWorkbookPart">The workbook part that will contain the cloned worksheet part</param>
        /// <param name="sourceWorksheetPart">The the sheet to clone</param>
        /// <param name="sourceState">State of the source.</param>
        /// <param name="targetSheetName">The name of the cloned sheet</param>
        /// <returns>
        /// The cloned worksheet part
        /// </returns>
        public static WorksheetPart InsertClonedWorksheetPartAtFirst(this WorkbookPart targetWorkbookPart, WorksheetPart sourceWorksheetPart, EnumValue<SheetStateValues> sourceState, string targetSheetName)
        {
            Guard.IsNotNull(targetWorkbookPart, "targetWorkbookPart");
            Guard.IsNotNull(sourceWorksheetPart, "sourceWorksheetPart");

            return InsertAppendClonedWorksheetPart(InsertAppend.Insert, targetWorkbookPart, sourceWorksheetPart, sourceState, targetSheetName);
        }

        /// <summary>
        /// Inserts a cloned worksheet part based on the worksheet with the provided source sheet name
        /// </summary>
        /// <param name="targetWorkbookPart">The workbook part that will contain the cloned worksheet part</param>
        /// <param name="sourceWorksheetPart">The the sheet to clone</param>
        /// <param name="sourceState">State of the source.</param>
        /// <param name="targetSheetName">The name of the cloned sheet</param>
        /// <returns>
        /// The cloned worksheet part
        /// </returns>
        public static WorksheetPart InsertClonedWorksheetPart(this WorkbookPart targetWorkbookPart, WorksheetPart sourceWorksheetPart, EnumValue<SheetStateValues> sourceState, string targetSheetName)
        {

            // I dont know why this method was originally called InsertClonedWorksheetPart, as
            // it does  an append...? see  InsertClonedWorksheetPartAtFirst for a method  that 
            // actually does the insert.
            Guard.IsNotNull(targetWorkbookPart, "targetWorkbookPart");
            Guard.IsNotNull(sourceWorksheetPart, "sourceWorksheetPart");

            return InsertAppendClonedWorksheetPart(InsertAppend.Append, targetWorkbookPart, sourceWorksheetPart, sourceState, targetSheetName);
        }

        /// <summary>
        /// Inserts the append cloned worksheet part.
        /// </summary>
        /// <param name="action">The action.</param>
        /// <param name="targetWorkbookPart">The target workbook part.</param>
        /// <param name="sourceWorksheetPart">The source worksheet part.</param>
        /// <param name="sourceState">State of the source.</param>
        /// <param name="targetSheetName">Name of the target sheet.</param>
        /// <returns></returns>
        /// <exception cref="ArgumentNullException">
        /// workbookPart
        /// or
        /// sourceWorksheetPart
        /// </exception>
        private static WorksheetPart InsertAppendClonedWorksheetPart(this InsertAppend action, WorkbookPart targetWorkbookPart, WorksheetPart sourceWorksheetPart, EnumValue<SheetStateValues> sourceState, string targetSheetName)
        {
            if (targetWorkbookPart == null)
            {
                throw new ArgumentNullException("workbookPart");
            }
            if (sourceWorksheetPart == null)
            {
                throw new ArgumentNullException("sourceWorksheetPart");
            }

            // take advantage of AddPart for deep cloning
            SpreadsheetDocument tempSheet = SpreadsheetDocument.Create(new MemoryStream(), SpreadsheetDocumentType.Workbook);
            WorkbookPart tempWorkbookPart = tempSheet.AddWorkbookPart();
            WorksheetPart tempWorksheetPart = tempWorkbookPart.AddPart<WorksheetPart>(sourceWorksheetPart);

            // add cloned sheet and all associated parts to workbook
            WorksheetPart clonedSheet = targetWorkbookPart.AddPart<WorksheetPart>(tempWorksheetPart);

            // add new sheet to main workbook part
            Sheets sheets = targetWorkbookPart.Workbook.GetFirstChild<Sheets>();
            var nextSheetId = sheets.GetNextSheetId();

            targetSheetName = sheets.GetSafeSheetName(targetSheetName ?? "Sheet");

            var newSheet = new DocumentFormat.OpenXml.Spreadsheet.Sheet
            {
                Id = targetWorkbookPart.GetIdOfPart(clonedSheet),
                // make sure the name is set to the provided name
                Name = targetSheetName,
                SheetId = nextSheetId,
                State = sourceState
            };

            if (action == InsertAppend.Insert)
            {
                sheets.InsertAt(newSheet, 0);
            }
            else
            {
                sheets.Append(newSheet);
            }

            return clonedSheet;
        }

        /// <summary>
        /// Deletes the sheet.
        /// </summary>
        /// <param name="workbookPart">The workbook part.</param>
        /// <param name="sheetName">Name of the sheet.</param>
        public static void DeleteSheet(this WorkbookPart workbookPart, string sheetName)
        {
            var match = (from s in workbookPart.Workbook.Descendants<DocumentFormat.OpenXml.Spreadsheet.Sheet>()
                         where s.Name.HasValue &&
                         s.Name.Value.CompareTo(sheetName) == 0
                         select s).FirstOrDefault();

            if (match != null)
            {
                workbookPart.DeleteSheet(match);
            }
        }

        /// <summary>
        /// Deletes the supplied sheet instance from the workbook part
        /// </summary>
        /// <param name="workbookPart">The workbook part to delete the sheet from</param>
        /// <param name="sheet">The sheet to delete</param>
        /// <exception cref="ArgumentNullException">
        /// workbookPart
        /// or
        /// sheet
        /// </exception>
        public static void DeleteSheet(this WorkbookPart workbookPart, DocumentFormat.OpenXml.Spreadsheet.Sheet sheet)
        {
            if (workbookPart == null)
            {
                throw new ArgumentNullException("workbookPart");
            }
            if (sheet == null)
            {
                throw new ArgumentNullException("sheet");
            }

            // get the id of the sheet for deletion
            Int32Value sheetId = Int32Value.FromInt32((int)sheet.SheetId.Value);

            // Remove the sheet reference from the workbook.
            WorksheetPart worksheetPart = (WorksheetPart)(workbookPart.GetPartById(sheet.Id));

            SheetViews views = worksheetPart.Worksheet.GetFirstChild<SheetViews>();
            if (views != null)
            {
                views.Remove();
                worksheetPart.Worksheet.Save();
            }

            sheet.Remove();

            // Delete the worksheet part.
            workbookPart.DeletePart(worksheetPart);

            // Get the CalculationChainPart 
            // Note: An instance of this part type contains an ordered set of references to all cells in all worksheets in the 
            // workbook whose value is calculated from any formula

            CalculationChainPart calChainPart = workbookPart.CalculationChainPart;
            if (calChainPart != null)
            {
                List<CalculationCell> forRemoval = new List<CalculationCell>();

                var calChainEntries = calChainPart.CalculationChain.Descendants<CalculationCell>().ToList();

                foreach (CalculationCell item in calChainEntries)
                {
                    if (item.SheetId == null)
                    {
                        item.Remove();
                    }
                    else if (item.SheetId.HasValue && item.SheetId.Value.Equals(sheetId))
                    {
                        item.Remove();
                    }
                }
                if (calChainPart.CalculationChain.Count() == 0)
                {
                    workbookPart.DeletePart(calChainPart);
                }
            }

            workbookPart.Workbook.Save();
        }
    }
}
