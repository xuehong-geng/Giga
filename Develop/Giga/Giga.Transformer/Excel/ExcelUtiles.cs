using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Giga.Transformer.Excel
{
    /// <summary>
    /// Utilities for handling excel sheet
    /// </summary>
    public class ExcelUtiles
    {
        /// <summary>
        /// Delete a worksheet from excel file
        /// </summary>
        /// <param name="fileName">Path of excel file</param>
        /// <param name="sheetToDelete">Name of sheet to be deleted</param>
        public static void DeleteWorksheet(string fileName, string sheetToDelete)
        {
            using (SpreadsheetDocument doc = SpreadsheetDocument.Open(fileName, true))
            {
                DeleteWorksheet(doc, sheetToDelete);
                doc.Close();
            }
        }

        /// <summary>
        /// Delete a worksheet from excel doc
        /// </summary>
        /// <param name="doc">Spreadsheet document</param>
        /// <param name="sheetToDelete">Name of sheet to be deleted</param>
        public static void DeleteWorksheet(SpreadsheetDocument doc, string sheetToDelete)
        {
            WorkbookPart wbPart = doc.WorkbookPart;
            // Get the pivot Table Parts
            IEnumerable<PivotTableCacheDefinitionPart> pvtTableCacheParts = wbPart.PivotTableCacheDefinitionParts;
            var pvtTableCacheDefinationPart = new Dictionary<PivotTableCacheDefinitionPart, string>();
            foreach (PivotTableCacheDefinitionPart item in pvtTableCacheParts)
            {
                PivotCacheDefinition pvtCacheDef = item.PivotCacheDefinition;
                //Check if this CacheSource is linked to SheetToDelete
                var pvtCahce = pvtCacheDef.Descendants<CacheSource>().Where(s => s.WorksheetSource.Sheet == sheetToDelete);
                if (pvtCahce.Count() > 0)
                {
                    pvtTableCacheDefinationPart.Add(item, item.ToString());
                }
            }
            foreach (var Item in pvtTableCacheDefinationPart)
            {
                wbPart.DeletePart(Item.Key);
            }
            //Get the SheetToDelete from workbook.xml
            Sheet theSheet = wbPart.Workbook.Descendants<Sheet>().FirstOrDefault(s => s.Name == sheetToDelete);
            if (theSheet == null)
                return;
            //Store the SheetID for the reference
            var sheetid = theSheet.SheetId;
            // Remove the sheet reference from the workbook.
            var worksheetPart = (WorksheetPart)(wbPart.GetPartById(theSheet.Id));
            theSheet.Remove();
            // Delete the worksheet part.
            wbPart.DeletePart(worksheetPart);
            //Get the DefinedNames
            var definedNames = wbPart.Workbook.Descendants<DefinedNames>().FirstOrDefault();
            if (definedNames != null)
            {
                var defNamesToDelete = new List<DefinedName>();
                foreach (DefinedName Item in definedNames)
                {
                    // This condition checks to delete only those names which are part of Sheet in question
                    if (Item.Text.Contains(sheetToDelete + "!"))
                        defNamesToDelete.Add(Item);
                }
                foreach (DefinedName Item in defNamesToDelete)
                {
                    Item.Remove();
                }
            }
            // Get the CalculationChainPart 
            //Note: An instance of this part type contains an ordered set of references to all cells in all worksheets in the 
            //workbook whose value is calculated from any formula
            CalculationChainPart calChainPart;
            calChainPart = wbPart.CalculationChainPart;
            if (calChainPart != null)
            {
                var iSheetId = new Int32Value((int)sheetid.Value);
                var calChainEntries = calChainPart.CalculationChain.Descendants<CalculationCell>().Where(c => c.SheetId != null && c.SheetId.HasValue && c.SheetId.Value == iSheetId).ToList();
                var calcsToDelete = new List<CalculationCell>();
                foreach (CalculationCell Item in calChainEntries)
                {
                    calcsToDelete.Add(Item);
                }
                foreach (CalculationCell Item in calcsToDelete)
                {
                    Item.Remove();
                }
                if (calChainPart.CalculationChain.Count() == 0)
                {
                    wbPart.DeletePart(calChainPart);
                }
            }
            // Save the workbook.
            wbPart.Workbook.Save();
        }

        /// <summary>
        /// Copy sheet from one file to another
        /// </summary>
        /// <param name="srcFile">Path of source excel file</param>
        /// <param name="srcSheet">Name of source sheet</param>
        /// <param name="tgtFile">Path of target excel file</param>
        /// <param name="tgtPos">Position of copied sheet in target file</param>
        public static void CopyWorksheet(String srcFile, String srcSheet, String tgtFile, uint tgtPos)
        {
            var srcDoc = SpreadsheetDocument.Open(srcFile, false);
            var tgtDoc = SpreadsheetDocument.Open(tgtFile, true);
            CopyWorksheet(srcDoc, srcSheet, tgtDoc, tgtPos);
            tgtDoc.Close();
            srcDoc.Close();
        }

        /// <summary>
        /// Copy sheet to another document
        /// </summary>
        /// <param name="sourceDoc">Source document</param>
        /// <param name="srcSheetName">Name of source sheet</param>
        /// <param name="targetDoc">Spreadsheet document to copied</param>
        /// <param name="targetIndex">Index of copied sheet in target document</param>
        public static void CopyWorksheet(SpreadsheetDocument sourceDoc, String srcSheetName, SpreadsheetDocument targetDoc, uint targetIndex)
        {
            // Locate the source sheet
            if (sourceDoc.WorkbookPart == null)
                throw new InvalidOperationException("WorkbookPart is not exist in sourceDoc!");
            if (sourceDoc.WorkbookPart.Workbook.Sheets == null)
                throw new InvalidOperationException("No sheets exist in sourceDoc!");
            var srcSheet =
                sourceDoc.WorkbookPart.Workbook.Sheets.Descendants<Sheet>()
                    .FirstOrDefault(a => a.Name == srcSheetName);
            if (srcSheet == null)
                throw new InvalidOperationException(String.Format("No sheet found with name {0}!", srcSheetName));
            var srcSheetPart = sourceDoc.WorkbookPart.GetPartById(srcSheet.Id) as WorksheetPart;
            if (srcSheetPart == null)
                throw new InvalidOperationException(String.Format("Cannot find worksheet part with Id {0}!", srcSheet.Id));
            var srcWorkSheet = srcSheetPart.Worksheet;
            if (srcWorkSheet == null)
                throw new InvalidOperationException("Worksheet not exist in source worksheet part!");
            // Locate the position of target sheet
            WorkbookPart tgtWbPart = targetDoc.WorkbookPart ?? targetDoc.AddWorkbookPart();
            Sheets tgtSheets = tgtWbPart.Workbook.Sheets ?? tgtWbPart.Workbook.AppendChild<Sheets>(new Sheets());
            if (targetIndex > tgtSheets.Count()) targetIndex = (uint)tgtSheets.Count();
            // Create a new worksheet and clone data from original worksheet
            var newSheetPart = tgtWbPart.AddNewPart<WorksheetPart>();
            newSheetPart.Worksheet = new Worksheet(); //srcWorkSheet.Clone() as Worksheet;
            // Create a unique ID for the new worksheet.
            uint sheetId = 1;
            if (tgtSheets.Elements<Sheet>().Any())
            {
                sheetId = tgtSheets.Elements<Sheet>().Select(s => s.SheetId.Value).Max() + 1;
            }
            // Add cloned worksheet to target workbook
            var newSheet = new Sheet()
            {
                Id = tgtWbPart.GetIdOfPart(newSheetPart),
                SheetId = sheetId,
                Name = srcSheet.Name
            };
            tgtSheets.InsertAt(newSheet, (int)targetIndex);
            // Import data from source sheet to target sheet
            ImportWorksheet(sourceDoc, srcWorkSheet, targetDoc, newSheetPart.Worksheet);
            // Import all necessary resources into target document that referenced by cloned sheet
            //ImportResources(sourceDoc, newSheetPart.Worksheet, targetDoc);
            // Import defined names
            ImportDefinedNames(sourceDoc, srcSheetName, targetDoc);
            // Import calculate chain
            ImportCalculateChain(sourceDoc, (int)srcSheet.SheetId.Value, targetDoc, (int)sheetId);
            // Save it
            tgtWbPart.Workbook.Save();
        }

        /// <summary>
        /// Import worksheet part from one document to another
        /// </summary>
        /// <param name="sourceDoc">Source excel document</param>
        /// <param name="srcSheet">Source worksheet</param>
        /// <param name="targetDoc">Target excel document</param>
        /// <param name="tgtSheet">Position of imported worksheet</param>
        /// <returns></returns>
        protected static void ImportWorksheet(SpreadsheetDocument sourceDoc, Worksheet srcSheet, SpreadsheetDocument targetDoc, Worksheet tgtSheet)
        {
            // Import sheet format properties
            tgtSheet.SheetFormatProperties = srcSheet.SheetFormatProperties.Clone() as SheetFormatProperties;
            // Import dimension
            tgtSheet.SheetDimension = srcSheet.SheetDimension.Clone() as SheetDimension;
            // Imported style buffer
            var dicStyles = new Dictionary<uint, uint>();
            // Import columns
            var srcCols = srcSheet.GetFirstChild<Columns>();
            if (srcCols != null)
            {
                var tgtCols = tgtSheet.GetFirstChild<Columns>() ?? tgtSheet.AppendChild(new Columns());
                foreach (var srcCol in srcCols.Descendants<Column>())
                {
                    var tgtCol = (Column)srcCol.Clone();
                    tgtCols.AppendChild(tgtCol);
                    // Handle column style
                    if (srcCol.Style != null)
                    {
                        uint iNewIdx = 0;
                        if (dicStyles.ContainsKey(srcCol.Style.Value))
                            iNewIdx = dicStyles[srcCol.Style.Value];
                        else
                        {
                            iNewIdx = ImportStyle(sourceDoc, srcCol.Style.Value, targetDoc);
                            dicStyles[srcCol.Style.Value] = iNewIdx;
                        }
                        tgtCol.Style = new UInt32Value(iNewIdx);
                    }
                }
            }
            // Import sheet data
            var srcSheetData = srcSheet.GetFirstChild<SheetData>();
            if (srcSheetData != null)
            {
                var tgtSheetData = tgtSheet.GetFirstChild<SheetData>();
                if (tgtSheetData == null)
                {
                    tgtSheetData = tgtSheet.AppendChild(new SheetData()) as SheetData;
                }
                // Import rows
                foreach (var row in srcSheetData.Elements<Row>())
                {
                    var newRow = row.CloneNode(false) as Row;
                    tgtSheetData.AppendChild(newRow);
                    // Import row style
                    if (newRow.StyleIndex != null)
                    {
                        var id = newRow.StyleIndex.Value;
                        var newId = dicStyles.ContainsKey(id) ? dicStyles[id] : ImportStyle(sourceDoc, id, targetDoc);
                        dicStyles[id] = newId;
                        newRow.StyleIndex = new UInt32Value(newId);
                    }
                    // Import cells
                    foreach (var cell in row.Elements<Cell>())
                    {
                        var newCell = cell.Clone() as Cell;
                        if (newCell == null) throw new InvalidCastException("Cloned cell is not Cell!");
                        // Handle shared string
                        if (cell.CellValue != null && cell.DataType != null &&
                            cell.DataType.Value == CellValues.SharedString)
                        {
                            uint id = uint.Parse(cell.CellValue.InnerText);
                            uint newId = ImportSharedString(sourceDoc, id, targetDoc);
                            newCell.CellValue = new CellValue(newId.ToString(CultureInfo.InvariantCulture));
                        }
                        // Handle style
                        if (cell.StyleIndex != null)
                        {
                            uint id = uint.Parse(cell.StyleIndex.InnerText);
                            uint newId = dicStyles.ContainsKey(id) ? dicStyles[id] : ImportStyle(sourceDoc, id, targetDoc);
                            dicStyles[id] = newId;
                            newCell.StyleIndex = new UInt32Value(newId);
                        }
                        // Handle cell metadata
                        if (cell.CellMetaIndex != null)
                        {}
                        // Handle value metadata
                        if (cell.ValueMetaIndex != null)
                        {}
                        newRow.AppendChild(newCell);
                    }
                }
            }
            // Import merge cells
            var srcMergeCells = srcSheet.GetFirstChild<MergeCells>();
            if (srcMergeCells != null)
            {
                tgtSheet.AppendChild(srcMergeCells.Clone() as MergeCells);
            }
        }

        /// <summary>
        /// Import defined names into target document
        /// </summary>
        /// <param name="sourceDoc"></param>
        /// <param name="sheetName"></param>
        /// <param name="targetDoc"></param>
        protected static void ImportDefinedNames(SpreadsheetDocument sourceDoc, String sheetName, SpreadsheetDocument targetDoc)
        {
            var srcNames = sourceDoc.WorkbookPart.Workbook.Descendants<DefinedNames>().FirstOrDefault();
            if (srcNames == null || !srcNames.Any())
                return; // No names
            var tgtNames = targetDoc.WorkbookPart.Workbook.Descendants<DefinedNames>().FirstOrDefault();
            if (tgtNames == null)
            {
                // Note: The defined names must be insert after sheets to make sure it is previous to calcPr.
                // if defined names are inserted after calcPr, excel will treat this document as damaged one.
                var sheets = targetDoc.WorkbookPart.Workbook.Sheets;
                tgtNames = targetDoc.WorkbookPart.Workbook.InsertAfter(new DefinedNames(), sheets);
            }
            foreach (var srcName in srcNames.Descendants<DefinedName>())
            {
                if (srcName.Text.Contains(sheetName))
                {   // Found one
                    var tgtName = tgtNames.Descendants<DefinedName>().FirstOrDefault(a => a.Name == srcName.Name);
                    if (tgtName == null)
                    {
                        tgtName = (DefinedName)srcName.Clone();
                        tgtNames.AppendChild(tgtName);
                    }
                    else
                    {
                        tgtNames.ReplaceChild((DefinedName) srcName.Clone(), tgtName);
                    }
                }
            }
        }

        /// <summary>
        /// Import calculate chain
        /// </summary>
        /// <param name="sourceDoc"></param>
        /// <param name="srcSheetId"></param>
        /// <param name="targetDoc"></param>
        /// <param name="tgtSheetId"></param>
        protected static void ImportCalculateChain(SpreadsheetDocument sourceDoc, int srcSheetId, SpreadsheetDocument targetDoc, int tgtSheetId)
        {
            var srcChainPart = sourceDoc.WorkbookPart.CalculationChainPart;
            if (srcChainPart != null)
            {
                var tgtChainPart = targetDoc.WorkbookPart.CalculationChainPart;
                if (tgtChainPart == null)
                    tgtChainPart = targetDoc.WorkbookPart.AddNewPart<CalculationChainPart>();
                if (tgtChainPart.CalculationChain == null)
                    tgtChainPart.CalculationChain = new CalculationChain();
                foreach (
                    var srcCell in
                        srcChainPart.CalculationChain.Descendants<CalculationCell>()
                            .Where(c => c.SheetId.HasValue && c.SheetId.Value == srcSheetId))
                {
                    var tgtCell = (CalculationCell) srcCell.Clone();
                    tgtCell.SheetId = new Int32Value(tgtSheetId);
                    tgtChainPart.CalculationChain.AppendChild(tgtCell);
                }
            }
        }

        /// <summary>
        /// Import referenced resource for worksheet from source document to target document
        /// </summary>
        /// <param name="sourceDoc">Source excel document</param>
        /// <param name="worksheet">Worksheet to be processed</param>
        /// <param name="targetDoc">Target document</param>
        protected static void ImportResources(SpreadsheetDocument sourceDoc, Worksheet worksheet, SpreadsheetDocument targetDoc)
        {
            var sheetData = worksheet.GetFirstChild<SheetData>();
            if (sheetData == null)
                return; // No Sheet data exist, no need to import resource

            // Enumerate rows
            foreach (var row in sheetData.Elements<Row>())
            {
                // Enumerate cells
                foreach (var cell in row.Elements<Cell>())
                {
                    // Handle shared string
                    if (cell.CellValue != null &&
                        cell.DataType == new EnumValue<CellValues>(CellValues.SharedString))
                    {
                        uint id = uint.Parse(cell.CellValue.InnerText);
                        uint newId = ImportSharedString(sourceDoc, id, targetDoc);
                        cell.CellValue = new CellValue(newId.ToString(CultureInfo.InvariantCulture));
                    }
                }
            }
        }

        /// <summary>
        /// Import shared string item from one document to another
        /// </summary>
        /// <param name="sourceDoc">Source excel document</param>
        /// <param name="id">Id of shared string item</param>
        /// <param name="targetDoc">Target excel document</param>
        /// <returns></returns>
        protected static uint ImportSharedString(SpreadsheetDocument sourceDoc, uint id, SpreadsheetDocument targetDoc)
        {
            // Prepare resource parts 
            var srcSharedStringPart = sourceDoc.WorkbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
            if (srcSharedStringPart == null)
                throw new InvalidOperationException("SharedStringPart was not found in Source document!");
            if (srcSharedStringPart.SharedStringTable == null)
                throw new InvalidOperationException("SharedStringTable was not found in Source document!");
            var tgtSharedStringPart = targetDoc.WorkbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
            if (tgtSharedStringPart == null)
            {   // Create one
                tgtSharedStringPart = targetDoc.WorkbookPart.AddNewPart<SharedStringTablePart>();
            }
            if (tgtSharedStringPart.SharedStringTable == null)
            {
                tgtSharedStringPart.SharedStringTable = new SharedStringTable();
            }
            // Get original shared string item
            var srcItem = srcSharedStringPart.SharedStringTable.ElementAt((int)id) as SharedStringItem;
            if (srcItem == null)
                throw new InvalidOperationException(
                    String.Format("SharedStringItem {0} was not found in source document!", id));
            // Try to find same shared string item in target
            uint i = 0;
            foreach (var item in tgtSharedStringPart.SharedStringTable.Elements<SharedStringItem>())
            {
                if (item.InnerText == srcItem.InnerText)
                {   // Found same value
                    return i;
                }
                i++;
            }
            // Not found, import it
            var tgtItem = srcItem.Clone() as SharedStringItem;
            tgtSharedStringPart.SharedStringTable.AppendChild(tgtItem);
            return i;
        }

        /// <summary>
        /// Import style from source to target document
        /// </summary>
        /// <param name="sourceDoc"></param>
        /// <param name="id"></param>
        /// <param name="targetDoc"></param>
        /// <returns></returns>
        protected static uint ImportStyle(SpreadsheetDocument sourceDoc, uint id, SpreadsheetDocument targetDoc)
        {
            // Get style parts
            var srcStylePart = sourceDoc.WorkbookPart.GetPartsOfType<WorkbookStylesPart>().FirstOrDefault();
            if (srcStylePart == null)
                throw new InvalidOperationException("StylesPart was not found in source document!");
            if (srcStylePart.Stylesheet == null)
                throw new InvalidOperationException("Stylesheet was not found in source document!");
            var tgtStylePart = targetDoc.WorkbookPart.GetPartsOfType<WorkbookStylesPart>().FirstOrDefault();
            if (tgtStylePart == null)
                tgtStylePart = targetDoc.AddNewPart<WorkbookStylesPart>();
            if (tgtStylePart.Stylesheet == null)
                tgtStylePart.Stylesheet = new Stylesheet();
            // Get the source style
            var srcCellFormat = srcStylePart.Stylesheet.CellFormats.ElementAt((int) id) as CellFormat;
            if (srcCellFormat == null)
                throw new InvalidOperationException(String.Format("CellFormat {0} was not found in source document!", id));
            // Import the style to target
            return ImportCellFormat(srcStylePart.Stylesheet, srcCellFormat, tgtStylePart.Stylesheet);
        }

        /// <summary>
        /// Import cell format
        /// </summary>
        /// <param name="sourceStyleSheet"></param>
        /// <param name="srcFormat"></param>
        /// <param name="targetStyleSheet"></param>
        /// <returns></returns>
        protected static uint ImportCellFormat(Stylesheet sourceStyleSheet, CellFormat srcFormat, Stylesheet targetStyleSheet)
        {
            var tgtCellFormats = targetStyleSheet.CellFormats ??
                                 targetStyleSheet.AppendChild(new CellFormats());
            var tgtCellFormat = tgtCellFormats.AppendChild(srcFormat.Clone() as CellFormat);
            tgtCellFormats.Count ++;
            // Import Style
            if (srcFormat.FormatId != null)
            {
                var srcStyle = sourceStyleSheet.CellStyleFormats.ElementAt((int) srcFormat.FormatId.Value) as CellFormat;
                if (srcStyle == null)
                    throw new InvalidOperationException(String.Format("Style {0} was not found in source document!",
                        srcFormat.FormatId));
                // Style is complex, so just clone it
                var tgtStyle = srcStyle.Clone() as CellFormat;
                // Add style to target
                if (targetStyleSheet.CellStyleFormats == null)
                    targetStyleSheet.CellStyleFormats = new CellStyleFormats();
                targetStyleSheet.CellStyleFormats.AppendChild(tgtStyle);
                tgtCellFormat.FormatId = targetStyleSheet.CellStyleFormats.Count - 1;
                // Import details
                ImportCellFormatDetails(sourceStyleSheet, tgtStyle, targetStyleSheet);
            }
            // Import details
            ImportCellFormatDetails(sourceStyleSheet, tgtCellFormat, targetStyleSheet);

            return tgtCellFormats.Count - 1;
        }

        /// <summary>
        /// Import number, font, fill, border etc for cell format
        /// </summary>
        /// <param name="sourceStyleSheet"></param>
        /// <param name="format"></param>
        /// <param name="targetStyleSheet"></param>
        protected static void ImportCellFormatDetails(Stylesheet sourceStyleSheet, CellFormat format, Stylesheet targetStyleSheet)
        {
            // Merge number format
            #region NumberFormat
            if (format.NumberFormatId != null && format.NumberFormatId.Value != 0)
            {
                var numFmt =
                    sourceStyleSheet.NumberingFormats.Elements<NumberingFormat>()
                        .FirstOrDefault(a => a.NumberFormatId == format.NumberFormatId);
                if (numFmt != null)
                {
                    // Try to find same number format in target
                    if (targetStyleSheet.NumberingFormats == null)
                        targetStyleSheet.NumberingFormats = new NumberingFormats();
                    var existFmt =
                        targetStyleSheet.NumberingFormats.Elements<NumberingFormat>()
                            .FirstOrDefault(a => a.FormatCode == numFmt.FormatCode);
                    if (existFmt == null)
                    {
                        // Not found, add one
                        var newNumId = targetStyleSheet.NumberingFormats.Any()
                            ? targetStyleSheet.NumberingFormats.Elements<NumberingFormat>()
                                .Max(a => a.NumberFormatId).Value + 1
                            : 1;
                        targetStyleSheet.NumberingFormats.AppendChild(new NumberingFormat
                        {
                            NumberFormatId = newNumId,
                            FormatCode = numFmt.FormatCode
                        });
                        format.NumberFormatId = newNumId;
                        targetStyleSheet.NumberingFormats.Count ++;
                    }
                    else
                    {
                        // Found, use exist one
                        format.NumberFormatId = existFmt.NumberFormatId;
                    }
                }
            }
            #endregion
            // Merge font format
            #region Font
            if (format.FontId != null)
            {
                var srcFont = sourceStyleSheet.Fonts.ElementAt((int)format.FontId.Value);
                if (srcFont == null)
                    throw new InvalidOperationException(String.Format("Font {0} was not found in source document!",
                        format.FontId));
                // Since font is complex, we just clone it
                var tgtFont = srcFont.Clone() as Font;
                // Add font to target stylesheet
                if (targetStyleSheet.Fonts == null)
                    targetStyleSheet.Fonts = new Fonts();
                targetStyleSheet.Fonts.AppendChild(tgtFont);
                targetStyleSheet.Fonts.Count ++;
                format.FontId = targetStyleSheet.Fonts.Count - 1;
            }
            #endregion
            // Merge fill
            #region Fill
            if (format.FillId != null)
            {
                var srcFill = sourceStyleSheet.Fills.ElementAt((int)format.FillId.Value);
                if (srcFill == null)
                    throw new InvalidOperationException(String.Format("Fill {0} was not found in source document!",
                        format.FillId));
                // Since fill is complex, we just clone it
                var tgtFill = srcFill.Clone() as Fill;
                // Add fill to target stylesheet
                if (targetStyleSheet.Fills == null)
                    targetStyleSheet.Fills = new Fills();
                targetStyleSheet.Fills.AppendChild(tgtFill);
                targetStyleSheet.Fills.Count ++;
                format.FillId = targetStyleSheet.Fills.Count - 1;
            }
            #endregion
            // Merge border
            #region Border
            if (format.BorderId != null)
            {
                var srcBorder = sourceStyleSheet.Borders.ElementAt((int)format.BorderId.Value);
                if (srcBorder == null)
                    throw new InvalidOperationException(String.Format("Border {0} was not found in source document!",
                        format.BorderId));
                // Since border is complex, we just clone it
                var tgtBorder = srcBorder.Clone() as Border;
                // Add border to target stylesheet
                if (targetStyleSheet.Borders == null)
                    targetStyleSheet.Borders = new Borders();
                targetStyleSheet.Borders.AppendChild(tgtBorder);
                targetStyleSheet.Borders.Count ++;
                format.BorderId = targetStyleSheet.Borders.Count - 1;
            }
            #endregion
        }
    }
}
