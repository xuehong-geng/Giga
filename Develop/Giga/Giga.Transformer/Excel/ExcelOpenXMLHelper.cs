using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Giga.Transformer.Excel
{
    /// <summary>
    /// Extension helper for accessing Excel data via Open XML SDK
    /// </summary>
    public static class ExcelOpenXMLHelper
    {
        public const String REG_REF = @"(?i)(?<Sheet>.+)!(?<Range>.+)";
        public const String REG_CELL_REF = @"(?i)(?<COL>\$?[a-zA-Z]+)(?<ROW>\$?[1-9][0-9]*)";
        public const String REG_RANGE_REF = @"(?i)(?<COL1>\$?[a-zA-Z]+)(?<ROW1>\$?[1-9][0-9]*)\:(?<COL2>\$?[a-zA-Z]+)(?<ROW2>\$?[1-9][0-9]*)";
        public const String REG_RANGE_PART_REF = @"(?i)((?<COL1>\$?[a-zA-Z]+))?((?<ROW1>\$?[1-9][0-9]*))?\:((?<COL2>\$?[a-zA-Z]+))?((?<ROW2>\$?[1-9][0-9]*))?";

        /// <summary>
        /// Get real reference of defined name
        /// </summary>
        /// <param name="doc">Spread sheet document</param>
        /// <param name="name">Defined name</param>
        /// <returns>Real reference</returns>
        public static String GetDefinedName(this SpreadsheetDocument doc, String name)
        {
            DefinedName dn = doc.WorkbookPart.Workbook.Descendants<DefinedName>().FirstOrDefault(a => a.Name == name);
            if (dn == null)
                return null;
            return dn.InnerText;
        }

        /// <summary>
        /// Get a range from excel document
        /// </summary>
        /// <param name="doc">Spread sheet document</param>
        /// <param name="referenceName">Reference name</param>
        /// <returns></returns>
        public static ExcelOpenXMLRange GetRange(this SpreadsheetDocument doc, String referenceName)
        {
            // Get sheet name and range in sheet
            var regex = new Regex(REG_REF);
            var match = regex.Match(referenceName);
            if (!match.Success)
            {   // This reference might be named
                var realRef = GetDefinedName(doc, referenceName);
                match = regex.Match(realRef);
                if (!match.Success)
                    throw new InvalidDataException(String.Format("Defined name refered to invalid cell reference {0}!",
                        realRef));
            }
            // Find the sheet
            var sheetName = match.Groups["Sheet"].Value;
            if (sheetName.StartsWith("'") && sheetName.EndsWith("'"))
                sheetName = sheetName.Substring(1, sheetName.Length - 2);
            var sheet =
                doc.WorkbookPart.Workbook.Descendants<Sheet>().FirstOrDefault(a => a.Name == sheetName);
            if (sheet == null)
                throw new InvalidDataException(String.Format("Sheet {0} not exist!", sheetName));
            var sheetPart = doc.WorkbookPart.GetPartById(sheet.Id) as WorksheetPart;
            if (sheetPart == null)
                throw new InvalidDataException(String.Format("Sheet part {0} not exist!", sheet.SheetId));
            // Find the range
            var range = match.Groups["Range"].Value;
            return new ExcelOpenXMLRange(doc, sheetPart.Worksheet, range);
        }

        /// <summary>
        /// Check open range reference, fix it with sheet range
        /// </summary>
        public static String ExpandToSheetBound(this Worksheet sheet, String reference)
        {
            var regexRange = new Regex(REG_RANGE_PART_REF);
            var match = regexRange.Match(reference);
            if (!match.Success)
                return reference;
            var l = match.Groups["COL1"].Value;
            var r = match.Groups["COL2"].Value;
            var t = match.Groups["ROW1"].Value;
            var b = match.Groups["ROW2"].Value;
            if (String.IsNullOrEmpty(l) || String.IsNullOrEmpty(r) || String.IsNullOrEmpty(t) || String.IsNullOrEmpty(b))
            {
                var sheetRange = sheet.SheetDimension.Reference;
                var matchRange = regexRange.Match(sheetRange);
                if (matchRange.Success)
                {
                    var col1 = matchRange.Groups["COL1"].Value;
                    var col2 = matchRange.Groups["COL2"].Value;
                    var row1 = matchRange.Groups["ROW1"].Value;
                    var row2 = matchRange.Groups["ROW2"].Value;
                    if (String.IsNullOrEmpty(l)) l = col1;
                    if (String.IsNullOrEmpty(r)) r = col2;
                    if (String.IsNullOrEmpty(t)) t = row1;
                    if (String.IsNullOrEmpty(b)) b = row2;
                }
                return String.Format("{0}{1}:{2}{3}", l, t, r, b);
            }
            else
            {
                return reference;
            }
        }

        /// <summary>
        /// Calculate range of entity
        /// </summary>
        /// <param name="rangeFirst">Range in configuration</param>
        /// <param name="idx">Index of entity</param>
        /// <param name="isVertical">Whether entities arranged vertically</param>
        /// <returns></returns>
        public static String CalculateEntityRange(ExcelOpenXMLRange collectionRange, String rangeFirst, int idx, bool isVertical = true,ExcelOpenXMLRange endBefore = null)
        {
            var tl = new CellReference();
            var br = new CellReference();
            RangeReference.ParseRange(rangeFirst, ref tl, ref br);
            int height = br.Row - tl.Row + 1;
            int width = br.Col - tl.Col + 1;
            int rangeH = collectionRange.Height;
            int rangeW = collectionRange.Width;
            if (height > rangeH) height = rangeH;   // Make sure height of entity not bigger than collection
            if (width > rangeW) width = rangeW;     // Make sure width of entity not bigger than collection
            int c = 0, r = 0;
            if (isVertical)
            {   // Entities arranged by rows
                int entPerRow = rangeW / width; // How many entities could be in one row
                int rowIdx = idx / entPerRow;
                int colIdx = idx % entPerRow;
                r = rowIdx * height;
                c = colIdx * width;
            }
            else
            {   // Entities arranged by columns
                int entPerCol = rangeH / height; // How many entities could be in one column
                int colIdx = idx / entPerCol;
                int rowIdx = idx % entPerCol;
                r = rowIdx * height;
                c = colIdx * width;
            }
            tl.Col = c + 1;
            tl.Row = r + 1;
            // Check abort flag
            if (endBefore != null)
            {
                if (isVertical)
                {   // Vertical
                    if (tl.Row + collectionRange.Top - 1 >= endBefore.Top)
                        return null;
                }
                else
                {   // Horizontal
                    if (tl.Col + collectionRange.Left - 1 >= endBefore.Left)
                        return null;
                }
            }

            br = tl.Offset(width - 1, height - 1);
            return String.Format("{0}:{1}", tl, br);
        }
    }
}
