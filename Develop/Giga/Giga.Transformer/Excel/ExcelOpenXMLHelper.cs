using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Linq.Expressions;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Giga.Transformer.Excel
{
    public static class BuiltInNumberFormats
    {
        public static Dictionary<uint, String> Formats = new Dictionary<uint, string>
        {
            {0,     "General"},
            {1,     "0"},
            {2,     "0.00"},
            {3,     "#,##0"},
            {4,     "#,##0.00"},

            {9,     "0%"},
            {10,    "0.00%"},
            {11,    "0.00E+00"},
            {12,    "# ?/?"},
            {13,    "# ??/??"},
            {14,    "mm-dd-yy"},
            {15,    "d-mmm-yy"},
            {16,    "d-mmm"},
            {17,    "mmm-yy"},
            {18,    "h:mm AM/PM"},
            {19,    "h:mm:ss AM/PM"},
            {20,    "h:mm"},
            {21,    "h:mm:ss"},
            {22,    "m/d/yy h:mm"},
            {23,    "#,##0 ;(#,##0)"},
            
            {37,    "#,##0 ;(#,##0)"},
            {38,    "#,##0 ;[Red](#,##0)"},
            {39,    "#,##0.00;(#,##0.00)"},
            {40,    "#,##0.00;[Red](#,##0.00)"},
            {45,    "mm:ss"},
            {46,    "[h]:mm:ss"},
            {47,    "mmss.0"},
            {48,    "##0.0E+0"},
            {49,    "@"},

            {27,     "yyyy\"年\"m\"月\""},
            {28,    "m\"月\"d\"日\""},
            {29,    "m\"月\"d\"日\""},
            {30,    "m-d-yy"},
            {31,    "yyyy\"年\"m\"月\"d\"日\""},
            {32,    "h\"时\"mm\"分\""},
            {33,    "h\"时\"mm\"分\"ss\"秒\""},
            {34,    "上午/下午h\"时\"mm\"分\""},
            {35,    "上午/下午h\"时\"mm\"分\"ss\"秒\""},
            {36,    "yyyy\"年\"m\"月\""},
            {50,    "yyyy\"年\"m\"月\""},
            {51,    "m\"月\"d\"日\""},
            {52,    "yyyy\"年\"m\"月\""},
            {53,    "m\"月\"d\"日\""},
            {54,    "m\"月\"d\"日\""},
            {55,    "上午/下午h\"时\"mm\"分\""},
            {56,    "上午/下午h\"时\"mm\"分\"ss\"秒\""},
            {57,    "yyyy\"年\"m\"月\""},
            {58,    "m\"月\"d\"日\""},
        };

        public static uint Find(String fmt)
        {
            return (from format in Formats where format.Value == fmt select format.Key).FirstOrDefault();
        }

        public static String Get(uint id)
        {
            if (Formats.ContainsKey(id))
                return Formats[id];
            return "General";
        }
    }

    /// <summary>
    /// Extension helper for accessing Excel data via Open XML SDK
    /// </summary>
    public static class ExcelOpenXMLHelper
    {
        public const String REG_REF = @"(?i)(?<Sheet>.+)!(?<Range>.+)";
        public const String REG_CELL_REF = @"(?i)(?<COL>\$?[a-zA-Z]+)(?<ROW>\$?[1-9][0-9]*)";
        public const String REG_RANGE_REF = @"(?i)(?<COL1>\$?[a-zA-Z]+)(?<ROW1>\$?[1-9][0-9]*)\:(?<COL2>\$?[a-zA-Z]+)(?<ROW2>\$?[1-9][0-9]*)";
        public const String REG_RANGE_PART_REF = @"(?i)((?<COL1>\$?[a-zA-Z]+))?((?<ROW1>\$?[1-9][0-9]*))?\:((?<COL2>\$?[a-zA-Z]+))?((?<ROW2>\$?[1-9][0-9]*))?";


        public static bool IsSame(OpenXmlSimpleType val, OpenXmlSimpleType other)
        {
            if (val == null && other == null)
                return true;
            if ((val == null && other != null) || (val != null && other == null))
                return false;

            if (val.HasValue != other.HasValue)
                return false;
            if (val.HasValue)
            {
                return val.InnerText.Equals(other.InnerText);
            }
            return true;
        }

        public static bool IsSame<T>(OpenXmlSimpleValue<T> val, OpenXmlSimpleValue<T> other) where T : struct
        {
            if (val == null && other == null)
                return true;
            if ((val == null && other != null) || (val != null && other == null))
                return false;

            if (val.HasValue != other.HasValue)
                return false;
            if (val.HasValue)
            {
                return val.Value.Equals(other.Value);
            }
            return true;
        }

        public static bool IsSame<T>(EnumValue<T> val, EnumValue<T> other) where T : struct
        {
            if (val == null && other == null)
                return true;
            if ((val == null && other != null) || (val != null && other == null))
                return false;

            if (val.HasValue != other.HasValue)
                return false;
            if (val.HasValue)
            {
                return val.Value.Equals(other.Value);
            }
            return true;
        }

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
        /// Get style sheet of document
        /// </summary>
        /// <param name="doc"></param>
        /// <returns></returns>
        public static Stylesheet GetStyleSheet(this SpreadsheetDocument doc)
        {
            // Get style parts
            var stylePart = doc.WorkbookPart.GetPartsOfType<WorkbookStylesPart>().FirstOrDefault();
            if (stylePart == null)
                throw new InvalidOperationException("StylesPart was not found in document!");
            if (stylePart.Stylesheet == null)
                stylePart.Stylesheet = new Stylesheet();
            return stylePart.Stylesheet;
        }

        /// <summary>
        /// Get cell format
        /// </summary>
        /// <param name="styleSheet"></param>
        /// <param name="cell"></param>
        /// <returns></returns>
        public static CellFormat GetCellFormat(this Stylesheet styleSheet, Cell cell)
        {
            if (cell.StyleIndex == null)
                return null;
            return styleSheet.CellFormats.ElementAt((int)cell.StyleIndex.Value) as CellFormat;
        }

        /// <summary>
        /// Check if two cell format is same
        /// </summary>
        /// <param name="fmt"></param>
        /// <param name="other"></param>
        /// <returns></returns>
        public static bool IsSame(this CellFormat fmt, CellFormat other)
        {
            // Align
            if (!IsSame(fmt.ApplyAlignment, other.ApplyAlignment))
                return false;
            if (fmt.ApplyAlignment != null && fmt.ApplyAlignment.HasValue && fmt.ApplyAlignment.Value)
            {
                if (!IsSame(fmt.Alignment.JustifyLastLine, other.Alignment.JustifyLastLine) ||
                    !IsSame(fmt.Alignment.ShrinkToFit, other.Alignment.ShrinkToFit) ||
                    !IsSame(fmt.Alignment.WrapText, other.Alignment.WrapText) ||
                    !IsSame(fmt.Alignment.Horizontal, other.Alignment.Horizontal) ||
                    !IsSame(fmt.Alignment.Indent, other.Alignment.Indent) ||
                    !IsSame(fmt.Alignment.MergeCell, other.Alignment.MergeCell) ||
                    !IsSame(fmt.Alignment.ReadingOrder, other.Alignment.ReadingOrder) ||
                    !IsSame(fmt.Alignment.RelativeIndent, other.Alignment.RelativeIndent) ||
                    !IsSame(fmt.Alignment.TextRotation, other.Alignment.TextRotation) ||
                    !IsSame(fmt.Alignment.WrapText, other.Alignment.WrapText))
                    return false;
            }
            // Border
            if (!IsSame(fmt.ApplyBorder, other.ApplyBorder))
                return false;
            if (fmt.ApplyBorder != null && fmt.ApplyBorder.HasValue && fmt.ApplyBorder.Value)
            {
                if (!IsSame(fmt.BorderId, other.BorderId))
                    return false;
            }
            // Fill
            if (!IsSame(fmt.ApplyFill, other.ApplyFill))
                return false;
            if (fmt.ApplyFill != null && fmt.ApplyFill.HasValue && fmt.ApplyFill.Value)
            {
                if (!IsSame(fmt.FillId, other.FillId))
                    return false;
            }
            // Font
            if (!IsSame(fmt.ApplyFont, other.ApplyFont))
                return false;
            if (fmt.ApplyFont != null && fmt.ApplyFont.HasValue && fmt.ApplyFont.Value)
            {
                if (!IsSame(fmt.FontId, other.FontId))
                    return false;
            }
            // Number
            if (!IsSame(fmt.ApplyNumberFormat, other.ApplyNumberFormat))
                return false;
            if (fmt.ApplyNumberFormat != null && fmt.ApplyNumberFormat.HasValue && fmt.ApplyNumberFormat.Value)
            {
                if (!IsSame(fmt.NumberFormatId, other.NumberFormatId))
                    return false;
            }
            // Protection
            if (!IsSame(fmt.ApplyProtection, (other.ApplyProtection)))
                return false;
            if (fmt.ApplyProtection != null && fmt.ApplyProtection.HasValue && fmt.ApplyProtection.Value)
            {
                if (!IsSame(fmt.Protection.Hidden, other.Protection.Hidden) ||
                    !IsSame(fmt.Protection.Locked, other.Protection.Locked))
                    return false;
            }

            return true;
        }

        /// <summary>
        /// Format a cell as number
        /// </summary>
        /// <param name="styleSheet">Stylesheet</param>
        /// <param name="cell">Cell to be formatted</param>
        /// <param name="numberFmt">Number format</param>
        /// <returns>Style index</returns>
        public static uint FormatCellAsNumber(this Stylesheet styleSheet, Cell cell, String numberFmt)
        {
            // Create a new format for cell
            var oldFmt = styleSheet.GetCellFormat(cell);
            var newFmt = oldFmt == null ? new CellFormat() : (CellFormat)oldFmt.Clone();
            // Try to locate the number format
            var numFmtId = BuiltInNumberFormats.Find(numberFmt);
            if (numFmtId == 0 && !numberFmt.Equals("General", StringComparison.OrdinalIgnoreCase))
            {   // Format not found in built-in list, try style sheet
                var numFmt =
                    styleSheet.NumberingFormats.Descendants<NumberingFormat>()
                        .FirstOrDefault(a => a.FormatCode == numberFmt);
                if (numFmt != null)
                    numFmtId = numFmt.NumberFormatId.Value;
                else
                {   // Need to create new numbering format
                    var maxId =
                        styleSheet.NumberingFormats.Descendants<NumberingFormat>().Max(a => a.NumberFormatId.Value);
                    if (maxId == 0) maxId = 168;
                    numFmt = styleSheet.NumberingFormats.AppendChild(new NumberingFormat
                    {
                        NumberFormatId = new UInt32Value(maxId + 1),
                        FormatCode = numberFmt
                    });
                    numFmtId = numFmt.NumberFormatId.Value;
                }
            }
            newFmt.NumberFormatId = numFmtId;
            newFmt.ApplyNumberFormat = true;
            // Try to find the cell format
            var cnt = styleSheet.CellFormats.Count.Value;
            int found = -1;
            for (uint i = 0; i < cnt; i++)
            {
                var existFmt = styleSheet.CellFormats.ElementAt((int)i) as CellFormat;
                if (existFmt == null) continue;
                if (existFmt.IsSame(newFmt))
                {
                    found = (int)i;
                    break;
                }
            }
            if (found < 0)
            {   // No exist cell format found, create new one
                styleSheet.CellFormats.Count++;
                styleSheet.CellFormats.AppendChild(newFmt);
                cell.StyleIndex = new UInt32Value((uint)styleSheet.CellFormats.Count - 1);
            }
            else
            {
                cell.StyleIndex = new UInt32Value((uint)found);
            }
            return cell.StyleIndex.Value;
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
        /// <param name="collectionRange"></param>
        /// <param name="rangeFirst">Range in configuration</param>
        /// <param name="idx">Index of entity</param>
        /// <param name="isVertical">Whether entities arranged vertically</param>
        /// <returns></returns>
        public static String CalculateEntityRange(ExcelOpenXMLRange collectionRange, String rangeFirst, int idx, bool isVertical = true, ExcelOpenXMLRange endBefore = null)
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

        /// <summary>
        /// Get previous row of sheet data
        /// </summary>
        /// <param name="sheetData"></param>
        /// <param name="rowNumber"></param>
        /// <returns></returns>
        public static Row GetPrevRow(this SheetData sheetData, int rowNumber)
        {
            for (int i = rowNumber - 1; i > 0; i--)
            {
                var prevRow = sheetData.Descendants<Row>().FirstOrDefault(a => a.RowIndex == i);
                if (prevRow != null)
                    return prevRow;
            }
            return null;
        }

        /// <summary>
        /// Get next row of sheet data
        /// </summary>
        /// <param name="sheetData"></param>
        /// <param name="rowNumber"></param>
        /// <returns></returns>
        public static Row GetNextRow(this SheetData sheetData, int rowNumber)
        {
            int i = rowNumber + 1;
            int cnt = sheetData.Descendants<Row>().Count();
            while (i <= cnt)
            {
                var nextRow = sheetData.Descendants<Row>().FirstOrDefault(a => a.RowIndex == i);
                if (nextRow != null)
                    return nextRow;
                i++;
            }
            return null;
        }

        /// <summary>
        /// Get previous cell of row
        /// </summary>
        /// <param name="row"></param>
        /// <param name="cell"></param>
        /// <returns></returns>
        public static Cell GetPrevCell(this Row row, CellReference cell)
        {
            if (cell.Col <= 1)
                return null; // This is left most one
            cell = cell.Offset(-1, 0);
            while (cell.Col > 0)
            {
                var prevCell = row.Descendants<Cell>().FirstOrDefault(a => a.CellReference == cell.ToString());
                if (prevCell != null)
                    return prevCell;
                if (cell.Col == 1)
                    break;
                cell = cell.Offset(-1, 0);
            }
            return null;
        }

        /// <summary>
        /// Get next cell of row
        /// </summary>
        /// <param name="row"></param>
        /// <param name="cell"></param>
        /// <returns></returns>
        public static Cell GetNextCell(this Row row, CellReference cell)
        {
            cell = cell.Offset(1, 0);
            var cnt = row.Descendants<Cell>().Count();
            while (cell.Col <= cnt)
            {
                var nextCell = row.Descendants<Cell>().FirstOrDefault(a => a.CellReference == cell.ToString());
                if (nextCell != null)
                    return nextCell;
                cell = cell.Offset(1, 0);
            }
            return null;
        }

        /// <summary>
        /// Create new cell
        /// </summary>
        /// <param name="sheet">Worksheet</param>
        /// <param name="cellRef">Cell address</param>
        /// <param name="dataType"></param>
        /// <param name="useRowStyle">Try to use style of previous row</param>
        /// <returns></returns>
        public static Cell CreateCell(this Worksheet sheet, CellReference cellRef, EnumValue<CellValues> dataType, bool useRowStyle = true)
        {
            var sheetData = sheet.Descendants<SheetData>().FirstOrDefault();
            if (sheetData == null)
            {
                sheetData = new SheetData();
                sheet.AppendChild(sheetData);
            }
            var row = sheetData.Descendants<Row>().FirstOrDefault(a => a.RowIndex == cellRef.Row);
            if (row == null)
            {
                row = new Row
                {
                    RowIndex = new UInt32Value((uint)cellRef.Row),
                    Spans = new ListValue<StringValue>(new[] { new StringValue(sheet.GetDefaultRowSpan()) })
                };
                var prevRow = sheetData.GetPrevRow(cellRef.Row);
                if (prevRow == null)
                    sheetData.InsertAt(row, 0);
                else
                    sheetData.InsertAfter(row, prevRow);
            }
            var cell = new Cell
            {
                CellReference = cellRef.ToString(),
            };
            // Handle cell format
            if (useRowStyle && cellRef.Row > 1)
            {   // Use previous row style
                var prevRow = sheetData.Descendants<Row>().FirstOrDefault(a => a.RowIndex == cellRef.Row - 1);
                if (prevRow != null)
                {
                    var prevCellRef = new CellReference(cellRef.Col, (int)prevRow.RowIndex.Value);
                    var styleCell =
                        prevRow.Descendants<Cell>().FirstOrDefault(a => a.CellReference == prevCellRef.ToString());
                    if (styleCell != null && styleCell.StyleIndex != null)
                    {
                        cell.StyleIndex = new UInt32Value(styleCell.StyleIndex);
                    }
                }
            }
            else if (cellRef.Col > 1)
            {   // Use previous column style
                var prevCellRef = new CellReference(cellRef.Col - 1, cellRef.Row);
                var styleCell =
                    row.Descendants<Cell>().FirstOrDefault(a => a.CellReference == prevCellRef.ToString());
                if (styleCell != null && styleCell.StyleIndex != null)
                {
                    cell.StyleIndex = new UInt32Value(styleCell.StyleIndex);
                }
            }
            if (dataType.Value == CellValues.Date)
            {   // Handle date
                if (cell.StyleIndex == null)
                {   // Cell has no style, try to format it with default date format.
                    var wbPart = (WorkbookPart)sheet.WorksheetPart.GetParentParts().First();
                    wbPart.WorkbookStylesPart.Stylesheet.FormatCellAsNumber(cell,
                        CultureInfo.CurrentCulture.DateTimeFormat.ShortDatePattern);
                }
                //cell.DataType = new EnumValue<CellValues>(CellValues.Number);
            }
            else
            {
                cell.DataType = dataType;
            }
            var prevCell = row.GetPrevCell(cellRef);
            if (prevCell == null)
                row.InsertAt(cell, 0);
            else
                row.InsertAfter(cell, prevCell);
            return cell;
        }

        /// <summary>
        /// Get default row span for sheet
        /// </summary>
        /// <param name="sheet"></param>
        /// <returns></returns>
        public static String GetDefaultRowSpan(this Worksheet sheet)
        {
            // Get columns
            var columns = sheet.Descendants<Columns>().FirstOrDefault();
            String span = "1:1";
            if (columns != null)
            {
                span = String.Format("1:{0}", columns.Count());
            }
            return span;
        }

        /// <summary>
        /// Insert rows into worksheet
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="cellRef"></param>
        /// <param name="count"></param>
        public static void InsertRow(this Worksheet sheet, CellReference cellRef, uint count = 1)
        {
            var sheetData = sheet.Descendants<SheetData>().FirstOrDefault();
            if (sheetData == null)
            {
                sheetData = new SheetData();
                sheet.AppendChild(sheetData);
            }
            var row = sheetData.Descendants<Row>().FirstOrDefault(a => a.RowIndex == cellRef.Row) ??
                      sheetData.GetNextRow(cellRef.Row);
            if (row == null)
            {   // Append to end
                var idx = (uint)cellRef.Row;
                for (uint i = 0; i < count; i++)
                {
                    var newRow = new Row
                    {
                        RowIndex = new UInt32Value(idx + i),
                        Spans = new ListValue<StringValue>(new[] { new StringValue(sheet.GetDefaultRowSpan()) })
                    };
                    sheetData.AppendChild(newRow);
                }
            }
            else
            {   // Insert before the row
                foreach (var rowBelow in sheetData.Descendants<Row>().Where(a => a.RowIndex >= cellRef.Row))
                {   // Update row's index below
                    rowBelow.RowIndex = new UInt32Value(rowBelow.RowIndex.Value + count);
                    foreach (var cell in rowBelow.Descendants<Cell>())
                    {
                        var cf = new CellReference(cell.CellReference.Value);
                        cf.Move(0, (int)count);
                        cell.CellReference = new StringValue(cf.ToString());
                    }
                }
                // Insert new rows
                var idx = (uint)cellRef.Row;
                for (uint i = 0; i < count; i++)
                {
                    var newRow = new Row
                    {
                        RowIndex = new UInt32Value(idx + i),
                        Spans = new ListValue<StringValue>(new[] { new StringValue(sheet.GetDefaultRowSpan()) })
                    };
                    sheetData.InsertBefore(newRow, row);
                }
            }

            // Update named ranges
            #region NAMEDRANGES
            var regex = new Regex(REG_REF);
            var wbPart = (WorkbookPart)sheet.WorksheetPart.GetParentParts().First();
            var nameToUpdate = new List<KeyValuePair<DefinedName, String>>();
            var definedNames = wbPart.Workbook.Descendants<DefinedNames>().FirstOrDefault();
            if (definedNames != null)
            {
                foreach (var definedName in definedNames.Descendants<DefinedName>())
                {
                    var addr = definedName.InnerText;
                    var match = regex.Match(addr);
                    if (match.Success)
                    {
                        var sheetName = match.Groups["Sheet"].Value;
                        if (sheetName.StartsWith("'") && sheetName.EndsWith("'"))
                            sheetName = sheetName.Substring(1, sheetName.Length - 2);
                        var sheetObj =
                            wbPart.Workbook.Descendants<Sheet>().FirstOrDefault(a => a.Name == sheetName);
                        if (sheetObj != null && sheet.WorksheetPart == wbPart.GetPartById(sheetObj.Id))
                        {   // This name is in my sheet
                            var rangeRef = new RangeReference(match.Groups["Range"].Value);
                            if (rangeRef.Top <= cellRef.Row && cellRef.Row <= rangeRef.Bottom)
                            {   // This range is crossing insert point, expand its range
                                if (rangeRef.Width == 1 && rangeRef.Height == 1)
                                {   // Is a cell, move it
                                    rangeRef.Move(0, (int)count);
                                }
                                else
                                {   // Is a range, expand it
                                    rangeRef.Expand(0, 0, 0, (int)count);
                                }
                                var newRef = String.Format("{0}!{1}", match.Groups["Sheet"].Value, rangeRef.AsAbsolute());
                                nameToUpdate.Add(new KeyValuePair<DefinedName, string>(definedName, newRef));
                            }
                        }
                    }
                }
                foreach (var pair in nameToUpdate)
                {
                    definedNames.ReplaceChild(new DefinedName(pair.Value) { Name = pair.Key.Name }, pair.Key);
                }
            }
            #endregion

            // Update formulars
            #region FORMULAR
            var calChainPart = wbPart.CalculationChainPart;
            if (calChainPart != null)
            {
                foreach (var calCell in calChainPart.CalculationChain.Descendants<CalculationCell>())
                {
                    var sheetObj =
                        wbPart.Workbook.Descendants<Sheet>().FirstOrDefault(a => a.SheetId - calCell.SheetId == 0);
                    if (sheetObj != null && sheet.WorksheetPart == wbPart.GetPartById(sheetObj.Id))
                    {   // Calculation cell is in my sheet
                        var calRef = new CellReference(calCell.CellReference);
                        if (calRef.Row >= cellRef.Row)
                        {
                            calRef.Move(0, (int)count);
                            calCell.CellReference = new StringValue(calRef.ToString());
                            // Process formular
                            var formularCell =
                                sheetData.Descendants<Cell>().FirstOrDefault(a => a.CellReference == calRef.ToString());
                            if (formularCell != null && formularCell.CellFormula != null)
                            {   // Try to update ranges in formular
                                var reg = new Regex(REG_RANGE_REF);
                                var formular = formularCell.CellFormula.InnerText;
                                var match = reg.Match(formular);
                                while (match.Success)
                                {
                                    var rangeRef = new RangeReference(match.Value);
                                    int startPos = match.Index + match.Length;
                                    if (rangeRef.Top <= cellRef.Row && cellRef.Row <= rangeRef.Bottom + 1)
                                    {   // This range is crossing or adjacent to insert point, expand its range
                                        rangeRef.Expand(0, 0, 0, (int)count);
                                        var newRangeRef = rangeRef.ToString();
                                        formular = formular.Remove(match.Index, match.Length)
                                            .Insert(match.Index, newRangeRef);
                                        startPos = match.Index + newRangeRef.Length;
                                    }
                                    match = reg.Match(formular, startPos);
                                }
                                formularCell.CellFormula.Text = formular;
                                formularCell.CellValue = null; // Remove value of formular, so excel will calculate it next time.
                            }
                        }
                    }
                }
            }
            #endregion

            // Update Dimension
            #region DIMENSION
            if (sheet.SheetDimension != null)
            {
                var dimRef = new RangeReference(sheet.SheetDimension.Reference);
                dimRef.Expand(0, 0, 0, (int)count);
                sheet.SheetDimension.Reference = new StringValue(dimRef.ToString());
            }
            #endregion
        }

        /// <summary>
        /// Insert columns into worksheet
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="cellRef"></param>
        /// <param name="count"></param>
        public static void InsertColumn(this Worksheet sheet, CellReference cellRef, uint count = 1)
        {
            var sheetData = sheet.Descendants<SheetData>().FirstOrDefault();
            var columns = sheet.Descendants<Columns>().FirstOrDefault() ?? sheet.AppendChild(new Columns());
            if (cellRef.Col > columns.Count())
            {   // Append new column
                for (uint i = 0; i < count; i++)
                    columns.AppendChild(new Column());
            }
            else
            {   // Insert new column
                for (uint i = 0; i < count; i++)
                    columns.InsertAt(new Column(), cellRef.Col - 1);
                // Update all cells that is at right of insert point
                if (sheetData != null)
                {
                    foreach (var cell in sheetData.Descendants<Cell>())
                    {
                        var cf = new CellReference(cell.CellReference.Value);
                        if (cf.Col >= cellRef.Col)
                        {
                            cf.Move((int)count, 0);
                            cell.CellValue = new CellValue(cf.ToString());
                        }
                    }
                }
            }
            // Update named ranges
            #region NAMEDRANGES
            var regex = new Regex(REG_REF);
            var wbPart = (WorkbookPart)sheet.WorksheetPart.GetParentParts().First();
            var nameToUpdate = new List<KeyValuePair<DefinedName, String>>();
            var definedNames = wbPart.Workbook.Descendants<DefinedNames>().FirstOrDefault();
            if (definedNames != null)
            {
                foreach (var definedName in definedNames.Descendants<DefinedName>())
                {
                    var addr = definedName.InnerText;
                    var match = regex.Match(addr);
                    if (match.Success)
                    {
                        var sheetName = match.Groups["Sheet"].Value;
                        if (sheetName.StartsWith("'") && sheetName.EndsWith("'"))
                            sheetName = sheetName.Substring(1, sheetName.Length - 2);
                        var sheetObj =
                            wbPart.Workbook.Descendants<Sheet>().FirstOrDefault(a => a.Name == sheetName);
                        if (sheetObj != null && sheet.WorksheetPart == wbPart.GetPartById(sheetObj.Id))
                        {   // This name is in my sheet
                            var rangeRef = new RangeReference(match.Groups["Range"].Value);
                            if (rangeRef.Top <= cellRef.Row && cellRef.Row <= rangeRef.Bottom)
                            {   // This range is crossing insert point, expand its range
                                if (rangeRef.Width == 1 && rangeRef.Height == 1)
                                {   // Is a cell, move it
                                    rangeRef.Move((int)count, 0);
                                }
                                else
                                {   // Is a range, expand it
                                    rangeRef.Expand(0, (int)count, 0, 0);
                                }
                                var newRef = String.Format("{0}!{1}", match.Groups["Sheet"].Value, rangeRef.AsAbsolute());
                                nameToUpdate.Add(new KeyValuePair<DefinedName, string>(definedName, newRef));
                            }
                        }
                    }
                }
                foreach (var pair in nameToUpdate)
                {
                    definedNames.ReplaceChild(new DefinedName(pair.Value) { Name = pair.Key.Name }, pair.Key);
                }
            }
            #endregion

            // Update formulars
            #region FORMULARS
            var calChainPart = wbPart.CalculationChainPart;
            if (calChainPart != null && sheetData != null)
            {
                foreach (var calCell in calChainPart.CalculationChain.Descendants<CalculationCell>())
                {
                    var sheetObj =
                        wbPart.Workbook.Descendants<Sheet>().FirstOrDefault(a => a.SheetId - calCell.SheetId == 0);
                    if (sheetObj != null && sheet.WorksheetPart == wbPart.GetPartById(sheetObj.Id))
                    {   // Calculation cell is in my sheet
                        var calRef = new CellReference(calCell.CellReference);
                        if (calRef.Col >= cellRef.Col)
                        {
                            calRef.Move(0, (int)count);
                            calCell.CellReference = new StringValue(calRef.ToString());
                            // Process formular
                            var formularCell =
                                sheetData.Descendants<Cell>().FirstOrDefault(a => a.CellReference == calRef.ToString());
                            if (formularCell != null && formularCell.CellFormula != null)
                            {   // Try to update ranges in formular
                                var reg = new Regex(REG_RANGE_REF);
                                var formular = formularCell.CellFormula.InnerText;
                                var match = reg.Match(formular);
                                while (match.Success)
                                {
                                    var rangeRef = new RangeReference(match.Value);
                                    int startPos = match.Index + match.Length;
                                    if (rangeRef.Left <= cellRef.Col && cellRef.Col <= rangeRef.Right + 1)
                                    {   // This range is crossing or adjacent to insert point, expand its range
                                        rangeRef.Expand(0, (int)count, 0, 0);
                                        var newRangeRef = rangeRef.ToString();
                                        formular = formular.Remove(match.Index, match.Length)
                                            .Insert(match.Index, newRangeRef);
                                        startPos = match.Index + newRangeRef.Length;
                                    }
                                    match = reg.Match(formular, startPos);
                                }
                                formularCell.CellFormula.Text = formular;
                                formularCell.CellValue = null; // Remove value of formular, so excel will calculate it next time.
                            }
                        }
                    }
                }
            }
            #endregion

            // Update Dimension
            #region DIMENSION
            if (sheet.SheetDimension != null)
            {
                var dimRef = new RangeReference(sheet.SheetDimension.Reference);
                dimRef.Expand(0, (int)count, 0, 0);
                sheet.SheetDimension.Reference = new StringValue(dimRef.ToString());
            }
            #endregion
        }
    }
}
