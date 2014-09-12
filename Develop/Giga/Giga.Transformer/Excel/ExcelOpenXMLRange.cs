using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Office2010.PowerPoint;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Giga.Transformer.Configuration;

namespace Giga.Transformer.Excel
{
    /// <summary>
    /// Exception : Cell not exist
    /// </summary>
    public class CellNotExistException : ApplicationException
    {
        public CellReference Reference { get; set; }

        public CellNotExistException(CellReference reference)
        {
            Reference = reference;
        }

        public override string Message
        {
            get { return String.Format("Cell {0} not exist!", Reference); }
        }
    }

    /// <summary>
    /// Simulating Range in Excel file with OpenXML SDK
    /// </summary>
    public class ExcelOpenXMLRange
    {
        public static void Swap<T>(ref T t1, ref T t2)
        {
            T tmp = t1;
            t1 = t2;
            t2 = tmp;
        }

        public const String REG_CELL_REF = @"(?i)(?<COL>\$?[a-zA-Z]+)(?<ROW>\$?[1-9][0-9]*)";
        public const String REG_RANGE_REF = @"(?i)(?<COL1>\$?[a-zA-Z]+)(?<ROW1>\$?[1-9][0-9]*)\:(?<COL2>\$?[a-zA-Z]+)(?<ROW2>\$?[1-9][0-9]*)";
        public const String REG_RANGE_PART_REF = @"(?i)((?<COL1>\$?[a-zA-Z]+))?((?<ROW1>\$?[1-9][0-9]*))?\:((?<COL2>\$?[a-zA-Z]+))?((?<ROW2>\$?[1-9][0-9]*))?";
        public const String REG_ANCHOR_CELL = @"(?i)(?<Anchor>.+)#(?<OffsetX>\d+),(?<OffsetY>\d+)";

        public static Regex _RegexRange = new Regex(REG_RANGE_PART_REF);
        public static Regex _RegexCell = new Regex(REG_CELL_REF);
        public static Regex _RegexAnchor = new Regex(REG_ANCHOR_CELL);

        protected SpreadsheetDocument _doc = null;
        protected Worksheet _sheet = null;
        protected String _reference = null;
        protected CellReference _topLeft = null;
        protected CellReference _bottomRight = null;

        /// <summary>
        /// Base excel document
        /// </summary>
        public SpreadsheetDocument Document
        {
            get { return _doc; }
        }
        /// <summary>
        /// Base worksheet
        /// </summary>
        public Worksheet Sheet
        {
            get { return _sheet; }
        }

        /// <summary>
        /// Initialize a range
        /// </summary>
        /// <param name="doc">Excel document</param>
        /// <param name="sheet">Worksheet</param>
        /// <param name="reference">Reference in work sheet</param>
        public ExcelOpenXMLRange(SpreadsheetDocument doc, Worksheet sheet, String reference)
        {
            _doc = doc;
            _sheet = sheet;
            _reference = reference;
            AnalyzeRange();
        }

        /// <summary>
        /// Initialize a range with top left and buttom right corner cells
        /// </summary>
        /// <param name="doc">Excel document</param>
        /// <param name="sheet">Worksheet</param>
        /// <param name="tl">Cell reference for top left corner</param>
        /// <param name="br">Cell reference for bottom right corner</param>
        public ExcelOpenXMLRange(SpreadsheetDocument doc, Worksheet sheet, CellReference tl, CellReference br)
        {
            _doc = doc;
            _sheet = sheet;
            _reference = String.Format("{0}:{1}", tl, br);
            int l = tl.Col;
            int r = br.Col;
            int t = tl.Row;
            int b = br.Row;
            if (l > r) Swap(ref l, ref r);
            if (t > b) Swap(ref t, ref b);
            _topLeft = new CellReference(l, t);
            _bottomRight = new CellReference(r, b);
        }


        /// <summary>
        /// Calculate range from string expression
        /// </summary>
        /// <param name="reference"></param>
        /// <param name="topLeft"></param>
        /// <param name="bottomRight"></param>
        public static void CalculateRange(String reference, ref CellReference topLeft, ref CellReference bottomRight)
        {
            var matchRange = _RegexRange.Match(reference);

            if (topLeft == null) topLeft = new CellReference();
            if (bottomRight == null) bottomRight = new CellReference();

            if (matchRange.Success)
            {   // It's a range
                var col1 = matchRange.Groups["COL1"].Value;
                var col2 = matchRange.Groups["COL2"].Value;
                var row1 = matchRange.Groups["ROW1"].Value;
                var row2 = matchRange.Groups["ROW2"].Value;
                var cell1 = new CellReference(col1 + row1);
                var cell2 = new CellReference(col2 + row2);
                int left = cell1.Col <= cell2.Col ? cell1.Col : cell2.Col;
                int right = left == cell1.Col ? cell2.Col : cell1.Col;
                int top = cell1.Row <= cell2.Row ? cell1.Row : cell2.Row;
                int bottom = top == cell1.Row ? cell2.Row : cell1.Row;
                topLeft.Col = left;
                topLeft.Row = top;
                bottomRight.Col = right;
                bottomRight.Row = bottom;
            }
            else
            {
                var matchCell = _RegexCell.Match(reference);
                if (matchCell.Success && matchCell.Length == reference.Trim().Length)
                {   // It's a cell
                    topLeft.Set(reference);
                    bottomRight.Set(reference);
                }
                else
                {
                    throw new InvalidDataException(String.Format("Range reference {0} is invalid!", reference));
                }
            }
        }

        /// <summary>
        /// Analyze range data
        /// </summary>
        protected void AnalyzeRange()
        {
            // Since the range might be open (ie. at least one direction is un limited, such as A:J, A1:J), we should
            // fix it by using sheet's diamention
            _reference = _sheet.ExpandToSheetBound(_reference);
            CalculateRange(_reference, ref _topLeft, ref _bottomRight);
        }

        /// <summary>
        /// Top cell of range
        /// </summary>
        public int Top
        {
            get { return _topLeft.Row; }
        }
        /// <summary>
        /// Left cell of range
        /// </summary>
        public int Left
        {
            get { return _topLeft.Col; }
        }
        /// <summary>
        /// Height of range
        /// </summary>
        public int Height
        {
            get { return _bottomRight.Row - _topLeft.Row + 1; }
        }
        /// <summary>
        /// Width of range
        /// </summary>
        public int Width
        {
            get { return _bottomRight.Col - _topLeft.Col + 1; }
        }

        /// <summary>
        /// Check if a cell reference is hitting in the range
        /// </summary>
        /// <param name="cell">Cell reference</param>
        /// <returns></returns>
        public bool IsInRange(CellReference cell)
        {
            if (cell.Col >= _topLeft.Col && cell.Col <= _bottomRight.Col &&
                cell.Row >= _topLeft.Row && cell.Row <= _bottomRight.Row)
                return true;
            else
                return false;
        }

        /// <summary>
        /// Create a sub range by using relative range descriptor
        /// </summary>
        /// <param name="relativeRange">Relative range descriptor</param>
        /// <param name="clipToRange">Whether to clip to parent range</param>
        /// <returns>Sub range</returns>
        /// <remarks>
        /// The relative range looks as same as normal range. For example, A1:B2 represent 2x2 cells 
        /// start from top left corner of parent range.
        /// </remarks>
        public ExcelOpenXMLRange GetSubRange(String relativeRange, bool clipToRange = true)
        {
            var tl = new CellReference();
            var br = new CellReference();
            CalculateRange(relativeRange, ref tl, ref br);
            var subTL = _topLeft.Offset(tl.Col - 1, tl.Row - 1);
            var subBR = subTL.Offset(br.Col - tl.Col, br.Row - tl.Row);
            if (clipToRange)
            {   // Clip to parent range
                if (subTL.Col > _bottomRight.Col ||
                    subTL.Row > _bottomRight.Row ||
                    subBR.Col < _topLeft.Col ||
                    subBR.Row < _topLeft.Row)
                    return null; // Out of range
                if (subTL.Col < _topLeft.Col) subTL.Col = _topLeft.Col;
                if (subTL.Row < _topLeft.Row) subTL.Row = _topLeft.Row;
                if (subBR.Col > _bottomRight.Col) subBR.Col = _bottomRight.Col;
                if (subBR.Row > _bottomRight.Row) subBR.Row = _bottomRight.Row;
            }
            return new ExcelOpenXMLRange(_doc, _sheet, subTL, subBR);
        }

        /// <summary>
        /// Calculate a reference of new cell that is related to the top left corner of range.
        /// </summary>
        /// <param name="col">Column offset</param>
        /// <param name="row">Row offset</param>
        /// <returns></returns>
        protected CellReference CalculateCellReference(int col, int row)
        {
            CellReference cell = _topLeft.Offset(col - 1, row - 1);
            if (!IsInRange(cell))
                throw new ArgumentException("Try to access cell that is out of range!");
            return cell;
        }
        /// <summary>
        /// Calculate reference of new cell with its relative address
        /// </summary>
        /// <param name="relativeRef">Address in A1 expression</param>
        /// <returns></returns>
        protected CellReference CalculateCellReference(String relativeRef)
        {
            var matchAchor = _RegexAnchor.Match(relativeRef);
            if (matchAchor.Success)
            {   // This cell reference is anchored to another cell
                try
                {
                    var anchor = matchAchor.Groups["Anchor"].Value;
                    var offx = int.Parse(matchAchor.Groups["OffsetX"].Value);
                    var offy = int.Parse(matchAchor.Groups["OffsetY"].Value);
                    var anchorCell = CalculateCellReference(anchor);
                    anchorCell.Move(offx, offy);
                    return anchorCell;
                }
                catch (Exception err)
                {
                    throw new InvalidDataException(String.Format("Cannot get reference of Anchored cell {0}! Err:{1}",
                        relativeRef, err.Message));
                }
            }
            String address = null;
            if (!_RegexCell.IsMatch(relativeRef))
            {   // The cell reference might be defined name
                address = _doc.GetDefinedName(relativeRef);
                if (address == null)
                    throw new InvalidDataException(String.Format("Defined name '{0}' is not exist!", relativeRef));
            }
            else
            {
                address = relativeRef;
            }
            CellReference cell = _topLeft.Offset(address);
            if (!IsInRange(cell))
                throw new ArgumentException("Try to access cell that is out of range!");
            return cell;
        }

        /// <summary>
        /// Get value of cell
        /// </summary>
        /// <param name="cellRef">Cell reference</param>
        /// <returns>Cell value</returns>
        public object GetCellValue(CellReference cellRef)
        {
            Cell cell = _sheet.Descendants<Cell>().FirstOrDefault(a => a.CellReference == cellRef.ToString());
            if (cell == null)
                throw new CellNotExistException(cellRef);
            var val = cell.CellValue.InnerText;
            if (cell.DataType != null)
            {   // Check data type
                switch (cell.DataType.Value)
                {
                    case CellValues.Boolean:
                        return val != "0";
                    case CellValues.Date:
                        {   // Convert serialized date to DateTime
                            long n = long.Parse(val);
                            return DateTime.FromBinary(n);
                        }
                    case CellValues.SharedString:
                        {   // Get from shared string table
                            var shareStrTbl = _doc.WorkbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
                            if (shareStrTbl == null)
                                return null;
                            else
                            {
                                return shareStrTbl.SharedStringTable.ElementAt(int.Parse(val)).InnerText;
                            }
                        }
                }
            }
            return val;
        }

        /// <summary>
        /// Cell value indexer
        /// </summary>
        /// <param name="col">Index of column, 0 based</param>
        /// <param name="row">Index of row, 0 based.</param>
        /// <returns></returns>
        public object this[int col, int row]
        {
            get
            {
                CellReference r = CalculateCellReference(col, row);
                return GetCellValue(r);
            }
        }
        /// <summary>
        /// Cell value indexer
        /// </summary>
        /// <param name="relRef">Relative address in A1 expression</param>
        /// <returns></returns>
        public object this[String relRef]
        {
            get
            {
                CellReference r = CalculateCellReference(relRef);
                return GetCellValue(r);
            }
        }

        /// <summary>
        /// Convert value from excel to dotNET standard
        /// </summary>
        /// <param name="val">value read from excel</param>
        /// <param name="tgtType">Target type to be converted to</param>
        /// <returns>Data converted</returns>
        protected object ConvertExcelValue(object val, Type tgtType)
        {
            if (tgtType == val.GetType())
            {
                return val;
            }
            else
            {
                if (tgtType.Name == "Nullable`1")
                {   // The tgtType is Nullable wrapper, must try to set real data
                    var baseType = tgtType.GenericTypeArguments[0];
                    return ConvertExcelValue(val, baseType);
                }
                // Need convert data type
                if (tgtType == typeof(DateTime))
                {
                    // Handle datetime specially
                    return DateTime.FromOADate(Convert.ToDouble(val));
                }
                else
                {
                    return Convert.ChangeType(val, tgtType);
                }
            }
        }

        /// <summary>
        /// Read one entity from range
        /// </summary>
        /// <typeparam name="T">Type of entity</typeparam>
        /// <param name="cfg">Configuration that defines fields of entity</param>
        /// <returns></returns>
        public T ReadEntity<T>(EntityConfigElement cfg) where T : class, new()
        {
            var ent = new T();
            Type t = typeof(T);
            bool entExist = false;
            foreach (FieldConfigElement field in cfg.Fields)
            {
                var pT = t.GetProperty(field.Name);
                if (pT != null)
                {   // Handle this property
                    object cellVal = null;
                    bool cellExist = true;
                    try
                    {
                        cellVal = this[field.Range];
                    }
                    catch (CellNotExistException)
                    {   // Cell not exist
                        cellExist = false;
                    }
                    if (cellExist && cellVal != null)
                    {
                        try
                        {
                            pT.SetValue(ent, ConvertExcelValue(cellVal, pT.PropertyType));
                            entExist = true;
                        }
                        catch (Exception err)
                        {
                            throw new InvalidDataException(String.Format("Cannot set property {0} to {1}! Err:{2}",
                                field.Name, cellVal, err.Message));
                        }
                    }
                }
                // TODO:Handle sub collections here.
            }

            return (cfg.AllowNull || entExist) ? ent : null;
        }
    }
}
