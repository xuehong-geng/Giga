using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Giga.Transformer.Excel
{
    /// <summary>
    /// Tool used to convert data between number and AA expression
    /// </summary>
    public class Alph26
    {
        private static readonly string[] Letters = { "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z" };

        /// <summary>
        /// Convert number to AA expression
        /// </summary>
        /// <param name="value"></param>
        /// <returns></returns>
        public static string N2A(int value)
        {
            value--;
            int remainder = value % 26;
            int front = (value - remainder) / 26;
            //if (remainder == 0 && front > 0) 
            //{
            //    front--;
            //    remainder = 25;
            //}
            if (front == 0)
            {
                return Letters[remainder];
            }
            else
            {
                return N2A(front) + Letters[remainder];
            }
        }

        /// <summary>
        /// Convert AA expression to number
        /// </summary>
        /// <param name="exp"></param>
        /// <returns></returns>
        public static int A2N(string exp)
        {
            exp = exp.ToUpper();
            var reg = new Regex(@"[A-Z]+");
            var match = reg.Match(exp);
            if (!match.Success)
                throw new ArgumentException(String.Format("Expression {0} is not AA expression!", exp));
            int val = 0;
            for (int i = 0; i < exp.Length; i++)
            {
                char c = exp[exp.Length - i - 1];
                int n = Convert.ToInt32(c) - Convert.ToInt32('A') + 1;
                val += (int)Math.Pow(26, i) * n;
            }
            return val;
        }
    }

    /// <summary>
    /// Tool to encapsulate the cell reference in excel
    /// </summary>
    public class CellReference
    {
        private const String REG_CELL_REF = @"(?i)(?<COL>\$?[a-zA-Z]+)(?<ROW>\$?[1-9][0-9]*)";
        private int _col;
        private int _row;

        public int Col
        {
            get { return _col; }
            set
            {
                if (value < 1)
                    throw new ArgumentException("Column number cannot be less than 1!");
                _col = value;
            }
        }

        public int Row
        {
            get { return _row; }
            set
            {
                if (value < 1)
                    throw new ArgumentException("Row number cannot be less than 1!");
                _row = value;
            }
        }

        public bool IsColumnAbsolute { get; set; }  // Whether column has $
        public bool IsRowAbsolute { get; set; }     // Whether row has $

        public CellReference()
        {
            Col = 1;
            Row = 1;
            IsColumnAbsolute = IsRowAbsolute = false;
        }

        public CellReference(String reference)
        {
            Col = 1;
            Row = 1;
            IsColumnAbsolute = IsRowAbsolute = false;
            Set(reference);
        }

        public CellReference(int col, int row)
        {
            Col = col;
            Row = row;
            IsColumnAbsolute = IsRowAbsolute = false;
        }

        public CellReference(CellReference other)
        {
            Col = other.Col;
            Row = other.Row;
            IsColumnAbsolute = other.IsColumnAbsolute;
            IsRowAbsolute = other.IsRowAbsolute;
        }

        public String ColumnName
        {
            get
            {
                return Alph26.N2A(Col);
            }
            set
            {
                String c = value.ToUpper();
                if (c.StartsWith("$"))
                {
                    c = c.Remove(0, 1);
                    IsColumnAbsolute = true;
                }
                Col = Alph26.A2N(c);
            }
        }

        public override string ToString()
        {
            return string.Format("{0}{1}{2}{3}", 
                (IsColumnAbsolute ? "$" : ""), 
                ColumnName, 
                (IsRowAbsolute ? "$" : ""),
                Row);
        }

        public void Set(String reference)
        {
            var reg = new Regex(REG_CELL_REF);
            var match = reg.Match(reference);
            if (!match.Success)
                throw new InvalidDataException(String.Format("Cell reference {0} is invalid!", reference));
            var col = match.Groups["COL"].Value;
            var row = match.Groups["ROW"].Value;
            ColumnName = col;
            if (row.StartsWith("$"))
            {
                row = row.Remove(0, 1);
                IsRowAbsolute = true;
            }
            Row = int.Parse(row);
        }

        public void Move(int x, int y)
        {
            int newCol = Col + x;
            int newRow = Row + y;
            if (newCol < 1)
                throw new InvalidOperationException("Cannot move column reference to left of 'A'!");
            if (newRow < 1)
                throw new InvalidOperationException("Cannot move row reference to top of 1!");
            Col = newCol;
            Row = newRow;
        }

        public void Move(String relativeRef)
        {
            var rel = new CellReference(relativeRef);
            if (!rel.IsColumnAbsolute && !rel.IsRowAbsolute)
            {
                Move(rel.Col - 1, rel.Row - 1);
            }
            else if (rel.IsColumnAbsolute && !rel.IsRowAbsolute)
            {
                Col = rel.Col;
                Move(0, rel.Row - 1);
            }
            else if (!rel.IsColumnAbsolute && rel.IsRowAbsolute)
            {
                Row = rel.Row;
                Move(rel.Col - 1, 0);
            }
            else
            {
                Col = rel.Col;
                Row = rel.Row;
            }
        }

        public CellReference Offset(int x, int y)
        {
            var newCell = new CellReference(Col, Row);
            newCell.Move(x, y);
            return newCell;
        }

        public CellReference Offset(String relativeRef)
        {
            var newCell = new CellReference(Col, Row);
            newCell.Move(relativeRef);
            return newCell;
        }
    }

    /// <summary>
    /// Tool to encapsulate the range reference in excel
    /// </summary>
    public class RangeReference
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

        protected CellReference _topLeft = null;
        protected CellReference _bottomRight = null;

        public RangeReference()
        {
            _topLeft = new CellReference();
            _bottomRight = new CellReference();
        }

        public RangeReference(String reference)
        {
            Set(reference);
        }

        public RangeReference(CellReference topLeft, CellReference bottomRight)
        {
            int l = topLeft.Col;
            int r = bottomRight.Col;
            int t = topLeft.Row;
            int b = bottomRight.Row;
            if (l > r) Swap(ref l, ref r);
            if (t > b) Swap(ref t, ref b);
            _topLeft = new CellReference(l, t);
            _bottomRight = new CellReference(r, b);
        }

        public RangeReference(RangeReference r)
        {
            _topLeft = new CellReference(r._topLeft);
            _bottomRight = new CellReference(r._bottomRight);
        }

        /// <summary>
        /// Parse range from string expression
        /// </summary>
        /// <param name="reference"></param>
        /// <param name="topLeft"></param>
        /// <param name="bottomRight"></param>
        /// <param name="boundary">Boundary range used when the range reference is open</param>
        public static void ParseRange(String reference, ref CellReference topLeft, ref CellReference bottomRight, RangeReference boundary = null)
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
                if (boundary != null)
                {   // Expand to boundary if range is open
                    if (String.IsNullOrEmpty(col1)) col1 = boundary._topLeft.ColumnName;
                    if (String.IsNullOrEmpty(col2)) col2 = boundary._bottomRight.ColumnName;
                    if (String.IsNullOrEmpty(row1)) row1 = boundary._topLeft.Row.ToString(CultureInfo.InvariantCulture);
                    if (String.IsNullOrEmpty(row2)) row2 = boundary._bottomRight.Row.ToString(CultureInfo.InvariantCulture);
                }
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
        /// Set reference of range.
        /// </summary>
        /// <param name="reference"></param>
        public void Set(String reference)
        {
            ParseRange(reference, ref _topLeft, ref _bottomRight);
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
        /// Move the range
        /// </summary>
        /// <param name="x">Delta X</param>
        /// <param name="y">Delta Y</param>
        public void Move(int x, int y)
        {
            _topLeft.Move(x, y);
            _bottomRight.Move(x, y);
        }

        /// <summary>
        /// Convert range reference to string expression
        /// </summary>
        /// <returns></returns>
        public override string ToString()
        {
            return String.Format("{0}:{1}", _topLeft, _bottomRight);
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
        /// Calculate a reference of new cell that is relative to the top left corner of range.
        /// </summary>
        /// <param name="col">Column offset</param>
        /// <param name="row">Row offset</param>
        /// <param name="throwIfNotInRange">Throw exception if the target cell is out of range</param>
        /// <returns></returns>
        public CellReference CalculateCellReference(int col, int row, bool throwIfNotInRange = true)
        {
            CellReference cell = _topLeft.Offset(col - 1, row - 1);
            if (!IsInRange(cell) && throwIfNotInRange)
                throw new ArgumentException("Try to access cell that is out of range!");
            return cell;
        }
        /// <summary>
        /// Calculate a reference of new cell that is relative to the top left corner of range.
        /// </summary>
        /// <param name="relativeRef">Relative reference</param>
        /// <param name="throwIfNotInRange">Throw exception if the target cell is out of range</param>
        /// <returns></returns>
        public CellReference CalculateCellReference(String relativeRef, bool throwIfNotInRange = true)
        {
            CellReference cell = _topLeft.Offset(relativeRef);
            if (!IsInRange(cell) && throwIfNotInRange)
                throw new ArgumentException("Try to access cell that is out of range!");
            return cell;
        }

        /// <summary>
        /// Get a sub range reference by using relative range descriptor
        /// </summary>
        /// <param name="relativeRange">Relative range descriptor</param>
        /// <param name="clipToRange">Whether to clip to parent range</param>
        /// <returns>Sub range</returns>
        /// <remarks>
        /// The relative range looks as same as normal range. For example, A1:B2 represent 2x2 cells 
        /// start from top left corner of parent range.
        /// </remarks>
        public RangeReference SubRange(String relativeRange, bool clipToRange = true)
        {
            var tl = new CellReference();
            var br = new CellReference();
            ParseRange(relativeRange, ref tl, ref br, this);
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
            return new RangeReference(subTL, subBR);
        }
    }
}
