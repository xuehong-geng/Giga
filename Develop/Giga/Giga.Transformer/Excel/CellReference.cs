using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Giga.Transformer.Excel
{
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

        public CellReference()
        {
            Col = 1;
            Row = 1;
        }

        public CellReference(String reference)
        {
            Col = 1;
            Row = 1;
            Set(reference);
        }

        public CellReference(int col, int row)
        {
            Col = col;
            Row = row;
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
                    c = c.Remove(0, 1);
                Col = Alph26.A2N(c);
            }
        }

        public override string ToString()
        {
            return string.Format("{0}{1}", ColumnName, Row);
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
            Move(rel.Col - 1, rel.Row - 1);
        }

        public CellReference Offset(int x, int y)
        {
            int newCol = Col + x;
            int newRow = Row + y;
            if (newCol < 1)
                throw new InvalidOperationException("Cannot move column reference to left of 'A'!");
            if (newRow < 1)
                throw new InvalidOperationException("Cannot move row reference to top of 1!");
            return new CellReference(newCol, newRow);
        }

        public CellReference Offset(String relativeRef)
        {
            var rel = new CellReference(relativeRef);
            return Offset(rel.Col - 1, rel.Row - 1);
        }
    }

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
}
