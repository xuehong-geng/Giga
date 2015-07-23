using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
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
    public class ExcelOpenXMLRange : RangeReference
    {
        protected SpreadsheetDocument _doc = null;
        protected Worksheet _sheet = null;

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
            // Since the range might be open (ie. at least one direction is un limited, such as A:J, A1:J), we should
            // fix it by using sheet's diamention
            reference = _sheet.ExpandToSheetBound(reference);
            Set(reference);
        }

        /// <summary>
        /// Initialize a range with top left and buttom right corner cells
        /// </summary>
        /// <param name="doc">Excel document</param>
        /// <param name="sheet">Worksheet</param>
        /// <param name="tl">Cell reference for top left corner</param>
        /// <param name="br">Cell reference for bottom right corner</param>
        public ExcelOpenXMLRange(SpreadsheetDocument doc, Worksheet sheet, CellReference tl, CellReference br)
            : base(tl, br)
        {
            _doc = doc;
            _sheet = sheet;
        }

        /// <summary>
        /// Initialize a range with range reference
        /// </summary>
        /// <param name="doc">Excel document</param>
        /// <param name="sheet">Worksheet</param>
        /// <param name="range">Range reference</param>
        public ExcelOpenXMLRange(SpreadsheetDocument doc, Worksheet sheet, RangeReference range)
            : base(range)
        {
            _doc = doc;
            _sheet = sheet;
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
            if (String.IsNullOrEmpty(relativeRange))
                throw new ArgumentNullException("relativeRange");
            return new ExcelOpenXMLRange(_doc, _sheet, SubRange(relativeRange, clipToRange));
        }

        /// <summary>
        /// Calculate reference of new cell with its relative address
        /// </summary>
        /// <param name="relativeRef">Address in A1 expression</param>
        /// <returns></returns>
        protected CellReference TranslateCellReference(String relativeRef)
        {
            var matchAchor = _RegexAnchor.Match(relativeRef);
            if (matchAchor.Success)
            {   // This cell reference is anchored to another cell
                try
                {
                    var anchor = matchAchor.Groups["Anchor"].Value;
                    var offx = int.Parse(matchAchor.Groups["OffsetX"].Value);
                    var offy = int.Parse(matchAchor.Groups["OffsetY"].Value);
                    var anchorCell = TranslateCellReference(anchor);
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
            return CalculateCellReference(address);
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
                    case CellValues.Number:
                        {
                            long n = 0;
                            double d = 0.0;
                            if (long.TryParse(val, out n))
                                return n;
                            if (double.TryParse(val, out d))
                                return d;
                            return 0;
                        }
                    case CellValues.Date:
                        {   // Convert serialized date to DateTime
                            var dbl = double.Parse(val);
                            return DateTime.FromOADate(dbl);
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
        /// Get corresponding cell value type from system type
        /// </summary>
        /// <param name="type"></param>
        /// <returns></returns>
        protected CellValues GetCellValues(Type type)
        {
            if (type == typeof(bool))
            {
                return CellValues.Boolean;
            }
            else if (type == typeof(DateTime))
            {
                return CellValues.Date;
            }
            else if (type == typeof(short) || type == typeof(ushort) ||
                type == typeof(int) || type == typeof(uint) ||
                type == typeof(long) || type == typeof(ulong) ||
                type == typeof(float) || type == typeof(double) ||
                type == typeof(byte) || type == typeof(sbyte) ||
                type == typeof(Int16) || type == typeof(UInt16) ||
                type == typeof(Int32) || type == typeof(UInt32) ||
                type == typeof(Int64) || type == typeof(UInt64))
            {
                return CellValues.Number;
            }
            else
            {
                return CellValues.SharedString;
            }
        }

        // Given text and a SharedStringTablePart, creates a SharedStringItem with the specified text 
        // and inserts it into the SharedStringTablePart. If the item already exists, returns its index.
        private int InsertSharedStringItem(string text)
        {
            SharedStringTablePart shareStringPart;
            shareStringPart = _doc.WorkbookPart.GetPartsOfType<SharedStringTablePart>().Any()
                ? _doc.WorkbookPart.GetPartsOfType<SharedStringTablePart>().First()
                : _doc.WorkbookPart.AddNewPart<SharedStringTablePart>();
            // If the part does not contain a SharedStringTable, create one.
            if (shareStringPart.SharedStringTable == null)
            {
                shareStringPart.SharedStringTable = new SharedStringTable();
            }

            int i = 0;

            // Iterate through all the items in the SharedStringTable. If the text already exists, return its index.
            foreach (SharedStringItem item in shareStringPart.SharedStringTable.Elements<SharedStringItem>())
            {
                if (item.InnerText == text)
                {
                    return i;
                }

                i++;
            }

            // The text does not exist in the part. Create the SharedStringItem and return its index.
            shareStringPart.SharedStringTable.AppendChild(new SharedStringItem(new DocumentFormat.OpenXml.Spreadsheet.Text(text)));
            shareStringPart.SharedStringTable.Save();

            return i;
        }

        protected Row GetPrevRow(int rowNumber, SheetData sheetData)
        {
            for(int i=rowNumber-1;i>0;i--)
            {
                var prevRow = sheetData.Descendants<Row>().FirstOrDefault(a => a.RowIndex == i);
                if (prevRow != null)
                    return prevRow;
            }
            return null;
        }

        protected Cell GetPrevCell(CellReference cell, Row row)
        {
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

        protected Cell CreateCell(CellReference cellRef, EnumValue<CellValues> dataType)
        {
            var sheetData = _sheet.Descendants<SheetData>().FirstOrDefault();
            if (sheetData == null)
            {
                sheetData = new SheetData();
                _sheet.AppendChild(sheetData);
            }
            var row = sheetData.Descendants<Row>().FirstOrDefault(a => a.RowIndex == cellRef.Row);
            if (row == null)
            {
                row = new Row();
                row.RowIndex = new UInt32Value((uint)cellRef.Row);
                var prevRow = GetPrevRow(cellRef.Row, sheetData);
                if (prevRow == null)
                    sheetData.InsertAt(row, 0);
                else
                    sheetData.InsertAfter(row, prevRow);
            }
            var cell = new Cell
            {
                CellReference = cellRef.ToString(),
                DataType = dataType
            };
            var prevCell = GetPrevCell(cellRef, row);
            if (prevCell == null)
                row.InsertAt(cell, 0);
            else
                row.InsertAfter(cell, prevCell);
            return cell;
        }

        /// <summary>
        /// Set value of cell
        /// </summary>
        /// <param name="cellRef">Cell reference</param>
        /// <param name="value">Cell value</param>
        public void SetCellValue(CellReference cellRef, Object value)
        {
            var cellType = GetCellValues(value == null ? null : value.GetType());
            Cell cell = _sheet.Descendants<Cell>().FirstOrDefault(a => a.CellReference == cellRef.ToString());
            if (cell == null)
            {   // Cell not exist, must create new one
                cell = CreateCell(cellRef, cellType);
            }
            else
            {
                if (cell.CellValue != null)
                    cell.RemoveChild(cell.CellValue);
                if (cell.DataType != null)
                    cellType = cell.DataType; // Don't change original data type
            }
            switch (cellType)
            {
                case CellValues.Boolean:
                    {
                        if ((bool)value)
                            cell.AppendChild(new CellValue("1"));
                        else
                            cell.AppendChild(new CellValue("0"));
                        break;
                    }
                case CellValues.Number:
                    {
                        cell.AppendChild(new CellValue(value.ToString()));
                        break;
                    }
                case CellValues.Date:
                    {   // Convert serialized date to DateTime
                        var dbl = ((DateTime)value).ToOADate();
                        cell.AppendChild(new CellValue(dbl.ToString(CultureInfo.InvariantCulture)));
                        break;
                    }
                case CellValues.SharedString:
                    {
                        if (value != null)
                        {
                            var i = InsertSharedStringItem(value.ToString());
                            if (cell.CellValue == null)
                                cell.CellValue = new CellValue(i.ToString());
                            if (cell.DataType == null)
                            {
                                cell.DataType = new EnumValue<CellValues>(cellType);
                            }
                        }
                        break;
                    }
            }
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
            set
            {
                CellReference r = CalculateCellReference(col, row);
                SetCellValue(r, value);
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
                CellReference r = TranslateCellReference(relRef);
                return GetCellValue(r);
            }
            set
            {
                CellReference r = TranslateCellReference(relRef);
                SetCellValue(r, value);
            }
        }

        /// <summary>
        /// Convert value from excel to dotNET standard
        /// </summary>
        /// <param name="val">value read from excel</param>
        /// <param name="tgtType">Target type to be converted to</param>
        /// <returns>Data converted</returns>
        protected static object ConvertExcelValue(object val, Type tgtType)
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
            }
            // Handle sub collections here
            foreach (CollectionConfigElement colCfg in cfg.Collections)
            {
                var pColT = t.GetProperty(colCfg.Name);
                if (pColT != null)
                {
                    var list = pColT.GetValue(ent) as IList;
                    if (list == null)
                    {
                        throw new InvalidOperationException(
                            String.Format(
                                "To support load embeded collection, property {0} must be type that implements IList interface and must not be null!",
                                colCfg.Name));
                    }
                    var listType = list.GetType();
                    if (!listType.IsGenericType)
                    {
                        throw new InvalidOperationException(
                            String.Format("Type of property {0} must be a generic collection!", colCfg.Name));
                    }
                    // Get the item type
                    if (listType.GenericTypeArguments.Count() != 1)
                    {
                        throw new InvalidOperationException(
                            String.Format(
                                "Type of property {0} must be a generic collection with only one type parameter!",
                                colCfg.Name));
                    }
                    var itemType = listType.GenericTypeArguments[0];
                    // Instantiate enumerable for item type
                    var enumType = typeof(ExcelEntityEnumerable<>);
                    Type[] typeArgs = { itemType };
                    var eType = enumType.MakeGenericType(typeArgs);
                    // Embeded collection may has its ranged defined as relative reference, so we should pass the range of main entity
                    // as a parent container.
                    var enumerable = Activator.CreateInstance(eType, _doc, colCfg, this) as IEnumerable;
                    if (enumerable == null)
                        throw new InvalidOperationException(
                            String.Format("Cannot create ExcelEntityEnumerable<{0}>!", itemType.Name));
                    foreach (var item in enumerable)
                    {
                        list.Add(item);
                    }
                }
            }

            return (cfg.AllowNull || entExist) ? ent : null;
        }

        /// <summary>
        /// Write one entity to range
        /// </summary>
        /// <typeparam name="T">Type of entity</typeparam>
        /// <param name="cfg">Configuration that defines fields of entity</param>
        /// <param name="ent">Entity objet to be written</param>
        public void WriteEntity<T>(EntityConfigElement cfg, T ent) where T : class
        {
            Type t = typeof(T);
            foreach (FieldConfigElement field in cfg.Fields)
            {
                var pT = t.GetProperty(field.Name);
                if (pT != null)
                {   // Handle this property
                    object cellVal = pT.GetValue(ent);
                    this[field.Range] = cellVal;
                }
            }
            // Handle sub collections here
            foreach (CollectionConfigElement colCfg in cfg.Collections)
            {
                var pColT = t.GetProperty(colCfg.Name);
                if (pColT != null)
                {
                    var list = pColT.GetValue(ent) as IList;
                    if (list == null)
                    {
                        throw new InvalidOperationException(
                            String.Format(
                                "To support write embeded collection, property {0} must be type that implements IList interface and must not be null!",
                                colCfg.Name));
                    }
                    var listType = list.GetType();
                    if (!listType.IsGenericType)
                    {
                        throw new InvalidOperationException(
                            String.Format("Type of property {0} must be a generic collection!", colCfg.Name));
                    }
                    // Get the item type
                    if (listType.GenericTypeArguments.Count() != 1)
                    {
                        throw new InvalidOperationException(
                            String.Format(
                                "Type of property {0} must be a generic collection with only one type parameter!",
                                colCfg.Name));
                    }
                    var itemType = listType.GenericTypeArguments[0];
                    var writterType = typeof(ExcelEntityWriter<>);
                    Type[] typeArgs = { itemType };
                    var wType = writterType.MakeGenericType(typeArgs);
                    var writter = Activator.CreateInstance(wType, _doc, colCfg, this);
                    var writeFunc = wType.GetMethod("Write");
                    foreach (var item in list)
                    {
                        writeFunc.Invoke(writter, new object[] { item });
                    }
                }
            }
        }
    }
}
