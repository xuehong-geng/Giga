using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Giga.Transformer.Configuration;

namespace Giga.Transformer.Excel
{
    /// <summary>
    /// Exception : No more entity exist!
    /// </summary>
    public class NoMoreEntityException : ApplicationException
    {
        public Type EntityType { get; set; }

        public NoMoreEntityException(Type entType)
        {
            EntityType = entType;
        }

        public override string Message
        {
            get { return String.Format("No more entity of type {0} exist!", EntityType.FullName); }
        }
    }

    /// <summary>
    /// Data parser for handling excel files
    /// </summary>
    public class ExcelParser : IDataParser
    {
        private SpreadsheetDocument _doc = null;

        /// <summary>
        /// Open a file to parse
        /// </summary>
        /// <param name="filePath"></param>
        /// <returns></returns>
        public bool Open(String filePath)
        {
            // Open excel file
            try
            {
                _doc = SpreadsheetDocument.Open(filePath, false);
            }
            catch (Exception err)
            {
                Trace.TraceError("Cannot open excel file {0}! Err:{1}", filePath, err.Message);
                _doc = null;
                return false;
            }
            return true;
        }

        /// <summary>
        /// Close file currently opened
        /// </summary>
        public void Close()
        {
            if (_doc != null)
            {
                _doc.Close();
                _doc = null;
            }
        }

        /// <summary>
        /// Load data object from file with specific configuration and template
        /// </summary>
        /// <param name="config">Configuration used to define behavior of data loading</param>
        /// <returns></returns>
        public IEnumerable<T> Parse<T>(TemplateConfigElement config) where T : class, new()
        {
            // Go throw template and parse data from excel file
            if (config.Collections == null || config.Collections.Count < 1)
                throw new InvalidOperationException(
                    "No valid collection found in TemplateConfigElement for ExcelParser!");
            // This method only handle one collection
            CollectionConfigElement colCfg = config.Collections[0];
            // Create enumerator for this type of entities
            var parser = new ExcelEntityEnumerable<T>(_doc, colCfg);
            return parser;
        }
    }

    /// <summary>
    /// Enumerable for entities in excel file
    /// </summary>
    internal class ExcelEntityEnumerable<T> : IEnumerable<T> where T : class, new()
    {
        private readonly SpreadsheetDocument _doc = null;
        private readonly CollectionConfigElement _colCfg = null;

        public ExcelEntityEnumerable(SpreadsheetDocument doc, CollectionConfigElement colCfg)
        {
            _doc = doc;
            _colCfg = colCfg;
        }

        IEnumerator<T> IEnumerable<T>.GetEnumerator()
        {
            return new ExcelEntityEnumerator<T>(_doc, _colCfg);
        }

        public IEnumerator GetEnumerator()
        {
            return new ExcelEntityEnumerator<T>(_doc, _colCfg);
        }
    }

    /// <summary>
    /// Enumerator for entities in excel file
    /// </summary>
    internal class ExcelEntityEnumerator<T> : IEnumerator<T> where T : class, new()
    {
        private SpreadsheetDocument _doc = null;
        private CollectionConfigElement _colCfg = null;
        private int _currentIdx = 0;
        private T _current = null;
        private ExcelOpenXMLRange _collectionRange = null;
        private ExcelOpenXMLRange _entityRange = null;

        /// <summary>
        /// Initialize entity enumerator
        /// </summary>
        /// <param name="doc">Excel document</param>
        /// <param name="colCfg">Configuration for entity collection template</param>
        public ExcelEntityEnumerator(SpreadsheetDocument doc, CollectionConfigElement colCfg)
        {
            _doc = doc;
            _colCfg = colCfg;
            Reset();
        }

        /// <summary>
        /// Move to next entity
        /// </summary>
        /// <returns></returns>
        public bool MoveNext()
        {
            _currentIdx++;
            _entityRange = null;
            try
            {
                _current = ReadCurrent();
                return _current != null;
            }
            catch (NoMoreEntityException)
            {
                return false;
            }
        }

        /// <summary>
        /// Reset enumerator, move to first entity
        /// </summary>
        public void Reset()
        {
            _currentIdx = 0;
            _entityRange = null;
            _current = ReadCurrent();
            if (_current == null)
                throw new InvalidDataException(String.Format("There is no valid data for entity {0}",
                    typeof(T).FullName));
        }

        /// <summary>
        /// Get current entity
        /// </summary>
        T IEnumerator<T>.Current
        {
            get
            {
                return _current;
            }
        }

        /// <summary>
        /// Get current entity
        /// </summary>
        object IEnumerator.Current
        {
            get
            {
                return _current;
            }
        }

        /// <summary>
        /// Calculate range of entity
        /// </summary>
        /// <param name="rangeFirst">Range in configuration</param>
        /// <param name="idx">Index of entity</param>
        /// <param name="isVertical">Whether entities arranged vertically</param>
        /// <returns></returns>
        protected String CalculateEntityRange(String rangeFirst, int idx, bool isVertical = true)
        {
            var tl = new CellReference();
            var br = new CellReference();
            ExcelOpenXMLRange.CalculateRange(rangeFirst, ref tl, ref br);
            int height = br.Row - tl.Row + 1;
            int width = br.Col - tl.Col + 1;
            int rangeH = _collectionRange.Height;
            int rangeW = _collectionRange.Width;
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
            tl.Col = c+1;
            tl.Row = r+1;
            br = tl.Offset(width-1, height-1);
            return String.Format("{0}:{1}", tl, br);
        }

        /// <summary>
        /// Check if there is a valid entity in file
        /// </summary>
        /// <returns></returns>
        protected bool Validate()
        {
            // Get range of collection
            if (_collectionRange == null)
            {
                try
                {
                    String rangeCol = _colCfg.Range;
                    _collectionRange = _doc.GetRange(rangeCol);
                    if (_collectionRange == null)
                        return false;
                }
                catch (Exception)
                {
                    return false;
                }
            }
            // Calculate range of current entity
            if (_entityRange == null)
            {
                try
                {
                    EntityConfigElement entCfg = _colCfg.ItemTemplate.Entity;
                    String entRange = CalculateEntityRange(entCfg.Range, _currentIdx,
                        _colCfg.Orientation.Equals("vertical", StringComparison.OrdinalIgnoreCase));
                    _entityRange = _collectionRange.GetSubRange(entRange);
                }
                catch (Exception)
                {
                    return false;
                }
            }
            return true;
        }

        /// <summary>
        /// Read entity at specific index
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <returns></returns>
        protected T ReadCurrent()
        {
            // Get range of collection
            if (_collectionRange == null)
            {
                String rangeCol = _colCfg.Range;
                _collectionRange = _doc.GetRange(rangeCol);
                if (_collectionRange == null)
                    throw new InvalidDataException(String.Format("Cannot find valid range for collection of {0}!",
                        typeof(T).FullName));
            }
            // Calculate range of current entity
            if (_entityRange == null)
            {
                EntityConfigElement entCfg = _colCfg.ItemTemplate.Entity;
                String entRange = CalculateEntityRange(entCfg.Range, _currentIdx,
                    _colCfg.Orientation.Equals("vertical", StringComparison.OrdinalIgnoreCase));
                _entityRange = _collectionRange.GetSubRange(entRange);
                if (_entityRange == null)
                    throw new NoMoreEntityException(typeof(T));
            }
            // Read data for entity
            var ent = _entityRange.ReadEntity<T>(_colCfg.ItemTemplate.Entity);
            if (ent == null)
                throw new NoMoreEntityException(typeof(T));
            return ent;
        }

        /// <summary>
        /// Dispose enumerator, release resources.
        /// </summary>
        public void Dispose()
        {
            // Do nothing right now.
        }
    }
}
