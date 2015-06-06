﻿using System;
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
        IEnumerable<T> IDataParser.Read<T>(TemplateConfigElement config)
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

        /// <summary>
        /// Write one object to file with specific configuration and template
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="config">Configuration used to define behavior of data writting</param>
        /// <param name="obj"></param>
        void IDataParser.Write<T>(TemplateConfigElement config, T obj)
        {
            // Go throw template and parse data from excel file
            if (config.Collections == null || config.Collections.Count < 1)
                throw new InvalidOperationException(
                    "No valid collection found in TemplateConfigElement for ExcelParser!");
            // This method only handle one collection
            CollectionConfigElement colCfg = config.Collections[0];
            var writter = new ExcelEntityWriter<T>(_doc, colCfg);
            writter.Write(obj);
        }

        /// <summary>
        /// Write multiple objects to file with specific configuration and template
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="config">Configuration used to define behavior of data writting</param>
        /// <param name="objs"></param>
        void IDataParser.Write<T>(TemplateConfigElement config, IEnumerable<T> objs)
        {
            // Go throw template and parse data from excel file
            if (config.Collections == null || config.Collections.Count < 1)
                throw new InvalidOperationException(
                    "No valid collection found in TemplateConfigElement for ExcelParser!");
            // This method only handle one collection
            CollectionConfigElement colCfg = config.Collections[0];
            var writter = new ExcelEntityWriter<T>(_doc, colCfg);
            foreach (var obj in objs)
            {
                writter.Write(obj);
            }
        }

    }

    /// <summary>
    /// Enumerable for entities in excel file
    /// </summary>
    internal class ExcelEntityEnumerable<T> : IEnumerable<T> where T : class, new()
    {
        private readonly SpreadsheetDocument _doc = null;
        private readonly CollectionConfigElement _colCfg = null;
        private ExcelOpenXMLRange _parentRange = null;

        public ExcelEntityEnumerable(SpreadsheetDocument doc, CollectionConfigElement colCfg, ExcelOpenXMLRange parentRange = null)
        {
            _doc = doc;
            _colCfg = colCfg;
            _parentRange = parentRange;
        }

        IEnumerator<T> IEnumerable<T>.GetEnumerator()
        {
            return new ExcelEntityEnumerator<T>(_doc, _colCfg, _parentRange);
        }

        public IEnumerator GetEnumerator()
        {
            return new ExcelEntityEnumerator<T>(_doc, _colCfg, _parentRange);
        }
    }

    /// <summary>
    /// Enumerator for entities in excel file
    /// </summary>
    internal class ExcelEntityEnumerator<T> : IEnumerator<T> where T : class, new()
    {
        private SpreadsheetDocument _doc = null;
        private CollectionConfigElement _colCfg = null;
        private ExcelOpenXMLRange _parentRange = null;
        private int _currentIdx = 0;
        private T _current = null;
        private ExcelOpenXMLRange _collectionRange = null;
        private ExcelOpenXMLRange _entityRange = null;
        private ExcelOpenXMLRange _endBefore = null;

        /// <summary>
        /// Initialize entity enumerator
        /// </summary>
        /// <param name="doc">Excel document</param>
        /// <param name="colCfg">Configuration for entity collection template</param>
        /// <param name="parentRange">Parent range that contains collection</param>
        /// <remarks>
        /// When parentRange is not null, all reference should be related to it.
        /// </remarks>
        public ExcelEntityEnumerator(SpreadsheetDocument doc, CollectionConfigElement colCfg, ExcelOpenXMLRange parentRange = null)
        {
            _doc = doc;
            _colCfg = colCfg;
            _parentRange = parentRange;
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
            _currentIdx = -1;
            _current = null;
            _entityRange = null;
            _endBefore = null;
            if (!String.IsNullOrEmpty(_colCfg.EndBefore))
            {   // There is abort flag cell, using defined name to mark cell dynamically.
                _endBefore = _doc.GetRange(_colCfg.EndBefore);
            }
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
        /// Read entity at specific index
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <returns></returns>
        protected T ReadCurrent()
        {
            // Get range of collection
            if (_collectionRange == null)
            {
                if (_parentRange != null)
                {
                    _collectionRange = _parentRange.GetSubRange(_colCfg.Range);
                }
                else
                {
                    _collectionRange = _doc.GetRange(_colCfg.Range);
                }
                if (_collectionRange == null)
                    throw new InvalidDataException(String.Format("Cannot find valid range for collection of {0}!",
                        typeof(T).FullName));
                // For dynamic collection, it's possible to use a defined name to mark the end of co
            }
            // Calculate range of current entity
            if (_entityRange == null)
            {
                EntityConfigElement entCfg = _colCfg.ItemTemplate.Entity;
                String entRef = _collectionRange.Sheet.ExpandToSheetBound(entCfg.Range);
                String entRange = ExcelOpenXMLHelper.CalculateEntityRange(_collectionRange, entRef, _currentIdx,
                    _colCfg.Orientation.Equals("vertical", StringComparison.OrdinalIgnoreCase), _endBefore);
                if (entRange == null)
                    throw new NoMoreEntityException(typeof(T));
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

    /// <summary>
    /// Writer for writting entities into excel file
    /// </summary>
    /// <typeparam name="T"></typeparam>
    internal class ExcelEntityWriter<T> where T : class
    {
        private readonly SpreadsheetDocument _doc = null;
        private readonly CollectionConfigElement _colCfg = null;
        private ExcelOpenXMLRange _parentRange = null;
        private ExcelOpenXMLRange _collectionRange = null;
        private ExcelOpenXMLRange _entityRange = null;

        public ExcelEntityWriter(SpreadsheetDocument doc, CollectionConfigElement colCfg, ExcelOpenXMLRange parentRange = null)
        {
            _doc = doc;
            _colCfg = colCfg;
            _parentRange = parentRange;           
        }

        /// <summary>
        /// Write one object to excel at current location and move to next location for next object
        /// </summary>
        /// <param name="obj"></param>
        public void Write(T obj)
        {
            // Calculate the entity range
            if (_entityRange == null)
            {   // No entity range calculated, this is the first entity
                if (_collectionRange == null)
                {   // First time, calculate collection range
                    if (_parentRange != null)
                    {
                        _collectionRange = _parentRange.GetSubRange(_colCfg.Range);
                    }
                    else
                    {
                        _collectionRange = _doc.GetRange(_colCfg.Range);
                    }
                    if (_collectionRange == null)
                        throw new InvalidDataException(String.Format("Cannot find valid range for collection of {0}!",
                            typeof(T).FullName));
                }
                EntityConfigElement entCfg = _colCfg.ItemTemplate.Entity;
                String entRef = _collectionRange.Sheet.ExpandToSheetBound(entCfg.Range);
                String entRange = ExcelOpenXMLHelper.CalculateEntityRange(_collectionRange, entRef, 0,
                    _colCfg.Orientation.Equals("vertical", StringComparison.OrdinalIgnoreCase));
                if (entRange == null)
                    throw new NoMoreEntityException(typeof(T));
                _entityRange = _collectionRange.GetSubRange(entRange);
                if (_entityRange == null)
                    throw new NoMoreEntityException(typeof(T));
            }
            // Write object to range
            _entityRange.WriteEntity(_colCfg.ItemTemplate.Entity, obj);
            // Move range to next
            if (_colCfg.Orientation.Equals("vertical", StringComparison.OrdinalIgnoreCase))
            {   // Move vertically
                int h = _entityRange.Height;
                _entityRange.Move(0, h);
            }
            else
            {   // Move horizontally
                int w = _entityRange.Width;
                _entityRange.Move(w, 0);
            }
        }
    }
}
