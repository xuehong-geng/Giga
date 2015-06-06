using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Giga.Transformer.Configuration;

namespace Giga.Transformer
{
    /// <summary>
    /// Interface of data parser
    /// </summary>
    public interface IDataParser
    {
        /// <summary>
        /// Open a file to parse
        /// </summary>
        /// <param name="filePath"></param>
        /// <param name="readOnly"></param>
        /// <returns></returns>
        bool Open(String filePath, bool readOnly = true);

        /// <summary>
        /// Close file currently opened
        /// </summary>
        void Close();

        /// <summary>
        /// Load data object from file with specific configuration and template
        /// </summary>
        /// <param name="config">Configuration used to define behavior of data loading</param>
        /// <returns></returns>
        IEnumerable<T> Read<T>(TemplateConfigElement config) where T : class, new();

        /// <summary>
        /// Write one object to file with specific configuration and template
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="config">Configuration used to define behavior of data writting</param>
        /// <param name="obj"></param>
        void Write<T>(TemplateConfigElement config, T obj) where T : class;

        /// <summary>
        /// Write multiple objects to file with specific configuration and template
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="config">Configuration used to define behavior of data writting</param>
        /// <param name="objs"></param>
        void Write<T>(TemplateConfigElement config, IEnumerable<T> objs) where T : class;
    }
}
