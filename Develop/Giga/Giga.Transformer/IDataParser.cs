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
        /// <returns></returns>
        bool Open(String filePath);

        /// <summary>
        /// Close file currently opened
        /// </summary>
        void Close();

        /// <summary>
        /// Load data object from file with specific configuration and template
        /// </summary>
        /// <param name="config">Configuration used to define behavior of data loading</param>
        /// <returns></returns>
        IEnumerable<T> Parse<T>(TemplateConfigElement config) where T : class, new();
    }
}
