using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Giga.Transformer.Configuration;

namespace Giga.Transformer
{
    /// <summary>
    /// Main class of transformer tool set
    /// </summary>
    public class Transformer
    {
        private TransformerConfigSection _cfg = null;

        /// <summary>
        /// Initialize transformer with specific configuration
        /// </summary>
        /// <param name="config"></param>
        public Transformer(TransformerConfigSection config)
        {
            _cfg = config;
        }

        /// <summary>
        /// Load data entities from file with specific template
        /// </summary>
        /// <typeparam name="T">Entity class</typeparam>
        /// <param name="filePath">Path of file that contains entity data</param>
        /// <param name="templateName">Name of template used to parse data from file</param>
        /// <returns>List of all entity data</returns>
        public IEnumerable<T> Load<T>(String filePath, String templateName) where T : class, new()
        {
            // Find parser
            TemplateConfigElement templateCfg = _cfg.Templates[templateName];
            if (templateCfg == null)
                throw new ConfigurationErrorsException(String.Format("Template {0} not exist!", templateName));
            var parserCfg = _cfg.Parsers[templateCfg.Parser] as ParserConfigElement;
            if (parserCfg == null)
                throw new ConfigurationErrorsException(String.Format("Parser {0} not configured!", templateCfg.Parser));
            Type parserType = Type.GetType(parserCfg.Type);
            if (parserType == null)
                throw new ConfigurationErrorsException(String.Format("Parser {0} not exist!", parserCfg.Type));
            // Create parser
            var parser = parserType.Assembly.CreateInstance(parserType.FullName) as IDataParser;
            if (parser == null)
                throw new ConfigurationErrorsException(String.Format("Cannot create instance of parser {0}!",
                    parserType.FullName));
            // Open file
            if (!parser.Open(filePath))
                throw new InvalidOperationException(String.Format("Cannot open file {0} with parser {1}!", filePath,
                    parserType.FullName));
            // Parse file and get all entities
            var lst = new List<T>();
            try
            {
                IEnumerable<T> set = parser.Parse<T>(templateCfg);
                lst.AddRange(set);
            }
            finally 
            {
                parser.Close();
            }
            return lst;
        }

        /// <summary>
        /// Load first entity from file with specific template
        /// </summary>
        /// <remarks>This method is usually used to load form data</remarks>
        /// <typeparam name="T">Entity class</typeparam>
        /// <param name="filePath">Path of file that contains entity data</param>
        /// <param name="templateName">Name of template used to parse data from file</param>
        /// <returns>The entity data</returns>
        public T LoadOne<T>(String filePath, String templateName) where T : class, new()
        {
            // Find parser
            TemplateConfigElement templateCfg = _cfg.Templates[templateName];
            if (templateCfg == null)
                throw new ConfigurationErrorsException(String.Format("Template {0} not exist!", templateName));
            var parserCfg = _cfg.Parsers[templateCfg.Parser] as ParserConfigElement;
            if (parserCfg == null)
                throw new ConfigurationErrorsException(String.Format("Parser {0} not configured!", templateCfg.Parser));
            Type parserType = Type.GetType(parserCfg.Type);
            if (parserType == null)
                throw new ConfigurationErrorsException(String.Format("Parser {0} not exist!", parserCfg.Type));
            // Create parser
            var parser = parserType.Assembly.CreateInstance(parserType.FullName) as IDataParser;
            if (parser == null)
                throw new ConfigurationErrorsException(String.Format("Cannot create instance of parser {0}!",
                    parserType.FullName));
            // Open file
            if (!parser.Open(filePath))
                throw new InvalidOperationException(String.Format("Cannot open file {0} with parser {1}!", filePath,
                    parserType.FullName));
            // Parse file and get all entities
            T ent = null;
            try
            {
                IEnumerable<T> set = parser.Parse<T>(templateCfg);
                ent = set.FirstOrDefault();
            }
            finally
            {
                parser.Close();
            }
            return ent;
        }
    }
}
