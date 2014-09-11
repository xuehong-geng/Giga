using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Giga.Transformer.Configuration
{
    /// <summary>
    /// Configuration section for <Giga.Transformer>
    /// </summary>
    public class TransformerConfigSection : ConfigurationSection
    {
        [ConfigurationProperty("Parsers", IsRequired = false)]
        [ConfigurationCollection(typeof(ParserConfigElement), AddItemName = "Parser")]
        public ParserConfigCollection Parsers
        {
            get
            {
                return (ParserConfigCollection) base["Parsers"];
            }
            set
            {
                base["Parsers"] = value;
            }
        }

        [ConfigurationProperty("Templates", IsRequired = false)]
        [ConfigurationCollection(typeof(TemplateConfigElement), AddItemName = "Template")]
        public TemplateConfigCollection Templates
        {
            get
            {
                return (TemplateConfigCollection) base["Templates"];
            }
            set
            {
                base["Templates"] = value;
            }
        }
    }
}
