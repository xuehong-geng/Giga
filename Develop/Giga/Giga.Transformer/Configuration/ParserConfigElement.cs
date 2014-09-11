using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Giga.Transformer.Configuration
{
    /// <summary>
    /// Configuration element for Transormer parsers
    /// </summary>
    public class ParserConfigElement : ConfigurationElement
    {
        [ConfigurationProperty("name", IsKey = true)]
        public String Name
        {
            get
            {
                return (String)base["name"];
            }
            set
            {
                base["name"] = value;
            }
        }

        [ConfigurationProperty("type", IsRequired = true)]
        public String Type
        {
            get
            {
                return (String)base["type"];
            }
            set
            {
                base["type"] = value;
            }
        }
    }

    /// <summary>
    /// Configuration collection for Transformer parsers
    /// </summary>
    public class ParserConfigCollection : ConfigurationElementCollection
    {
        protected override ConfigurationElement CreateNewElement()
        {
            return new ParserConfigElement();
        }

        protected override object GetElementKey(ConfigurationElement element)
        {
            return ((ParserConfigElement)element).Name;
        }

        public new ParserConfigElement this[String name]
        {
            get
            {
                return (ParserConfigElement) base.BaseGet(name);
            }
        }
    }
}
