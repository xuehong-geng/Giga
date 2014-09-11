using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Giga.Transformer.Configuration
{
    /// <summary>
    /// Configuration element for Transformer data template
    /// </summary>
    public class TemplateConfigElement : ConfigurationElement
    {
        [ConfigurationProperty("name",IsKey = true)]
        public String Name
        {
            get
            {
                return (String) base["name"];
            }
            set
            {
                base["name"] = value;
            }
        }

        [ConfigurationProperty("parser", IsRequired = true)]
        public String Parser
        {
            get
            {
                return (String) base["parser"];
            }
            set
            {
                base["parser"] = value;
            }
        }

        [ConfigurationProperty("Collections", IsRequired = false)]
        [ConfigurationCollection(typeof (CollectionConfigElement), AddItemName = "Collection")]
        public CollectionConfigCollection Collections
        {
            get
            {
                return (CollectionConfigCollection) base["Collections"];
            }
            set
            {
                base["Collections"] = value;
            }
        }
    }

    /// <summary>
    /// Configuration collection of Transformer templates
    /// </summary>
    public class TemplateConfigCollection : ConfigurationElementCollection
    {
        protected override ConfigurationElement CreateNewElement()
        {
            return new TemplateConfigElement();
        }

        protected override object GetElementKey(ConfigurationElement element)
        {
            return ((TemplateConfigElement)element).Name;
        }

        public new TemplateConfigElement this[String name]
        {
            get
            {
                return (TemplateConfigElement) base.BaseGet(name);
            }
        }
    }
}
