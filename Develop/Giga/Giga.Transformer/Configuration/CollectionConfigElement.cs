using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Giga.Transformer.Configuration
{
    /// <summary>
    /// Configuration Element for Transformer entity collections
    /// </summary>
    public class CollectionConfigElement : ConfigurationElement
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

        [ConfigurationProperty("range", IsRequired = true)]
        public String Range
        {
            get
            {
                return (String) base["range"];
            }
            set
            {
                base["range"] = value;
            }
        }

        [ConfigurationProperty("orientation", IsRequired = false, DefaultValue = "vertical")]
        public String Orientation
        {
            get
            {
                return (String) base["orientation"];
            }
            set
            {
                base["orientation"] = value;
            }
        }

        [ConfigurationProperty("ItemTemplate", IsRequired = true)]
        public ItemTemplateConfigElement ItemTemplate
        {
            get
            {
                return (ItemTemplateConfigElement) base["ItemTemplate"];
            }
            set
            {
                base["ItemTemplate"] = value;
            }
        }
    }

    /// <summary>
    /// Configuration collection for transformer entity collection
    /// </summary>
    public class CollectionConfigCollection : ConfigurationElementCollection
    {
        protected override ConfigurationElement CreateNewElement()
        {
            return new CollectionConfigElement();
        }

        protected override object GetElementKey(ConfigurationElement element)
        {
            return ((CollectionConfigElement)element).Name;
        }

        public CollectionConfigElement this[int idx]
        {
            get
            {
                return (CollectionConfigElement)base.BaseGet(idx);
            }
        }

        public new CollectionConfigElement this[String name]
        {
            get
            {
                return (CollectionConfigElement) base[name];
            }
            set
            {
                base[name] = value;
            }
        }
    }
}
