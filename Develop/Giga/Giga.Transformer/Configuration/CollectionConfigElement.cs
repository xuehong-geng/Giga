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

        [ConfigurationProperty("range", IsRequired = false, DefaultValue = "")]
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

        /// <summary>
        /// A cell reference that acts as the signal of the begining of collection
        /// </summary>
        [ConfigurationProperty("startFrom", IsRequired = false, DefaultValue = "")]
        public String StartFrom
        {
            get
            {
                return (String) base["startFrom"];
            }
            set
            {
                base["startFrom"] = value;
            }
        }
        /// <summary>
        /// A cell reference that acts as the signal of the end of collection.
        /// </summary>
        /// <remarks>This property is usually used with a collection that is embeded in a form and its count is dynamical.</remarks>
        [ConfigurationProperty("endBefore", IsRequired = false, DefaultValue = "")]
        public String EndBefore
        {
            get
            {
                return (String)base["endBefore"];
            }
            set
            {
                base["endBefore"] = value;
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
