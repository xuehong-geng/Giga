using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Giga.Transformer.Configuration
{
    /// <summary>
    /// Configuration element for transformer entity fields
    /// </summary>
    public class FieldConfigElement : ConfigurationElement
    {
        [ConfigurationProperty("name", IsKey = true)]
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

        [ConfigurationProperty("type", IsRequired = false, DefaultValue = "System.String")]
        public String Type
        {
            get
            {
                return (String) base["type"];
            }
            set
            {
                base["type"] = value;
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
    }

    /// <summary>
    /// Configuration collection for transformer entity field collection
    /// </summary>
    public class FieldConfigCollection : ConfigurationElementCollection
    {
        protected override ConfigurationElement CreateNewElement()
        {
            return new FieldConfigElement();
        }

        protected override object GetElementKey(ConfigurationElement element)
        {
            return ((FieldConfigElement)element).Name;
        }
    }
}
