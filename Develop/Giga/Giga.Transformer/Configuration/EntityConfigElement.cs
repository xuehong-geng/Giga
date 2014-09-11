using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Giga.Transformer.Configuration
{
    /// <summary>
    /// Configuration of transformer entity
    /// </summary>
    public class EntityConfigElement : ConfigurationElement
    {
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

        [ConfigurationProperty("Fields", IsRequired = false)]
        [ConfigurationCollection(typeof (FieldConfigElement), AddItemName = "Field")]
        public FieldConfigCollection Fields
        {
            get
            {
                return (FieldConfigCollection) base["Fields"];
            }
            set
            {
                base["Fields"] = value;
            }
        }
    }
}
