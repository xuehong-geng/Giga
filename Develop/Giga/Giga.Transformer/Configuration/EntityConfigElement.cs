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

        [ConfigurationProperty("allowNull", IsRequired = false, DefaultValue = "true")]
        public bool AllowNull
        {
            get
            {
                return (bool) base["allowNull"];
            }
            set
            {
                base["allowNull"] = value;
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
}
