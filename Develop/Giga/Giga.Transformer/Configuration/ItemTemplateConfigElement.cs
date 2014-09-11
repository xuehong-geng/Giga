using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Giga.Transformer.Configuration
{
    /// <summary>
    /// Configuration element for transformer collection item template
    /// </summary>
    public class ItemTemplateConfigElement : ConfigurationElement
    {
        [ConfigurationProperty("Entity", IsRequired = true)]
        public EntityConfigElement Entity
        {
            get
            {
                return (EntityConfigElement) base["Entity"];
            }
            set
            {
                base["Entity"] = value;
            }
        }
    }
}
