using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Giga.User.Configuration
{
    /// <summary>
    /// Configuration section for user management
    /// </summary>
    public class UserConfigurationSection : ConfigurationSection
    {
        /// <summary>
        /// Account providers collection
        /// </summary>
        [ConfigurationProperty("AccountProviders", IsDefaultCollection=false)]
        [ConfigurationCollection(typeof(AccountProviderConfigurationElement),
            AddItemName="Provider",
            CollectionType=ConfigurationElementCollectionType.AddRemoveClearMap)]
        public AccountProviderConfigurationCollection AccountProviders
        {
            get { return (AccountProviderConfigurationCollection)this["AccountProviders"]; }
            set { this["AccountProviders"] = value; }
        }
    }
}
