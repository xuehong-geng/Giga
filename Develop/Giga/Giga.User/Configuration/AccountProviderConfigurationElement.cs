using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Giga.Configuration;

namespace Giga.User.Configuration
{
    /// <summary>
    /// Configuration element for account provider
    /// </summary>
    public class AccountProviderConfigurationElement : ConfigurationElement
    {
        /// <summary>
        /// Name of provider
        /// </summary>
        [ConfigurationProperty("name", IsKey=true)]
        public String Name
        {
            get { return (String)this["name"]; }
            set { this["name"] = value; }
        }

        /// <summary>
        /// Type of account provider
        /// </summary>
        [ConfigurationProperty("type", IsRequired=true)]
        public String Type
        {
            get { return (String)this["type"]; }
            set { this["type"] = value; }
        }

        /// <summary>
        /// Connect string name of account provider
        /// </summary>
        [ConfigurationProperty("connectStringName", IsRequired = true)]
        public String ConnectStringName
        {
            get { return (String)this["connectStringName"]; }
            set { this["connectStringName"] = value; }
        }

        /// <summary>
        /// Parameters of account provider
        /// </summary>
        [ConfigurationProperty("Parameters", IsDefaultCollection=false)]
        [ConfigurationCollection(typeof(ParameterConfigurationElement),
            AddItemName="Parameter",
            CollectionType=ConfigurationElementCollectionType.AddRemoveClearMap)]
        public ParameterConfigurationCollection Parameters
        {
            get { return (ParameterConfigurationCollection)this["Parameters"]; }
            set { this["Parameters"] = value; }
        }
    }

    /// <summary>
    /// Configuration collection of account providers
    /// </summary>
    public class AccountProviderConfigurationCollection : ConfigurationElementCollection
    {
        protected override ConfigurationElement CreateNewElement()
        {
            return new AccountProviderConfigurationElement();
        }

        protected override object GetElementKey(ConfigurationElement element)
        {
            return (element as AccountProviderConfigurationElement).Name;
        }

        /// <summary>
        /// Check if provider exists
        /// </summary>
        /// <param name="providerName">Name of provider</param>
        /// <returns></returns>
        public bool Contains(String providerName)
        {
            foreach (AccountProviderConfigurationElement elem in this)
            {
                if (elem.Name.Equals(providerName, StringComparison.Ordinal))
                    return true;
            }
            return false;
        }

        /// <summary>
        /// Get provider configuration
        /// </summary>
        /// <param name="providerName">Name of provider</param>
        /// <returns></returns>
        public AccountProviderConfigurationElement Get(String providerName)
        {
            ConfigurationElement elem = this.BaseGet(providerName);
            if (elem == null)
                return null;
            else
                return elem as AccountProviderConfigurationElement;
        }
    }
}
