using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Giga.Configuration;

namespace Giga.Log.Configuration
{
    /// <summary>
    /// Configuration element for event logger
    /// </summary>
    public class LoggerConfigurationElement : ConfigurationElement
    {
        /// <summary>
        /// Name of logger
        /// </summary>
        [ConfigurationProperty("name", IsKey=true)]
        public String Name
        {
            get { return (String)this["name"]; }
            set { this["name"] = value; }
        }

        /// <summary>
        /// Type of logger
        /// </summary>
        [ConfigurationProperty("type", IsKey = false, IsRequired = true)]
        public String Type
        {
            get { return (String)this["type"]; }
            set { this["type"] = value; }
        }

        /// <summary>
        /// Severity of logger
        /// </summary>
        [ConfigurationProperty("severity", IsKey = false, IsRequired = false, DefaultValue = "Error")]
        public String Severity
        {
            get { return (String)this["severity"]; }
            set { this["severity"] = value; }
        }

        /// <summary>
        /// Whether the logger is enabled
        /// </summary>
        [ConfigurationProperty("enabled", IsKey = false, IsRequired=false, DefaultValue=true)]
        public bool Enabled
        {
            get { return (bool)this["enabled"]; }
            set { this["enabled"] = value; }
        }

        /// <summary>
        /// Whether log event in Synchronize Mode
        /// </summary>
        [ConfigurationProperty("synchronize", IsKey = false, IsRequired=false, DefaultValue=true)]
        public bool Synchronize
        {
            get { return (bool)this["synchronize"]; }
            set { this["synchronize"] = value; }
        }


        /// <summary>
        /// Parameters collection
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
    /// Collection of logger configuration elements
    /// </summary>
    public class LoggerConfigurationCollection : ConfigurationElementCollection
    {
        protected override ConfigurationElement CreateNewElement()
        {
            return new LoggerConfigurationElement();
        }

        protected override object GetElementKey(ConfigurationElement element)
        {
            return (element as LoggerConfigurationElement).Name;
        }

        /// <summary>
        /// Check if logger is exist in the collection
        /// </summary>
        /// <param name="name">Name of logger</param>
        /// <returns></returns>
        public bool Contains(String name)
        {
            foreach (LoggerConfigurationElement elem in this)
            {
                if (elem.Name.Equals(name))
                    return true;
            }
            return false;
        }

        /// <summary>
        /// Get logger configuration element
        /// </summary>
        /// <param name="name">Name of logger</param>
        /// <returns></returns>
        public LoggerConfigurationElement Get(String name)
        {
            if (!Contains(name))
                return null;
            return this.BaseGet(name) as LoggerConfigurationElement;
        }
    }
}
