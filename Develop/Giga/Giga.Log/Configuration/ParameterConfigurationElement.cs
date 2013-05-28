using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Giga.Log.Configuration
{
    /// <summary>
    /// Configuration element for parameter
    /// </summary>
    public class ParameterConfigurationElement : ConfigurationElement
    {
        /// <summary>
        /// Parameter name
        /// </summary>
        [ConfigurationProperty("name", IsKey=true)]
        public String Name
        {
            get { return (String)this["name"]; }
            set { this["name"] = value; }
        }

        /// <summary>
        /// Parameter value
        /// </summary>
        [ConfigurationProperty("value", IsRequired=true)]
        public String Value
        {
            get { return (String)this["value"]; }
            set { this["value"] = value; }
        }

        /// <summary>
        /// Get value
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <returns></returns>
        public T Get<T>()
        {
            String val = Value;
            return (T)Convert.ChangeType(val, typeof(T));
        }
    }

    /// <summary>
    /// Collection of parameter configuration elements
    /// </summary>
    public class ParameterConfigurationCollection : ConfigurationElementCollection
    {
        protected override ConfigurationElement CreateNewElement()
        {
            return new ParameterConfigurationElement();
        }

        protected override object GetElementKey(ConfigurationElement element)
        {
            return (element as ParameterConfigurationElement).Name;
        }

        /// <summary>
        /// Check if a parameter is exist
        /// </summary>
        /// <param name="name">Name of parameter</param>
        /// <returns></returns>
        public bool Contains(String name)
        {
            foreach (ParameterConfigurationElement elem in this)
            {
                if (elem.Name.Equals(name, StringComparison.OrdinalIgnoreCase))
                    return true;
            }
            return false;
        }

        /// <summary>
        /// Get parameter from collection
        /// </summary>
        /// <typeparam name="T">Parameter type</typeparam>
        /// <param name="name">Parameter name</param>
        /// <returns>Parameter value</returns>
        public T Get<T>(String name)
        {
            if (!Contains(name))
                throw new InvalidOperationException(String.Format("Parameter {0} not exist!", name));
            ParameterConfigurationElement elem = this.BaseGet(name) as ParameterConfigurationElement;
            if (elem == null)
                throw new InvalidOperationException(String.Format("Parameter {0} not exist!", name));
            return elem.Get<T>();
        }
    }
}
