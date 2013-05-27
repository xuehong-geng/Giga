using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Giga.Log.Configuration
{
    /// <summary>
    /// Configuration section for Event Log
    /// </summary>
    public class LogConfigurationSection : ConfigurationSection
    {
        [ConfigurationProperty("Loggers", IsDefaultCollection=false)]
        [ConfigurationCollection(typeof(LoggerConfigurationElement),
            AddItemName="Logger",
            CollectionType=ConfigurationElementCollectionType.AddRemoveClearMap)]
        public LoggerConfigurationCollection Loggers
        {
            get { return (LoggerConfigurationCollection)this["Loggers"]; }
            set { this["Loggers"] = value; }
        }
    }
}
