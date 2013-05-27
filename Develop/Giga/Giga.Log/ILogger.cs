using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Giga.Log
{
    /// <summary>
    /// Interface of a event logger
    /// </summary>
    public interface ILogger
    {
        /// <summary>
        /// Log a event into log storage
        /// </summary>
        /// <param name="eventLog">Event to be logged</param>
        void Log(EventLog eventLog);
    }
}
