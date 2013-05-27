using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.Text;
using System.Threading.Tasks;

namespace Giga.Log
{
    /// <summary>
    /// Event severity
    /// </summary>
    [DataContract]
    public enum EventSeverity : int
    {
        [EnumMember]
        Verbose = 0,
        [EnumMember]
        Info,
        [EnumMember]
        Warning,
        [EnumMember]
        Error,
        [EnumMember]
        Fatal,
    }

    /// <summary>
    /// Event log object
    /// </summary>
    [DataContract]
    public class EventLog
    {
        /// <summary>
        /// Time when event happened
        /// </summary>
        [DataMember]
        public DateTime LogTime { get; set; }

        /// <summary>
        /// Source of this event log
        /// </summary>
        [DataMember]
        public String Source { get; set; }

        /// <summary>
        /// Severity of this event log
        /// </summary>
        [DataMember]
        public EventSeverity Severity { get; set; }

        /// <summary>
        /// Message of this event log
        /// </summary>
        [DataMember]
        public String Message { get; set; }

        /// <summary>
        /// Exception captured if exists
        /// </summary>
        [DataMember]
        public Exception CapturedException { get; set; }
    }
}
