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
        public EventLog()
        {
            LogTime = DateTime.Now;
        }
        public EventLog(String source, EventSeverity severity, String messageFormat, params object[] args)
        {
            LogTime = DateTime.Now;
            Source = source;
            Severity = severity;
            Message = String.Format(messageFormat, args);
            CapturedException = null;
        }
        public EventLog(String source, EventSeverity severity, Exception exception, String messageFormat, params object[] args)
        {
            LogTime = DateTime.Now;
            Source = source;
            Severity = severity;
            Message = String.Format(messageFormat, args);
            CapturedException = exception;
        }

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

        /// <summary>
        /// Convert to string
        /// </summary>
        /// <returns></returns>
        public override string ToString()
        {
            StringBuilder str = new StringBuilder();
            str.AppendFormat("{0} [{1}] [{2}] :: {3}", LogTime, Source, Enum.GetName(typeof(EventSeverity),Severity), Message);
            if (CapturedException != null)
            {
                str.AppendFormat(" Exception: {0}", CapturedException.Message);
            }
            return str.ToString();
        }
    }
}
