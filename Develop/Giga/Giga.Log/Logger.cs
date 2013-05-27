using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Giga.Log.Configuration;

namespace Giga.Log
{
    /// <summary>
    /// Base class of loggers
    /// </summary>
    public abstract class Logger : ILogger
    {
        /// <summary>
        /// Name of logger
        /// </summary>
        public String Name { get; set; }
        /// <summary>
        /// Logger severity
        /// </summary>
        public EventSeverity Severity { get; set; }
        /// <summary>
        /// If is enabled
        /// </summary>
        public bool Enabled { get; set; }
        /// <summary>
        /// If log event synchronized
        /// </summary>
        public bool Synchronize { get; set; }

        /// <summary>
        /// Initialize logger
        /// </summary>
        /// <param name="elemLogger">Logger configuration element</param>
        public virtual void Initialize(LoggerConfigurationElement cfg)
        {
            Enabled = cfg.Enabled;
            Synchronize = cfg.Synchronize;
            try
            {
                Severity = (EventSeverity)Enum.Parse(typeof(EventSeverity), cfg.Severity);
            }
            catch (Exception err)
            {
                throw new ConfigurationException(String.Format("Invalid value of Severity: {0}!", cfg.Severity), err);
            }
            if (!Synchronize)
            {   // Asynchronize mode, start backend thread
                StartBackgroundThread();
            }
        }

        /// <summary>
        /// Background worker thread
        /// </summary>
        private Thread _thread = null;

        /// <summary>
        /// Start background logger thread
        /// </summary>
        private void StartBackgroundThread()
        {
            if (_thread != null)
                return; // Already started
            _thread = new Thread(LoggerThread);
            _thread.Start();
        }

        /// <summary>
        /// Log event
        /// </summary>
        /// <param name="eventLog"></param>
        public void Log(EventLog eventLog)
        {
            if (eventLog == null || eventLog.Severity < Severity)
                return;
            if (Synchronize)
            {   // Write the log synchronizely
                try
                {
                    WriteLog(eventLog);
                }
                catch (Exception err)
                {
                    System.Diagnostics.Trace.TraceError(String.Format("Write log failed! {0}. \nException: {1}\n", eventLog.ToString(), err.ToString()));
                }
            }
            else
            {   // Write log asynchronizely, put log to the queue
                _logQueue.Add(eventLog);
            }
        }

        private BlockingCollection<EventLog> _logQueue = new BlockingCollection<EventLog>();

        /// <summary>
        /// Main procedure of background logger thread
        /// </summary>
        private void LoggerThread()
        {
            while (true)
            {
                EventLog evt = null;
                if (_logQueue.TryTake(out evt, 1000))
                {   // Has event to log
                    try
                    {
                        WriteLog(evt);
                    }
                    catch (Exception err)
                    {
                        System.Diagnostics.Trace.TraceError("Write log failed! {0}\nException:{1}\n", evt, err);
                    }
                }
            }
        }

        /// <summary>
        /// Write event log. Implementation should override this method to do real logging operation
        /// </summary>
        /// <param name="log">Event log to be written</param>
        protected abstract void WriteLog(EventLog log);
    }
}
