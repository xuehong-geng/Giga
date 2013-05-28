using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Giga.Log.Configuration;

namespace Giga.Log
{
    /// <summary>
    /// Factory that used to create logger
    /// </summary>
    public class LogManager
    {
        /// <summary>
        /// Singletone object
        /// </summary>
        private static LogManager _instance = new LogManager();
        public static LogManager GetInstance() { return _instance; }

        public LogManager()
        {
            LoadLoggers();
        }

        private bool _outputToConsole = true;
        private Dictionary<String, ILogger> _loggers = new Dictionary<string, ILogger>();

        /// <summary>
        /// Load loggers according to configuration
        /// </summary>
        private void LoadLoggers()
        {
            LogConfigurationSection sec = ConfigurationManager.GetSection("Giga.Log") as LogConfigurationSection;
            if (sec == null)
            {   // No logger configured
                System.Diagnostics.Trace.TraceWarning("Giga.Log is not configured!\n");
                return;
            }
            _outputToConsole = sec.OutputToConsole;
            foreach (LoggerConfigurationElement elemLogger in sec.Loggers)
            {
                if (_loggers.ContainsKey(elemLogger.Name))
                {   // Logger already exist
                    System.Diagnostics.Trace.TraceWarning(String.Format("Logger {0} already loaded!\n", elemLogger.Name));
                }
                else
                {
                    Type loggerType = Type.GetType(elemLogger.Type);
                    if (loggerType == null)
                    {
                        throw new TypeLoadException(String.Format("Cannot load logger type {0}!", elemLogger.Type));
                    }
                    Logger logger = loggerType.Assembly.CreateInstance(loggerType.FullName) as Logger;
                    if (logger == null)
                    {
                        throw new InvalidCastException(String.Format("Cannot cast object of type {0} to Logger!", elemLogger.Type));
                    }
                    logger.Initialize(elemLogger);
                    _loggers.Add(elemLogger.Name, logger);
                }
            }
        }

        /// <summary>
        /// Write event log to loggers
        /// </summary>
        /// <param name="log"></param>
        public void Write(EventLog log)
        {
            foreach (ILogger logger in _loggers.Values)
            {
                try
                {
                    logger.Log(log);
                }
                catch (Exception err)
                {
                    System.Diagnostics.Trace.TraceError("Log event to logger failed! Event: {0}; Exception: {1}", log, err.ToString());
                }
            }
            System.Diagnostics.Trace.WriteLineIf(log != null, log.ToString());
            if (_outputToConsole)
            {
                Console.WriteLine(log.ToString());
            }
        }

        /// <summary>
        /// Helper method to log events
        /// </summary>
        /// <param name="log">Event log object</param>
        public static void Log(EventLog log)
        {
            GetInstance().Write(log);
        }

        /// <summary>
        /// Helper method to log events
        /// </summary>
        /// <param name="source">Event log source</param>
        /// <param name="severity">Severity of event log</param>
        /// <param name="exception">Exception related, if has</param>
        /// <param name="messageFmt">Message formatting string</param>
        /// <param name="args">Arguments for message formatting</param>
        public static void Log(String source, EventSeverity severity, Exception exception, String messageFmt, params object[] args)
        {
            EventLog evt = new EventLog(source, severity, exception, messageFmt, args);
            Log(evt);
        }

        /// <summary>
        /// Helper method to log events
        /// </summary>
        /// <param name="source">Event log source</param>
        /// <param name="severity">Severity of event log</param>
        /// <param name="messageFmt">Message formatting string</param>
        /// <param name="args">Arguments for message formatting</param>
        public static void Log(String source, EventSeverity severity, String messageFmt, params object[] args)
        {
            EventLog evt = new EventLog(source, severity, messageFmt, args);
            GetInstance().Write(evt);
        }

        /// <summary>
        /// Helper method to log verbose information
        /// </summary>
        /// <param name="source">Event log source</param>
        /// <param name="messageFmt">Message formatting string</param>
        /// <param name="args">Arguments for message formatting</param>
        public static void Verbose(String source, String messageFmt, params object[] args)
        {
            Log(source, EventSeverity.Verbose, messageFmt, args);
        }

        /// <summary>
        /// Helper method to log information
        /// </summary>
        /// <param name="source">Event log source</param>
        /// <param name="messageFmt">Message formatting string</param>
        /// <param name="args">Arguments for message formatting</param>
        public static void Info(String source, String messageFmt, params object[] args)
        {
            Log(source, EventSeverity.Info, messageFmt, args);
        }

        /// <summary>
        /// Helper method to log warning
        /// </summary>
        /// <param name="source">Event log source</param>
        /// <param name="messageFmt">Message formatting string</param>
        /// <param name="args">Arguments for message formatting</param>
        public static void Warning(String source, String messageFmt, params object[] args)
        {
            Log(source, EventSeverity.Warning, messageFmt, args);
        }

        /// <summary>
        /// Helper method to log error
        /// </summary>
        /// <param name="source">Event log source</param>
        /// <param name="messageFmt">Message formatting string</param>
        /// <param name="args">Arguments for message formatting</param>
        public static void Error(String source, String messageFmt, params object[] args)
        {
            Log(source, EventSeverity.Error, messageFmt, args);
        }

        /// <summary>
        /// Helper method to log error
        /// </summary>
        /// <param name="source">Event log source</param>
        /// <param name="err">Related exception</param>
        /// <param name="messageFmt">Message formatting string</param>
        /// <param name="args">Arguments for message formatting</param>
        public static void Error(String source, Exception err, String messageFmt, params object[] args)
        {
            Log(source, EventSeverity.Error, err, messageFmt, args);
        }

        /// <summary>
        /// Helper method to log fatal error
        /// </summary>
        /// <param name="source">Event log source</param>
        /// <param name="messageFmt">Message formatting string</param>
        /// <param name="args">Arguments for message formatting</param>
        public static void Fatal(String source, String messageFmt, params object[] args)
        {
            Log(source, EventSeverity.Fatal, messageFmt, args);
        }

        /// <summary>
        /// Helper method to log fatal error
        /// </summary>
        /// <param name="source">Event log source</param>
        /// <param name="err">Related exception</param>
        /// <param name="messageFmt">Message formatting string</param>
        /// <param name="args">Arguments for message formatting</param>
        public static void Fatal(String source, Exception err, String messageFmt, params object[] args)
        {
            Log(source, EventSeverity.Fatal, err, messageFmt, args);
        }

    }
}
