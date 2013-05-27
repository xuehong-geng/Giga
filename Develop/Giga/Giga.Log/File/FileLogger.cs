using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Giga.Log.File
{
    /// <summary>
    /// File logger
    /// </summary>
    public class FileLogger : Logger
    {
        private String _rootPath = null;
        private String _baseName = null;
        private String _curLogFile = null;
        private long _maxSize = 1024000;

        public override void Initialize(Configuration.LoggerConfigurationElement cfg)
        {
            base.Initialize(cfg);
            // Setup root path
            String appBaseDir = AppDomain.CurrentDomain.BaseDirectory;
            _rootPath = cfg.Parameters.Get<String>("RootPath");
            if (String.IsNullOrEmpty(_rootPath))
            {   // No root path configured, use application's base dir
                _rootPath = Path.Combine(appBaseDir, "/logs");
            }
            else
            {   // Has root path
                if (!Path.IsPathRooted(_rootPath))
                {   // Relative path configured
                    _rootPath = Path.Combine(appBaseDir, _rootPath);
                }
            }
            // Get base name
            _baseName = cfg.Parameters.Get<String>("BaseName");
            if (String.IsNullOrEmpty(_baseName))
                _baseName = "LOG";
            // Get max size
            _maxSize = cfg.Parameters.Get<long>("MaxSize");
            // Calculate file name
            CalculateFileName();
        }

        /// <summary>
        /// Calculate current file name
        /// </summary>
        private void CalculateFileName()
        {
            bool needNew = false;
            if (!String.IsNullOrEmpty(_curLogFile))
            {   // Current file name exist
                if (System.IO.File.Exists(_curLogFile))
                {   // File exists
                    System.IO.FileInfo info = new FileInfo(_curLogFile);
                    if (info.Length > _maxSize)
                        needNew = true;
                }
            }
            else
                needNew = true;
            if (needNew)
            {
                String newFile = _baseName + DateTime.Now.ToString("yyyyMMddhhmmss") + ".log";
                _curLogFile = Path.Combine(_rootPath, newFile);
            }
        }

        /// <summary>
        /// Write log to file
        /// </summary>
        /// <param name="log"></param>
        protected override void WriteLog(EventLog log)
        {
            // Ensure file size
            CalculateFileName();
            // Open file for appending log
            FileStream strm = new FileStream(_curLogFile, FileMode.Append, FileAccess.ReadWrite, FileShare.ReadWrite);
            StreamWriter writer = new StreamWriter(strm, Encoding.UTF8);
            writer.WriteLine(log.ToString());
            writer.Close();
        }
    }
}
