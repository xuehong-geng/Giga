using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Giga.Log;

namespace Giga.Test.Log
{
    [TestClass]
    public class EventLogTest
    {
        [TestMethod]
        public void TestEventLogs()
        {
            // Normal event log helpers
            LogManager.Fatal("Test", "This is a fatal error test at {0}!", DateTime.Now);
            LogManager.Error("Test", "This is a fatal error test at {0}!", DateTime.Now);
            LogManager.Warning("Test", "This is a fatal error test at {0}!", DateTime.Now);
            LogManager.Info("Test", "This is a fatal error test at {0}!", DateTime.Now);
            LogManager.Verbose("Test", "This is a fatal error test at {0}!", DateTime.Now);
            // Advance event log
            EventLog log = new EventLog("Advance Test", EventSeverity.Error, "This is advance event log test at {0}!", DateTime.Now);
            log.CapturedException = new ArgumentNullException();
            log.SetEnvironment("User", "Test user");
            log.SetEnvironment("Object", this);
            log.SetEnvironment("Method", GetType().GetMethod("TestEventLogs"));
            LogManager.Log(log);
        }
    }
}
