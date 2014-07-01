using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using AutomationServer;
using log4net;

namespace AutomationService
{
    public partial class AutomationService : ServiceBase
    {
        private static readonly ILog log = LogManager.GetLogger(typeof (AutomationService));

        private Server server;

        public AutomationService()
        {
            InitializeComponent();
            if (!System.Diagnostics.EventLog.SourceExists("AutomationServiceLogSource"))
                System.Diagnostics.EventLog.CreateEventSource("AutomationServiceLogSource",
                                                              "AutomationServiceLog");

            eventLog1.Source = "AutomationServiceLogSource";
            // the event log source by which the application is registered on the computer
            eventLog1.Log = "AutomationServiceLog";
        }

        protected override void OnStart(string[] args)
        {
            try
            {
                server = new Server();
                server.run();

                log.Info("Automation service started");
                eventLog1.WriteEntry("Automation service started");
            }
            catch (Exception e)
            {
                log.Error(e.Message, e);
                eventLog1.WriteEntry("Exception thrown from IISServer: " + e.Message + " Stacktrace:" + e.StackTrace + " 2" + e.InnerException.StackTrace, EventLogEntryType.Error);
                if (e.InnerException.InnerException != null)
                {
                    eventLog1.WriteEntry("33: " + e.InnerException.InnerException.StackTrace, EventLogEntryType.Error);
                }
                eventLog1.WriteEntry("Woo !!!", EventLogEntryType.Information);
                throw new Exception("Exception thrown from AutomationServer" + e.Message + " Stacktrace:" + e.StackTrace, e);
            }
        }

        protected override void OnStop()
        {
            server.withdrawService();
            log.Info("Automation service stopped");
            eventLog1.WriteEntry("Automation service stopped");
        }
    }
}