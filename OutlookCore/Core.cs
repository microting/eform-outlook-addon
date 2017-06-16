using OutlookOffice;
using OutlookSql;
using OutlookShared;
using eFormShared;

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using eFormCore;

namespace OutlookCore
{
    public class Core
    {
        #region events
        public event EventHandler HandleEventException;
        #endregion

        #region var
        SqlController sqlController;
        OutlookController outlookController;
        Tools t = new Tools();

        string connectionString;
        bool logEvents;

        bool coreRunning = false;
        bool coreRestarting = false;
        bool coreStatChanging = false;
        bool coreThreadAlive = false;

        bool syncOutlookConvertRunning = false;
        bool syncOutlookAppsRunning = false;
        bool syncInteractionCaseRunning = false;

        List<ExceptionClass> exceptionLst = new List<ExceptionClass>();
        #endregion

        #region con
        public Core()
        {

        }
        #endregion

        #region public state
        public bool     Start(string connectionString)
        {
            try
            {
                if (!coreRunning && !coreStatChanging)
                {
                    coreStatChanging = true;


                    //sqlController
                    sqlController = new SqlController(connectionString);
                    logEvents = bool.Parse(sqlController.SettingRead(Settings.logLevel));

                    sqlController.LogCritical("Core.Start() at:" + DateTime.Now.ToShortDateString() + " " + DateTime.Now.ToLongTimeString());
                    sqlController.LogVariable(nameof(connectionString), connectionString);
                    this.connectionString = connectionString;

                    if (!sqlController.SettingCheckAll())
                        throw new ArgumentException("Use AdminTool to setup database correct");
                    sqlController.LogStandard("SqlController started");


                    //outlookController
                    outlookController = new OutlookController(sqlController);
                    sqlController.LogStandard("OutlookController started");


                    //coreThread
                    Thread coreThread = new Thread(() => CoreThread());
                    coreThread.Start();
                    sqlController.LogStandard("CoreThread started");

                    sqlController.LogStandard("Core started");
                    coreStatChanging = false;
                }
            }
            catch (Exception ex)
            {
                coreRunning = false;
                coreStatChanging = false;
                throw new Exception("FATAL Exception. Core failed to 'Start'", ex);
            }
            return true;
        }

        public bool     Close()
        {
            try
            {
                if (coreRunning && !coreStatChanging)
                {
                    coreStatChanging = true;

                    coreThreadAlive = false;
                    sqlController.LogCritical("Core.Close() at:" + DateTime.Now.ToShortDateString() + " " + DateTime.Now.ToLongTimeString());


                    while (coreRunning)
                        Thread.Sleep(200);

                    while (syncOutlookAppsRunning)
                        Thread.Sleep(200);

                    while (syncInteractionCaseRunning)
                        Thread.Sleep(200);

                    sqlController.LogStandard("Core closed");

                    outlookController = null;
                    sqlController = null;

                    coreStatChanging = false;
                }
            }
            catch (Exception ex)
            {
                coreRunning = false;
                coreThreadAlive = false;
                coreStatChanging = false;
                throw new Exception("FATAL Exception. Core failed to 'Close'", ex);
            }
            return true;
        }

        public bool     Running()
        {
            return coreRunning;
        }
        #endregion

        public void     Test_Reset(string connectionString)
        {
            sqlController = new SqlController(connectionString);
            outlookController = new OutlookController(sqlController);
            AdminTools at = new AdminTools(sqlController.SettingRead(Settings.microtingDb));

            try
            {
                if (!coreRunning && !coreStatChanging)
                {
                    coreStatChanging = true;
                    sqlController.LogStandard("Reset!");

                    List<Appointment> lstAppointments;

                    DateTime now = DateTime.Now;
                    DateTime rollBackTo = now.AddDays(-0);
                    lstAppointments = outlookController.UnitTest_CalendarItemGetAllNonRecurring(rollBackTo.AddDays(-7), now);
                    foreach (var item in lstAppointments)
                        outlookController.CalendarItemUpdate(item, WorkflowState.Planned, true);

                    sqlController.SettingUpdate(Settings.checkLast_At, rollBackTo.ToString());

                    at.RetractEforms();

                    sqlController.UnitTest_TruncateTable_Outlook("appointment_versions");
                    sqlController.UnitTest_TruncateTable_Outlook("appointments");
                    sqlController.UnitTest_TruncateTable_Microting("a_interaction_case_lists");
                    sqlController.UnitTest_TruncateTable_Microting("a_interaction_cases");
                    sqlController.UnitTest_TruncateTable_Microting("notifications");
                    sqlController.UnitTest_TruncateTable_Microting("cases");

                    coreStatChanging = false;
                }
            }
            catch (Exception ex)
            {
                coreRunning = false;
                coreStatChanging = false;
                throw new Exception("FATAL Exception. Core failed to 'Reset'", ex);
            }
            Close();
        }

        #region private
        private void    CoreThread()
        {
            bool firstRun = true;

            coreRunning = true;
            coreThreadAlive = true;

            while (coreThreadAlive)
            {
                try
                {
                    if (coreRunning)
                    {
                        #region warm up
                        sqlController.LogEverything(t.GetMethodName() + " initiated");

                        if (firstRun)
                        {
                            outlookController.CalendarItemConvertRecurrences();
                            firstRun = false;
                            sqlController.LogStandard(t.GetMethodName() + " warm up completed");
                        }
                        #endregion

                        Thread syncOutlookConvertThread
                            = new Thread(() => SyncOutlookConvert());
                        syncOutlookConvertThread.Start();

                        Thread syncOutlookAppsThread
                            = new Thread(() => SyncOutlookApps());
                        syncOutlookAppsThread.Start();

                        Thread syncInteractionCaseCreateThread
                            = new Thread(() => SyncInteractionCase());
                        syncInteractionCaseCreateThread.Start();

                        Thread.Sleep(1500);
                    }

                    Thread.Sleep(500);
                }
                catch (ThreadAbortException)
                {
                    coreRunning = false;
                    coreStatChanging = false;
                    sqlController.LogWarning(t.GetMethodName() + " catch of ThreadAbortException");
                }
                catch (Exception ex)
                {
                    coreRunning = false;
                    coreStatChanging = false;
                    throw new Exception("FATAL Exception. " + t.GetMethodName() + " failed", ex);
                }
            }
            coreRunning = false;
        }

        private void    SyncOutlookConvert()
        {
            try
            {
                if (!syncOutlookConvertRunning)
                {
                    syncOutlookConvertRunning = true;

                    if (coreRunning)
                    {
                        while (outlookController.CalendarItemConvertRecurrences()) { }

                        sqlController.LogEverything("outlookController.CalendarItemIntrepid() completed");

                        for (int i = 0; i < 20 && coreRunning; i++)
                            Thread.Sleep(1000);
                    }
                    
                    syncOutlookConvertRunning = false;
                }
            }
            catch (Exception ex)
            {
                syncOutlookConvertRunning = false;

                if (Running())
                    TriggerHandleExpection(t.GetMethodName() + " failed", ex, false);
            }
        }

        private void    SyncOutlookApps()
        {
            try
            {
                if (!syncOutlookAppsRunning)
                {
                    syncOutlookAppsRunning = true;

                    while (coreRunning && outlookController.CalendarItemIntrepid())
                        sqlController.LogEverything("outlookController.CalendarItemIntrepid() completed");

                    while (coreRunning && outlookController.CalendarItemReflecting(null))
                        sqlController.LogEverything("outlookController.CalendarItemReflecting() completed");

                    syncOutlookAppsRunning = false;
                }
            }
            catch (Exception ex)
            {
                syncOutlookAppsRunning = false;
                TriggerHandleExpection(t.GetMethodName() + " failed", ex, false);
            }
        }

        private void    SyncInteractionCase()
        {
            try
            {
                if (!syncInteractionCaseRunning)
                {
                    syncInteractionCaseRunning = true;

                    while (coreRunning && sqlController.SyncInteractionCase())
                        sqlController.LogEverything(t.GetMethodName() + " completed");

                    syncInteractionCaseRunning = false;
                }
            }
            catch (Exception ex)
            {
                syncInteractionCaseRunning = false;
                TriggerHandleExpection(t.GetMethodName() + " failed", ex, false);
            }
        }
        #endregion

        #region private outwards triggers
        private void    TriggerHandleExpection(string exceptionDescription, Exception exception, bool restartCore)
        {
            try
            {
                HandleEventException?.Invoke(exception, EventArgs.Empty);

                string fullExceptionDescription = t.PrintException(exceptionDescription, exception);

                if (sqlController != null)
                    sqlController.LogCritical(fullExceptionDescription);

                ExceptionClass exCls = new ExceptionClass(fullExceptionDescription, DateTime.Now);
                exceptionLst.Add(exCls);

                int secondsDelay = CheckExceptionLst(exCls);

                if (restartCore)
                {
                    Thread coreRestartThread = new Thread(() => Restart(secondsDelay));
                    coreRestartThread.Start();
                }
            }
            catch
            {
                coreRunning = false;
                throw new Exception("FATAL Exception. Core failed to handle an Expection", exception);
            }
        }

        private int         CheckExceptionLst(ExceptionClass exceptionClass)
        {
            int secondsDelay = 1;

            int count = 0;
            #region find count
            try
            {
                //remove Exceptions older than an hour
                for (int i = exceptionLst.Count; i < 0; i--)
                {
                    if (exceptionLst[i].Time < DateTime.Now.AddHours(-1))
                        exceptionLst.RemoveAt(i);
                }

                //keep only the last 10 Exceptions
                if (exceptionLst.Count > 10)
                {
                    exceptionLst.RemoveAt(0);
                }

                //find highest court of the same Exception
                if (exceptionLst.Count > 1)
                {
                    foreach (ExceptionClass exCls in exceptionLst)
                    {
                        if (exceptionClass.Description == exCls.Description)
                        {
                            count++;
                        }
                    }
                }
            }
            catch { }
            #endregion

            sqlController.LogCritical(count + ". time the same Exception, within the last hour");
            if (count == 2)
                secondsDelay = 6; // 1/10 min

            if (count == 3)
                secondsDelay = 60; // 1 min

            if (count == 4)
                secondsDelay = 600; // 10 min

            if (count > 4)
                throw new Exception("FATAL Exception. Same Exception repeated to many times within one hour");

            return secondsDelay;
        }

        private void        Restart(int secondsDelay)
        {
            try
            {
                if (coreRestarting == false)
                {
                    coreRestarting = true;

                    sqlController.LogCritical("Core.Restart() at:" + DateTime.Now.ToShortDateString() + " " + DateTime.Now.ToLongTimeString());
                    Close();
                    sqlController.LogCritical("Trying to restart the Core in " + secondsDelay + " seconds");
                    Thread.Sleep(secondsDelay * 1000);
                    Start(connectionString);

                    coreRestarting = false;
                }
            }
            catch (Exception ex)
            {
                coreRunning = false;
                throw new Exception("FATAL Exception. Core failed to restart", ex);
            }
        }
        #endregion
    }

    internal class ExceptionClass
    {
        private ExceptionClass()
        {

        }

        internal ExceptionClass(string description, DateTime time)
        {
            Description = description;
            Time = time;
        }

        public string Description { get; set; }
        public DateTime Time { get; set; }
    }
}
