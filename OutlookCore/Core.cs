/*
The MIT License (MIT)

Copyright (c) 2014 microting

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
*/

using OutlookOffice;
using OutlookSql;
using eFormShared;
using eFormCore;

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace OutlookCore
{
    public class Core : CoreBase
    {
        //events
        public event EventHandler HandleEventException;

        #region var
        SqlController sqlController;
        IOutlookController outlookController;
        Tools t = new Tools();

        public Log log;

        bool syncOutlookConvertRunning = false;
        bool syncOutlookAppsRunning = false;
        bool syncInteractionCaseRunning = false;

        bool coreRunning = false;
        bool coreRestarting = false;
        bool coreStatChanging = false;
        bool coreThreadAlive = false;

        string connectionString;
        #endregion

        //con

        #region public state
        public bool             Start(string connectionString)
        {
            try
            {
                if (!coreRunning && !coreStatChanging)
                {
                    coreStatChanging = true;

                    if (string.IsNullOrEmpty(connectionString))
                        throw new ArgumentException("serverConnectionString is not allowed to be null or empty");

                    //sqlController
                    sqlController = new SqlController(connectionString);

                    //check settings
                    if (sqlController.SettingCheckAll().Count > 0)
                        throw new ArgumentException("Use AdminTool to setup database correctly. 'SettingCheckAll()' returned with errors");

                    if (sqlController.SettingRead(Settings.microtingDb) == "...missing...")
                        throw new ArgumentException("Use AdminTool to setup database correctly. microtingDb(connection string) not set, only default value found");

                    if (sqlController.SettingRead(Settings.firstRunDone) != "true")
                        throw new ArgumentException("Use AdminTool to setup database correctly. FirstRunDone has not completed");

                    //log
                    log = sqlController.StartLog(this);
                    log.LogCritical("Not Specified", "###########################################################################");
                    log.LogCritical("Not Specified", t.GetMethodName() + " called");
                    log.LogStandard("Not Specified", "SqlController and Logger started");

                    //settings read
                    this.connectionString = connectionString;
                    log.LogStandard("Not Specified", "Settings read");

                    //outlookController
                    if (sqlController.SettingRead(Settings.calendarName) == "unittest")
                    {
                        outlookController = new OutlookController_Fake(sqlController, log);
                        log.LogStandard("Not Specified", "OutlookController_Fake started");
                    }
                    else
                    {
                        outlookController = new OutlookController(sqlController, log);
                        log.LogStandard("Not Specified", "OutlookController started");
                    }

                    //coreThread
                    Thread coreThread = new Thread(() => CoreThread());
                    coreThread.Start();
                    log.LogStandard("Not Specified", "CoreThread started");

                    log.LogStandard("Not Specified", "Core started");
                    coreRunning = true;
                    coreStatChanging = false;
                }
            }
            #region catch
            catch (Exception ex)
            {
                coreRunning = false;
                coreStatChanging = false;

                FatalExpection(t.GetMethodName() + " failed", ex);
                return false;
            }
            #endregion

            return true;
        }

        public override void    Restart(int secondsDelay)
        {
            try
            {
                if (coreRestarting == false)
                {
                    coreRestarting = true;

                    log.LogCritical("Not Specified", t.GetMethodName() + " called");
                    Close();
                    log.LogStandard("Not Specified", "Trying to restart the Core in " + secondsDelay + " seconds");
                    Thread.Sleep(secondsDelay * 1000);
                    Start(connectionString);

                    coreRestarting = false;
                }
            }
            catch (Exception ex)
            {
                FatalExpection(t.GetMethodName() + "failed. Core failed to restart", ex);
            }
        }

        public bool             Close()
        {
            try
            {
                if (coreRunning && !coreStatChanging)
                {
                    coreStatChanging = true;

                    coreThreadAlive = false;
                    log.LogCritical("Not Specified", t.GetMethodName() + " called");

                    int tries = 0;
                    while (coreRunning)
                    {
                        Thread.Sleep(100);
                        tries++;

                        if (tries > 600)
                            FatalExpection("Failed to close Core correct after 60 secs (coreRunning)", new Exception());
                    }

                    //tries = 0;
                    //while (syncOutlookAppsRunning)
                    //{
                    //    Thread.Sleep(100);
                    //    tries++;
                    //    if (tries > 300)
                    //        FatalExpection("Failed to close Core correct after 30 secs (syncOutlookAppsRunning)", new Exception());
                    //}

                    //tries = 0;
                    //while (syncInteractionCaseRunning)
                    //{
                    //    Thread.Sleep(100);
                    //    tries++;
                    //    if (tries > 300)
                    //        FatalExpection("Failed to close Core correct after 30 secs (syncInteractionCaseRunning)", new Exception());
                    //}

                    log.LogStandard("Not Specified", "Core closed");
                    outlookController = null;
                    sqlController = null;

                    coreStatChanging = false;
                }
            }
            catch (ThreadAbortException)
            {
                //"Even if you handle it, it will be automatically re-thrown by the CLR at the end of the try/catch/finally."
                Thread.ResetAbort(); //This ends the re-throwning
            }
            catch (Exception ex)
            {
                FatalExpection(t.GetMethodName() + "failed. Core failed to close", ex);
            }
            return true;
        }

        public bool             Running()
        {
            return coreRunning;
        }

        public override void    FatalExpection(string reason, Exception exception)
        {
            try
            {
                log.LogFatalException(t.GetMethodName() + " called for reason:'" + reason + "'", exception);
            }
            catch { }

            try
            {
                Thread coreRestartThread = new Thread(() => Close());
                coreRestartThread.Start();
            }
            catch { }

            coreRunning = false;
            coreStatChanging = false;

            try { HandleEventException?.Invoke(exception, EventArgs.Empty); } catch { }
            throw new Exception("FATAL exception, Core shutting down, due to:'" + reason + "'", exception);
        }
        #endregion

        #region private
        private void            CoreThread()
        {
            bool firstRun = true;

            coreRunning = true;
            coreThreadAlive = true;

            log.LogEverything("Not Specified", t.GetMethodName() + " initiated");
            while (coreThreadAlive)
            {
                try
                {
                    if (coreRunning)
                    {
                        #region warm up
                        log.LogEverything("Not Specified", t.GetMethodName() + " initiated");

                        if (firstRun)
                        {
                            outlookController.CalendarItemConvertRecurrences();
                            firstRun = false;
                            log.LogStandard("Not Specified", t.GetMethodName() + " warm up completed");
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

                        Thread.Sleep(2000);
                    }

                    Thread.Sleep(500);
                }
                catch (ThreadAbortException)
                {
                    log.LogWarning("Not Specified", t.GetMethodName() + " catch of ThreadAbortException");
                }
                catch (Exception ex)
                {
                    FatalExpection(t.GetMethodName() + "failed", ex);
                }
            }
            log.LogEverything("Not Specified", t.GetMethodName() + " completed");

            coreRunning = false;
            coreStatChanging = false;
        }

        private void            SyncOutlookConvert()
        {
            try
            {
                if (!syncOutlookConvertRunning)
                {
                    syncOutlookConvertRunning = true;

                    if (coreRunning)
                    {
                        while (coreRunning && outlookController.CalendarItemConvertRecurrences()) { }

                        log.LogEverything("Not Specified", "outlookController.CalendarItemIntrepid() completed");

                        for (int i = 0; i < 6 && coreRunning; i++)
                            Thread.Sleep(1000);
                    }
                    
                    syncOutlookConvertRunning = false;
                }
            }
            catch (ThreadAbortException)
            {
                log.LogWarning("Not Specified", t.GetMethodName() + " catch of ThreadAbortException");
            }
            catch (Exception ex)
            {
                syncOutlookConvertRunning = false;

                if (Running())
                    log.LogException("Not Specified", t.GetMethodName() + " failed", ex, true);
            }
        }

        private void            SyncOutlookApps()
        {
            try
            {
                if (!syncOutlookAppsRunning)
                {
                    syncOutlookAppsRunning = true;

                    while (coreRunning && outlookController.CalendarItemIntrepid())
                        log.LogEverything("Not Specified", "outlookController.CalendarItemIntrepid() completed");

                    while (coreRunning && outlookController.CalendarItemReflecting(null))
                        log.LogEverything("Not Specified", "outlookController.CalendarItemReflecting() completed");

                    syncOutlookAppsRunning = false;
                }
            }
            catch (ThreadAbortException)
            {
                log.LogWarning("Not Specified", t.GetMethodName() + " catch of ThreadAbortException");
            }
            catch (Exception ex)
            {
                syncOutlookAppsRunning = false;
                log.LogException("Not Specified", t.GetMethodName() + " failed", ex, true);
            }
        }

        private void            SyncInteractionCase()
        {
            try
            {
                if (!syncInteractionCaseRunning)
                {
                    syncInteractionCaseRunning = true;

                    while (coreRunning && sqlController.SyncInteractionCase())
                        log.LogEverything("Not Specified", t.GetMethodName() + " completed");

                    syncInteractionCaseRunning = false;
                }
            }
            catch (ThreadAbortException)
            {
                log.LogWarning("Not Specified", t.GetMethodName() + " catch of ThreadAbortException");
            }
            catch (Exception ex)
            {
                syncInteractionCaseRunning = false;
                log.LogException("Not Specified", t.GetMethodName() + " failed", ex, true);
            }
        }
        #endregion

        public void             Test_Reset(string connectionString)
        {
            sqlController = new SqlController(connectionString);
            Log log = sqlController.StartLog(this);
            outlookController = new OutlookController(sqlController, log);
            AdminTools at = new AdminTools(sqlController.SettingRead(Settings.microtingDb));

            try
            {
                if (!coreRunning && !coreStatChanging)
                {
                    coreStatChanging = true;
                    log.LogStandard("Not Specified", "Reset!");

                    List<Appointment> lstAppointments;

                    DateTime now = DateTime.Now;
                    DateTime rollBackTo__ = now.AddDays(+2);
                    DateTime rollBackFrom = now.AddDays(-3);

                    lstAppointments = outlookController.UnitTest_CalendarItemGetAllNonRecurring(rollBackFrom, rollBackTo__);
                    foreach (var item in lstAppointments)
                        outlookController.CalendarItemUpdate(item, WorkflowState.Planned, true);

                    sqlController.SettingUpdate(Settings.checkLast_At, now.ToString());

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
                FatalExpection(t.GetMethodName() + "failed. Core failed to restart", ex);
            }
            Close();
        }
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