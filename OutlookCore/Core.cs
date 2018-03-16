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

//using OutlookOffice;
using OutlookSql;
using eFormShared;

using System;
using System.Linq;
using System.Threading;
using OutlookOfficeOnline;
using OutlookExchangeOnlineAPI;
using Rebus.Bus;
using Castle.MicroKernel.Registration;
using Castle.Windsor;
using Microting.OutlookAddon.Installers;
using Microting.OutlookAddon.Messages;

namespace OutlookCore
{
    public class Core : CoreBase
    {
        //events
        public event EventHandler HandleEventException;

        #region var
        SqlController sqlController;
        Tools t = new Tools();
        private eFormCore.Core sdkCore;
        OutlookExchangeOnlineAPIClient outlookExchangeOnlineAPI;
        public IOutlookOnlineController outlookOnlineController;
        public Log log;
        IWindsorContainer container;
        public IBus bus;

        bool coreThreadRunning = false;
        bool coreRestarting = false;
        bool coreStatChanging = false;
        bool coreAvailable = false;

        bool skipRestartDelay = false;

        string connectionString;

        int sameExceptionCountTried = 0;
        string serviceLocation;
        #endregion

        //con
        #region public state
        public bool Start(string connectionString, string service_location)
        {
            try
            {
                if (!coreAvailable && !coreStatChanging)
                {
                    serviceLocation = service_location;
                    coreStatChanging = true;

                    if (string.IsNullOrEmpty(serviceLocation))
                        throw new ArgumentException("serviceLocation is not allowed to be null or empty");

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
                    if (log == null)
                        log = sqlController.StartLog(this);

                    log.LogCritical("Not Specified", "###########################################################################");
                    log.LogCritical("Not Specified", t.GetMethodName() + " called");
                    log.LogStandard("Not Specified", "SqlController and Logger started");

                    //settings read
                    this.connectionString = connectionString;
                    log.LogStandard("Not Specified", "Settings read");
                    log.LogStandard("Not Specified", "this.serviceLocation is " + serviceLocation);

                    //Initialise Outlook API client's object
                    if (sqlController.SettingRead(Settings.calendarName) == "unittest")
                    {
                        outlookOnlineController = new OutlookOnlineController_Fake(sqlController, log);
                        log.LogStandard("Not Specified", "OutlookController_Fake started");
                    }
                    else
                    {
                        outlookExchangeOnlineAPI = new OutlookExchangeOnlineAPIClient(serviceLocation, log);
                        outlookOnlineController = new OutlookOnlineController(sqlController, log, outlookExchangeOnlineAPI);
                        log.LogStandard("Not Specified", "OutlookController started");
                    }
                    log.LogStandard("Not Specified", "OutlookController started");

                    log.LogCritical("Not Specified", t.GetMethodName() + " started");
                    coreAvailable = true;
                    coreStatChanging = false;

                    //coreThread
                    string sdkCoreConnectionString = sqlController.SettingRead(Settings.microtingDb);
                    startSdkCore(sdkCoreConnectionString);

                    container = new WindsorContainer();
                    container.Register(Component.For<SqlController>().Instance(sqlController));
                    container.Register(Component.For<eFormCore.Core>().Instance(sdkCore));
                    container.Register(Component.For<Log>().Instance(log));
                    container.Register(Component.For<IOutlookOnlineController>().Instance(outlookOnlineController));
                    container.Install(
                        new RebusHandlerInstaller()
                        , new RebusInstaller(connectionString)
                    );


                    this.bus = container.Resolve<IBus>();

                    Thread coreThread = new Thread(() => CoreThread(sdkCoreConnectionString));
                    coreThread.Start();
                    log.LogStandard("Not Specified", "CoreThread started");
                }
            }
            #region catch
            catch (Exception ex)
            {
                FatalExpection(t.GetMethodName() + " failed", ex);
                return false;
            }
            #endregion

            return true;
        }

        public override void Restart(int sameExceptionCount, int sameExceptionCountMax)
        {
            try
            {
                if (coreRestarting == false)
                {
                    coreRestarting = true;
                    log.LogCritical("Not Specified", t.GetMethodName() + " called");
                    log.LogVariable("Not Specified", nameof(sameExceptionCount), sameExceptionCount);
                    log.LogVariable("Not Specified", nameof(sameExceptionCountMax), sameExceptionCountMax);

                    sameExceptionCountTried++;

                    if (sameExceptionCountTried > sameExceptionCountMax)
                        sameExceptionCountTried = sameExceptionCountMax;

                    if (sameExceptionCountTried > 4)
                        throw new Exception("The same Exception repeated to many times (5+) within one hour");

                    int secondsDelay = 0;
                    switch (sameExceptionCountTried)
                    {
                        case 1: secondsDelay = 001; break;
                        case 2: secondsDelay = 008; break;
                        case 3: secondsDelay = 064; break;
                        case 4: secondsDelay = 512; break;
                        default: throw new ArgumentOutOfRangeException("sameExceptionCount should be above 0");
                    }
                    log.LogVariable("Not Specified", nameof(sameExceptionCountTried), sameExceptionCountTried);
                    log.LogVariable("Not Specified", nameof(secondsDelay), secondsDelay);

                    Close();

                    log.LogStandard("Not Specified", "Trying to restart the Core in " + secondsDelay + " seconds");

                    if (!skipRestartDelay)
                        Thread.Sleep(secondsDelay * 1000);
                    else
                        log.LogStandard("Not Specified", "Delay skipped");
                    sdkCore.Close();

                    Start(connectionString, serviceLocation);
                    coreRestarting = false;
                }
            }
            catch (Exception ex)
            {
                FatalExpection(t.GetMethodName() + "failed. Core failed to restart", ex);
            }
        }

        public bool Close()
        {
            try
            {
                if (coreAvailable && !coreStatChanging)
                {
                    coreStatChanging = true;

                    coreAvailable = false;
                    log.LogCritical("Not Specified", t.GetMethodName() + " called");

                    int tries = 0;
                    while (coreThreadRunning)
                    {
                        Thread.Sleep(100);
                        tries++;

                        if (tries > 600)
                            FatalExpection("Failed to close Core correct after 60 secs (coreRunning)", new Exception());
                    }

                    log.LogStandard("Not Specified", "Core closed");
                    outlookOnlineController = null;
                    sqlController = null;
                    sdkCore.Close();

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

        public bool Running()
        {
            return coreAvailable;
        }

        public void FatalExpection(string reason, Exception exception)
        {
            coreAvailable = false;
            coreThreadRunning = false;
            coreStatChanging = false;
            coreRestarting = false;

            try
            {
                log.LogFatalException(t.GetMethodName() + " called for reason:'" + reason + "'", exception);
            }
            catch { }

            try { HandleEventException?.Invoke(exception, EventArgs.Empty); } catch { }
            throw new Exception("FATAL exception, Core shutting down, due to:'" + reason + "'", exception);
        }
        #endregion

        /// <summary>
        /// No summary
        /// </summary>
        public Appointment AppointmentRead(string globalId)
        {
            try
            {
                log.LogStandard("Not Specified", t.GetMethodName() + " called");
                log.LogVariable("Not Specified", nameof(globalId), globalId);

                return sqlController.AppointmentsFind(globalId);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public bool MarkAppointmentRetrived(string caseId)
        {
            Appointment appo = sqlController.AppointmentFindCaseId(caseId);
            bool result = false;
            try
            {
                result = outlookOnlineController.CalendarItemUpdate(appo.GlobalId, appo.Start, ProcessingStateOptions.Retrived, appo.Body);
            }
            catch (Exception ex)
            {
                if (ex.Message.Equals("Item not found"))
                {
                    result = true;
                }
            }

            if (result)
            {
                sqlController.AppointmentsUpdate(appo.GlobalId, ProcessingStateOptions.Retrived, appo.Body, "", "", false, appo.Start, appo.End, appo.Duration);
                sqlController.AppointmentSiteUpdate((int)appo.AppointmentSites.First().Id, caseId, ProcessingStateOptions.Retrived);
                return true;
            }
            else
            {
                return false;
            }

        }

        public bool MarkAppointmentCompleted(string caseId)
        {
            Appointment appo = sqlController.AppointmentFindCaseId(caseId);
            bool result = false;

            try
            {
                result = outlookOnlineController.CalendarItemUpdate(appo.GlobalId, appo.Start, ProcessingStateOptions.Completed, appo.Body);
            }
            catch (Exception ex)
            {
                if (ex.Message.Equals("Item not found"))
                {
                    result = true;
                }
            }
            if (result)
            {
                sqlController.AppointmentsUpdate(appo.GlobalId, ProcessingStateOptions.Completed, appo.Body, "", "", true, appo.Start, appo.End, appo.Duration);
                sqlController.AppointmentSiteUpdate((int)appo.AppointmentSites.First().Id, caseId, ProcessingStateOptions.Completed);
                return true;
            }
            else
            {
                return false;
            }
        }

        #region parsing threads
        private void CoreThread(string sdkCoreConnectionString)
        {
            bool firstRun = true;
            coreThreadRunning = true;

            log.LogEverything("Not Specified", t.GetMethodName() + " initiated");
            try
            {
                #region warm up
                if (firstRun)
                {
                    outlookOnlineController.CalendarItemConvertRecurrences();
                    int? currentId = null;
                    while (firstRun)
                    {
                        if (sdkCore == null)
                        {
                            startSdkCore(sdkCoreConnectionString);
                        }
                        log.LogEverything("Not Specified", t.GetMethodName() + " checking Appointments which are sent and currentId is now " + currentId.ToString());
                        Appointment appo = sqlController.AppointmentsFindOne(ProcessingStateOptions.Sent, false, currentId);
                        if (appo != null)
                        {
                            foreach (AppoinntmentSite appo_site in appo.AppointmentSites)
                            {
                                log.LogEverything("Not Specified", t.GetMethodName() + " checking appointment_site with MicrotingUuId : " + appo_site.MicrotingUuId.ToString());
                                Case_Dto kase = sdkCore.CaseReadByCaseId(int.Parse(appo_site.MicrotingUuId));
                                if (kase == null)
                                {
                                    log.LogEverything("Not Specified", t.GetMethodName() + " kase IS NULL!");
                                    //firstRun = false;
                                }

                                if (kase.Stat == "Retrived")
                                {
                                    MarkAppointmentRetrived(kase.CaseId.ToString());
                                }
                                else if (kase.Stat == "Completed")
                                {
                                    MarkAppointmentCompleted(kase.CaseId.ToString());
                                }
                                else
                                {
                                    currentId = appo_site.Id;
                                }
                            }

                        }
                        else
                        {
                            firstRun = false;
                        }
                    }

                    log.LogStandard("Not Specified", t.GetMethodName() + " warm up completed");
                }
                #endregion

                Thread syncOutlookConvertThread
                    = new Thread(() => SyncOutlookConvert());
                syncOutlookConvertThread.Start(); // This thread takes recurring events and convert the needed ones into single events.

                Thread syncOutlookAppsThread
                    = new Thread(() => SyncOutlookApps());
                syncOutlookAppsThread.Start(); // This thread takes single events and create the corresponding Appointment

                #region TODO
                Thread syncAppointmentsToSdk
                    = new Thread(() => SyncAppointmentsToSdk(sdkCoreConnectionString));
                syncAppointmentsToSdk.Start();
                #endregion

                Thread.Sleep(2500);
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

        private void SyncOutlookConvert()
        {
            try
            {

                while (coreThreadRunning)
                {
                    outlookOnlineController.CalendarItemConvertRecurrences();
                    log.LogEverything("Not Specified", "outlookController.CalendarItemIntrepid() done and sleeping for 2 seconds");
                    Thread.Sleep(2000);
                }
            }
            catch (ThreadAbortException)
            {
                log.LogWarning("Not Specified", t.GetMethodName() + " catch of ThreadAbortException");
            }
            catch (Exception ex)
            {
                log.LogException("Not Specified", t.GetMethodName() + " failed", ex, true);
            }
        }

        private void SyncOutlookApps()
        {
            try
            {
                while (coreThreadRunning)
                {
                    outlookOnlineController.ParseCalendarItems();
                    log.LogEverything("Not Specified", "outlookController.CalendarItemIntrepid() completed");
                    outlookOnlineController.CalendarItemReflecting(null);
                    log.LogEverything("Not Specified", "outlookController.CalendarItemReflecting() completed");
                    log.LogEverything("Not Specified", "outlookController.SyncOutlookApps() done and sleeping for 2 seconds");
                    Thread.Sleep(2000);
                }
            }
            catch (ThreadAbortException)
            {
                log.LogWarning("Not Specified", t.GetMethodName() + " catch of ThreadAbortException");
            }
            catch (Exception ex)
            {
                log.LogException("Not Specified", t.GetMethodName() + " failed", ex, true);
            }
        }

        private void SyncAppointmentsToSdk(string sdkConnectionString)
        {
            try
            {

                if (sdkCore == null)
                {
                    startSdkCore(sdkConnectionString);
                }

                string serverAddress = sdkCore.GetHttpServerAddress();

                Appointment appo = null;
                while (coreThreadRunning)
                {
                    if (appo != null)
                    {
                        appo = sqlController.AppointmentsFindOne(ProcessingStateOptions.Processed, true, appo.Id);
                    }
                    else
                    {
                        appo = sqlController.AppointmentsFindOne(ProcessingStateOptions.Processed, true, null);
                    }

                    if (appo != null)
                    {
                        bus.SendLocal(new AppointmentCreatedInOutlook(appo)).Wait();
                    }
                    else
                    {
                        Thread.Sleep(5000); // This is done, so if we don't find an appointment, we don't hammer the db
                                            // TODO find better way of solving this.
                    }
                    log.LogEverything("Not Specified", t.GetMethodName() + " completed");
                }
            }
            catch (ThreadAbortException)
            {
                log.LogWarning("Not Specified", t.GetMethodName() + " catch of ThreadAbortException");
            }
            catch (Exception ex)
            {
                log.LogException("Not Specified", t.GetMethodName() + " failed", ex, true);
            }
        }
        #endregion

        public void startSdkCore(string sdkConnectionString)
        {
            this.sdkCore = new eFormCore.Core();

            sdkCore.StartSqlOnly(sdkConnectionString);
        }
    }
}