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
using eFormCore;

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using OutlookOfficeOnline;
using OutlookExchangeOnlineAPI;

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
        //public OutlookController outlookController;
        public Log log;

        bool syncOutlookConvertRunning = false;
        bool syncOutlookAppsRunning = false;
        bool syncAppointmentsToSdkRunning = false;

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

                    //outlookController
                    //outlookController = new OutlookController(sqlController, log);
                    //outlookController
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

                    syncOutlookConvertRunning = false;
                    syncOutlookAppsRunning = false;
                    syncAppointmentsToSdkRunning = false;

                    log.LogStandard("Not Specified", "Core closed");
                    //outlookController = null;
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
        /// IMPORTANT: templateId, sites, startTime, duration, outlookTitle and eFormConnected are mandatory. Rest are optional, and should be passed 'null' if not wanted to use
        /// </summary>
        public string AppointmentCreate(int templateId, List<int> sites, DateTime startTime, int duration,
            string outlookTitle, string outlookCommentary, bool? outlookColorRuleOverride,
            bool eFormConnected, string eFormTitle, string eFormDescription, string eFormInfo, int? eFormDaysToExpire, List<string> eFormReplacements)
        {
            try
            {
                #region log everything...
                log.LogStandard("Not Specified", t.GetMethodName() + " called");
                log.LogVariable("Not Specified", nameof(templateId), templateId.ToString());
                log.LogVariable("Not Specified", nameof(startTime), startTime);
                log.LogVariable("Not Specified", nameof(duration), duration);
                log.LogVariable("Not Specified", nameof(outlookTitle), outlookTitle);
                log.LogVariable("Not Specified", nameof(outlookCommentary), outlookCommentary);
                log.LogVariable("Not Specified", nameof(outlookColorRuleOverride), outlookColorRuleOverride);
                log.LogVariable("Not Specified", nameof(eFormConnected), eFormConnected);
                log.LogVariable("Not Specified", nameof(eFormTitle), eFormTitle);
                log.LogVariable("Not Specified", nameof(eFormDescription), eFormDescription);
                log.LogVariable("Not Specified", nameof(eFormInfo), eFormInfo);
                log.LogVariable("Not Specified", nameof(eFormDaysToExpire), eFormDaysToExpire);
                #endregion

                #region needed
                if (templateId < 1)
                    throw new ArgumentException("templateId needs to be minimum 1");

                if (sites == null)
                    throw new ArgumentException("sites needs to be not null");
                if (sites.Count < 1)
                    throw new ArgumentException("sites.Count needs to be minimum 1");

                //---

                if (startTime == null)
                    throw new ArgumentException("startTime needs to be not null");
                if (startTime < DateTime.Now)
                    throw new ArgumentException("startTime needs to be a future DateTime");

                if (duration < 1)
                    throw new ArgumentException("duration needs to be minimum 1");

                if (string.IsNullOrWhiteSpace(outlookTitle))
                    throw new ArgumentException("outlookTitle needs to be not empty");

                if (eFormDaysToExpire != null)
                    if (eFormDaysToExpire < 1)
                        throw new ArgumentException("eFormDaysToExpire needs to be minimum 1");

                if (eFormReplacements != null)
                    foreach (var item in eFormReplacements)
                        if (!item.Contains("=="))
                            throw new ArgumentException("All eFormReplacements needs to contain '=='");
                #endregion

                #region body = ...
                string body = "";

                if (!string.IsNullOrWhiteSpace(outlookCommentary))
                    body = outlookCommentary + Environment.NewLine + Environment.NewLine;

                if (true)
                    body = body + "Template# " + templateId
                    + Environment.NewLine + "Sites# " + string.Join(",", sites)
                    + Environment.NewLine;

                if (!string.IsNullOrWhiteSpace(eFormTitle))
                    body += Environment.NewLine + "Title# " + eFormTitle;

                if (!string.IsNullOrWhiteSpace(eFormDescription))
                    body += Environment.NewLine + "Description# " + eFormDescription;

                if (!string.IsNullOrWhiteSpace(eFormInfo))
                    body += Environment.NewLine + "Info# " + eFormInfo;

                if (eFormConnected)
                    body += Environment.NewLine + "Connected# " + eFormConnected;

                if (eFormDaysToExpire != null)
                    body += Environment.NewLine + "Expire# " + eFormDaysToExpire;

                bool colorRule = t.Bool(sqlController.SettingRead(Settings.colorsRule));
                if (outlookColorRuleOverride != null)
                {
                    colorRule = (bool)outlookColorRuleOverride;
                    body += Environment.NewLine + "Color# " + colorRule.ToString();
                }

                if (eFormReplacements != null)
                    if (eFormReplacements.Count > 0)
                        foreach (var replacement in eFormReplacements)
                            body += Environment.NewLine + "Replacements# " + replacement;
                #endregion

                //string globalId = outlookController.CalendarItemCreate("Planned", startTime, duration, outlookTitle, body);
                TimeZone localZone = TimeZone.CurrentTimeZone;

                string globalId = outlookOnlineController.CalendarItemCreate("Planned", startTime, duration, outlookTitle, body, localZone.StandardName, localZone.StandardName);
                return globalId;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

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
                //if (appo == null)
                //{
                //    log.LogStandard("Not Specified", "No match found");
                //    return null;
                //}

                //Appointment appointment = 
                //    new Appointment(appo.GlobalId, t.Date(appo.start_at), t.Int(appo.duration), appo.subject, appo.processing_state, appo.body, t.Bool(appo.color_rule), true);

                //return appointment;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        /// <summary>
        /// No summary
        /// </summary>
        //public bool? AppointmentCancel(string globalId)
        //{
        //    try
        //    {
        //        log.LogStandard("Not Specified", t.GetMethodName() + " called");
        //        log.LogVariable("Not Specified", nameof(globalId), globalId);

        //        var appo = sqlController.AppointmentsFind(globalId);
        //        if (appo == null)
        //        {
        //            log.LogStandard("Not Specified", "No match found");
        //            return null;
        //        }

        //        return sqlController.AppointmentsUpdate(appo.GlobalId, ProcessingStateOptions.Canceled, appo.Body, null, null, appo.Completed);
        //    }
        //    catch (Exception ex)
        //    {
        //        throw ex;
        //    }
        //}

        /// <summary>
        /// No summary
        /// </summary>
        //public bool? AppointmentDelete(string globalId)
        //{
        //    try
        //    {
        //        log.LogStandard("Not Specified", t.GetMethodName() + " called");
        //        log.LogVariable("Not Specified", nameof(globalId), globalId);

        //        var appo = sqlController.AppointmentsFind(globalId);
        //        if (appo == null)
        //        {
        //            log.LogStandard("Not Specified", "No match found");
        //            return null;
        //        }

        //        if (outlookOnlineController.CalendarItemDelete(appo.GlobalId))
        //            return sqlController.AppointmentsUpdate(appo.GlobalId, ProcessingStateOptions.Canceled, appo.Body, null, null, appo.Completed);
        //        else
        //            return false;

        //        //if (outlookController.CalendarItemDelete(appo.global_id, (DateTime)appo.start_at))
        //        //    return sqlController.AppointmentsUpdate(appo.global_id, LocationOptions.Canceled, appo.body, null, null);
        //        //else
        //        //    return false;
        //    }
        //    catch (Exception ex)
        //    {
        //        throw ex;
        //    }
        //}

        public bool MarkAppointmentRetrived(string caseId)
        {
            //log.LogEverything("Not Specified", "outlookController.MarkAppointmentRetrived() L471");
            //log.LogEverything("Not Specified", "outlookController.MarkAppointmentRetrived() L471 caseId is " + caseId);
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
                //log.LogEverything("Not Specified", "outlookController.MarkAppointmentRetrived() L476");
                sqlController.AppointmentsUpdate(appo.GlobalId, ProcessingStateOptions.Retrived, appo.Body, "", "", false, appo.Start, appo.End, appo.Duration);
                sqlController.AppointmentSiteUpdate((int)appo.AppointmentSites.First().Id, caseId, ProcessingStateOptions.Retrived);
                return true;
            }
            else
            {
                //log.LogEverything("Not Specified", "outlookController.MarkAppointmentRetrived() L482");
                return false;
            }

        }

        public bool MarkAppointmentCompleted(string caseId)
        {
            //log.LogEverything("Not Specified", "outlookController.MarkAppointmentCompleted() L490");
            Appointment appo = sqlController.AppointmentFindCaseId(caseId);
            bool result = false;

            //bool result = outlookOnlineController.CalendarItemUpdate(appo.GlobalId, appo.Start, ProcessingStateOptions.Completed, appo.Body);
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
                //log.LogEverything("Not Specified", "outlookController.MarkAppointmentCompleted() L495");
                sqlController.AppointmentsUpdate(appo.GlobalId, ProcessingStateOptions.Completed, appo.Body, "", "", true, appo.Start, appo.End, appo.Duration);
                sqlController.AppointmentSiteUpdate((int)appo.AppointmentSites.First().Id, caseId, ProcessingStateOptions.Completed);
                return true;
            }
            else
            {
                //log.LogEverything("Not Specified", "outlookController.MarkAppointmentCompleted() L502");
                return false;
            }
        }


        #region parsing threads
        private void CoreThread(string sdkCoreConnectionString)
        {
            bool firstRun = true;
            coreThreadRunning = true;

            log.LogEverything("Not Specified", t.GetMethodName() + " initiated");
            //while (coreAvailable)
            //{
            try
            {
                //if (coreThreadRunning)
                //{
                #region warm up
                //log.LogEverything("Not Specified", t.GetMethodName() + " initiated");

                if (firstRun)
                {
                    //outlookController.CalendarItemConvertRecurrences();
                    outlookOnlineController.CalendarItemConvertRecurrences();
                    //bool 
                    //while ()
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
                                //string result = sdkCore.CaseCheck(appo_site.MicrotingUuId);
                                //log.LogEverything("Not Specified", t.GetMethodName() + " kase IS NULL!");
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
                                //sdkCore.CaseCheck(appo_site.MicrotingUuId);
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

                Thread.Sleep(2000);
                //}

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
            //}
            //log.LogEverything("Not Specified", t.GetMethodName() + " completed");

            //coreThreadRunning = false;
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
                //if (!syncOutlookConvertRunning)
                //{
                //    syncOutlookConvertRunning = true;

                //    if (coreThreadRunning)
                //    {
                //        while (coreThreadRunning && outlookOnlineController.CalendarItemConvertRecurrences()) { }
                //        //while (coreThreadRunning && outlookController.CalendarItemConvertRecurrences()) { }

                //        log.LogEverything("Not Specified", "outlookController.CalendarItemIntrepid() completed");

                //        for (int i = 0; i < 6 && coreThreadRunning; i++)
                //            Thread.Sleep(1000);
                //    }

                //    syncOutlookConvertRunning = false;
                //}
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
                //if (!syncOutlookAppsRunning)
                //{
                //    syncOutlookAppsRunning = true;

                //    //while (coreThreadRunning && outlookController.CalendarItemIntrepid())
                //    //    log.LogEverything("Not Specified", "outlookController.CalendarItemIntrepid() completed");

                //    //while (coreThreadRunning && outlookController.CalendarItemReflecting(null))
                //    //    log.LogEverything("Not Specified", "outlookController.CalendarItemReflecting() completed");

                //    while (coreThreadRunning && outlookOnlineController.ParseCalendarItems())
                //        log.LogEverything("Not Specified", "outlookController.CalendarItemIntrepid() completed");

                //    while (coreThreadRunning && outlookOnlineController.CalendarItemReflecting(null))
                //        log.LogEverything("Not Specified", "outlookController.CalendarItemReflecting() completed");

                //    syncOutlookAppsRunning = false;
                //}
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
            //log.LogEverything("Not Specified", "outlookController.SyncAppointmentsToSdk() L594 sdkConnectionString is " + sdkConnectionString);
            try
            {
                //log.LogEverything("Not Specified", "outlookController.SyncAppointmentsToSdk() L597");

                //if (!syncAppointmentsToSdkRunning)
                //{
                //    log.LogEverything("Not Specified", "outlookController.SyncAppointmentsToSdk() L601");

                //    syncAppointmentsToSdkRunning = true;

                if (sdkCore == null)
                {
                    //log.LogEverything("Not Specified", "outlookController.SyncAppointmentsToSdk() L607");

                    startSdkCore(sdkConnectionString);
                }
                //log.LogEverything("Not Specified", "outlookController.SyncAppointmentsToSdk() L611");

                string serverAddress = sdkCore.GetHttpServerAddress();
                //log.LogEverything("Not Specified", "outlookController.SyncAppointmentsToSdk() L614");

                while (coreThreadRunning)
                {
                    //log.LogEverything("Not Specified", "outlookController.SyncAppointmentsToSdk() L618");
                    Appointment appo = sqlController.AppointmentsFindOne(ProcessingStateOptions.Processed, true, null);
                    //log.LogEverything("Not Specified", "outlookController.SyncAppointmentsToSdk() L620");
                    if (appo != null)
                    {
                        //log.LogEverything("Not Specified", "outlookController.SyncAppointmentsToSdk() L623");
                        eFormData.MainElement mainElement = sdkCore.TemplateRead((int)appo.TemplateId);
                        if (mainElement == null)
                        {
                            log.LogEverything("Not Specified", "outlookController.SyncAppointmentsToSdk() L625 mainElement is NULL!!!");
                        }

                        mainElement.Repeated = 1;
                        DateTime startDt = new DateTime(appo.Start.Year, appo.Start.Month, appo.Start.Day, 0, 0, 0);
                        DateTime endDt = new DateTime(appo.End.AddDays(1).Year, appo.End.AddDays(1).Month, appo.End.AddDays(1).Day, 23, 59, 59);
                        //mainElement.StartDate = ((DateTime)appo.Start).ToUniversalTime();
                        mainElement.StartDate = startDt;
                        //mainElement.EndDate = ((DateTime)appo.End.AddDays(1)).ToUniversalTime();
                        mainElement.EndDate = endDt;
                        //log.LogEverything("Not Specified", "outlookController.SyncAppointmentsToSdk() L629");
                        log.LogEverything("Not Specified", "outlookController.SyncAppointmentsToSdk() StartDate is " + mainElement.StartDate);
                        log.LogEverything("Not Specified", "outlookController.SyncAppointmentsToSdk() EndDate is " + mainElement.EndDate);

                        bool allGood = false;
                        List<AppoinntmentSite> appoSites = appo.AppointmentSites;
                        if (appoSites == null)
                        {
                            log.LogEverything("Not Specified", "outlookController.SyncAppointmentsToSdk() L635 appoSites is NULL!!! for appo.GlobalId" + appo.GlobalId);
                        }
                        else
                        {
                            foreach (AppoinntmentSite appo_site in appoSites)
                            {
                                log.LogEverything("Not Specified", "outlookController.foreach AppoinntmentSite appo_site is : " + appo_site.MicrotingSiteUid);
                                //List<int> siteIds = new List<int>();
                                //siteIds.Add(appo_site.MicrotingSiteUid);
                                string resultId = sdkCore.CaseCreate(mainElement, "", appo_site.MicrotingSiteUid);
                                log.LogEverything("Not Specified", "outlookController.foreach resultId is : " + resultId);
                                int localCaseId = (int)sdkCore.CaseLookupMUId(resultId).CaseId;

                                //string localCaseId = .First().CaseId.ToString();
                                if (!string.IsNullOrEmpty(resultId))
                                {
                                    //log.LogEverything("Not Specified", "outlookController.SyncAppointmentsToSdk() L649");
                                    appo_site.MicrotingUuId = resultId;
                                    sqlController.AppointmentSiteUpdate((int)appo_site.Id, localCaseId.ToString(), ProcessingStateOptions.Sent);
                                    allGood = true;
                                }
                                else
                                {
                                    log.LogEverything("Not Specified", "outlookController.SyncAppointmentsToSdk() L656");
                                    allGood = false;
                                }
                            }

                            if (allGood)
                            {
                                //log.LogEverything("Not Specified", "outlookController.SyncAppointmentsToSdk() L663");
                                bool updateStatus = outlookOnlineController.CalendarItemUpdate(appo.GlobalId, (DateTime)appo.Start, ProcessingStateOptions.Sent, appo.Body);
                                if (updateStatus)
                                {
                                    //log.LogEverything("Not Specified", "outlookController.SyncAppointmentsToSdk() L667");
                                    sqlController.AppointmentsUpdate(appo.GlobalId, ProcessingStateOptions.Sent, appo.Body, "", "", appo.Completed, appo.Start, appo.End, appo.Duration);
                                }
                            }
                            else
                            {
                                //log.LogEverything("Not Specified", "outlookController.SyncAppointmentsToSdk() L673");
                                //syncAppointmentsToSdkRunning = false;
                            }
                        }
                    }
                    else
                    {
                        //log.LogEverything("Not Specified", "outlookController.SyncAppointmentsToSdk() L679");
                        Thread.Sleep(5000); // This is done, so if we don't find an appointment, we don't hammer the db
                                            // TODO find better way of solving this.
                    }
                    log.LogEverything("Not Specified", t.GetMethodName() + " completed");
                }
                //syncAppointmentsToSdkRunning = false;
                //}
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

        #region unit test
        //internal void UnitTest_SetUnittest()
        //{
        //    skipRestartDelay = true;
        //}

        //internal bool UnitTest_CoreDead()
        //{
        //    if (!coreAvailable)
        //        if (!coreStatChanging)
        //            if (!coreRestarting)
        //                if (!coreThreadRunning)
        //                    return true;
        //    return false;
        //}

        //public void UnitTest_Reset(string connectionString)
        //{
        //    sqlController = new SqlController(connectionString);
        //    Log log = sqlController.StartLog(this);
        //    //outlookController = new OutlookController(sqlController, log);
        //    outlookOnlineController = new OutlookOnlineController(sqlController, log, outlookExchangeOnlineAPI);
        //    AdminTools at = new AdminTools(sqlController.SettingRead(Settings.microtingDb));

        //    try
        //    {
        //        if (!coreThreadRunning && !coreStatChanging)
        //        {
        //            coreStatChanging = true;
        //            log.LogStandard("Not Specified", "Reset!");

        //            List<Appointment> lstAppointments;

        //            DateTime now = DateTime.Now;
        //            DateTime rollBackTo__ = now.AddDays(+2);
        //            DateTime rollBackFrom = now.AddDays(-3);

        //            //lstAppointments = outlookController.UnitTest_CalendarItemGetAllNonRecurring(rollBackFrom, rollBackTo__);
        //            lstAppointments = outlookOnlineController.UnitTest_CalendarItemGetAllNonRecurring(rollBackFrom, rollBackTo__);

        //            foreach (var item in lstAppointments)
        //                //outlookController.CalendarItemUpdate(item.GlobalId, item.Start, LocationOptions.Planned, item.Body);
        //                outlookOnlineController.CalendarItemUpdate(item.GlobalId, item.Start, ProcessingStateOptions.Planned, item.Body);

        //            sqlController.SettingUpdate(Settings.checkLast_At, now.ToString());

        //            at.RetractEforms();

        //            sqlController.UnitTest_TruncateTable("appointment_versions");
        //            sqlController.UnitTest_TruncateTable("appointments");
        //            //sqlController.UnitTest_TruncateTable_Microting("a_interaction_case_lists");
        //            //sqlController.UnitTest_TruncateTable_Microting("a_interaction_cases");
        //            //sqlController.UnitTest_TruncateTable_Microting("notifications");
        //            //sqlController.UnitTest_TruncateTable_Microting("cases");

        //            coreStatChanging = false;
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        FatalExpection(t.GetMethodName() + "failed. Core failed to restart", ex);
        //    }
        //    Close();
        //}
        #endregion

        public void startSdkCore(string sdkConnectionString)
        {
            //string[] lines;
            //try
            //{
            //    lines =
            //        System.IO.File.ReadAllLines(System.Web.Hosting.HostingEnvironment.MapPath("~/bin/Input.txt"));

            //    if (lines[0].IsEmpty())
            //    {
            //        throw new Exception();
            //    }
            //}
            //catch (Exception)
            //{
            //    throw new HttpResponseException(HttpStatusCode.Unauthorized);
            //}


            //string connectionStr = lines.First();

            this.sdkCore = new eFormCore.Core();
            //bool running = false;
            //_core.HandleCaseCreated += EventCaseCreated;
            //_core.HandleCaseRetrived += EventCaseRetrived;
            //_core.HandleCaseCompleted += EventCaseCompleted;
            //_core.HandleCaseDeleted += EventCaseDeleted;
            //_core.HandleFileDownloaded += EventFileDownloaded;
            //_core.HandleSiteActivated += EventSiteActivated;
            //_core.HandleEventLog += EventLog;
            //_core.HandleEventMessage += EventMessage;
            //_core.HandleEventWarning += EventWarning;
            //_core.HandleEventException += EventException;

            //try
            //{
            sdkCore.StartSqlOnly(sdkConnectionString);
            //}
            //catch (Exception ex)
            //{
            //    AdminTools adminTools = new AdminTools(connectionStr);
            //    adminTools.MigrateDb();
            //    adminTools.DbSettingsReloadRemote();
            //    running = _core.StartSqlOnly(connectionStr);
            //}

            //if (running)
            //{
            //    return this.sdkCore;
            //}
            ////Logger.Error("Core is not running");
            //throw new Exception("Core is not running");
            //return null;
        }
    }
}