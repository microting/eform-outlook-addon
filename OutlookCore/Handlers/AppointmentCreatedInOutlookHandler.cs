using eFormCore;
using eFormShared;
using Microting.OutlookAddon.Messages;
using OutlookOfficeOnline;
using OutlookSql;
using Rebus.Handlers;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Microting.OutlookAddon.Handlers
{
    class AppointmentCreatedInOutlookHandler : IHandleMessages<AppointmentCreatedInOutlook>
    {
        private readonly SqlController sqlController;
        private readonly Log log;
        private readonly Core sdkCore;
        private readonly IOutlookOnlineController outlookOnlineController;

        public AppointmentCreatedInOutlookHandler(SqlController sqlController, Log log, Core sdkCore, IOutlookOnlineController outlookOnlineController)
        {
            this.sqlController = sqlController;
            this.log = log;
            this.sdkCore = sdkCore;
            this.outlookOnlineController = outlookOnlineController;

        }

#pragma warning disable 1998
        public async Task Handle(AppointmentCreatedInOutlook message)
        {
            try
            {
                bool sqlUpdate = sqlController.AppointmentsUpdate(message.Appo.GlobalId, ProcessingStateOptions.Created, message.Appo.Body, "", "", message.Appo.Completed, message.Appo.Start, message.Appo.End, message.Appo.Duration);
                if (sqlUpdate)
                {
                    bool updateStatus = outlookOnlineController.CalendarItemUpdate(message.Appo.GlobalId, (DateTime)message.Appo.Start, ProcessingStateOptions.Created, message.Appo.Body);
                }
                else
                {
                    throw new Exception("Unable to update Appointment in AppointmentCreatedInOutlookHandler");
                }
                //log.LogEverything("Not Specified", "outlookController.SyncAppointmentsToSdk() L623");
                eFormData.MainElement mainElement = sdkCore.TemplateRead((int)message.Appo.TemplateId);
                if (mainElement == null)
                {
                    log.LogEverything("Not Specified", "outlookController.SyncAppointmentsToSdk() L625 mainElement is NULL!!!");
                }

                mainElement.Repeated = 1;
                DateTime startDt = new DateTime(message.Appo.Start.Year, message.Appo.Start.Month, message.Appo.Start.Day, 0, 0, 0);
                DateTime endDt = new DateTime(message.Appo.End.AddDays(1).Year, message.Appo.End.AddDays(1).Month, message.Appo.End.AddDays(1).Day, 23, 59, 59);
                //mainElement.StartDate = ((DateTime)appo.Start).ToUniversalTime();
                mainElement.StartDate = startDt;
                //mainElement.EndDate = ((DateTime)appo.End.AddDays(1)).ToUniversalTime();
                mainElement.EndDate = endDt;
                //log.LogEverything("Not Specified", "outlookController.SyncAppointmentsToSdk() L629");
                log.LogEverything("Not Specified", "outlookController.SyncAppointmentsToSdk() StartDate is " + mainElement.StartDate);
                log.LogEverything("Not Specified", "outlookController.SyncAppointmentsToSdk() EndDate is " + mainElement.EndDate);

                bool allGood = false;
                List<AppoinntmentSite> appoSites = message.Appo.AppointmentSites;
                if (appoSites == null)
                {
                    log.LogEverything("Not Specified", "outlookController.SyncAppointmentsToSdk() L635 appoSites is NULL!!! for appo.GlobalId" + message.Appo.GlobalId);
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
                        bool updateStatus = outlookOnlineController.CalendarItemUpdate(message.Appo.GlobalId, (DateTime)message.Appo.Start, ProcessingStateOptions.Sent, message.Appo.Body);
                        if (updateStatus)
                        {
                            //log.LogEverything("Not Specified", "outlookController.SyncAppointmentsToSdk() L667");
                            sqlController.AppointmentsUpdate(message.Appo.GlobalId, ProcessingStateOptions.Sent, message.Appo.Body, "", "", message.Appo.Completed, message.Appo.Start, message.Appo.End, message.Appo.Duration);
                        }
                    }
                    else
                    {
                        //log.LogEverything("Not Specified", "outlookController.SyncAppointmentsToSdk() L673");
                        //syncAppointmentsToSdkRunning = false;
                    }
                }
            }
            catch (Exception ex)
            {
                log.LogEverything("Exception", "Got the following exception : " + ex.Message + " and stacktrace is : " + ex.StackTrace);
                throw ex;
            }

        }
    }
}
