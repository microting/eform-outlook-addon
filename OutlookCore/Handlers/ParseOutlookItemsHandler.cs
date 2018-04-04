using Microting.OutlookAddon.Messages;
using System.Threading.Tasks;
using Rebus.Handlers;
using System;
using eFormCore;
using OutlookOfficeOnline;
using OutlookSql;
using System.Collections.Generic;
using eFormShared;
using OutlookExchangeOnlineAPI;
using Rebus.Bus;

namespace Microting.OutlookAddon.Handlers
{
    public class ParseOutlookItemsHandler : IHandleMessages<ParseOutlookItem>
    {
        private readonly SqlController sqlController;
        private readonly Log log;
        private readonly Core sdkCore;
        private readonly IOutlookOnlineController outlookOnlineController;
        Tools t = new Tools();
        OutlookExchangeOnlineAPIClient outlookExchangeOnlineAPIClient;
        IBus bus;

        #region var
        DateTime checkLast_At; //DateTime.Parse(sqlController.SettingRead(Settings.checkLast_At));
        double checkPreSend_Hours;// = double.Parse(sqlController.SettingRead(Settings.checkPreSend_Hours));
        double checkRetrace_Hours;// = double.Parse(sqlController.SettingRead(Settings.checkRetrace_Hours));
        int checkEvery_Mins;// = int.Parse(sqlController.SettingRead(Settings.checkEvery_Mins));
        bool includeBlankLocations;// = bool.Parse(sqlController.SettingRead(Settings.includeBlankLocations));
        string userEmailAddess;


        DateTime timeOfRun;// = DateTime.Now;
        DateTime tLimitTo;// = timeOfRun.AddHours(+checkPreSend_Hours);
        DateTime tLimitFrom;// = checkLast_At.AddHours(-checkRetrace_Hours);
        #endregion

        public ParseOutlookItemsHandler(SqlController sqlController, Log log, Core sdkCore, IOutlookOnlineController outlookOnlineController, OutlookExchangeOnlineAPIClient outlookExchangeOnlineAPIClient, IBus bus)
        {
            this.sqlController = sqlController;
            this.log = log;
            this.sdkCore = sdkCore;
            this.outlookOnlineController = outlookOnlineController;
            this.outlookExchangeOnlineAPIClient = outlookExchangeOnlineAPIClient;
            this.bus = bus;
            //checkLast_At = DateTime.Parse(sqlController.SettingRead(Settings.checkLast_At));
            checkPreSend_Hours = double.Parse(sqlController.SettingRead(Settings.checkPreSend_Hours));
            checkRetrace_Hours = double.Parse(sqlController.SettingRead(Settings.checkRetrace_Hours));
            checkEvery_Mins = int.Parse(sqlController.SettingRead(Settings.checkEvery_Mins));
            includeBlankLocations = bool.Parse(sqlController.SettingRead(Settings.includeBlankLocations));

            userEmailAddess = outlookOnlineController.GetUserEmailAddress();

            timeOfRun = DateTime.Now;
            tLimitTo = timeOfRun.AddHours(+checkPreSend_Hours);
            tLimitFrom = timeOfRun.AddHours(-checkRetrace_Hours);
        }

#pragma warning disable 1998
        public async Task Handle(ParseOutlookItem message)
        {

            //#region processingState "planned"?
            string processingState = null;
            try
            {
                processingState = message.Item.Location.DisplayName;
            }
            catch { }


            if (string.IsNullOrEmpty(processingState))
            {
                if (includeBlankLocations)
                    processingState = "planned";
                else
                    processingState = "";
            }

            processingState = processingState.ToLower();
            //#endregion

            //if (processingState.ToLower() == "planned")
            //#region planned
            //{
                log.LogVariable("Not Specified", nameof(processingState), processingState);

                if (message.Item.BodyPreview != null)
                    if (message.Item.BodyPreview.Contains("<<< "))
                        if (message.Item.BodyPreview.Contains("End >>>"))
                        {
                            message.Item.BodyPreview = t.ReplaceAtLocationAll(message.Item.BodyPreview, "<<< ", "End >>>", "", true);
                            message.Item.BodyPreview = message.Item.BodyPreview.Replace("<<< End >>>", "");
                            message.Item.BodyPreview = message.Item.BodyPreview.Trim();
                        }

                log.LogStandard("Not Specified", "Trying to do UpdateEvent on item.Id:" + message.Item.Id + " to have new location location : " + processingState);
                Event updatedItem = outlookExchangeOnlineAPIClient.UpdateEvent(userEmailAddess, message.Item.Id, "{\"Location\": {\"DisplayName\": \"" + processingState + "\"},\"Body\": {\"ContentType\": \"HTML\",\"Content\": \"" + ReplaceLinesInBody(message.Item.BodyPreview) + "\"}}");

                //if (updatedItem == null)
                //{
                //    return false;
                //}

                log.LogStandard("Not Specified", "Trying create new appointment for item.Id : " + message.Item.Id + " and the UpdateEvent returned Updateditem: " + updatedItem.ToString());

                Appointment appo;
                int appoId = 0;
                appo = sqlController.AppointmentsFind(message.Item.Id);
                if (appo == null)
                {
                    appo = new Appointment(message.Item.Id, message.Item.Start.DateTime, (message.Item.End.DateTime - message.Item.Start.DateTime).Minutes, message.Item.Subject, "planned", updatedItem.BodyPreview, t.Bool(sqlController.SettingRead(Settings.colorsRule)), true, null);
                    appoId = sqlController.AppointmentsCreate(appo);
                }

                //log.LogStandard("Not Specified", "Before calling CalendarItemIntrepret.AppointmentsCreate");
                //log.LogStandard("Not Specified", "After calling CalendarItemIntrepret.AppointmentsCreate");

                if (appoId > 0)
                {
                    log.LogStandard("Not Specified", "Appointment created successfully for item.Id : " + message.Item.Id);
                    outlookOnlineController.CalendarItemUpdate(appo.GlobalId, appo.Start, ProcessingStateOptions.Processed, appo.Body);
                    bus.SendLocal(new AppointmentCreatedInOutlook(appo)).Wait();
            }
                else
                {
                    if (appoId == 0)
                    {
                        outlookOnlineController. CalendarItemUpdate(appo.GlobalId, appo.Start, ProcessingStateOptions.Exception, appo.Body);
                    }
                    if (appoId == -1)
                    {
                        log.LogStandard("Not Specified", "Appointment not created successfully for item.Id : " + message.Item.Id);

                        #region appo.Body = 'text'
                        appo.Body = "<<< Parsing error: Start >>>" +
                            Environment.NewLine + "Global ID already exists in the database." +
                            Environment.NewLine + "Indicating that this appointment has already been created." +
                            Environment.NewLine + "Likely course, is that you set the Appointment’s location to 'planned'/[blank] again." +
                            Environment.NewLine + "" +
                            Environment.NewLine + "If you wanted to a create a new appointment in the calendar:" +
                            Environment.NewLine + "- Create a new appointment in the calendar" +
                            Environment.NewLine + "- Create or copy the wanted details to the new appointment" +
                            Environment.NewLine + "" +
                            Environment.NewLine + "Item.Id :" + message.Item.Id +
                            Environment.NewLine + "<<< Parsing error: End >>>" +
                            Environment.NewLine + "" +
                            Environment.NewLine + appo.Body;
                        #endregion
                        outlookOnlineController.CalendarItemUpdate(appo.GlobalId, appo.Start, ProcessingStateOptions.ParsingFailed, appo.Body);
                    }
                }

                //AllParsed = true;
            //}
            //#endregion

            //if (processingState.ToLower() == "cancel")
            //#region cancel
            //{
            //    log.LogVariable("Not Specified", nameof(processingState), processingState);

            //    Appointment appo = new Appointment(message.Item.Id, message.Item.Start.DateTime, (message.Item.End.DateTime - message.Item.Start.DateTime).Minutes, message.Item.Subject, message.Item.Location.DisplayName, ReplaceLinesInBody(message.Item.BodyPreview), t.Bool(sqlController.SettingRead(Settings.colorsRule)), true, null);

            //    if (sqlController.AppointmentsCancel(appo.GlobalId))
            //        outlookOnlineController.CalendarItemUpdate(appo.GlobalId, appo.Start, ProcessingStateOptions.Canceled, appo.Body);
            //    else
            //        outlookOnlineController.CalendarItemUpdate(appo.GlobalId, appo.Start, ProcessingStateOptions.ParsingFailed, appo.Body);

            //    //AllParsed = true;
            //}
            //#endregion

            //if (processingState.ToLower() == "processed")
            //#region processed
            //{
            //    Appointment appo = sqlController.AppointmentsFind(message.Item.Id);

            //    log.LogStandard("Not Specified", "ParseCalendarItems appo start is : " + appo.Start.ToString());
            //    log.LogStandard("Not Specified", "ParseCalendarItems item start is : " + message.Item.Start.DateTime.ToString());
            //    log.LogStandard("Not Specified", "ParseCalendarItems appo end is : " + appo.End.ToString());
            //    log.LogStandard("Not Specified", "ParseCalendarItems item end is : " + message.Item.End.DateTime.ToString());
            //    if (appo.Start != message.Item.Start.DateTime)
            //    {
            //        log.LogStandard("Not Specified", "ParseCalendarItems updating calendar entry with globalId : " + appo.GlobalId);
            //        sqlController.AppointmentsUpdate(appo.GlobalId, ProcessingStateOptions.Processed, appo.Body, "", "", appo.Completed, message.Item.Start.DateTime, message.Item.End.DateTime, (message.Item.End.DateTime - message.Item.Start.DateTime).Minutes);
            //    }
            //}
            //#endregion

        }
        private string ReplaceLinesInBody(string BodyPreview)
        {
            return BodyPreview.Replace("\r\n", "<br/>"); ;
        }
    }
}
