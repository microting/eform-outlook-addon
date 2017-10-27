using eFormShared;
using OutlookSql;

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OutlookOffice
{
    public class OutlookController
    {
        #region var
        SqlController sqlController;
        IOutlookCommunicator communicator;
        Log log;
        Tools t = new Tools();
        #endregion

        #region con
        public                      OutlookController(SqlController sqlController, Log log)
        {
            this.sqlController = sqlController;
            this.log = log;

            string calendarName = sqlController.SettingRead(Settings.calendarName);



            if (calendarName == "unittest")
            {
                communicator = new OutlookCommunicator_Fake(sqlController, log);
                log.LogStandard("Not Specified", "OutlookController_Fake started");
            }
            else
            {
                communicator = new OutlookCommunicator_OutlookClient(calendarName, log);
                log.LogStandard("Not Specified", "OutlookController started");
            }
        }
        #endregion

        #region public
        public bool                 CalendarItemConvertRecurrences()
        {
            try
            {
                bool ConvertedAny = false;
                #region var
                DateTime checkLast_At       = DateTime.Parse(sqlController.SettingRead(Settings.checkLast_At));
                double checkPreSend_Hours   = double.Parse(sqlController.SettingRead(Settings.checkPreSend_Hours));
                double checkRetrace_Hours   = double.Parse(sqlController.SettingRead(Settings.checkRetrace_Hours));
                int checkEvery_Mins         = int.Parse(sqlController.SettingRead(Settings.checkEvery_Mins));
                bool includeBlankLocations  = bool.Parse(sqlController.SettingRead(Settings.includeBlankLocations));

                DateTime timeOfRun          = DateTime.Now;
                DateTime tLimitTo           = timeOfRun.AddHours(+checkPreSend_Hours);
                DateTime tLimitFrom         = checkLast_At.AddHours(-checkRetrace_Hours);
                #endregion

                // convert recurrences
                ConvertedAny = communicator.ConvertRecurringAppointments(timeOfRun, tLimitFrom, tLimitTo, checkLast_At, checkPreSend_Hours, checkRetrace_Hours, checkEvery_Mins, includeBlankLocations);

                if (ConvertedAny)
                    log.LogStandard  ("Not Specified", t.GetMethodName() + " completed + converted appointment(s)");
                else
                    log.LogEverything("Not Specified", t.GetMethodName() + " completed");

                return ConvertedAny;
            }
            catch (Exception ex)
            {
                throw new Exception(t.GetMethodName() + " failed", ex);
            }
        }

        public bool                 CalendarItemIntrepid()
        {
            try
            {
                bool AllIntrepid = false;
                #region var
                DateTime checkLast_At = DateTime.Parse(sqlController.SettingRead(Settings.checkLast_At));
                double checkPreSend_Hours = double.Parse(sqlController.SettingRead(Settings.checkPreSend_Hours));
                double checkRetrace_Hours = double.Parse(sqlController.SettingRead(Settings.checkRetrace_Hours));
                int checkEvery_Mins = int.Parse(sqlController.SettingRead(Settings.checkEvery_Mins));
                bool includeBlankLocations = bool.Parse(sqlController.SettingRead(Settings.includeBlankLocations));

                DateTime timeOfRun = DateTime.Now;
                DateTime tLimitTo = timeOfRun.AddHours(+checkPreSend_Hours);
                DateTime tLimitFrom = checkLast_At.AddHours(-checkRetrace_Hours);
                #endregion

                #region process appointments
                foreach (CalendarItem item in communicator.AppointmentItemReadAll(tLimitFrom, tLimitTo))
                {
                    #region item.Location "Planned"?
                    if (item.Location == null)
                        if (includeBlankLocations)
                            item.Location = "Planned";
                        else
                            item.Location = "";
                    #endregion
  
                    if (item.Location.ToLower() == "planned")
                    #region ...
                    {
                        log.LogVariable("Not Specified", nameof(item.Location), item.Location);

                        if (item.Body != null)
                            if (item.Body.Contains("<<< "))
                                if (item.Body.Contains("End >>>"))
                                {
                                    item.Body = t.ReplaceAtLocationAll(item.Body, "<<< ", "End >>>", "", true);
                                    item.Body = item.Body.Replace("<<< End >>>", "");
                                    item.Body = item.Body.Trim();
                                }
                  
                        Appointment appo = new Appointment(item.GlobalId, item.Start, item.Duration, item.Subject, item.Location, item.Body, t.Bool(sqlController.SettingRead(Settings.colorsRule)), true, sqlController.LookupRead);

                        if (appo.Location.ToLower() == "planned")
                        {
                            int count = sqlController.AppointmentsCreate(appo);

                            if (count > 0)
                                CalendarItemUpdate(appo.GlobalId, appo.Start, LocationOptions.Processed, appo.Body);
                            else
                            {
                                if (count == 0)
                                    CalendarItemUpdate(appo.GlobalId, appo.Start, LocationOptions.Exception, appo.Body);

                                if (count == -1)
                                {
                                    #region appo.Body = 'text'
                                    appo.Body =               "<<< Intrepid error: Start >>>" +
                                        Environment.NewLine + "Global ID already exists in the database." +
                                        Environment.NewLine + "Indicating that this appointment has already been created." +
                                        Environment.NewLine + "Likely course, is that you set the Appointment’s location to 'planned'/[blank] again." +
                                        Environment.NewLine + "" +
                                        Environment.NewLine + "If you wanted to a create a new appointment in the calendar:" +
                                        Environment.NewLine + "- Create a new appointment in the calendar" +
                                        Environment.NewLine + "- Create or copy the wanted details to the new appointment" +
                                        Environment.NewLine + "" +
                                        Environment.NewLine + "If you want to restore this appointment’s correct status:" +
                                        Environment.NewLine + "- Set the appointment’s location to 'check'" +
                                        Environment.NewLine + "<<< Intrepid error: End >>>" +
                                        Environment.NewLine + "" +
                                        Environment.NewLine + appo.Body;
                                    #endregion
                                    CalendarItemUpdate(appo.GlobalId, appo.Start, LocationOptions.Failed_to_intrepid, appo.Body);
                                }
                            }
                        }
                        else
                            CalendarItemUpdate(appo.GlobalId, appo.Start, LocationOptions.Failed_to_intrepid, appo.Body);

                        AllIntrepid = true;
                    }
                    #endregion

                    if (item.Location.ToLower() == "cancel")
                    #region ...
                    {
                        log.LogVariable("Not Specified", nameof(item.Location), item.Location);

                        Appointment appo = new Appointment(item.GlobalId, item.Start, item.Duration, item.Subject, item.Location, item.Body, t.Bool(sqlController.SettingRead(Settings.colorsRule)), true, sqlController.LookupRead);

                        if (sqlController.AppointmentsCancel(appo.GlobalId))
                            CalendarItemUpdate(appo.GlobalId, appo.Start, LocationOptions.Canceled, appo.Body);
                        else
                            CalendarItemUpdate(appo.GlobalId, appo.Start, LocationOptions.Failed_to_intrepid, appo.Body);

                        AllIntrepid = true;
                    }
                    #endregion

                    if (item.Location.ToLower() == "check")
                    #region ...
                    {
                        log.LogVariable("Not Specified", nameof(item.Location), item.Location);

                        eFormSqlController.SqlController sqlMicroting = new eFormSqlController.SqlController(sqlController.SettingRead(Settings.microtingDb));
                        eFormCommunicator.Communicator com = new eFormCommunicator.Communicator(sqlMicroting, log);

                        var temp = sqlController.AppointmentsFind(item.GlobalId);

                        var list = sqlMicroting.InteractionCaseListRead(int.Parse(temp.microting_uid));
                        foreach (var aCase in list)
                            com.CheckStatusUpdateIfNeeded(aCase.microting_uid);

                        CalendarItemReflecting(item.GlobalId);
                        AllIntrepid = true;
                    }
                    #endregion
                }
                #endregion

                sqlController.SettingUpdate(Settings.checkLast_At, timeOfRun.ToString());
                log.LogVariable("Not Specified", nameof(Settings.checkLast_At), timeOfRun.ToString());

                return AllIntrepid;
            }
            catch (Exception ex)
            {
                throw new Exception(t.GetMethodName() + " failed", ex);
            }
        }

        public bool                 CalendarItemReflecting(string globalId)
        {
            try
            {
                #region appointment = 'find one';
                appointments appointment = null;
     
                if (globalId == null)
                    appointment = sqlController.AppointmentsFindOne(0);
                else
                    appointment = sqlController.AppointmentsFind(globalId);

                if (appointment == null) //double checks status if no new found
                    appointment = sqlController.AppointmentsFindOne(1);
                #endregion

                if (appointment == null)
                    return false;
                log.LogVariable("Not Specified", nameof(appointments), appointment.ToString());

                CalendarItem item = new CalendarItem(appointment.global_id, appointment.location, (DateTime)appointment.start_at, (int)appointment.duration, appointment.subject, appointment.body);
                log.LogEverything("Not Specified", "CalenderItem created");
    
                #region item.Body = appointment.expectionString + item.Body + appointment.response ...
                if (!string.IsNullOrEmpty(appointment.response))
                {
                    if (t.Bool(sqlController.SettingRead(Settings.responseBeforeBody)))
                    {
                        item.Body = "<<< Response: Start >>>" +
                        Environment.NewLine +
                        Environment.NewLine + appointment.response +
                        Environment.NewLine +
                        Environment.NewLine + "<<< Response: End >>>" +
                        Environment.NewLine +
                        Environment.NewLine + item.Body;
                    }
                    else
                    {
                        item.Body = item.Body +
                        Environment.NewLine +
                        Environment.NewLine + "<<< Response: Start >>>" +
                        Environment.NewLine +
                        Environment.NewLine + appointment.response +
                        Environment.NewLine +
                        Environment.NewLine + "<<< Response: End >>>";
                    }
                }
                if (!string.IsNullOrEmpty(appointment.expectionString))
                {
                    item.Body = "<<< Exception: Start >>>" +
                    Environment.NewLine +
                    Environment.NewLine + appointment.expectionString +
                    Environment.NewLine +
                    Environment.NewLine + "<<< Exception: End >>>" +
                    Environment.NewLine +
                    Environment.NewLine + item.Body;
                }
                #endregion
                log.LogEverything("Not Specified", "Body composed");

                communicator.AppointmentItemUpdate(item);
                sqlController.AppointmentsReflected(appointment.global_id);
                log.LogStandard("Not Specified", "globalId:'" + appointment.global_id + "' reflected in database");

                return true;
            }
            catch (Exception ex)
            {
                throw new Exception(t.GetMethodName() + " failed", ex);
            }
        }

        public string               CalendarItemCreate(string location, DateTime start, int duration, string subject, string body)
        {
            try
            {
                CalendarItem newAppo = new CalendarItem("To be created", location, start, duration, subject, body);
                return communicator.AppointmentItemCreate(newAppo);
            }
            catch (Exception ex)
            {
                throw new Exception(t.GetMethodName() + " failed", ex);
            }
        }

        public bool                 CalendarItemUpdate(string globalId, DateTime start, LocationOptions location, string body)
        {
            try
            {
                var item = communicator.AppointmentItemRead(globalId, start);
                item.Location = location.ToString();
                item.Body = body;

                communicator.AppointmentItemUpdate(item);
                log.LogStandard("Not Specified", AppointmentPrint(item) + " updated to " + location.ToString());
                return true;
            }
            catch (Exception ex)
            {
                throw new Exception(t.GetMethodName() + " failed", ex);
            }
        }

        public bool                 CalendarItemDelete(string globalId, DateTime start)
        {
            try
            {
                return communicator.AppointmentItemDelete(globalId, start);
            }
            catch (Exception ex)
            {
                throw new Exception(t.GetMethodName() + " failed", ex);
            }
        }
        #endregion

        #region private
        private DateTime            RoundTime(DateTime dTime)
        {
            dTime = dTime.AddMinutes(1);
            dTime = new DateTime(dTime.Year, dTime.Month, dTime.Day, dTime.Hour, 0, 0);
            log.LogVariable("Not Specified", nameof(dTime), dTime);
            return dTime;
        }

        private string              AppointmentPrint(CalendarItem appItem)
        {
            return "GlobalId:" + appItem.GlobalId + " / Start:" + appItem.Start + " / Title:" + appItem.Subject;
        }


        


   
        #endregion

        #region unit test
        public List<Appointment>    UnitTest_CalendarItemGetAllNonRecurring(DateTime startPoint, DateTime endPoint)
        {
            try
            {
                List<Appointment> lstAppoint = new List<Appointment>();
                
                foreach (CalendarItem item in communicator.AppointmentItemReadAll(endPoint, startPoint))
                    if (item.Location != null)
                        if (startPoint <= item.Start && item.Start <= endPoint)
                            lstAppoint.Add(new Appointment(item.GlobalId, item.Start, item.Duration, item.Subject, item.Location, item.Body, t.Bool(sqlController.SettingRead(Settings.colorsRule)), true, sqlController.LookupRead));
         
                return lstAppoint;
            }
            catch (Exception ex)
            {
                throw new Exception(t.GetMethodName() + " failed", ex);
            }
        }

        public bool                 UnitTest_ForceException(string exceptionType)
        {
            throw new NotImplementedException();
        }

        private string              UnitTest_CalendarBody()
        {
            return
                                            "TempLate# " + "’Besked’"
                    + Environment.NewLine + "Sites# " + "’All’"
                    + Environment.NewLine + "title# " + "Outlook appointment eForm test"
                    + Environment.NewLine + "info# " + "Tekst fra Outlook appointment";
        }
        #endregion
    }

    public enum CalendarItemCategory
    {
        Completed,
        Error,
        Processing,
        Retrived,
        Revoked,
        Sent
    }
}