using eFormShared;
using OutlookSql;
using Outlook = Microsoft.Office.Interop.Outlook;

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OutlookOffice
{
    public class OutlookController : IOutlookController
    {
        #region var
        string calendarName;
        Outlook.MAPIFolder calendarFolder = null;
        SqlController sqlController;
        Log log;
        Tools t = new Tools();
        object _lockOutlook = new object();
        #endregion

        #region con
        public                      OutlookController(SqlController sqlController, Log log)
        {
            this.sqlController = sqlController;
            this.log = log;
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

                #region convert recurrences
                foreach (Outlook.AppointmentItem item in GetCalendarItems(tLimitTo, tLimitFrom))
                {
                    if (item.IsRecurring) //is recurring, otherwise ignore
                    {
                        #region location "planned"?
                        string location = item.Location;

                        if (location == null)
                        {
                            if (includeBlankLocations)
                                location = "planned";
                            else
                                location = "";
                        }

                        location = location.ToLower();
                        #endregion

                        if (location == "planned")
                        #region ...
                        {
                            Outlook.RecurrencePattern rp = item.GetRecurrencePattern();
                            Outlook.AppointmentItem recur = null;

                            DateTime startPoint = item.Start;
                            while (startPoint.AddYears(1) <= tLimitFrom)
                                startPoint = startPoint.AddYears(1);
                            while (startPoint.AddMonths(1) <= tLimitFrom)
                                startPoint = startPoint.AddMonths(1);
                            while (startPoint.AddDays(1) <= tLimitFrom)
                                startPoint = startPoint.AddDays(1);

                            log.LogVariable("Not Specified", nameof(startPoint), startPoint);

                            for (DateTime testPoint = RoundTime(startPoint); testPoint <= tLimitTo; testPoint = testPoint.AddMinutes(checkEvery_Mins)) //KEY POINT
                            {
                                if (testPoint >= tLimitFrom)
                                {
                                    try
                                    {
                                        recur = rp.GetOccurrence(testPoint);

                                        try
                                        {
                                            Appointment appo_Dto = new Appointment(recur.GlobalAppointmentID, recur.Start, item.Duration, recur.Subject, recur.Location, recur.Body, t.Bool(sqlController.SettingRead(Settings.colorsRule)), false, sqlController.LookupRead);
                                            appo_Dto = CreateAppointment(appo_Dto);
                                            recur.Delete();
                                            log.LogStandard("Not Specified", recur.GlobalAppointmentID + " / " + recur.Start + " converted to non-recurence appointment");
                                        }
                                        catch (Exception ex)
                                        {
                                            log.LogWarning("Not Specified", t.PrintException(t.GetMethodName() + " failed. The OutlookController will keep the Expection contained", ex));
                                        }
                                        ConvertedAny = true;
                                    }
                                    catch { }
                                }
                            }
                        }
                        #endregion
                    }
                }
                #endregion

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
                DateTime checkLast_At       = DateTime.Parse(sqlController.SettingRead(Settings.checkLast_At));
                double checkPreSend_Hours   = double.Parse(sqlController.SettingRead(Settings.checkPreSend_Hours));
                double checkRetrace_Hours   = double.Parse(sqlController.SettingRead(Settings.checkRetrace_Hours));
                int checkEvery_Mins         = int.Parse(sqlController.SettingRead(Settings.checkEvery_Mins));
                bool includeBlankLocations  = bool.Parse(sqlController.SettingRead(Settings.includeBlankLocations));

                DateTime timeOfRun          = DateTime.Now;
                DateTime tLimitTo           = timeOfRun.AddHours(+checkPreSend_Hours);
                DateTime tLimitFrom         = checkLast_At.AddHours(-checkRetrace_Hours);
                #endregion

                #region process appointments
                foreach (Outlook.AppointmentItem item in GetCalendarItems(tLimitTo, tLimitFrom))
                {
                    if (!item.IsRecurring) //is NOT recurring, otherwise ignore
                    {
                        if (tLimitFrom <= item.Start && item.Start <= tLimitTo)
                        {
                            #region location "planned"?
                            string location = item.Location;

                            if (location == null)
                            {
                                if (includeBlankLocations)
                                    location = "planned";
                                else
                                    location = "";
                            }

                            location = location.ToLower();
                            #endregion

                            if (location.ToLower() == "planned")
                            #region ...
                            {
                                if (item.Body != null)
                                    if (item.Body.Contains("<<< "))
                                        if (item.Body.Contains("End >>>"))
                                        {
                                            item.Body = t.LocateReplaceAll(item.Body, "<<< ", "End >>>", "");
                                            item.Body = item.Body.Replace("<<< End >>>", "");
                                            item.Body = item.Body.Trim();
                                        }

                                item.Save();

                                Appointment appo = new Appointment(item.GlobalAppointmentID, item.Start, item.Duration, item.Subject, item.Location, item.Body, t.Bool(sqlController.SettingRead(Settings.colorsRule)), true, sqlController.LookupRead);

                                if (appo.Location == null)
                                {
                                    if (includeBlankLocations)
                                        appo.Location = "planned";
                                    else
                                        appo.Location = "";
                                }

                                if (appo.Location.ToLower() == "planned")
                                {
                                    int count = sqlController.AppointmentsCreate(appo);

                                    if (count == 1)
                                        CalendarItemUpdate(appo, WorkflowState.Processed, false);
                                    else
                                    {
                                        if (count == 0)
                                            CalendarItemUpdate(appo, WorkflowState.Failed_to_expection, false);

                                        if (count == 2)
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
                                            CalendarItemUpdate(appo, WorkflowState.Failed_to_intrepid, false);
                                        }
                                    }
                                }
                                else
                                    CalendarItemUpdate(appo, WorkflowState.Failed_to_intrepid, false);

                                AllIntrepid = true;
                            }
                            #endregion

                            if (location.ToLower() == "cancel")
                            #region ...
                            {
                                Appointment appo = new Appointment(item.GlobalAppointmentID, item.Start, item.Duration, item.Subject, item.Location, item.Body, t.Bool(sqlController.SettingRead(Settings.colorsRule)), true, sqlController.LookupRead);

                                if (sqlController.AppointmentsCancel(appo.GlobalId))
                                    CalendarItemUpdate(appo, WorkflowState.Canceled, false);
                                else
                                    CalendarItemUpdate(appo, WorkflowState.Failed_to_intrepid, false);

                                AllIntrepid = true;
                            }
                            #endregion

                            if (location.ToLower() == "check")
                            #region ...
                            {
                                eFormSqlController.SqlController sqlMicroting = new eFormSqlController.SqlController(sqlController.SettingRead(Settings.microtingDb));
                                eFormCommunicator.Communicator com = new eFormCommunicator.Communicator(sqlMicroting);

                                var temp = sqlController.AppointmentsFind(item.GlobalAppointmentID);

                                var list = sqlMicroting.InteractionCaseListRead(int.Parse(temp.microting_uid));
                                foreach (var aCase in list)
                                    com.CheckStatusUpdateIfNeeded(aCase.microting_uid);

                                CalendarItemReflecting(item.GlobalAppointmentID);
                                AllIntrepid = true;
                            }
                            #endregion
                        }
                    }
                }
                #endregion

                sqlController.SettingUpdate(Settings.checkLast_At, timeOfRun.ToString());
                log.LogVariable("Not Specified", nameof(Settings.checkLast_At), Settings.checkLast_At.ToString());

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

                if (appointment == null) //double check status if no new found
                    appointment = sqlController.AppointmentsFindOne(1);
     
                if (appointment == null)
                    return false;
                #endregion
                Outlook.AppointmentItem item = AppointmentItemFind(appointment.global_id, appointment.start_at.Value);

                item.Location = appointment.workflow_state;
                #region item.Categories = 'workflowState'...
                switch (appointment.workflow_state)
                {
                    case "Planned":
                        item.Categories = null;
                        break;
                    case "Processed":
                        item.Categories = "Processing";
                        break;
                    case "Created":
                        item.Categories = "Processing";
                        break;
                    case "Sent":
                        item.Categories = "Sent";
                        break;
                    case "Retrived":
                        item.Categories = "Retrived";
                        break;
                    case "Completed":
                        item.Categories = "Completed";
                        break;
                    case "Canceled":
                        item.Categories = "Canceled";
                        break;
                    case "Revoked":
                        item.Categories = "Revoked";
                        break;
                    case "Failed_to_expection":
                        item.Categories = "Error";
                        break;
                    case "Failed_to_intrepid":
                        item.Categories = "Error";
                        break;
                    default:
                        item.Categories = "Error";
                        break;
                }
                #endregion
                item.Body = appointment.body;
                #region item.Body = appointment.expectionString + item.Body + appointment.response ...
                if (!string.IsNullOrEmpty(appointment.response))
                {
                    if (t.Bool(sqlController.SettingRead(Settings.responseBeforeBody)))
                    {
                        item.Body =           "<<< Response: Start >>>" + 
                        Environment.NewLine +
                        Environment.NewLine + appointment.response + 
                        Environment.NewLine +
                        Environment.NewLine + "<<< Response: End >>>" + 
                        Environment.NewLine +
                        Environment.NewLine + item.Body;
                    }
                    else
                    {
                        item.Body =           item.Body + 
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
                    item.Body =           "<<< Exception: Start >>>" + 
                    Environment.NewLine +
                    Environment.NewLine + appointment.expectionString + 
                    Environment.NewLine +
                    Environment.NewLine + "<<< Exception: End >>>" + 
                    Environment.NewLine +
                    Environment.NewLine + item.Body;
                }
                #endregion
                item.Save();

                sqlController.AppointmentsReflected(appointment.global_id);
                return true;
            }
            catch (Exception ex)
            {
                throw new Exception(t.GetMethodName() + " failed", ex);
            }
        }

        private Outlook.AppointmentItem AppointmentItemFind(string globalId, DateTime start)
        {
            try
            {
                string filter = "[Start] = '" + start.ToString("g") + "'";
                log.LogVariable("Not Specified", nameof(filter), filter.ToString());

                Outlook.MAPIFolder calendarFolder = GetCalendarFolder();
                Outlook.Items calendarItemsAll = calendarFolder.Items;
                calendarItemsAll.IncludeRecurrences = false;
                Outlook.Items calendarItemsRes = calendarItemsAll.Restrict(filter);

                foreach (Outlook.AppointmentItem item in calendarItemsRes)
                    if (item.GlobalAppointmentID == globalId)
                        return item;

                foreach (Outlook.AppointmentItem item in calendarItemsAll)
                    if (item.GlobalAppointmentID == globalId)
                        return item;

                throw new Exception(t.GetMethodName() + " failed. Due to no match found global id:" + globalId);
            }
            catch (Exception ex)
            {
                throw new Exception(t.GetMethodName() + " failed", ex);
            }
        }

        public void                 CalendarItemUpdate(Appointment appointment, WorkflowState workflowState, bool resetBody)
        {
            Outlook.AppointmentItem item = AppointmentItemFind(appointment.GlobalId, appointment.Start);

            item.Body = appointment.Body;
            item.Location = workflowState.ToString();
            #region item.Categories = 'workflowState'...
            switch (workflowState)
            {
                case WorkflowState.Planned:
                    item.Categories = null;
                    break;
                case WorkflowState.Processed:
                    item.Categories = "Processing";
                    break;
                case WorkflowState.Created:
                    item.Categories = "Processing";
                    break;
                case WorkflowState.Sent:
                    item.Categories = "Sent";
                    break;
                case WorkflowState.Retrived:
                    item.Categories = "Retrived";
                    break;
                case WorkflowState.Completed:
                    item.Categories = "Completed";
                    break;
                case WorkflowState.Canceled:
                    item.Categories = "Canceled";
                    break;
                case WorkflowState.Revoked:
                    item.Categories = "Revoked";
                    break;
                case WorkflowState.Failed_to_expection:
                    item.Categories = "Error";
                    break;
                case WorkflowState.Failed_to_intrepid:
                    item.Categories = "Error";
                    break;
            }
            #endregion

            if (resetBody)
                item.Body = UnitTest_CalendarBody();

            item.Save();

            log.LogStandard("Not Specified", PrintAppointment(item) + " updated to " + workflowState.ToString());
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

        private Appointment         CreateAppointment(Appointment appointment)
        {
            try
            {
                Outlook.Application outlookApp = new Outlook.Application(); // creates new outlook app
                Outlook.AppointmentItem newAppo = (Outlook.AppointmentItem)outlookApp.CreateItem(Outlook.OlItemType.olAppointmentItem); // creates a new appointment

                newAppo.AllDayEvent = false;
                newAppo.ReminderSet = false;

                newAppo.Location = appointment.Location;
                newAppo.Start = appointment.Start;
                newAppo.Duration = appointment.Duration;
                newAppo.Subject = appointment.Subject;
                newAppo.Body = appointment.Body;

                newAppo.Save();
                Appointment returnAppo = new Appointment(newAppo.GlobalAppointmentID, newAppo.Start, newAppo.Duration, newAppo.Subject, newAppo.Location, newAppo.Body, t.Bool(sqlController.SettingRead(Settings.colorsRule)), true, sqlController.LookupRead);

                Outlook.MAPIFolder calendarFolderDestination = GetCalendarFolder();
                Outlook.NameSpace mapiNamespace = outlookApp.GetNamespace("MAPI");
                Outlook.MAPIFolder oDefault = mapiNamespace.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderCalendar);

                if (calendarFolderDestination.Name != oDefault.Name)
                    newAppo.Move(calendarFolderDestination);

                return returnAppo;
            }
            catch (Exception ex)
            {
                throw new Exception("The following error occurred: " + ex.Message);
            }
        }

        private string              PrintAppointment(Outlook.AppointmentItem appItem)
        {
            return "GlobalId:" + appItem.GlobalAppointmentID + " / Start:" + appItem.Start + " / Title:" + appItem.Subject;
        }

        private Outlook.Items       GetCalendarItems(DateTime tLimitTo, DateTime tLimitFrom)
        {
            lock (_lockOutlook)
            {
                string filter = "[Start] >= '" + tLimitFrom.ToString("g") + "' AND [Start] <= '" + tLimitTo.ToString("g") + "'";
                log.LogVariable("Not Specified", nameof(filter), filter.ToString());

                Outlook.MAPIFolder CalendarFolder = GetCalendarFolder();
                Outlook.Items outlookCalendarItems = CalendarFolder.Items;
                outlookCalendarItems = outlookCalendarItems.Restrict(filter);
                log.LogVariable("Not Specified", "outlookCalendarItems.Count", outlookCalendarItems.Count);
                return outlookCalendarItems;
            }
        }

        private Outlook.MAPIFolder  GetCalendarFolder()
        {
            if (calendarName == sqlController.SettingRead(Settings.calendarName))
                return calendarFolder;
            else
            {
                calendarName = sqlController.SettingRead(Settings.calendarName);

                Outlook.Application oApp = new Outlook.Application();
                Outlook.NameSpace mapiNamespace = oApp.GetNamespace("MAPI");
                Outlook.MAPIFolder oDefault = mapiNamespace.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox).Parent;

                try
                {
                    calendarFolder = GetCalendarFolderByName(oDefault.Folders, calendarName);

                    if (calendarFolder == null)
                        throw new Exception(t.GetMethodName() + " failed, for calendarName:'" + calendarName + "'. No such calendar found");
                }
                catch (Exception ex)
                {
                    throw new Exception(t.GetMethodName() + " failed, for calendarName:'" + calendarName +"'", ex);
                }

                return calendarFolder;
            }
        }

        private Outlook.MAPIFolder  GetCalendarFolderByName(Outlook._Folders folder, string name)
        {
            foreach (Outlook.MAPIFolder Folder in folder)
            {
                if (Folder.Name == name)
                    return Folder;
                else
                {
                    Outlook.MAPIFolder rtrnFolder = GetCalendarFolderByName(Folder.Folders, name);

                    if (rtrnFolder != null)
                        return rtrnFolder;
                }
            }

            return null;
        }
        #endregion

        public List<Appointment>    UnitTest_CalendarItemGetAllNonRecurring(DateTime startPoint, DateTime endPoint)
        {
            try
            {
                #region var
                List<Appointment> lstAppoint = new List<Appointment>();

                Outlook.MAPIFolder CalendarFolder = GetCalendarFolder();
                Outlook.Items outlookCalendarItems = CalendarFolder.Items;
                outlookCalendarItems.IncludeRecurrences = true;
                #endregion

                foreach (Outlook.AppointmentItem item in outlookCalendarItems)
                {
                    if (item.Location != null)
                    {
                        if (item.IsRecurring)
                        {
                            //ignore
                        }
                        else
                        {
                            if (startPoint <= item.Start && item.Start <= endPoint)
                                lstAppoint.Add(new Appointment(item.GlobalAppointmentID, item.Start, item.Duration, item.Subject, item.Location, item.Body, t.Bool(sqlController.SettingRead(Settings.colorsRule)), true, sqlController.LookupRead));
                        }
                    }
                }

                return lstAppoint;
            }
            catch (Exception ex)
            {
                throw new Exception(t.GetMethodName() + " failed", ex);
            }
        }

        private string              UnitTest_CalendarBody()
        {
            return
                                            "TempLate# "+ "’Besked’"
                    + Environment.NewLine + "Sites# "   + "’All’"
                    + Environment.NewLine + "title# "   + "Outlook appointment eForm test"
                    + Environment.NewLine + "info# "    + "Tekst fra Outlook appointment";
        }
    }
}