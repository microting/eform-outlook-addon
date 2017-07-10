using eFormShared;
using OutlookShared;
using OutlookSql;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;

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
        string calendarName;
        Outlook.MAPIFolder calendarFolder = null;
        SqlController sqlController;
        Tools t = new Tools();
        #endregion

        #region con
        public                      OutlookController(SqlController sqlController)
        {
            this.sqlController = sqlController;
        }
        #endregion

        #region public
        public bool                 CalendarItemConvertRecurrences()
        {
            try
            {
                bool ConvertedAny = false;

                #region var
                int checkEvery_Mins = int.Parse(sqlController.SettingRead(Settings.checkEvery_Mins));
                int checkRetrace_Hours = int.Parse(sqlController.SettingRead(Settings.checkRetrace_Hours));
                DateTime checkLast_At = DateTime.Parse(sqlController.SettingRead(Settings.checkLast_At));
                int preSend_Mins = int.Parse(sqlController.SettingRead(Settings.preSend_Mins));
                bool includeBlankLocations = bool.Parse(sqlController.SettingRead(Settings.includeBlankLocations));

                DateTime tLimitTo__ = DateTime.Now.AddMinutes(preSend_Mins);
                DateTime tLimitFrom = checkLast_At.AddHours(-checkRetrace_Hours);

                string filter = "[Start] >= '" + tLimitFrom.ToString("g") + "' AND [Start] <= '" + tLimitTo__.ToString("g") + "'";
                sqlController.LogVariable(nameof(filter), filter.ToString());

                Outlook.MAPIFolder CalendarFolder = GetCalendarFolder();
                Outlook.Items outlookCalendarItems = CalendarFolder.Items;
                outlookCalendarItems.IncludeRecurrences = true;
                outlookCalendarItems = outlookCalendarItems.Restrict(filter);
                #endregion

                #region convert recurrences
                foreach (Outlook.AppointmentItem item in outlookCalendarItems)
                {
                    if (item.Location != null && item.IsRecurring)
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

                        for (DateTime testPoint = startPoint; testPoint <= tLimitTo__; testPoint = testPoint.AddMinutes(checkEvery_Mins)) //KEY POINT
                        {
                            if (testPoint >= tLimitFrom)
                            {
                                try
                                {
                                    recur = rp.GetOccurrence(testPoint);

                                    try
                                    {
                                        Appointment appo_Dto = new Appointment(recur.GlobalAppointmentID, recur.Start, item.Duration, recur.Subject, recur.Location, recur.Body, t.Bool(sqlController.SettingRead(Settings.colorsRule)), false, sqlController.Lookup);
                                        appo_Dto = CreateAppointment(appo_Dto);
                                        recur.Delete();
                                        sqlController.LogStandard(recur.GlobalAppointmentID + " / " + recur.Start + " converted to non-recurence appointment");
                                    }
                                    catch (Exception ex)
                                    {
                                        sqlController.LogWarning(t.PrintException(t.GetMethodName() + " failed. The OutlookController will keep the Expection contained", ex));
                                    }
                                    ConvertedAny = true;
                                }
                                catch { }
                            }
                        }
                    }
                }
                #endregion

                if (ConvertedAny)
                    sqlController.LogStandard(t.GetMethodName() + " completed + converted appointment(s)");
                else
                    sqlController.LogEverything(t.GetMethodName() + " completed");

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
                int checkRetrace_Hours = int.Parse(sqlController.SettingRead(Settings.checkRetrace_Hours));
                int preSend_Mins = int.Parse(sqlController.SettingRead(Settings.preSend_Mins));
                bool includeBlankLocations = bool.Parse(sqlController.SettingRead(Settings.includeBlankLocations));
                DateTime timeOfRun_ = DateTime.Now;
                DateTime checkLast_At = DateTime.Parse(sqlController.SettingRead(Settings.checkLast_At));

                DateTime tLimitTo__ = RoundTime(timeOfRun_).AddMinutes(preSend_Mins);
                DateTime tLimitFrom = checkLast_At.AddHours(-checkRetrace_Hours);

                if (tLimitFrom.AddDays(7).AddHours(checkRetrace_Hours) < tLimitTo__)
                    tLimitTo__ = tLimitFrom.AddDays(7).AddHours(checkRetrace_Hours);

                string filter = "[Start] >= '" + tLimitFrom.ToString("g") + "' AND [Start] <= '" + tLimitTo__.ToString("g") + "'";
                sqlController.LogVariable(nameof(filter), filter.ToString());

                Outlook.MAPIFolder CalendarFolder = GetCalendarFolder();
                Outlook.Items outlookCalendarItems = CalendarFolder.Items;
                outlookCalendarItems.IncludeRecurrences = false;
                outlookCalendarItems = outlookCalendarItems.Restrict(filter);
                #endregion

                #region process appointments
                foreach (Outlook.AppointmentItem item in outlookCalendarItems)
                {
                    if (tLimitFrom <= item.Start && item.Start <= tLimitTo__)
                    {
                        string location = item.Location;

                        if (location == null)
                        {
                            if (includeBlankLocations)
                                location = "planned";
                            else
                                location = "";
                        }

                        if (location.ToLower() == "planned")
                        #region ...
                        {
                            if (item.Body == null)
                                item.Body = "";

                            if (item.Body.Contains("<<Info field:"))
                                if (item.Body.Contains("End>>"))
                                    item.Body = t.LocateReplaceAll(item.Body, "<<Info field:", "End>>", "").Trim();

                            item.Save();

                            Appointment appo = new Appointment(item.GlobalAppointmentID, item.Start, item.Duration, item.Subject, item.Location, item.Body, t.Bool(sqlController.SettingRead(Settings.colorsRule)), true, sqlController.Lookup);

                            if (appo.Location == null)
                            {
                                if (includeBlankLocations)
                                    appo.Location = "planned";
                                else
                                    appo.Location = "";
                            }

                            if (appo.Location.ToLower() == "planned")
                            {
                                if (sqlController.OutlookEfromCreate(appo))
                                    CalendarItemUpdate(appo, WorkflowState.Processed, false);
                                else
                                    CalendarItemUpdate(appo, WorkflowState.Failed_to_expection, false);
                            }
                            else
                                CalendarItemUpdate(appo, WorkflowState.Failed_to_intrepid, false);
           
                            AllIntrepid = true;
                        }
                        #endregion

                        if (location.ToLower() == "cancel")
                        #region ...
                        {
                            Appointment appo = new Appointment(item.GlobalAppointmentID, item.Start, item.Duration, item.Subject, item.Location, item.Body, t.Bool(sqlController.SettingRead(Settings.colorsRule)), true, sqlController.Lookup);

                            if (sqlController.OutlookEformCancel(appo))
                                CalendarItemUpdate(appo, WorkflowState.Canceled, false);
                            else
                                CalendarItemUpdate(appo, WorkflowState.Failed_to_intrepid, false);

                            AllIntrepid = true;
                        }
                        #endregion

                        if (location.ToLower() == "check")
                        #region ...
                        {
                            CalendarItemReflecting(item.GlobalAppointmentID);
                            AllIntrepid = true;
                        }
                        #endregion
                    }
                }
                #endregion

                sqlController.SettingUpdate(Settings.checkLast_At, tLimitTo__.ToString());
                sqlController.LogVariable("Settings.checkLast_At", Settings.checkLast_At.ToString());

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
                appointments appointment = null;
     
                if (globalId == null)
                    appointment = sqlController.AppointmentsFindOne(0);
                else
                    appointment = sqlController.AppointmentsFind(globalId);

                if (appointment == null) //double check status if no new found
                    appointment = sqlController.AppointmentsFindOne(1);
     
                if (appointment == null)
                    return false;

                Outlook.AppointmentItem item = AppointmentItemFind(appointment.global_id, appointment.start_at.Value);

                item.Location = appointment.workflow_state;
                #region item.Categories = 'workflowState'...
                switch (appointment.workflow_state)
                {
                    case "Planned":
                        item.Categories = null;
                        break;
                    case "Processed":
                        item.Categories = "1";
                        break;
                    case "Created":
                        item.Categories = "2";
                        break;
                    case "Sent":
                        item.Categories = "3";
                        break;
                    case "Retrived":
                        item.Categories = "4";
                        break;
                    case "Completed":
                        item.Categories = "5";
                        break;
                    case "Canceled":
                        item.Categories = "6";
                        break;
                    case "Revoked":
                        item.Categories = "7";
                        break;
                    case "Failed_to_expection":
                        item.Categories = "0";
                        break;
                    case "Failed_to_intrepid":
                        item.Categories = "0";
                        break;
                    default:
                        item.Categories = "0";
                        break;
                }
                #endregion
                item.Body = appointment.body;
                #region item.Body = appointment.expectionString + item.Body + appointment.response ...
                if (!string.IsNullOrEmpty(appointment.expectionString))
                {
                    item.Body =
                    "<<Info field: Exception: Start>>" + Environment.NewLine +
                    appointment.expectionString + Environment.NewLine +
                    "<<Info field: Exception: End>>" + Environment.NewLine +
                    Environment.NewLine +
                    item.Body;
                }
                if (!string.IsNullOrEmpty(appointment.response))
                {
                    item.Body =
                    "<<Info field: Response: Start>>" + Environment.NewLine +
                    appointment.response + Environment.NewLine +
                    "<<Info field: Response: End>>" + Environment.NewLine +
                    Environment.NewLine +
                    item.Body;
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
                sqlController.LogVariable(nameof(filter), filter.ToString());

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
                    item.Categories = "1";
                    break;
                case WorkflowState.Created:
                    item.Categories = "2";
                    break;
                case WorkflowState.Sent:
                    item.Categories = "3";
                    break;
                case WorkflowState.Retrived:
                    item.Categories = "4";
                    break;
                case WorkflowState.Completed:
                    item.Categories = "5";
                    break;
                case WorkflowState.Canceled:
                    item.Categories = "6";
                    break;
                case WorkflowState.Revoked:
                    item.Categories = "7";
                    break;
                case WorkflowState.Failed_to_expection:
                    item.Categories = "0";
                    break;
                case WorkflowState.Failed_to_intrepid:
                    item.Categories = "0";
                    break;
            }
            #endregion

            if (resetBody)
                item.Body = UnitTest_CalendarBody();

            item.Save();

            sqlController.LogStandard(PrintAppointment(item) + " updated to " + workflowState.ToString());
        }
        #endregion

        #region private
        private DateTime            RoundTime(DateTime dTime)
        {
            dTime = dTime.AddMinutes(1);
            return new DateTime(dTime.Year, dTime.Month, dTime.Day, dTime.Hour, dTime.Minute, 0);
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
                Appointment returnAppo = new Appointment(newAppo.GlobalAppointmentID, newAppo.Start, newAppo.Duration, newAppo.Subject, newAppo.Location, newAppo.Body, t.Bool(sqlController.SettingRead(Settings.colorsRule)), true, sqlController.Lookup);

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
                                lstAppoint.Add(new Appointment(item.GlobalAppointmentID, item.Start, item.Duration, item.Subject, item.Location, item.Body, t.Bool(sqlController.SettingRead(Settings.colorsRule)), true, sqlController.Lookup));
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
                                            "TempLate# "        + "’Test Template’"
                    + Environment.NewLine + "title# "           + "Outlook appointment eForm test"
                    + Environment.NewLine + "info# "            + "1: Udfyldt besked linje 1"
                    + Environment.NewLine + "info# "            + "2: Udfyldt besked linje 2"
                    + Environment.NewLine + "info# "            + "3: Udfyldt besked linje 3"
                    + Environment.NewLine + "connected# "       + "0"
                    + Environment.NewLine + "expirE# "          + "4"
                    + Environment.NewLine + "replacements# "    + "Gem knap==Save"
                    + Environment.NewLine + "replacements# "    + "Numerisk==Tal"
                    + Environment.NewLine + "Sites# "           + "’salg’, 3913";
        }
    }
}