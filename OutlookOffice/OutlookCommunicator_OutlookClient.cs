using eFormShared;
using Outlook = Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;

namespace OutlookOffice
{
    public class OutlookCommunicator_OutlookClient : IOutlookCommunicator
    {
        #region var
        Log log;
        string calendarName;
        Outlook.MAPIFolder calendarFolder = null;
        Tools t = new Tools();
        object _lockOutlook = new object();
        #endregion

        #region con
        public                      OutlookCommunicator_OutlookClient(string calendarName, Log log)
        {
            this.calendarName = calendarName;
            this.log = log;
        }
        #endregion

        #region public
        public string               AppointmentItemCreate(CalendarItem calendarItem)
        {
            Outlook.Application outlookApp = new Outlook.Application();                                                             // creates new outlook app
            Outlook.AppointmentItem newAppo = (Outlook.AppointmentItem)outlookApp.CreateItem(Outlook.OlItemType.olAppointmentItem); // creates a new appointment

            newAppo.AllDayEvent = false;
            newAppo.ReminderSet = false;
            newAppo.Location = calendarItem.Location;
            if (calendarItem.Location == "Created")
                newAppo.Categories = "Processing";
            newAppo.Start = calendarItem.Start;
            newAppo.Duration = calendarItem.Duration;
            newAppo.Subject = calendarItem.Subject;
            newAppo.Body = calendarItem.Body;
            newAppo.Save();
            log.LogStandard("Not Specified", "Calendar item created in default folder"); //Only place for .Com class to create, hence the need for the move

            Outlook.MAPIFolder calendarFolderDestination = GetCalendarFolder();
            Outlook.NameSpace mapiNamespace = outlookApp.GetNamespace("MAPI");
            Outlook.MAPIFolder oDefault = mapiNamespace.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderCalendar);
            if (calendarFolderDestination.Name != oDefault.Name)
            {
                newAppo.Move(calendarFolderDestination);
                log.LogStandard("Not Specified", "Calendar item moved to " + calendarFolderDestination.Name);
            }

            return newAppo.GlobalAppointmentID;
        }

        public CalendarItem         AppointmentItemRead(string globalId, DateTime start)
        {
            var item = AppointmentItemFind(globalId, start);
            CalendarItem cItem = new CalendarItem(item.GlobalAppointmentID, item.Location, item.Start, item.Duration, item.Subject, item.Body);
            return cItem;
        }

        public List<CalendarItem>   AppointmentItemReadAll(DateTime tLimitFrom, DateTime tLimitTo)
        {
            List<CalendarItem> rtrnLst = new List<CalendarItem>();

            foreach (Outlook.AppointmentItem item in GetCalendarItems(tLimitFrom, tLimitTo))
                if (!item.IsRecurring) //is NOT recurring, otherwise ignore
                    rtrnLst.Add(new CalendarItem(item.GlobalAppointmentID, item.Location, item.Start, item.Duration, item.Subject, item.Body));

            return rtrnLst;
        }

        public bool                 AppointmentItemUpdate(CalendarItem calendarItem)
        {
            Outlook.AppointmentItem item = AppointmentItemFind(calendarItem.GlobalId, calendarItem.Start);
            if (item != null)
            {
                item.Location = calendarItem.Location;
                #region item.Categories = 'workflowState'...
                switch (calendarItem.Location)
                {
                    case "Planned":
                        item.Categories = null;
                        break;
                    case "Processed":
                        item.Categories = CalendarItemCategory.Processing.ToString();
                        break;
                    case "Created":
                        item.Categories = CalendarItemCategory.Processing.ToString();
                        break;
                    case "Sent":
                        item.Categories = CalendarItemCategory.Sent.ToString();
                        break;
                    case "Retrived":
                        item.Categories = CalendarItemCategory.Retrived.ToString();
                        break;
                    case "Completed":
                        item.Categories = CalendarItemCategory.Completed.ToString();
                        break;
                    case "Canceled":
                        item.Categories = CalendarItemCategory.Revoked.ToString();
                        break;
                    case "Revoked":
                        item.Categories = CalendarItemCategory.Revoked.ToString();
                        break;
                    case "Exception":
                        item.Categories = CalendarItemCategory.Error.ToString();
                        break;
                    case "Failed_to_intrepid":
                        item.Categories = CalendarItemCategory.Error.ToString();
                        break;
                    default:
                        item.Categories = CalendarItemCategory.Error.ToString();
                        break;
                }
                #endregion
                item.Body = calendarItem.Body;

                item.Save();
                log.LogStandard("Not Specified", "globalId:'" + calendarItem.GlobalId + "' updated in calendar");
            }
            else
                log.LogWarning("Not Specified", "globalId:'" + calendarItem.GlobalId + "' no longer in calendar, so hence is considered to be updated in calendar");

            return true;
        }

        public bool                 AppointmentItemDelete(string globalId, DateTime start)
        {
            Outlook.AppointmentItem item = AppointmentItemFind(globalId, start);

            if (item == null)
                return false;

            item.Delete();

            log.LogStandard("Not Specified", "globalId:'" + globalId + "' deleted");
            return true;
        }

        public bool                 ConvertRecurringAppointments(DateTime timeOfRun, DateTime tLimitFrom, DateTime tLimitTo, DateTime checkLast_At, 
                                        double checkPreSend_Hours, double checkRetrace_Hours, int checkEvery_Mins, bool includeBlankLocations)
        {
            bool ConvertedAny = false;

            #region convert recurrences
            foreach (Outlook.AppointmentItem item in GetCalendarItems(tLimitFrom, tLimitTo))
            {
                if (item.IsRecurring) //is NOT recurring, otherwise ignore
                {
                    #region location "planned"?
                    string location = item.Location;

                    if (location == null)
                        if (includeBlankLocations)
                            location = "planned";
                        else
                            location = "";
   
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

                        for (DateTime testPoint = startPoint; testPoint <= tLimitTo; testPoint = testPoint.AddMinutes(checkEvery_Mins)) //KEY POINT
                            if (testPoint >= tLimitFrom)
                                try
                                {
                                    recur = rp.GetOccurrence(testPoint);

                                    try
                                    {
                                        AppointmentItemCreate(new CalendarItem("", recur.Location, recur.Start, item.Duration, recur.Subject, recur.Body));
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
                    #endregion
                }
            }
            #endregion

            return ConvertedAny;
        }
        #endregion

        #region private
        private Outlook.AppointmentItem AppointmentItemFind(string globalId, DateTime start)
        {
            try
            {
                log.LogEverything("Not Specified", "OutlookCommunicator_OutlookClient.AppointmentItemFind called");
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

                log.LogEverything("Not Specified", "No match found for " + nameof(globalId) + ":" + globalId);
                return null;
            }
            catch (Exception ex)
            {
                throw new Exception(t.GetMethodName() + " failed. Due to no match found global id:" + globalId, ex);
            }
        }

        private Outlook.Items       GetCalendarItems(DateTime tLimitFrom, DateTime tLimitTo)
        {
            lock (_lockOutlook)
            {

                string filter = "GetCalendarItems [After] '" + tLimitTo.ToString("g") + "' AND [before] <= '" + tLimitFrom.ToString("g") + "'";
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
            log.LogEverything("Not Specified", "GetCalendarFolder called");

            if (calendarFolder != null)
            {
                return calendarFolder;
            }
            else
            {
                log.LogVariable("Not Specified", nameof(calendarName), calendarName);

                Outlook.Application oApp = new Outlook.Application();
                log.LogEverything("Not Specified", "Found oApp");

                Outlook.NameSpace mapiNamespace = oApp.GetNamespace("MAPI");
                log.LogEverything("Not Specified", "Found mapiNamespace");

                Outlook.MAPIFolder oDefault = mapiNamespace.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderCalendar);
                log.LogEverything("Not Specified", "Found oDefault");

                try
                {
                    calendarFolder = GetCalendarFolderByName(mapiNamespace.Folders, calendarName);
                    log.LogVariable("Not Specified", nameof(calendarFolder), calendarFolder.FolderPath);

                    if (calendarFolder == null)
                        throw new Exception(t.GetMethodName() + " failed, for calendarName:'" + calendarName + "'. No such calendar found");
                }
                catch (Exception ex)
                {
                    throw new Exception(t.GetMethodName() + " failed, for calendarName:'" + calendarName + "'", ex);
                }

                return calendarFolder;
            }
        }

        private Outlook.MAPIFolder  GetCalendarFolderByName(Outlook._Folders folder, string name)
        {
            log.LogEverything("Not Specified", "GetCalendarFolderByName called");
            foreach (Outlook.MAPIFolder Folder in folder)
            {
                log.LogEverything("Not Specified", "current folder is : " + Folder.Name);
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
    }
}