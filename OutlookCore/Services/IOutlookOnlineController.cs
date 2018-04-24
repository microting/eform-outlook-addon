using OutlookExchangeOnlineAPI;
using OutlookSql;
using System;
using System.Collections.Generic;

namespace OutlookOfficeOnline
{
    public interface IOutlookOnlineController
    {
        bool CalendarItemConvertRecurrences();

        bool ParseCalendarItems();

        bool CalendarItemReflecting(string globalId);

        Event CalendarItemCreate(string location, DateTime start, int duration, string subject, string body, string originalStartTimeZone, string originalEndTimeZone);

        bool CalendarItemUpdate(string globalId, DateTime start, ProcessingStateOptions workflowState, string body);

        bool CalendarItemDelete(string globalId);

        string GetUserEmailAddress();
    }
}