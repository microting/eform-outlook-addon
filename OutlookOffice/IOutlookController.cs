﻿using OutlookSql;
using System;
using System.Collections.Generic;

namespace OutlookOffice
{
    public interface IOutlookController
    {
        bool CalendarItemConvertRecurrences();

        bool CalendarItemIntrepid();

        bool CalendarItemReflecting(string globalId);

        string CalendarItemCreate(string location, DateTime start, int duration, string subject, string body);
        
        bool CalendarItemUpdate(string globalId, DateTime start, LocationOptions workflowState, string body);

        bool CalendarItemDelete(string globalId, DateTime start);

        List<Appointment> UnitTest_CalendarItemGetAllNonRecurring(DateTime startPoint, DateTime endPoint);

        bool UnitTest_ForceException(string exceptionType);
    }
}