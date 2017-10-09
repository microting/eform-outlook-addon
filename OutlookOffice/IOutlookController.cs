using OutlookSql;
using System;
using System.Collections.Generic;

namespace OutlookOffice
{
    public interface IOutlookController
    {
        bool CalendarItemConvertRecurrences();

        bool CalendarItemIntrepid();

        bool CalendarItemReflecting(string globalId);

        void CalendarItemUpdate(Appointment appointment, WorkflowState workflowState, bool resetBody);

        List<Appointment> UnitTest_CalendarItemGetAllNonRecurring(DateTime startPoint, DateTime endPoint);

        bool UnitTest_ForceException(string exceptionType);
    }
}