using System;
using System.Collections.Generic;

namespace OutlookOffice
{
    public interface IOutlookCommunicator
    {
        string              AppointmentItemCreate(CalendarItem calendarItem);

        CalendarItem        AppointmentItemRead(string globalId, DateTime start);

        List<CalendarItem>  AppointmentItemReadAll(DateTime tLimitFrom, DateTime tLimitTo);

        bool                AppointmentItemUpdate(CalendarItem calendarItem);

        bool                AppointmentItemDelete(string globalId, DateTime start);

        bool                ConvertRecurringAppointments(DateTime timeOfRun, DateTime tLimitFrom, DateTime tLimitTo, DateTime checkLast_At, 
                                double checkPreSend_Hours, double checkRetrace_Hours, int checkEvery_Mins, bool includeBlankLocations);
    }

    public class CalendarItem
    {
        #region var/pop
        public string GlobalId { get; set; }
        public DateTime Start { get; set; }
        public int Duration { get; set; }
        public string Subject { get; set; }
        public string Location { get; set; }
        public string Body { get; set; }
        #endregion

        #region con
        public CalendarItem()
        {

        }

        public CalendarItem(string globalId, string location, DateTime start, int duration, string subject, string body)
        {
            GlobalId = globalId;
            Start = start;
            Duration = duration;
            Subject = subject;
            Location = location;
            Body = body;
        }
        #endregion

        public override string ToString()
        {
            string globalId = "";
            string start = "";
            string subject = "";
            string location = "";

            if (GlobalId != null)
                globalId = GlobalId;

            if (Start != null)
                start = Start.ToString();

            if (Subject != null)
                subject = Subject;

            if (Location != null)
                location = Location;

            return "GlobalId:" + globalId + " / Start:" + start + " / Subject:" + subject + " / Location:" + location;
        }
    }
}