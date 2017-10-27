using eFormShared;
using OutlookSql;

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OutlookOffice
{
    public class OutlookCommunicator_Fake : IOutlookCommunicator
    {
        #region var
        SqlController sqlController;
        Log log;
        Tools t = new Tools();
        object _lockOutlook = new object();
        Random rndm = new Random();
        List<CalendarItem> calendarLst;

        string forceException = "";
        #endregion

        //con
        public OutlookCommunicator_Fake(SqlController sqlController, Log log)
        {
            this.sqlController = sqlController;
            this.log = log;
            calendarLst = new List<CalendarItem>();
        }

        public string AppointmentItemCreate(CalendarItem calendarItem)
        {
            string globalId = "Faked GlobalId:" + t.GetRandomInt(8);
            calendarItem.GlobalId = globalId;
            calendarLst.Add(calendarItem);
            return globalId;
        }

        public CalendarItem AppointmentItemRead(string globalId, DateTime start)
        {
            foreach (var item in calendarLst)
                if (item.GlobalId == globalId)
                    if (item.Start == start)
                        return item;

            return null;
        }

        public List<CalendarItem> AppointmentItemReadAll(DateTime tLimitFrom, DateTime tLimitTo)
        {
            return calendarLst;
        }

        public bool AppointmentItemUpdate(CalendarItem calendarItem)
        {
            foreach (var item in calendarLst)
                if (item.GlobalId == calendarItem.GlobalId)
                {
                    item.Body = calendarItem.Body;
                    item.Location = calendarItem.Location;
                    return true;
                }

            return false;
        }

        public bool AppointmentItemDelete(string globalId, DateTime start)
        {
            foreach (var item in calendarLst)
                if (item.GlobalId == globalId)
                    if (item.Start == start)
                    {
                        calendarLst.Remove(item);
                        return true;
                    }

            return false;
        }

        public bool ConvertRecurringAppointments(DateTime timeOfRun, DateTime tLimitFrom, DateTime tLimitTo, DateTime checkLast_At, 
                        double checkPreSend_Hours, double checkRetrace_Hours, int checkEvery_Mins, bool includeBlankLocations)
        {
            log.LogEverything("Unit test", t.GetMethodName() + " called");

            int temp = rndm.Next(0, 2);
            if (temp == 2)
                temp = 1;
            bool flag = t.Bool(temp + "");

            log.LogVariable("Unit test", nameof(flag), flag);
            return flag;
        }
    }
}