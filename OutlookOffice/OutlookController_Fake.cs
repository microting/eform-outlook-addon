using eFormShared;
using OutlookSql;

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OutlookOffice
{
    public class OutlookController_Fake : IOutlookController
    {
        #region var
        string calendarName;
        SqlController sqlController;
        Log log;
        Tools t = new Tools();
        object _lockOutlook = new object();
        #endregion

        #region con
        public                      OutlookController_Fake(SqlController sqlController, Log log)
        {
            this.sqlController = sqlController;
            this.log = log;
        }
        #endregion

        #region public
        public bool                 CalendarItemConvertRecurrences()
        {
            return true;
        }

        public bool                 CalendarItemIntrepid()
        {
            return false;
        }

        public bool                 CalendarItemReflecting(string globalId)
        {
            return true;
        }

        public void                 CalendarItemUpdate(Appointment appointment, WorkflowState workflowState, bool resetBody)
        {
            log.LogStandard("Not Specified", appointment.GlobalId + " updated to " + workflowState.ToString());
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
            return new Appointment();
        }
        #endregion

        public List<Appointment>    UnitTest_CalendarItemGetAllNonRecurring(DateTime startPoint, DateTime endPoint)
        {
            return new List<Appointment>();
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