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
        SqlController sqlController;
        Log log;
        Tools t = new Tools();
        object _lockOutlook = new object();
        Random rndm = new Random();
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
            log.LogEverything("Unit test", t.GetMethodName() + " called");

            int temp  = rndm.Next(0, 2);
            if (temp == 2)
                temp = 1;
            bool flag = t.Bool(temp + "");

            log.LogVariable("Unit test", nameof(flag), flag);
            return flag;
        }

        public bool                 CalendarItemIntrepid()
        {
            log.LogEverything("Unit test", t.GetMethodName() + " called");

            int temp = rndm.Next(0, 2);
            if (temp == 2)
                temp = 1;
            bool flag = t.Bool(temp + "");

            log.LogVariable("Unit test", nameof(flag), flag);
            return flag;
        }

        public bool                 CalendarItemReflecting(string globalId)
        {
            log.LogStandard("Unit test", t.GetMethodName() + " called");
            log.LogVariable("Unit test", (nameof(globalId)), globalId);
            return true;
        }

        public void                 CalendarItemUpdate(Appointment appointment, WorkflowState workflowState, bool resetBody)
        {
            log.LogStandard("Unit test", appointment.GlobalId + " updated to " + workflowState.ToString());
        }
        #endregion

        #region private
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